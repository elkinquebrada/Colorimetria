using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Tesseract;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Color.Services;

namespace Color
{
    // ─────────────────────────────────────────────────────────────────────────
    // Extractor de receta por columnas geométricas independientes
    // ─────────────────────────────────────────────────────────────────────────

    public sealed class DynamicSplitGridExtractor : IDisposable
    {
        // ── Rutas de Tesseract ────────────────────────────────────────────────
        private readonly string _tessDataPath;
        private TesseractEngine _engine;
        private readonly object _lock = new object();

        // ── Proporciones geométricas (fracciones de la altura total) ──────────

        private const double RECIPE_ZONE_HEIGHT_COMBINED = 0.22;
        private const double RECIPE_ZONE_TOP_FLAT   = 0.18;
        private const double RECIPE_ZONE_HEIGHT_FLAT = 0.25;

        // ── Constructor ───────────────────────────────────────────────────────

        public DynamicSplitGridExtractor(string tessDataPath = @".\tessdata")
        {
            _tessDataPath = tessDataPath;
        }

        // ── API pública ───────────────────────────────────────────────────────

        public List<RecipeItem> ExtractRecipePositional(
            string imagePath,
            System.Drawing.Rectangle? tableBounds = null)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return new List<RecipeItem>();

            using (Bitmap bmp = ReportFormatRouter.LoadUniversalImage24bpp(imagePath))
                return ExtractFromBitmap(bmp, tableBounds);
        }

        // ── API estática (permite invocar sin instancia pre-creada) ───────────

        public static List<RecipeItem> ExtractRecipePositional(
            string imagePath,
            System.Drawing.Rectangle? tableBounds,
            string tessDataPath)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return new List<RecipeItem>();

            string path = string.IsNullOrWhiteSpace(tessDataPath) ? @".\tessdata" : tessDataPath;
            using (var instance = new DynamicSplitGridExtractor(path))
            using (Bitmap bmp = ReportFormatRouter.LoadUniversalImage24bpp(imagePath))
                return instance.ExtractFromBitmap(bmp, tableBounds);
        }

        /// Extrae la lista de colorantes directamente desde un Bitmap ya cargado.
        public List<RecipeItem> ExtractFromBitmap(
            Bitmap bmp,
            System.Drawing.Rectangle? tableBounds = null)
        {
            var recipeItems = new List<RecipeItem>();
            if (bmp == null) return recipeItems;

            int imgWidth  = bmp.Width;
            int imgHeight = bmp.Height;

            // Procesamos la imagen COMPLETA ya que la posición de los datos varía drásticamente
            var region = new System.Drawing.Rectangle(0, 0, imgWidth, imgHeight);
            string rawText = ExtractTextWithOpenCv(bmp, region);
            var lines = rawText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();

            // ──  Parseo y limpieza de líneas ────────────────────────────────
            return FilterAndCleanRecipeLines(lines);
        }

        // ── Procesamiento con OpenCV + OCR ────────────────────────────────────

        private string ExtractTextWithOpenCv(Bitmap originalBmp, System.Drawing.Rectangle region)
        {
            region = System.Drawing.Rectangle.Intersect(
                region,
                new System.Drawing.Rectangle(0, 0, originalBmp.Width, originalBmp.Height));
            if (region.IsEmpty) return string.Empty;

            using (Bitmap regionBmp = originalBmp.Clone(region, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                using (Mat mat = BitmapConverter.ToMat(regionBmp))
                using (Mat gray = new Mat())
                using (Mat scaled = new Mat())
                using (Mat bin = new Mat())
                {
                    // 1. Convertir a grises
                    Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                    // 2. Escalar 3x para nitidez
                    Cv2.Resize(gray, scaled, new OpenCvSharp.Size(gray.Width * 3, gray.Height * 3), 0, 0, InterpolationFlags.Cubic);
                    // 3. Umbral adaptativo Otsu
                    Cv2.Threshold(scaled, bin, 0, 255, ThresholdTypes.Otsu | ThresholdTypes.Binary);

                    using (Bitmap processedBmp = BitmapConverter.ToBitmap(bin))
                    {
                        return RunTesseractOnBitmap(processedBmp);
                    }
                }
            }
        }

        // ── Motor Tesseract compartido ────────────────────────────────────────

        private string RunTesseractOnBitmap(Bitmap bmp)
        {
            try
            {
                lock (_lock)
                {
                    if (_engine == null)
                        _engine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default);

                    using (var page = _engine.Process(bmp, PageSegMode.SingleBlock))
                        return page.GetText() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        // ── Parseo Estructural por Expresiones Regulares ──────────────────────

        private List<RecipeItem> FilterAndCleanRecipeLines(List<string> rawLines)
        {
            List<RecipeItem> validRecipe = new List<RecipeItem>();
         
            string pattern = @"^\s*(\d{5,})\s+(.+?)\s+(\d+[.,]\d{2,})";

            foreach (string rawLine in rawLines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;
                
                string upperLine = line.ToUpper();
                if (upperLine.Contains("DYE CODE") || upperLine.Contains("DYE NAME") ||
                    upperLine.Contains("CONCENTRATION") || upperLine.Contains("DYES") ||
                    upperLine.Contains("PROCESS") || upperLine.Contains("RECIPE") ||
                    upperLine.Contains("TOTAL"))
                {
                    continue; 
                }
                
                // Exclusión inquebrantable de los Químicos y Sales: si la fila menciona "g/l" o nombres evidentes
                if (Regex.IsMatch(upperLine, @"G\s*[/\\]\s*[LI1|]")) continue;
                if (upperLine.Contains("SULFATO") || upperLine.Contains("ACIDO") || upperLine.Contains("ÁCIDO") ||
                    upperLine.Contains("CINDYE") || upperLine.Contains("INVADINE")) continue;
                
                // Limpieza secuencial profunda para despejar el número central
                line = Regex.Replace(line, @"0[.,]00\s*[A-Za-z/|]*$", "", RegexOptions.IgnoreCase);
                line = Regex.Replace(line, @"(?:%|9\/?6)\s*$", "", RegexOptions.IgnoreCase).Trim();
                
                line = line.Replace("|", " ").Trim();
                
                Match m = Regex.Match(line, pattern);
                if (m.Success)
                {
                    string code = m.Groups[1].Value.Trim();
                    string name = m.Groups[2].Value.Trim();
                    string percentage = m.Groups[3].Value.Trim().Replace(",", ".");
                    
                    // Si la captura del nombre arrastró un símbolo de porcentaje (a menos que realmente sea parte del nombre como 100%)
                    if (name.EndsWith("%") && Regex.IsMatch(name, @"\s+%$"))
                        name = name.Substring(0, name.Length - 1).Trim();

                    // TRATAMIENTO QUIRÚRGICO DE DAÑOS OCR:
                    if (percentage.StartsWith("3001.") && percentage.Length > 6)
                    {
                        name += " 300"; 
                        percentage = "1." + percentage.Substring(5); 
                    }
                    
                    // Si el remanente comienza en punto, ponerle un cero.
                    if (percentage.StartsWith(".")) percentage = "0" + percentage;

                    validRecipe.Add(new RecipeItem {
                        Code = code,
                        Name = name,
                        Percentage = percentage
                    });
                }
            }
            return validRecipe;
        }

        // ── IDisposable ───────────────────────────────────────────────────────

        public void Dispose()
        {
            lock (_lock)
            {
                _engine?.Dispose();
                _engine = null;
            }
        }
    }
}
