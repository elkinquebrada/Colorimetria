using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Tesseract;
using Color.Services;

namespace Color
{
    // ─────────────────────────────────────────────────────────────────────────
    // Extractor de receta por columnas geométricas independientes - CORREGIDO
    // ─────────────────────────────────────────────────────────────────────────

    public sealed class DynamicSplitGridExtractor : IDisposable
    {
        // ── Rutas de Tesseract ────────────────────────────────────────────────
        private readonly string _tessDataPath;
        private TesseractEngine _engine;
        private readonly object _lock = new object();

        // ── Proporciones geométricas (fracciones del ancho total) ─────────────
        private const double COL_CODE_WIDTH = 0.15;
        private const double COL_NAME_WIDTH = 0.38;
        private const double COL_PCT_WIDTH = 0.15;

        // ── Proporciones geométricas (fracciones de la altura total) ──────────
        private const double RECIPE_ZONE_HEIGHT_COMBINED = 0.22;
        private const double RECIPE_ZONE_TOP_FLAT = 0.18;
        private const double RECIPE_ZONE_HEIGHT_FLAT = 0.25;

        // ── Constructor ───────────────────────────────────────────────────────
        public DynamicSplitGridExtractor(string tessDataPath = @".\\tessdata")
        {
            _tessDataPath = tessDataPath;
        }

        // ── API Pública de Extracción ────────────────────────────────────────
        public List<ColoranteData> ExtractRecipe(Bitmap sourceBmp)
        {
            var rawLines = ExtractRawLinesFromGrid(sourceBmp);
            return FilterAndCleanRecipeLines(rawLines);
        }

        // ── Segmentación Geométrica de la Imagen ──────────────────────────────
        private List<string> ExtractRawLinesFromGrid(Bitmap bmp)
        {
            List<string> lines = new List<string>();

            // Calculamos el área aproximada donde se encuentra la tabla de colorantes
            int startY = (int)(bmp.Height * RECIPE_ZONE_TOP_FLAT);
            int zoneHeight = (int)(bmp.Height * RECIPE_ZONE_HEIGHT_FLAT);

            if (startY + zoneHeight > bmp.Height)
                zoneHeight = bmp.Height - startY;

            // Definimos el ancho de corte de la zona de interés
            int targetWidth = (int)(bmp.Width * (COL_CODE_WIDTH + COL_NAME_WIDTH + COL_PCT_WIDTH + 0.05));
            if (targetWidth > bmp.Width) targetWidth = bmp.Width;

            Rectangle recipeRegion = new Rectangle(0, startY, targetWidth, zoneHeight);

            using (Bitmap croppedZone = CropImage(bmp, recipeRegion))
            {
                string rawText = RunTesseractOnBitmap(croppedZone);
                if (!string.IsNullOrEmpty(rawText))
                {
                    var split = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    lines.AddRange(split);
                }
            }

            return lines;
        }

        // ── MÉTODO NUEVO: Filtro y limpieza estricta de datos (OCR Sanitization) ──
        private List<ColoranteData> FilterAndCleanRecipeLines(List<string> rawLines)
        {
            List<ColoranteData> validRecipe = new List<ColoranteData>();

            foreach (var line in rawLines)
            {
                string cleanLine = line.Trim();

                // 1. Descartar líneas vacías o excesivamente cortas (ruido)
                if (string.IsNullOrEmpty(cleanLine) || cleanLine.Length < 5)
                    continue;

                // 2. Descartar cabeceras y la fila totalizadora '[Dyes]' / 'Total'
                string upperLine = cleanLine.ToUpper();
                if (upperLine.Contains("DYE CODE") ||
                    upperLine.Contains("DYE NAME") ||
                    upperLine.Contains("CONCENTRATION") ||
                    upperLine.Contains("PROPORTION") ||
                    upperLine.Contains("DYES") ||
                    upperLine.Contains("TOTAL") ||
                    upperLine.Contains("[DYES]"))
                {
                    continue; // Saltar fila basura
                }

                // 3. Segmentar los componentes (Código, Nombre, Porcentaje)
                // Buscamos separar por múltiples espacios o tabulaciones del OCR
                var parts = Regex.Split(cleanLine, @"\s{2,}")
                                 .Where(p => !string.IsNullOrWhiteSpace(p))
                                 .Select(p => p.Trim())
                                 .ToList();

                // Si no se separó correctamente por espacios amplios, intentamos por espacio simple
                if (parts.Count < 2)
                {
                    parts = cleanLine.Split(' ')
                                     .Where(p => !string.IsNullOrWhiteSpace(p))
                                     .Select(p => p.Trim())
                                     .ToList();
                }

                if (parts.Count >= 2)
                {
                    string code = parts[0];

                    // VALIDACIÓN DE ORO: El código de colorante debe contener obligatoriamente dígitos numéricos
                    if (!Regex.IsMatch(code, @"\d+"))
                        continue; // No es una fila de datos válida

                    string name = parts[1];
                    string percentage = "0%";

                    // Si logramos recuperar la columna del porcentaje
                    if (parts.Count >= 3)
                    {
                        percentage = parts[parts.Count - 1];
                    }

                    validRecipe.Add(new ColoranteData
                    {
                        Codigo = code,
                        Nombre = name,
                        Porcentaje = percentage
                    });
                }
            }

            return validRecipe;
        }

        // ── Utilidades de Corte de Imagen ─────────────────────────────────────
        private Bitmap CropImage(Bitmap src, Rectangle rect)
        {
            Bitmap bmp = new Bitmap(rect.Width, rect.Height, src.PixelFormat);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.DrawImage(src, new Rectangle(0, 0, bmp.Width, bmp.Height), rect, GraphicsUnit.Pixel);
            }
            return bmp;
        }

        // ── Motor Tesseract Compartido con Escala 3x Optimizada ───────────────
        private string RunTesseractOnBitmap(Bitmap bmp)
        {
            try
            {
                lock (_lock)
                {
                    if (_engine == null)
                        _engine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default);

                    // Escalar 3x para mejorar radicalmente la detección de números y puntos decimales
                    int scaledW = bmp.Width * 3;
                    int scaledH = bmp.Height * 3;

                    using (Bitmap scaled = new Bitmap(scaledW, scaledH))
                    using (Graphics g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(bmp, 0, 0, scaledW, scaledH);

                        using (var page = _engine.Process(scaled, PageSegMode.SingleBlock))
                            return page.GetText() ?? string.Empty;
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
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

    // Clase auxiliar para el tipado de los datos devueltos
    public class ColoranteData
    {
        public string Codigo { get; set; }
        public string Nombre { get; set; }
        public string Porcentaje { get; set; }
    }
}