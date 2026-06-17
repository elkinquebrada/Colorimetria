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
    // Extractor de receta por columnas geométricas independientes
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Segmenta la franja de colorantes de la imagen en columnas verticales
    /// independientes (Código | Nombre | Concentración) y retorna únicamente
    /// filas de colorantes puros, sin ruido de químicos auxiliares.
    ///
    /// Opera bajo dos modos:
    ///  • tableBounds != null → Formato combinado (LegacyCombinedFormat):
    ///    la receta útil termina exactamente donde empieza la tabla CMC inferior.
    ///  • tableBounds == null → Ticket plano (DynamicSplitGridFormat):
    ///    se usa el cuadrante central superior estándar.
    /// </summary>
    public sealed class DynamicSplitGridExtractor
    {
        // ── Rutas de Tesseract ────────────────────────────────────────────────
        private readonly string _tessDataPath;
        private TesseractEngine _engine;
        private readonly object _lock = new object();

        // ── Proporciones geométricas (fracciones del ancho total) ─────────────

        /// <summary>
        /// Columna 0: Código de colorante (primeros 15 % del ancho).
        /// </summary>
        private const double COL_CODE_WIDTH  = 0.15;

        /// <summary>
        /// Columna 1: Nombre del colorante (siguiente 38 % del ancho).
        /// </summary>
        private const double COL_NAME_WIDTH  = 0.38;

        /// <summary>
        /// Columna 2: Concentración % (siguiente 15 % del ancho).
        /// </summary>
        private const double COL_PCT_WIDTH   = 0.15;

        // ── Proporciones geométricas (fracciones de la altura total) ──────────

        /// <summary>
        /// Modo LegacyCombinedFormat: la franja de colorantes ocupa el
        /// 22 % de la altura, terminando en el borde superior de la tabla CMC.
        /// </summary>
        private const double RECIPE_ZONE_HEIGHT_COMBINED = 0.22;

        /// <summary>
        /// Modo DynamicSplitGridFormat: zona superior estándar (18 % – 43 %).
        /// </summary>
        private const double RECIPE_ZONE_TOP_FLAT   = 0.18;
        private const double RECIPE_ZONE_HEIGHT_FLAT = 0.25;

        // ── Constructor ───────────────────────────────────────────────────────

        public DynamicSplitGridExtractor(string tessDataPath = @".\tessdata")
        {
            _tessDataPath = tessDataPath;
        }

        // ── API pública ───────────────────────────────────────────────────────

        /// <summary>
        /// Extrae la lista de colorantes de la imagen usando segmentación
        /// geométrica de columnas.
        /// </summary>
        /// <param name="imagePath">Ruta al archivo PNG.</param>
        /// <param name="tableBounds">
        ///   Coordenadas de la tabla CMC inferior en el espacio de la imagen
        ///   original. Pase <c>null</c> para modo ticket plano.
        /// </param>
        public List<RecipeItem> ExtractRecipePositional(
            string imagePath,
            System.Drawing.Rectangle? tableBounds = null)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return new List<RecipeItem>();

            using (Bitmap bmp = ReportFormatRouter.LoadUniversalImage24bpp(imagePath))
                return ExtractFromBitmap(bmp, tableBounds);
        }

        /// <summary>
        /// Extrae la lista de colorantes directamente desde un Bitmap ya cargado.
        /// </summary>
        public List<RecipeItem> ExtractFromBitmap(
            Bitmap bmp,
            System.Drawing.Rectangle? tableBounds = null)
        {
            var recipeItems = new List<RecipeItem>();
            if (bmp == null) return recipeItems;

            int imgWidth  = bmp.Width;
            int imgHeight = bmp.Height;

            // ── 1. Calcular la franja de colorantes ───────────────────────────
            int zoneTop, zoneHeight;

            if (tableBounds.HasValue && tableBounds.Value.Y > 0)
            {
                // Formato combinado: la receta termina donde empieza la tabla CMC.
                // Tomamos el 22 % de altura hacia arriba desde ese límite.
                int recipeBottom = tableBounds.Value.Y;
                int rawHeight    = (int)(imgHeight * RECIPE_ZONE_HEIGHT_COMBINED);
                zoneTop    = Math.Max(0, recipeBottom - rawHeight);
                zoneHeight = recipeBottom - zoneTop;
            }
            else
            {
                // Ticket plano: cuadrante central superior estándar.
                zoneTop    = (int)(imgHeight * RECIPE_ZONE_TOP_FLAT);
                zoneHeight = (int)(imgHeight * RECIPE_ZONE_HEIGHT_FLAT);
            }

            // Protección de límites
            zoneTop    = Math.Max(0, zoneTop);
            zoneHeight = Math.Min(zoneHeight, imgHeight - zoneTop);
            if (zoneHeight <= 0) return recipeItems;

            // ── 2. Definir los rectángulos de columna ─────────────────────────
            int colCodeW = (int)(imgWidth * COL_CODE_WIDTH);
            int colNameW = (int)(imgWidth * COL_NAME_WIDTH);
            int colPctW  = (int)(imgWidth * COL_PCT_WIDTH);

            var rectCodes = new System.Drawing.Rectangle(0,                        zoneTop, colCodeW, zoneHeight);
            var rectNames = new System.Drawing.Rectangle(colCodeW,                 zoneTop, colNameW, zoneHeight);
            var rectPcts  = new System.Drawing.Rectangle(colCodeW + colNameW,      zoneTop, colPctW,  zoneHeight);

            // ── 3. OCR por columna (zonas limpias e independientes) ───────────
            List<string> codesList = ExtractLinesFromRegion(bmp, rectCodes);
            List<string> namesList = ExtractLinesFromRegion(bmp, rectNames);
            List<string> pctsList  = ExtractLinesFromRegion(bmp, rectPcts);

            // ── 4. Consolidar filas por índice ────────────────────────────────
            int maxRows = Math.Max(namesList.Count, pctsList.Count);

            for (int i = 0; i < maxRows; i++)
            {
                string rawCode = i < codesList.Count ? codesList[i] : "S/C";
                string rawName = i < namesList.Count ? namesList[i] : "";
                string rawPct  = i < pctsList.Count  ? pctsList[i]  : "";

                // Limpiar código: solo dígitos y guion
                rawCode = Regex.Replace(rawCode, @"[^0-9\-]", "").Trim();
                if (rawCode.StartsWith("-") || string.IsNullOrWhiteSpace(rawCode))
                    rawCode = "S/C";

                // Limpiar nombre
                rawName = rawName.Replace("|", "").Replace("\"", "").Trim();

                // Extraer el primer decimal válido de la columna de concentración
                var pctMatch = Regex.Match(rawPct, @"([0-9]+\.[0-9]+)");
                if (!pctMatch.Success || string.IsNullOrWhiteSpace(rawName))
                    continue;

                string pctStr = pctMatch.Groups[1].Value;
                if (!double.TryParse(pctStr,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out _))
                    continue;

                recipeItems.Add(new RecipeItem
                {
                    Code       = rawCode,
                    Name       = rawName,
                    Percentage = pctStr + "%"
                });
            }

            return recipeItems;
        }

        // ── OCR por región ────────────────────────────────────────────────────

        private List<string> ExtractLinesFromRegion(
            Bitmap originalBmp,
            System.Drawing.Rectangle region)
        {
            var lines = new List<string>();
            if (region.Width <= 0 || region.Height <= 0) return lines;

            // Intersectar con límites de la imagen para evitar desbordamientos
            region = System.Drawing.Rectangle.Intersect(
                region,
                new System.Drawing.Rectangle(0, 0, originalBmp.Width, originalBmp.Height));
            if (region.IsEmpty) return lines;

            using (Bitmap regionBmp = new Bitmap(region.Width, region.Height,
                originalBmp.PixelFormat == System.Drawing.Imaging.PixelFormat.Format24bppRgb
                    ? System.Drawing.Imaging.PixelFormat.Format24bppRgb
                    : System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(regionBmp))
            {
                g.Clear(System.Drawing.Color.White);
                g.DrawImage(
                    originalBmp,
                    new System.Drawing.Rectangle(0, 0, region.Width, region.Height),
                    region,
                    GraphicsUnit.Pixel);

                string text = RunTesseractOnBitmap(regionBmp);

                using (var reader = new StringReader(text))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        // Filtrar encabezados de columna residuales del OCR
                        if (line.IndexOf("Dye code", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                        if (line.IndexOf("Dye name", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                        lines.Add(line);
                    }
                }
            }

            return lines;
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

                    // Escalar 3x para mejorar la detección de puntos decimales pequeños
                    int scaledW = bmp.Width  * 3;
                    int scaledH = bmp.Height * 3;

                    using (Bitmap scaled = new Bitmap(scaledW, scaledH))
                    using (Graphics g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode =
                            System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
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
}
