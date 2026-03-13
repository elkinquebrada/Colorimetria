using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging; // NO usar ImageFormat directamente: usar calificación completa al guardar
using System.IO;
using System.Text.RegularExpressions;
using Tesseract;

namespace Color
{
    // ── Modelos de datos ─────────────────────────────────────────────────────────
    public class RecipeItem
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Percentage { get; set; }
        public override string ToString()
        {
            return string.Format("{0} {1,-30} {2}", Code, Name, Percentage);
        }
    }

    public class LabValues
    {
        public string L { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string dL { get; set; }
        public string da { get; set; }
        public string dB { get; set; }
        public string cde { get; set; }
        public string PF { get; set; }
        public override string ToString()
        {
            return string.Format(
                "L={0} A={1} B={2} dL={3} da={4} dB={5} CDE={6} P/F={7}",
                L, A, B, dL, da, dB, cde, PF);
        }
    }

    /// <summary>
    /// Resultado completo de procesar la imagen de RECETA (Shade History Report).
    /// </summary>
    public class ShadeExtractionResult
    {
        public List<RecipeItem> Recipe { get; set; } = new List<RecipeItem>();
        public LabValues Lab { get; set; }
        public BatchMeasure Batch { get; set; }
        public string RawText { get; set; }
        public bool Success
        {
            get { return (Recipe != null && Recipe.Count > 0) || Lab != null || Batch != null; }
        }
    }

    /// <summary>Fila de medición del lote (L A B dL dC dH dE P/F).</summary>
    public class BatchMeasure
    {
        public string L { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string dL { get; set; }
        public string dC { get; set; }
        public string dH { get; set; }
        public string dE { get; set; }
        public string PF { get; set; }
        public override string ToString()
        {
            return string.Format(
                "L={0} A={1} B={2} dL={3} dC={4} dH={5} dE={6} P/F={7}",
                L, A, B, dL, dC, dH, dE, PF);
        }
    }

    // ── Extractor ────────────────────────────────────────────────────────────────
    public class ShadeReportExtractor
    {
        private readonly string _tessdataPath;
        private const string OCR_LANG = "eng+spa";

        // ── Expresiones regulares (más tolerantes) ────────────────────────────
        // Receta: 8 dígitos + nombre + número (±, coma/punto) + % con/sin espacio
        private static readonly Regex RecipeRegex = new Regex(
            @"(?mi)^\s*(\d{8})\s+([A-Z0-9][A-Z0-9\s\-\(\)\/\.,_]+?)\s+([+-]?\d+(?:[.,]\d+)?)\s*%",
            RegexOptions.Compiled);

        // Cabecera de la tabla LAB
        private static readonly Regex LabHeaderRegex = new Regex(
            @"L\s+A\s+B\s+dL\s+da\s+dB\s+cde\s+P\/F",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Valores LAB (acepta signo y coma/punto)
        private static readonly Regex LabValuesRegex = new Regex(
            @"([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([FP])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public ShadeReportExtractor(string tessdataPath)
        {
            _tessdataPath = tessdataPath;
        }

        // ── OCR desde ruta de archivo ──────────────────────────────────────────
        public string ExtractTextFromFile(string imagePath)
        {
            using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
            using (var img = Pix.LoadFromFile(imagePath))
            using (var page = engine.Process(img, PageSegMode.SingleBlock))
            {
                return page.GetText();
            }
        }

        // ── OCR desde Bitmap (integración directa con UI) ─────────────────────
        public string ExtractTextFromBitmap(Bitmap bmp)
        {
            // Guardamos a PNG temporal para evitar ambigüedad de ImageFormat
            string tmp = Path.Combine(Path.GetTempPath(), string.Format("ocr_{0:N}.png", Guid.NewGuid()));
            try
            {
                bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                // Reintentos PSM: SingleBlock → SingleColumn → SparseText
                string text = RunOcrWithRetries(tmp, new[] {
                    PageSegMode.SingleBlock, PageSegMode.SingleColumn, PageSegMode.SparseText
                });
                return text ?? string.Empty;
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* no-op */ }
            }
        }

        private string RunOcrWithRetries(string pngPath, PageSegMode[] modes)
        {
            foreach (var psm in modes)
            {
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var pix = Pix.LoadFromFile(pngPath))
                using (var page = engine.Process(pix, psm))
                {
                    string t = page.GetText();
                    if (!string.IsNullOrWhiteSpace(t)) return t;
                }
            }
            return string.Empty;
        }

        // ── Extracción completa (texto → receta + LAB + batch) ────────────────
        public ShadeExtractionResult ExtractAll(string ocrText)
        {
            return new ShadeExtractionResult
            {
                RawText = ocrText,
                Recipe = ExtractRecipe(ocrText),
                Lab = ExtractLabValues(ocrText),
                Batch = ExtractBatchMeasure(ocrText)
            };
        }

        // ── Método todo‑en‑uno desde Bitmap ───────────────────────────────────
        public ShadeExtractionResult ExtractFromBitmap(Bitmap bmp)
        {
            // OCR general primero
            string text = ExtractTextFromBitmap(bmp);
            var result = ExtractAll(text);

            // Receta con recorte dedicado (mejor precisión en %) con PSM retries
            try
            {
                var recipeFromCrop = ExtractRecipeFromCrop(bmp);
                if (recipeFromCrop != null && recipeFromCrop.Count > 0)
                    result.Recipe = recipeFromCrop;
            }
            catch { /* fallback al OCR completo ya en result.Recipe */ }

            // Batch con recorte dedicado (whitelist + PSM retries)
            try
            {
                var bm = ExtractBatchMeasureFromBitmap(bmp);
                if (bm != null) result.Batch = bm;
            }
            catch { /* no-op */ }

            return result;
        }

        /// <summary>Extrae la receta usando un recorte de la zona ~34–62% de la altura.</summary>
        public List<RecipeItem> ExtractRecipeFromCrop(Bitmap original)
        {
            int top = (int)(original.Height * 0.34);
            int bot = (int)(original.Height * 0.62);
            int cropH = bot - top;
            if (cropH <= 0) return null;

            int newW = original.Width * 4;
            int newH = cropH * 4;

            var crop = new Bitmap(newW, newH, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(crop))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original,
                    new Rectangle(0, 0, newW, newH),
                    new Rectangle(0, top, original.Width, cropH),
                    GraphicsUnit.Pixel);
            }
            var enhanced = ApplyContrast(crop, 1.5f);
            crop.Dispose();

            string tmp = Path.Combine(Path.GetTempPath(), string.Format("recipe_{0:N}.png", Guid.NewGuid()));
            try
            {
                enhanced.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                enhanced.Dispose();

                List<RecipeItem> best = null;

                // Intento 1: SingleBlock
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var pixImg = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(pixImg, PageSegMode.SingleBlock))
                {
                    best = ExtractRecipe(page.GetText());
                }

                // Intento 2: SingleColumn
                if (best == null || best.Count == 0)
                {
                    using (var engine2 = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                    using (var pixImg2 = Pix.LoadFromFile(tmp))
                    using (var page2 = engine2.Process(pixImg2, PageSegMode.SingleColumn))
                    {
                        var best2 = ExtractRecipe(page2.GetText());
                        if (best2 != null && best2.Count > 0) best = best2;
                    }
                }

                // Intento 3: SparseText
                if (best == null || best.Count == 0)
                {
                    using (var engine3 = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                    using (var pixImg3 = Pix.LoadFromFile(tmp))
                    using (var page3 = engine3.Process(pixImg3, PageSegMode.SparseText))
                    {
                        var best3 = ExtractRecipe(page3.GetText());
                        if (best3 != null && best3.Count > 0) best = best3;
                    }
                }

                return best;
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* no-op */ }
            }
        }

        private static Bitmap ApplyContrast(Bitmap src, float contrast)
        {
            // Matriz de contraste+brillo leve
            float b = -0.2f;
            float[][] m =
            {
                new float[] { contrast, 0f, 0f, 0f, 0f },
                new float[] { 0f, contrast, 0f, 0f, 0f },
                new float[] { 0f, 0f, contrast, 0f, 0f },
                new float[] { 0f, 0f, 0f, 1f, 0f },
                new float[] { b, b, b, 0f, 1f }
            };
            var result = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(result))
            using (var ia = new ImageAttributes())
            {
                ia.SetColorMatrix(new ColorMatrix(m));
                g.DrawImage(src,
                    new Rectangle(0, 0, src.Width, src.Height),
                    0, 0, src.Width, src.Height,
                    GraphicsUnit.Pixel, ia);
            }
            return result;
        }

        // ── Método todo‑en‑uno desde archivo ───────────────────────────────────
        public ShadeExtractionResult ExtractFromFile(string imagePath)
        {
            string text = ExtractTextFromFile(imagePath);
            return ExtractAll(text);
        }

        // ═══════════════════════════════════════════════════════════════════════
        // A. EXTRAER RECETA
        // ═══════════════════════════════════════════════════════════════════════
        public List<RecipeItem> ExtractRecipe(string ocrText)
        {
            var recipe = new List<RecipeItem>();
            if (string.IsNullOrEmpty(ocrText)) return recipe;

            var matches = RecipeRegex.Matches(ocrText);
            foreach (Match match in matches)
            {
                string code = match.Groups[1].Value.Trim();
                string name = match.Groups[2].Value.Trim();
                string rawPct = match.Groups[3].Value.Trim();

                // Normalizaciones: coma → punto
                rawPct = rawPct.Replace(',', '.');

                // Corrección típica: OCR confunde "1." con "L" al principio
                if (Regex.IsMatch(rawPct, @"^[Ll]\d{4}$"))
                    rawPct = "1.1" + rawPct.Substring(1); // L7826 → 1.17826
                else if (Regex.IsMatch(rawPct, @"^[Ll]\d{5}$"))
                    rawPct = "1." + rawPct.Substring(1);  // L17826 → 1.17826
                else if (Regex.IsMatch(rawPct, @"^[Ll]\d"))
                    rawPct = "1" + rawPct.Substring(1);

                // Si no tiene punto, intenta insertar según longitud (5→X.XXXX, 6→X.XXXXX)
                string formattedPct = rawPct;
                if (rawPct.IndexOf('.') < 0)
                {
                    string s = rawPct.TrimStart('-');
                    int len = s.Length;
                    if (len == 5)
                        formattedPct = rawPct.Substring(0, rawPct.Length - 4) + "." + rawPct.Substring(rawPct.Length - 4);
                    else if (len >= 6)
                        formattedPct = rawPct.Substring(0, rawPct.Length - 5) + "." + rawPct.Substring(rawPct.Length - 5);
                }

                double pctVal;
                if (double.TryParse(formattedPct, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out pctVal))
                {
                    formattedPct = pctVal.ToString("0.#####", System.Globalization.CultureInfo.InvariantCulture);
                }

                recipe.Add(new RecipeItem
                {
                    Code = code,
                    Name = name,
                    Percentage = formattedPct + "%"
                });
            }
            return recipe;
        }

        // ═══════════════════════════════════════════════════════════════════════
        // B. EXTRAER L A B dL da dB cde P/F
        // ═══════════════════════════════════════════════════════════════════════
        public LabValues ExtractLabValues(string ocrText)
        {
            if (string.IsNullOrEmpty(ocrText)) return null;

            var headerMatch = LabHeaderRegex.Match(ocrText);
            if (!headerMatch.Success) return null;

            string afterHeader = ocrText.Substring(headerMatch.Index + headerMatch.Length);
            var valuesMatch = LabValuesRegex.Match(afterHeader);
            if (!valuesMatch.Success) return null;

            // Normaliza comas a puntos
            string G(int i) { return valuesMatch.Groups[i].Value.Replace(',', '.'); }

            return new LabValues
            {
                L = G(1),
                A = G(2),
                B = G(3),
                dL = G(4),
                da = G(5),
                dB = G(6),
                cde = G(7),
                PF = valuesMatch.Groups[8].Value.ToUpperInvariant()
            };
        }

        // ═══════════════════════════════════════════════════════════════════════
        // C. EXTRAER STD LAB + BATCH (L A B dL dC dH dE P/F)
        // ═══════════════════════════════════════════════════════════════════════

        // ── Extrae batch desde la imagen con recorte específico ───────────────
        public BatchMeasure ExtractBatchMeasureFromBitmap(Bitmap original)
        {
            // Recortar zona del batch: entre 70% y 78% de la altura
            int top = (int)(original.Height * 0.70);
            int bot = (int)(original.Height * 0.78);
            int cropH = bot - top;
            if (cropH <= 0) return null;

            // Escalar 4x y aumentar contraste
            int newW = original.Width * 4;
            int newH = cropH * 4;

            var crop = new Bitmap(newW, newH, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(crop))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original,
                    new Rectangle(0, 0, newW, newH),
                    new Rectangle(0, top, original.Width, cropH),
                    GraphicsUnit.Pixel);
            }

            var enhanced = ApplyHighContrast(crop);
            crop.Dispose();

            string tmp = Path.Combine(Path.GetTempPath(), string.Format("batch_{0:N}.png", Guid.NewGuid()));
            try
            {
                enhanced.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                enhanced.Dispose();

                BatchMeasure bm = null;

                // Intento 1: sin whitelist + SingleBlock
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var img = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(img, PageSegMode.SingleBlock))
                {
                    bm = ParseBatchLine(page.GetText());
                }

                // Intento 2: SingleColumn
                if (bm == null)
                {
                    using (var engine2 = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                    using (var img2 = Pix.LoadFromFile(tmp))
                    using (var page2 = engine2.Process(img2, PageSegMode.SingleColumn))
                    {
                        bm = ParseBatchLine(page2.GetText());
                    }
                }

                // Intento 3: SparseText
                if (bm == null)
                {
                    using (var engine3 = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                    using (var img3 = Pix.LoadFromFile(tmp))
                    using (var page3 = engine3.Process(img3, PageSegMode.SparseText))
                    {
                        bm = ParseBatchLine(page3.GetText());
                    }
                }

                return bm;
            }
            finally
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* no-op */ }
            }
        }

        private static Bitmap ApplyHighContrast(Bitmap src)
        {
            return ApplyContrast(src, 2.5f);
        }

        private static BatchMeasure ParseBatchLine(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            string[] lines = ocrText.Split('\n');
            for (int li = 0; li < lines.Length; li++)
            {
                string rawLine = lines[li];
                if (string.IsNullOrWhiteSpace(rawLine)) continue;

                string line = rawLine.Trim();

                // Buscar F o P (Pass/Fail) en la linea
                var pfm = Regex.Match(line, @"\b([FfPp])\b");
                if (!pfm.Success) continue;
                string pf = pfm.Groups[1].Value.ToUpperInvariant();

                // Extraer todos los números de la línea (con signo y decimales)
                var values = new List<double>();
                foreach (Match nm in Regex.Matches(line, @"-?\d+(?:\.\d+)?"))
                {
                    double v;
                    if (double.TryParse(nm.Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out v))
                        values.Add(v);
                }

                // Los últimos 7 números antes del P/F son: L A B dL dC dH dE
                if (values.Count < 7) continue;
                var tail = values.GetRange(values.Count - 7, 7);

                return new BatchMeasure
                {
                    L = FixLabToken(tail[0]),
                    A = FixLabToken(tail[1]),
                    B = FixLabToken(tail[2]),
                    dL = FixDeltaToken(tail[3]),
                    dC = FixDeltaToken(tail[4]),
                    dH = FixDeltaToken(tail[5]),
                    dE = FixDeltaToken(tail[6]),
                    PF = pf
                };
            }
            return null;
        }

        /// Corrige valores L*a*b* cuando OCR omitió decimal: 3224 → 32.24
        private static string FixLabToken(double v)
        {
            if (v >= -100 && v <= 100) return v.ToString(System.Globalization.CultureInfo.InvariantCulture);

            double sign = v < 0 ? -1 : 1;
            string s = ((long)Math.Abs(v)).ToString();
            if (s.Length >= 4)
            {
                double fixed_v = sign * double.Parse(
                    s.Substring(0, s.Length - 2) + "." + s.Substring(s.Length - 2),
                    System.Globalization.CultureInfo.InvariantCulture);
                if (fixed_v >= -100 && fixed_v <= 100)
                    return fixed_v.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            return v.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        /// Corrige deltas donde OCR omitió decimal: 118 → 1.18, 80 → 0.80
        private static string FixDeltaToken(double v)
        {
            if (v >= -9.99 && v <= 9.99)
                return v.ToString(System.Globalization.CultureInfo.InvariantCulture);

            double sign = v < 0 ? -1 : 1;
            string s = ((long)Math.Abs(v)).ToString();

            double candidate;
            if (s.Length >= 3)
                candidate = sign * double.Parse(s[0] + "." + s.Substring(1),
                    System.Globalization.CultureInfo.InvariantCulture);
            else
                candidate = sign * double.Parse("0." + s,
                    System.Globalization.CultureInfo.InvariantCulture);

            if (candidate >= -9.99 && candidate <= 9.99)
                return candidate.ToString(System.Globalization.CultureInfo.InvariantCulture);

            return v.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        public BatchMeasure ExtractBatchMeasure(string ocrText)
        {
            return ParseBatchLine(ocrText);
        }

        private static int FindBatchHeader(string text)
        {
            var rx = new Regex(
                @"Batch\s*Id|Lot\s*No|Dyelot\s*Date",
                RegexOptions.IgnoreCase);
            var m = rx.Match(text);
            return m.Success ? m.Index : -1;
        }

        /// Restaura decimal si el OCR lo omitió (ej: "3224" → "32.24")
        private static string RestoreDecimal(string token, int decimals)
        {
            if (string.IsNullOrWhiteSpace(token)) return token;
            token = token.Trim();
            if (token.Contains(".")) return token;
            if (token.Length <= decimals) return token;
            return token.Substring(0, token.Length - decimals) + "." + token.Substring(token.Length - decimals);
        }
    }
}