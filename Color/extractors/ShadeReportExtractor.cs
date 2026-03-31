using OCR;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using Tesseract;

namespace Color
{
    //-----------------------------------------------------------------------
    // MODELOS DE DATOS (RECETA + LAB + BATCH)
    //-----------------------------------------------------------------------

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
    }

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
    }

    //-----------------------------------------------------------------------
    // RESULTADO COMPLETO DEL EXTRACTOR
    //-----------------------------------------------------------------------

    public class ShadeExtractionResult
    {
        public string ShadeName { get; set; }
        public List<RecipeItem> Recipe { get; set; } = new List<RecipeItem>();
        public LabValues Lab { get; set; }
        public BatchMeasure Batch { get; set; }
        public string RawText { get; set; }

        public bool Success
        {
            get
            {
                return (Recipe != null && Recipe.Count > 0)
                        && Lab != null
                        && Batch != null;
            }
        }
    }

    //-----------------------------------------------------------------------
    // INICIO DEL EXTRACTOR
    //-----------------------------------------------------------------------

    public class ShadeReportExtractor
    {
        private readonly string _tessdataPath;
        private const string OCR_LANG = "eng+spa";

        // Regex del archivo original
        private static readonly Regex RecipeRegex = new Regex(
            @"(?mi)[^\d]*(\d{8})\s+([A-Za-z][A-Za-z0-9\s\-\(\)\/\.,_]+?)\s+([0-9OL][0-9OL.,]*)\s*%",
            RegexOptions.Compiled);

        private static readonly Regex LabHeaderRegex = new Regex(
            @"L\s+A\s+B\s+dL\s+da\s+dB\s+cde\s+P\/F",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LabValuesRegex = new Regex(
            @"([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([FP])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        //-------------------------------------------------------------------
        // Constructor
        //-------------------------------------------------------------------
        public ShadeReportExtractor(string tessdataPath)
        {
            _tessdataPath = tessdataPath;
        }

        //-------------------------------------------------------------------
        // OCR desde archivo — NO BORRAR PNG AQUÍ
        //-------------------------------------------------------------------
        public string ExtractTextFromFile(string imagePath)
        {
            using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
            using (var img = Pix.LoadFromFile(imagePath))
            using (var page = engine.Process(img, PageSegMode.SingleBlock))
            {
                return page.GetText() ?? string.Empty;
            }
        }

        //-------------------------------------------------------------------
        // OCR desde BITMAP — AQUÍ ES DONDE TUS PNG SE ESTABAN BORRANDO
        //-------------------------------------------------------------------
        public string ExtractTextFromBitmap(Bitmap bmp)
        {
            string tmp = Path.Combine(Path.GetTempPath(),
                                      "ocr_" + Guid.NewGuid().ToString("N") + ".png");

            bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);

            // ❗ PROTECCIÓN — asegurarse de que el archivo existe
            if (!File.Exists(tmp))
                throw new Exception("Archivo temporario no existe: " + tmp);

            // NUNCA BORRAR AQUÍ → Form1.cs lo borrará al final
            using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
            using (var pix = Pix.LoadFromFile(tmp))
            using (var page = engine.Process(pix))
            {
                return page.GetText() ?? string.Empty;
            }
        }

        //-------------------------------------------------------------------
        // Flujo completo: OCR → Recipe + Lab + Batch
        //-------------------------------------------------------------------
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

        //-------------------------------------------------------------------
        // Método todo-en-uno desde Bitmap original
        //-------------------------------------------------------------------
        public ShadeExtractionResult ExtractFromBitmap(Bitmap bmp)
        {
            string text = ExtractTextFromBitmap(bmp) ?? string.Empty;

            var result = ExtractAll(text);

            // DIAGNÓSTICO: guardar log en Desktop para identificar por qué falla la receta
            try
            {
                var dbg = new System.Text.StringBuilder();
                dbg.AppendLine("=== ShadeReportExtractor DEBUG ===");
                dbg.AppendLine(string.Format("OCR texto completo ({0} chars)", text.Length));
                dbg.AppendLine(string.Format("Recipe desde OCR completo: {0} items", result.Recipe != null ? result.Recipe.Count : 0));
                dbg.AppendLine("--- Primeras 800 chars del OCR ---");
                dbg.AppendLine(text.Length > 800 ? text.Substring(0, 800) : text);
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(
                        System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop),
                        "debug_shade_ocr.txt"),
                    dbg.ToString());
            }
            catch { }

            // Intentar mejorar la receta con OCR dirigido sobre el recorte de la zona de ingredientes.
            try
            {
                var shadeName = ExtractShadeNameFromBitmap(bmp);
                if (!string.IsNullOrWhiteSpace(shadeName))
                    result.ShadeName = shadeName;
            }
            catch { }

            try
            {
                var r2 = ExtractRecipeFromCrop(bmp);
                if (r2 != null && r2.Count > 0
                    && r2.Count >= (result.Recipe != null ? result.Recipe.Count : 0))
                    result.Recipe = r2;
            }
            catch { }

            try
            {
                var b2 = ExtractBatchMeasureFromBitmap(bmp);
                if (b2 != null)
                    result.Batch = b2;
            }
            catch { }

            return result;
        }

        //-------------------------------------------------------------------
        // EXTRACCIÓN DESDE EXCEL (NUEVO MÉTODO)
        //-------------------------------------------------------------------
        public ShadeExtractionResult ExtractFromExcel(string excelPath)
        {
            var recipe = ExcelReader.LoadRecipe(excelPath);

            return new ShadeExtractionResult
            {
                ShadeName = "[Desde EXCEL]",
                Recipe = recipe,
                Lab = new LabValues(),           
                Batch = new BatchMeasure(),     
                RawText = "[EXCEL importado]"
            };
        }

        public List<RecipeItem> ExtractRecipe(string ocrText)
        {
            var list = new List<RecipeItem>();
            if (string.IsNullOrWhiteSpace(ocrText)) return list;

            // Normalizar errores OCR comunes antes de aplicar el regex:
            string normalized = Regex.Replace(ocrText, @"\bL(\d)", "1$1");
            normalized = Regex.Replace(normalized, @"(\d)L(\d)", "$11$2");

            var matches = RecipeRegex.Matches(normalized);

            foreach (Match m in matches)
            {
                var code = m.Groups[1].Value.Trim();
                var name = m.Groups[2].Value.Trim();
                var pctRaw = m.Groups[3].Value.Trim()
                    .Replace('O', '0')   
                    .Replace('L', '1')   
                    .Replace(',', '.');  

                // Restaurar punto decimal si el OCR lo omitió
                string pctFixed = pctRaw;
                // Cambia la línea 253 por esta:
                if (!pctFixed.Contains(".") && pctFixed.Length >= 2)
                {
                    // Insertamos el punto después del primer carácter
                    pctFixed = pctFixed.Insert(1, ".");
                }
                list.Add(new RecipeItem
                {
                    Code = code,
                    Name = name,
                    Percentage = pctFixed + "%"
                });
            }

            return list;
        }

        public LabValues ExtractLabValues(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            var header = LabHeaderRegex.Match(ocrText);
            if (!header.Success) return null;

            var after = ocrText.Substring(header.Index + header.Length);
            var values = LabValuesRegex.Match(after);

            if (!values.Success) return null;

            string G(int i) => values.Groups[i].Value.Replace(',', '.');

            return new LabValues
            {
                L = G(1),
                A = G(2),
                B = G(3),
                dL = G(4),
                da = G(5),
                dB = G(6),
                cde = G(7),
                PF = values.Groups[8].Value
            };
        }

        public BatchMeasure ExtractBatchMeasure(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            string[] lines = ocrText.Split('\n');

            foreach (string raw in lines)
            {
                var line = raw.Trim();

                var pfMatch = Regex.Match(line, @"\b[FfPp]\b");
                if (!pfMatch.Success) continue;

                var nums = new List<double>();

                foreach (Match m in Regex.Matches(line, @"-?\d+(?:\.\d+)?"))
                    if (double.TryParse(m.Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double v))
                        nums.Add(v);

                if (nums.Count < 7) continue;

                int start = nums.Count - 7;

                return new BatchMeasure
                {
                    L = nums[start + 0].ToString(),
                    A = nums[start + 1].ToString(),
                    B = nums[start + 2].ToString(),
                    dL = nums[start + 3].ToString(),
                    dC = nums[start + 4].ToString(),
                    dH = nums[start + 5].ToString(),
                    dE = nums[start + 6].ToString(),
                    PF = pfMatch.Value.ToUpper()
                };
            }

            return null;
        }

        //-------------------------------------------------------------------
        // REEMPLAZO Y NUEVOS (Recorte dirigido de alta fidelidad)
        //-------------------------------------------------------------------
        private string ExtractShadeNameFromBitmap(Bitmap original)
        {
            // El Shade Name está en la parte superior izquierda (aprox 0% al 15% de altura y 0 al 60% de ancho)
            int h = (int)(original.Height * 0.15);
            int w = (int)(original.Width * 0.60);

            if (h <= 0 || w <= 0) return null;

            // Escalar 3x mejora consistentemente las lecturas alfanuméricas complejas (e.g. SYLB8792 vs sesm2)
            Bitmap crop = new Bitmap(w * 3, h * 3);
            using (var g = Graphics.FromImage(crop))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original,
                    new Rectangle(0, 0, crop.Width, crop.Height),
                    new Rectangle(0, 0, w, h),
                    GraphicsUnit.Pixel);
            }

            string tmp = Path.Combine(Path.GetTempPath(),
                "shade_" + Guid.NewGuid().ToString("N") + ".png");

            crop.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
            crop.Dispose();

            string ocrText = string.Empty;
            try
            {
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var pix = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                {
                    ocrText = page.GetText() ?? string.Empty;
                }
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }

            // Expresión regular para ubicar "Shade Name:" y capturar hasta antes de "Dyehouse:" o el fin de la línea
            var match = Regex.Match(ocrText, @"Shade Name:\s*(.+?)(?=\s*Dyehouse:|\s*$)", RegexOptions.IgnoreCase | RegexOptions.Multiline);

            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            return null;
        }

        private List<RecipeItem> ExtractRecipeFromCrop(Bitmap original)
        {
            // Zona de ingredientes en el Shade History Report:
            int top = (int)(original.Height * 0.24);
            int bot = (int)(original.Height * 0.50);
            int h = bot - top;

            if (h <= 0) return null;

            Bitmap crop = new Bitmap(original.Width * 4, h * 4);

            using (var g = Graphics.FromImage(crop))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original,
                    new Rectangle(0, 0, crop.Width, crop.Height),
                    new Rectangle(0, top, original.Width, h),
                    GraphicsUnit.Pixel);
            }

            string tmp = Path.Combine(Path.GetTempPath(),
                "recipe_" + Guid.NewGuid().ToString("N") + ".png");

            crop.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
            crop.Dispose();

            // FIX: ejecutar OCR sobre el recorte y pasar el TEXTO resultante
            string ocrText = string.Empty;
            try
            {
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var pix = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                {
                    ocrText = page.GetText() ?? string.Empty;
                }
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }

            return ExtractRecipe(ocrText);
        }

        // REEMPLAZO IMPLEMENTADO: Extraer BatchMeasure usando Recorte dirigido
        private BatchMeasure ExtractBatchMeasureFromBitmap(Bitmap original)
        {
            // La fila de Batch Measure está en la parte inferior (aprox 65% a 95% de la altura).
            int top = (int)(original.Height * 0.60);
            int h = (int)(original.Height * 0.35);

            if (h <= 0 || top + h > original.Height) return null;

            // Recortar y escalar 3x para mejorar calidad OCR numérico (igual que Recipe)
            Bitmap crop = new Bitmap(original.Width * 3, h * 3);

            using (var g = Graphics.FromImage(crop))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original,
                    new Rectangle(0, 0, crop.Width, crop.Height),
                    new Rectangle(0, top, original.Width, h),
                    GraphicsUnit.Pixel);
            }

            string tmp = Path.Combine(Path.GetTempPath(),
                "batch_" + Guid.NewGuid().ToString("N") + ".png");

            crop.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
            crop.Dispose();

            string ocrText = string.Empty;
            try
            {
                using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                using (var pix = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(pix, PageSegMode.SingleBlock))
                {
                    ocrText = page.GetText() ?? string.Empty;
                }
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }

            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            // Para proteger de comas decimales europeas o confusas:
            ocrText = ocrText.Replace(',', '.').Replace('O', '0').Replace('o', '0');
            string[] lines = ocrText.Split('\n');

            foreach (string raw in lines)
            {
                var line = raw.Trim();

                // Miremos si la línea termina en un 'P' o 'F' aislado
                var pfMatch = Regex.Match(line, @"\b([FfPp])\s*$");
                if (!pfMatch.Success) continue;

                // Extraemos todos los números, forzando un punto en caso de pérdida OCR a nivel interno si es necesario.
                var nums = new List<double>();
                var rawNums = new List<string>();

                foreach (Match m in Regex.Matches(line, @"-?\d{1,5}(?:\.\d+)?"))
                {
                    if (double.TryParse(m.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double v))
                    {
                        nums.Add(v);
                        rawNums.Add(m.Value);
                    }
                }

                // Necesitamos al menos los 7 valores de Lab + dL dC dH dE
                if (nums.Count >= 7)
                {
                    int start = nums.Count - 7;

                    // Función local para restaurar el decimal si OCR lo ignoró (ej: 3224 -> 32.24)
                    string FixNum(string s)
                    {
                        if (s.Contains(".")) return s;
                        bool neg = s.StartsWith("-");
                        string abs = neg ? s.Substring(1) : s;
                        if (abs.Length >= 3 && abs.Length <= 5)
                            return (neg ? "-" : "") + abs.Insert(abs.Length - 2, ".");

                        // Para deltas como 0.80 que se leen como 80 o 28
                        if (abs.Length == 2)
                            return (neg ? "-" : "") + "0." + abs;
                        if (abs.Length == 1)
                            return (neg ? "-" : "") + "0.0" + abs;

                        return s;
                    }

                    string rawL = FixNum(rawNums[start + 0]);
                    string rawA = FixNum(rawNums[start + 1]);

                    // Prueba de cordura, L* debe estar entre 0 y 105
                    double L = double.Parse(rawL, System.Globalization.CultureInfo.InvariantCulture);
                    if (L < 0 || L > 105) continue;

                    // A* debe estar entre -150 y 150
                    double A = double.Parse(rawA, System.Globalization.CultureInfo.InvariantCulture);
                    if (Math.Abs(A) > 150) continue;

                    return new BatchMeasure
                    {
                        L = rawL,
                        A = rawA,
                        B = FixNum(rawNums[start + 2]),
                        dL = FixNum(rawNums[start + 3]),
                        dC = FixNum(rawNums[start + 4]),
                        dH = FixNum(rawNums[start + 5]),
                        dE = FixNum(rawNums[start + 6]),
                        PF = pfMatch.Groups[1].Value.ToUpper()
                    };
                }
            }

            // Fallback original con el texto procesado por OCR dirigido si fallan los filtros agresivos
            return ExtractBatchMeasure(ocrText);
        }
    }
}