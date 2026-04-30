using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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
        public string DL { get; set; }
        public string DA { get; set; }
        public string DB { get; set; }
        public string CDE { get; set; }
        public string PF { get; set; }
    }

    public class BatchMeasure
    {
        public string L { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string DL { get; set; }
        public string DC { get; set; }
        public string DH { get; set; }
        public string DE { get; set; }
        public string PF { get; set; }
        public string LotNo { get; set; }
    }

    //-----------------------------------------------------------------------
    // RESULTADO COMPLETO DEL EXTRACTOR
    //-----------------------------------------------------------------------

    public class ShadeExtractionResult
    {
        public string ShadeName { get; set; }
        public string LotNo { get; set; }
        public string DtMain { get; set; }

        // NUEVOS: Valores L* a* b* del Estándar (Std L A B)
        public string StdL { get; set; }
        public string StdA { get; set; }
        public string StdB { get; set; }
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
            @"(?mi)[^\d]*(\d{8})\s+(.+?)\s+([0-9OL\|Ll\s\.,]{2,})(?:\s*%)?",
            RegexOptions.Compiled);

        private static readonly Regex LabHeaderRegex = new Regex(
            @"L\s+A\s+B\s+dL\s+(?:da|dC)\s+(?:dB|dH)\s+(?:cde|dE)\s+P\/F",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex LabValuesRegex = new Regex(
            @"([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([+-]?\d+(?:[.,]\d+)?)\s+([FP])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Regex para extraer el DT Main del encabezado (ej: "DT Main: DFC12")
        private static readonly Regex DtMainRegex = new Regex(
            @"DT\s*Main[:\s]+([A-Z0-9]+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Regex para capturar Std L A B (Estándar) - Permite espacios opcionales alrededor del punto/coma
        private static readonly Regex StdLabRegex = new Regex(
            @"(?i)Std\s?.*?\s*([-+]?\d+(?:\s?[.,]\s?\d+)?)\s+([-+]?\d+(?:\s?[.,]\s?\d+)?)\s+([-+]?\d+(?:\s?[.,]\s?\d+)?)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Variable estática para modo de emergencia / compatibilidad cruzada entre ventanas
        public static ShadeExtractionResult LastResult { get; set; }

        //-------------------------------------------------------------------
        // Constructor
        //-------------------------------------------------------------------
        public ShadeReportExtractor(string tessdataPath)
        {
            _tessdataPath = tessdataPath;
        }

        //-------------------------------------------------------------------
        //                OCR desde archivo 
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
        //                  OCR desde BITMAP 
        //-------------------------------------------------------------------
        public string ExtractTextFromBitmap(Bitmap bmp)
        {
            string tmp = Path.Combine(Path.GetTempPath(),
                                      "ocr_" + Guid.NewGuid().ToString("N") + ".png");

            // MEJORA: Si la imagen es pequeña, escalar 2x antes de procesar para mejorar OCR
            if (bmp.Width < 1200 || bmp.Height < 1000)
            {
                int newW = bmp.Width * 2;
                int newH = bmp.Height * 2;
                using (Bitmap scaled = new Bitmap(newW, newH))
                using (Graphics g = Graphics.FromImage(scaled))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(bmp, 0, 0, newW, newH);
                    scaled.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            else
            {
                bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
            }

            //  asegurarse de que el archivo existe
            if (!File.Exists(tmp))
                throw new Exception("Archivo temporario no existe: " + tmp);

            using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
            using (var pix = Pix.LoadFromFile(tmp))
            using (var page = engine.Process(pix))
            {
                var text = page.GetText() ?? string.Empty;
                if (File.Exists(tmp)) try { File.Delete(tmp); } catch { }
                return text;
            }
        }

        //-------------------------------------------------------------------
        // Flujo completo: OCR → Recipe + Lab + Batch
        //-------------------------------------------------------------------
        public ShadeExtractionResult ExtractAll(string ocrText)
        {
            var res = new ShadeExtractionResult
            {
                RawText = ocrText,
                DtMain = ExtractDtMain(ocrText),
                Recipe = ExtractRecipe(ocrText),
                Lab = ExtractLabValues(ocrText),
                Batch = ExtractBatchMeasure(ocrText)
            };

            // Extraer Std L A B del texto crudo (fallback)
            var std = ExtractStdLab(ocrText);
            if (std != null)
            {
                res.StdL = std.Value.L;
                res.StdA = std.Value.A;
                res.StdB = std.Value.B;
            }

            LastResult = res;
            return res;
        }

        /// Extrae el DT Main del texto OCR completo 
        public string ExtractDtMain(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;
            var m = DtMainRegex.Match(ocrText);
            return m.Success ? m.Groups[1].Value.Trim() : null;
        }

        public (string L, string A, string B)? ExtractStdLab(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            // 1. Intento con Regex principal (que busca la palabra "Std")
            var m = StdLabRegex.Match(ocrText);
            if (m.Success)
            {
                return (NormalizeLabValue(m.Groups[1].Value),
                        NormalizeLabValue(m.Groups[2].Value),
                        NormalizeLabValue(m.Groups[3].Value));
            }

            // 2. Fallback: Buscar los primeros 3 números en la parte superior del reporte (donde suele estar el estándar)
            var mNum = Regex.Matches(ocrText, @"[-+]?\d+(?:\s?[.,]\s?\d+)?");
            if (mNum.Count >= 3)
            {
                return (NormalizeLabValue(mNum[0].Value),
                        NormalizeLabValue(mNum[1].Value),
                        NormalizeLabValue(mNum[2].Value));
            }

            return null;
        }

        //-------------------------------------------------------------------
        // Método todo-en-uno desde Bitmap original
        //-------------------------------------------------------------------
        public ShadeExtractionResult ExtractFromBitmap(Bitmap bmp)
        {
            string text = ExtractTextFromBitmap(bmp) ?? string.Empty;

            var result = ExtractAll(text);

            // Intentar mejorar la receta con OCR dirigido sobre el recorte de la zona de ingredientes.
            try
            {
                var shadeName = ExtractShadeNameFromBitmap(bmp);
                if (!string.IsNullOrWhiteSpace(shadeName))
                    result.ShadeName = shadeName;
            }
            catch { }

            // Intentar mejorar DT Main con OCR dirigido sobre el encabezado central del reporte
            try
            {
                var dtMain = ExtractDtMainFromBitmap(bmp);
                if (!string.IsNullOrWhiteSpace(dtMain))
                    result.DtMain = dtMain;
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
                {
                    result.Batch = b2;
                    if (!string.IsNullOrWhiteSpace(b2.LotNo))
                        result.LotNo = b2.LotNo;
                }
            }
            catch { }

            try
            {
                var std2 = ExtractStdLabFromBitmap(bmp);
                if (std2 != null)
                {
                    result.StdL = std2.Value.L;
                    result.StdA = std2.Value.A;
                    result.StdB = std2.Value.B;
                }
            }
            catch { }

            if (result.Batch != null)
            {
                result.LotNo = result.Batch.LotNo;
            }

            LastResult = result;
            return result;
        }

        public List<RecipeItem> ExtractRecipe(string ocrText)
        {
            var list = new List<RecipeItem>();
            // --- PASO 0: ANCLAJE DE BLOQUE (Para evitar capturar datos de la tabla Batch) ---
            string recipeArea = ocrText;
            int startIdx = ocrText.IndexOf("Recipe Version", StringComparison.OrdinalIgnoreCase);
            if (startIdx < 0) startIdx = ocrText.IndexOf("Recipe Number", StringComparison.OrdinalIgnoreCase); // fallback

            int endIdx = ocrText.IndexOf("Dyelots for Recipe", StringComparison.OrdinalIgnoreCase);
            if (endIdx < 0) endIdx = ocrText.IndexOf("Prescreening", StringComparison.OrdinalIgnoreCase); // fallback

            if (startIdx >= 0 && endIdx > startIdx)
            {
                recipeArea = ocrText.Substring(startIdx, endIdx - startIdx);
            }
            else if (startIdx >= 0)
            {
                // Si no hay fin claro, tomamos un bloque razonable después del inicio
                recipeArea = ocrText.Substring(startIdx, Math.Min(ocrText.Length - startIdx, 2000));
            }

            // --- PASO 1: EXTRACCIÓN CRUDA Y DETECCIÓN DE CONTEXTO ---
            var matches = RecipeRegex.Matches(recipeArea);
            var rawResults = new List<(string code, string name, string pctRaw, string digits)>();
            var decimalCounts = new List<int>();

            foreach (Match m in matches)
            {
                var code = m.Groups[1].Value.Trim();
                var name = m.Groups[2].Value.Trim();
                var pctRaw = m.Groups[3].Value.Trim().Replace(" ", "");

                // Limpiar dígitos para detectar precisión real capturada
                string digitsOnly = Regex.Replace(pctRaw, @"[^\d]", "");

                // Detectar cuántos decimales leyó el OCR originalmente (si hay punto)
                int decPos = pctRaw.IndexOf('.');
                if (decPos >= 0)
                {
                    decimalCounts.Add(pctRaw.Length - decPos - 1);
                }

                // --- FILTRO DE DATOS BASURA ---
                if (IsGarbageRecipe(name)) continue;

                rawResults.Add((code, name, pctRaw, digitsOnly));
            }

            // Determinar precisión dominante (Moda)
            int dominantPrecision = 5;
            if (decimalCounts.Count > 0)
            {
                dominantPrecision = decimalCounts.GroupBy(n => n)
                                              .OrderByDescending(g => g.Count())
                                              .First().Key;
            }

            // --- PASO 2: REPARACIÓN INTELIGENTE BASADA EN CONTEXTO ---
            foreach (var item in rawResults)
            {
                string name = item.name;
                string pctRaw = item.pctRaw;
                string digits = item.digits;

                char lastCharName = name.Length > 0 ? name[name.Length - 1] : ' ';
                char firstCharPct = pctRaw.Length > 0 ? pctRaw[0] : ' ';
                bool hasConfusionEvidence = (lastCharName == 'L' || lastCharName == 'I' || lastCharName == '|' ||
                                           firstCharPct == 'L' || firstCharPct == 'I' || firstCharPct == '|');

                // ¿Faltan dígitos respecto al resto de la tabla?
                int missingDigits = (dominantPrecision + 1) - digits.Length;

                if (missingDigits > 0 && hasConfusionEvidence)
                {
                    // Reparación inteligente: solo si hay evidencia de colapso OCR (L -> 1.1)
                    if (missingDigits == 1 && digits.StartsWith("1"))
                    {
                        // Ej: L7826 -> 1.17826 (en contexto de 5 decimales)
                        pctRaw = "1.1" + digits.Substring(1);
                    }
                    else if (missingDigits == 2)
                    {
                        // Ej: L7826 (donde L colapsó "1.1")
                        pctRaw = "1.1" + digits;
                    }

                    // Limpiar el carácter de confusión del nombre si era una letra pegada
                    if (char.IsLetter(lastCharName)) name = name.Substring(0, name.Length - 1).Trim();
                }
                else if (!pctRaw.Contains(".") && digits.Length >= 4)
                {
                    // Inserción de punto automática basada en contexto si el OCR lo omitió
                    if (digits.Length > dominantPrecision)
                    {
                        pctRaw = digits.Insert(digits.Length - dominantPrecision, ".");
                    }
                    else
                    {
                        pctRaw = "0." + digits.PadLeft(dominantPrecision, '0');
                    }
                }

                // Normalización final conservadora
                pctRaw = pctRaw.Replace('O', '0').Replace('o', '0').Replace('S', '5').Replace('Z', '2').Replace(',', '.').Replace("..", ".");

                list.Add(new RecipeItem
                {
                    Code = item.code,
                    Name = name,
                    Percentage = pctRaw + (pctRaw.EndsWith("%") ? "" : "%")
                });
            }

            return list;
        }

        private bool IsGarbageRecipe(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return true;

            string n = name.ToLower();
            // Palabras clave de encabezado o etiquetas de reporte que NO son ingredientes
            if (n.Contains("to:") || n.Contains("de:") || n.Contains("d e:") || n.Contains("period:") ||
                n.Contains("dyehouse") || n.Contains("database") || n.Contains("server") ||
                n.Contains("standard type") || n.Contains("main:") || n.Contains("version") ||
                n.Contains("accuracy") || n.Contains("thread") || n.Contains("lot no"))
                return true;

            // Si el nombre contiene patrones de hilos (ej: "133x2", "G000")
            if (n.Contains("x") && Regex.IsMatch(n, @"\d+x\d+")) return true;
            if (n.Contains("g000")) return true;

            // Si el nombre es muy corto y contiene dos puntos (ej: "To:", "To :")
            if (name.Length <= 4 && name.Contains(":")) return true;

            return false;
        }

        public LabValues ExtractLabValues(string ocrText)
        {
            if (string.IsNullOrWhiteSpace(ocrText)) return null;

            var header = LabHeaderRegex.Match(ocrText);
            if (!header.Success) return null;

            var after = ocrText.Substring(header.Index + header.Length);
            string[] lines = after.Split('\n');

            foreach (string raw in lines)
            {
                var line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Buscamos P o F al final
                var pfMatch = Regex.Match(line, @"\b([FfPp])\b\s*$");
                if (!pfMatch.Success) continue;

                // Extraemos números
                var nums = new List<string>();
                foreach (Match m in Regex.Matches(line, @"-?\d+(?:\.\d+)?"))
                    nums.Add(m.Value.Replace(',', '.'));

                if (nums.Count < 7) continue;

                int start = nums.Count - 7;
                return new LabValues
                {
                    L = nums[start + 0],
                    A = nums[start + 1],
                    B = nums[start + 2],
                    DL = nums[start + 3],
                    DA = nums[start + 4],
                    DB = nums[start + 5],
                    CDE = nums[start + 6],
                    PF = pfMatch.Groups[1].Value.ToUpper()
                };
            }

            return null;
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

                var res = new BatchMeasure
                {
                    L = nums[start + 0].ToString(),
                    A = nums[start + 1].ToString(),
                    B = nums[start + 2].ToString(),
                    DL = nums[start + 3].ToString(),
                    DC = nums[start + 4].ToString(),
                    DH = nums[start + 5].ToString(),
                    DE = nums[start + 6].ToString(),
                    PF = pfMatch.Value.ToUpper()
                };

                // Extraer Lot No (segunda columna aprox o patrón letra+números)
                var lotMatch = Regex.Match(line, @"\b([A-Z]\d{6,9}|\d{6,10})\b");
                if (lotMatch.Success)
                {
                    res.LotNo = lotMatch.Value;
                }
                else
                {
                    var words = line.Split(new[] { ' ', '|', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (words.Length > 1) res.LotNo = words[1];
                }

                return res;
            }

            return null;
        }

        //-------------------------------------------------------------------
        // REEMPLAZO Y NUEVOS (Recorte dirigido de alta fidelidad)
        //-------------------------------------------------------------------
        private string ExtractDtMainFromBitmap(Bitmap original)
        {
            // La línea "DT Main:" está en el encabezado medio, aprox entre 20% y 38% de la altura
            // Ampliado: 15% a 45%
            int top = (int)(original.Height * 0.15);
            int bot = (int)(original.Height * 0.45);
            int h = bot - top;

            if (h <= 0 || top + h > original.Height) return null;

            // Escalar 3x para mejorar la lectura alfanumérica
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
                "dtmain_" + Guid.NewGuid().ToString("N") + ".png");

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

            return ExtractDtMain(ocrText);
        }
        private string ExtractShadeNameFromBitmap(Bitmap original)
        {
            // El Shade Name está en la parte superior izquierda (aprox 0% al 15% de altura y 0 al 60% de ancho)
            // Ampliado: 0% a 20%
            int h = (int)(original.Height * 0.20);
            int w = (int)(original.Width * 0.70);

            if (h <= 0 || w <= 0) return null;

            // Escalar 3x mejora consistentemente las lecturas alfanuméricas complejas 
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

        private (string L, string A, string B)? ExtractStdLabFromBitmap(Bitmap original)
        {
            // Localización de la banda Std L A B (mejorada con pre-procesamiento OpenCV)
            // Ampliado: 18% a 45%
            int top = (int)(original.Height * 0.18);
            int bot = (int)(original.Height * 0.45);
            int h = bot - top;
            if (h <= 0) return null;

            string ocrText = string.Empty;
            try
            {
                using (var mat = BitmapConverter.ToMat(original))
                using (var gray = new Mat())
                {
                    var roiRect = new OpenCvSharp.Rect(0, top, original.Width, h);
                    using (var roi = new Mat(mat, roiRect))
                    {
                        Cv2.CvtColor(roi, gray, ColorConversionCodes.BGR2GRAY);
                        using (var resized = new Mat())
                        {
                            // Escala 4x para capturar puntos decimales pequeños
                            Cv2.Resize(gray, resized, new OpenCvSharp.Size(0, 0), 4.0, 4.0, InterpolationFlags.Cubic);
                            using (var thresholded = new Mat())
                            {
                                // Cambio a AdaptiveThreshold: 
                                Cv2.AdaptiveThreshold(resized, thresholded, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, 11, 2);

                                using (Bitmap bmpToOcr = BitmapConverter.ToBitmap(thresholded))
                                {
                                    using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                                    using (var page = engine.Process(bmpToOcr, PageSegMode.SingleLine))
                                    {
                                        ocrText = (page.GetText() ?? string.Empty).Replace(',', '.');
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            return ExtractStdLab(ocrText);
        }

        private string NormalizeLabValue(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return "0.00";

            // Eliminar espacios y limpiar caracteres no numericos (excepto punto y signo)
            string clean = Regex.Replace(val, @"[^\d.-]", "");

            // Si tiene 4 dígitos sin punto (ej: 3093), asumimos que el punto va antes de los últimos 2 (30.93)
            // Si tiene 3 dígitos sin punto (ej: 860), asumimos 8.60
            if (!clean.Contains(".") && clean.Length >= 3)
            {
                int len = clean.Length;
                clean = clean.Insert(len - 2, ".");
            }

            return clean;
        }

        private List<RecipeItem> ExtractRecipeFromCrop(Bitmap original)
        {
            // MEJORA: Uso de OpenCV para pre-procesamiento avanzado (Estilo Dataextraxtor)
            string ocrText = string.Empty;
            try
            {
                // Ampliado: 18% a 65%
                int top = (int)(original.Height * 0.18);
                int bot = (int)(original.Height * 0.65);

                // Si la imagen es pequeña (< 900px de alto), probablemente ya es un recorte.
                // En este caso, buscamos en el 100% de la imagen.
                if (original.Height < 900) { top = 0; bot = original.Height; }

                int h = bot - top;
                if (h <= 0) return null;

                using (var mat = BitmapConverter.ToMat(original))
                using (var gray = new Mat())
                {
                    var roiRect = new OpenCvSharp.Rect(0, top, original.Width, h);
                    using (var roi = new Mat(mat, roiRect))
                    {
                        Cv2.CvtColor(roi, gray, ColorConversionCodes.BGR2GRAY);

                        using (var resized = new Mat())
                        {
                            // Escala 4x con interpolación cúbica (como Dataextraxtor) para nitidez decimal
                            Cv2.Resize(gray, resized, new OpenCvSharp.Size(0, 0), 4.0, 4.0, InterpolationFlags.Cubic);

                            using (var thresholded = new Mat())
                            {
                                // Binarización de Otsu para limpiar el ruido del fondo
                                Cv2.Threshold(resized, thresholded, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

                                using (Bitmap bmpToOcr = BitmapConverter.ToBitmap(thresholded))
                                {
                                    using (var engine = new TesseractEngine(_tessdataPath, OCR_LANG, EngineMode.Default))
                                    using (var page = engine.Process(bmpToOcr, PageSegMode.SingleBlock))
                                    {
                                        ocrText = page.GetText() ?? string.Empty;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            return ExtractRecipe(ocrText);
        }

        // REEMPLAZO IMPLEMENTADO: Extraer BatchMeasure usando Recorte dirigido
        private BatchMeasure ExtractBatchMeasureFromBitmap(Bitmap original)
        {
            // La fila de Batch Measure está en la parte inferior (aprox 65% a 95% de la altura).
            // Ampliado: 55% a 98%
            int top = (int)(original.Height * 0.55);
            int h = (int)(original.Height * 0.43);

            // Si la imagen es pequeña, usar toda la altura
            if (original.Height < 900) { top = 0; h = original.Height; }

            if (h <= 0 || top + h > original.Height) return null;

            // Recortar y escalar 3x para mejorar calidad OCR numérico 
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

                    var res = new BatchMeasure
                    {
                        L = rawL,
                        A = rawA,
                        B = FixNum(rawNums[start + 2]),
                        DL = FixNum(rawNums[start + 3]),
                        DC = FixNum(rawNums[start + 4]),
                        DH = FixNum(rawNums[start + 5]),
                        DE = FixNum(rawNums[start + 6]),
                        PF = pfMatch.Groups[1].Value.ToUpper()
                    };

                    // Extraer Lot No
                    var lotMatch = Regex.Match(line, @"\b([A-Z]\d{6,9}|\d{6,10})\b");
                    if (lotMatch.Success)
                    {
                        res.LotNo = lotMatch.Value;
                    }
                    else
                    {
                        var words = line.Split(new[] { ' ', '|', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        if (words.Length > 1) res.LotNo = words[1];
                    }

                    return res;
                }
            }

            // Fallback original con el texto procesado por OCR dirigido si fallan los filtros agresivos
            return ExtractBatchMeasure(ocrText);
        }
    }
}