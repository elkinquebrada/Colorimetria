// Dataextraxtor.cs — versión C# 7.3 para OCR de colorimetríca
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Tesseract;

namespace Colorimetria
{

    // MODELOS
    public class ColorimetricRow
    {
        public string Illuminant { get; set; }
        public string Type { get; set; }   // "Std" | "Lot"
        public double L { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double Chroma { get; set; }
        /// Hue en grados (0–360). Puede venir con 0, 1 o 2 decimales.</summary>
        public double Hue { get; set; }

        public int HueInt
        {
            get { return (int)Math.Round(Hue, MidpointRounding.AwayFromZero); }
        }

        /// true → algo fuera de rango o corregido; conviene revisar visualmente.</summary>
        public bool NeedsReview { get; set; }

        public override string ToString()
        {
            return string.Format("{0,-6} {1,-4} L={2,7:F2} a={3,7:F2} b={4,7:F2} C={5,7:F2} h={6,6:F2}",
                Illuminant, Type, L, A, B, Chroma, Hue);
        }
    }

    public class CmcDifferenceRow
    {
        public string Illuminant { get; set; }
        public double DeltaLightness { get; set; }   // ΔL*
        public double DeltaChroma { get; set; }      // ΔC*
        public double DeltaHue { get; set; }         // ΔH*
        public double? DeltaCMC { get; set; }        // Col Diff CMC(2:1) (si viene)

        // Texto crudo (del OCR) solo para diagnóstico
        public string LightnessFlagOcr { get; set; }
        public string ChromaFlagOcr { get; set; }
        public string HueFlagOcr { get; set; }

        // Derivados desde los deltas (no confiamos en el texto OCR para esto)
        private const double Threshold = 0.05;

        public string ChromaFlag
        {
            get
            {
                if (DeltaChroma > Threshold) return "Fuller";
                if (DeltaChroma < -Threshold) return "Duller";
                return "Same Chroma";
            }
        }

        public string LightnessFlag
        {
            get
            {
                if (DeltaLightness > Threshold) return "Brighter";
                if (DeltaLightness < -Threshold) return "Darker";
                return "Same Lightness";
            }
        }

        // Convención: ΔH+ → Yellower/Redder, ΔH− → Bluer/Greener
        public string HueFlag
        {
            get
            {
                if (DeltaHue > Threshold) return "Yellower";
                if (DeltaHue < -Threshold) return "Bluer/Greener";
                return "Same Hue";
            }
        }

        public string ChromaHueFlag
        {
            get { return (LightnessFlag + " " + HueFlag).Trim(); }
        }

        public bool NeedsReview { get; set; }
    }

    public class OcrReport
    {
        public List<ColorimetricRow> Measures { get; set; } = new List<ColorimetricRow>();
        public List<CmcDifferenceRow> CmcDifferences { get; set; } = new List<CmcDifferenceRow>();

        public double TolDL { get; set; }
        public double TolDC { get; set; }
        public double TolDH { get; set; }
        public double TolDE { get; set; }

        public string PrintDate { get; set; }

        public List<string> ParseLog { get; set; } = new List<string>();
    }

    // RANGOS
    internal static class ColorimetryRanges
    {
        public static bool IsValidL(double v) { return v >= 0 && v <= 100; }
        public static bool IsValidAB(double v) { return v >= -200 && v <= 200; }
        public static bool IsValidChroma(double v) { return v >= 0 && v <= 200; }
        public static bool IsValidHue(double v) { return v >= 0 && v <= 360; }
        public static bool IsValidDL(double v) { return v >= -10 && v <= 10; }
        public static bool IsValidDC(double v) { return v >= -10 && v <= 10; }
        public static bool IsValidDH(double v) { return v >= -50 && v <= 50; }
        public static bool IsValidDE(double v) { return v >= 0 && v <= 10; }
    }

    // EXTRACTOR PRINCIPAL
    public class ColorimetricDataExtractor
    {
        private readonly string _tessDataPath;
        private const int SCALE_FACTOR = 3;
        private const bool ENFORCE_ONE_PER_ILLUMINANT_TYPE = true;

        // (Opcional) autocorrección por coherencia (1 dígito) en C* y h°
        private const bool ENABLE_COHERENCE_FIX = true;

        public static readonly HashSet<string> KnownIlluminants =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "D65", "D50", "D55", "D75",
                "TL84", "TL83", "TL85",
                "CWF", "CWF2",
                "A",
                "F2", "F7", "F11", "F12",
                "UV"
            };

        private static readonly string IlluminantPattern =
            @"\b(" + string.Join("|",
                KnownIlluminants
                    .OrderByDescending(s => s.Length)
                    .Select(Regex.Escape)) + @")\b";

        /// Vacío = todos; si agregas elementos, se filtra por estos.
        public HashSet<string> AllowedIlluminants { get; } =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public ColorimetricDataExtractor(string tessDataPath = @".\tessdata")
        {
            _tessDataPath = tessDataPath;
        }

        // ── APIs públicas ──────────────────────────────────────
        public List<ColorimetricRow> ExtractFromFile(string imagePath)
        {
            using (var bmp = new Bitmap(imagePath))
            {
                return ExtractFromBitmap(bmp);
            }
        }

        public List<ColorimetricRow> ExtractFromBitmap(Bitmap original)
        {
            using (var processed = Preprocess(original))
            {
                return ParseColorimetricData(RunOCR(processed));
            }
        }

        public OcrReport ExtractReportFromFile(string imagePath)
        {
            using (var bmp = new Bitmap(imagePath))
            {
                return ExtractReportFromBitmap(bmp);
            }
        }

        public OcrReport ExtractReportFromBitmap(Bitmap original)
        {
            using (var processed = Preprocess(original))
            {
                string text = RunOCR(processed);
                // Diagnóstico (descomentar si quieres ver el OCR crudo)
                return ParseFullReport(text);
            }
        }

        // ── Preprocesamiento ───────────────────────────────────
        private Bitmap Preprocess(Bitmap src)
        {
            int nW = src.Width * SCALE_FACTOR, nH = src.Height * SCALE_FACTOR;
            var scaled = new Bitmap(nW, nH, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.DrawImage(src, 0, 0, nW, nH);
            }
            var gray = ToGrayscaleContrast(scaled, 1.8f, -0.1f);
            scaled.Dispose();
            return gray;
        }

        private Bitmap ToGrayscaleContrast(Bitmap src, float contrast, float brightness)
        {
            var result = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
            float c = contrast, b = brightness;
            float[][] m =
            {
                new[] { 0.299f*c, 0.299f*c, 0.299f*c, 0, 0 },
                new[] { 0.587f*c, 0.587f*c, 0.587f*c, 0, 0 },
                new[] { 0.114f*c, 0.114f*c, 0.114f*c, 0, 0 },
                new[] { 0f, 0f, 0f, 1, 0 },
                new[] { b, b, b, 0, 1 }
            };
            using (var g = Graphics.FromImage(result))
            using (var ia = new ImageAttributes())
            {
                ia.SetColorMatrix(new ColorMatrix(m));
                g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height),
                            0, 0, src.Width, src.Height, GraphicsUnit.Pixel, ia);
            }
            return result;
        }

        // OCR  (C# 7.3) — FIX: usa System.Drawing.Imaging.ImageFormat.Png
        private string RunOCR(Bitmap bmp)
        {
            string tmp = Path.Combine(Path.GetTempPath(), string.Format("ocr_{0:N}.png", Guid.NewGuid()));
            try
            {
                // 👇 FIX CS0104: 'ImageFormat' ambiguo (System.Drawing.Imaging vs Tesseract)
                bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);

                using (var engine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default))
                using (var img = Pix.LoadFromFile(tmp))
                using (var page = engine.Process(img, PageSegMode.SingleBlock))
                {
                    return page.GetText();
                }
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
        }

        // ── Parse general ─────────────────────────────────────
        private OcrReport ParseFullReport(string ocrText)
        {
            var report = new OcrReport();

            ParseCombinedTable(ocrText, report);

            var t = ParseTolerances(ocrText);
            report.TolDL = t.Item1;
            report.TolDC = t.Item2;
            report.TolDH = t.Item3;
            report.TolDE = t.Item4;

            report.PrintDate = ParsePrintDate(ocrText);

            return report;
        }

        // PARSER PRINCIPAL — tabla combinada
        private void ParseCombinedTable(string ocrText, OcrReport report)
        {
            var rawLines = ocrText.Split('\n');
            var normLines = new List<string>(rawLines.Length);
            foreach (var ln in rawLines) normLines.Add(NormalizeOCRLine(ln));

            int start = FindSectionStart(normLines);

            // 1) localizar inicio de cada bloque por iluminante
            var illuminantStarts = new List<Tuple<int, string>>();
            for (int i = start; i < normLines.Count; i++)
            {
                if (IsSectionEnd(normLines[i])) break;
                var m = Regex.Match(normLines[i], IlluminantPattern, RegexOptions.IgnoreCase);
                if (m.Success)
                {
                    string ill = m.Groups[1].Value.ToUpperInvariant();
                    if (!ShouldSkipIlluminant(ill))
                        illuminantStarts.Add(Tuple.Create(i, ill));
                }
            }
            if (illuminantStarts.Count == 0) return;

            // 2) procesar cada bloque
            for (int b = 0; b < illuminantStarts.Count; b++)
            {
                int lineStart = illuminantStarts[b].Item1;
                string illuminant = illuminantStarts[b].Item2;
                int lineEnd = (b + 1 < illuminantStarts.Count) ? illuminantStarts[b + 1].Item1 : normLines.Count;

                var sbNorm = new StringBuilder();
                var sbRaw = new StringBuilder();
                for (int i = lineStart; i < lineEnd && i < normLines.Count; i++)
                {
                    if (IsSectionEnd(normLines[i])) break;
                    sbNorm.Append(' ').Append(normLines[i]);
                    sbRaw.Append(' ').Append(i < rawLines.Length ? rawLines[i] : normLines[i]);
                }
                string bNorm = sbNorm.ToString().Trim();
                string bRaw = sbRaw.ToString().Trim();

                bool hasStd = Regex.IsMatch(bNorm, @"\b(STD|5TD)\b");
                bool hasLot = Regex.IsMatch(bNorm, @"\bLOT\b");

                if (hasStd && hasLot)
                {
                    int lotN = FindWordIndex(bNorm, "LOT");
                    int lotR = FindWordIndex(bRaw, "LOT");

                    string stdR = lotR >= 0 ? bRaw.Substring(0, lotR) : bRaw;
                    string lotR2 = lotR >= 0 ? bRaw.Substring(lotR) : "";

                    var stdRow = ParseMeasureLine(stdR, illuminant, "Std", report.ParseLog);
                    if (stdRow != null) report.Measures.Add(stdRow);

                    var lotRow = ParseMeasureLine(lotR2, illuminant, "Lot", report.ParseLog);
                    if (lotRow != null) report.Measures.Add(lotRow);

                    string lotTailNorm = "";
                    if (lotN >= 0 && lotN <= bNorm.Length) lotTailNorm = bNorm.Substring(lotN);
                    var cmc = ParseCmcFromStdPart(stdR, lotTailNorm, illuminant, report.ParseLog);
                    if (cmc != null) report.CmcDifferences.Add(cmc);
                }
                else if (hasStd)
                {
                    var stdRow = ParseMeasureLine(bRaw, illuminant, "Std", report.ParseLog);
                    if (stdRow != null) report.Measures.Add(stdRow);

                    var cmc = ParseCmcFromStdPart(bRaw, "", illuminant, report.ParseLog);
                    if (cmc != null) report.CmcDifferences.Add(cmc);
                }
                else if (hasLot)
                {
                    var lotRow = ParseMeasureLine(bRaw, illuminant, "Lot", report.ParseLog);
                    if (lotRow != null) report.Measures.Add(lotRow);

                    CmcDifferenceRow pending = null;
                    for (int k = report.CmcDifferences.Count - 1; k >= 0; k--)
                    {
                        if (string.Equals(report.CmcDifferences[k].Illuminant, illuminant, StringComparison.OrdinalIgnoreCase))
                        {
                            pending = report.CmcDifferences[k];
                            break;
                        }
                    }
                    if (pending != null) ExtractFlagsOcr(bNorm, pending);
                }
            }

            report.Measures = DedupAndSort(report.Measures);
            report.CmcDifferences = DedupCmc(report.CmcDifferences);
        }

        private static bool IsSectionEnd(string normLine)
        {
            return normLine.Contains("TOLERANCES") || normLine.Contains("PRINT DATE");
        }

        private static int FindWordIndex(string text, string word)
        {
            var m = Regex.Match(text, @"\b" + Regex.Escape(word) + @"\b", RegexOptions.IgnoreCase);
            return m.Success ? m.Index : -1;
        }

        /// Quita el prefijo "ILUMINANTE + STD/LOT" para evitar que sus dígitos contaminen los tokens.
        private static string StripRowPrefix(string rawPart)
        {
            if (string.IsNullOrWhiteSpace(rawPart)) return rawPart;

            string s = rawPart.Trim();
            s = Regex.Replace(
                s,
                @"^\s*(" + string.Join("|", KnownIlluminants.OrderByDescending(x => x.Length).Select(Regex.Escape))
                + @")\s*(STD|LOT|5TD|Std|Lot)?\s*",
                " ",
                RegexOptions.IgnoreCase);

            return s;
        }

        // PARSEO DE MEDICIÓN: L a b C h
        private static ColorimetricRow ParseMeasureLine(
            string rawPart, string illuminant, string type, List<string> log)
        {
            rawPart = StripRowPrefix(rawPart);
            var tokens = ExtractNumericTokens(rawPart);

            if (tokens.Count < 5) return null;

            int base_ = FindMeasureBase(tokens);
            if (base_ < 0 || base_ + 4 >= tokens.Count) return null;

            double vL = RestoreMeasureDecimal(tokens[base_]);
            double vA = RestoreMeasureDecimal(tokens[base_ + 1]);
            double vB = RestoreMeasureDecimal(tokens[base_ + 2]);
            double vC = RestoreMeasureDecimal(tokens[base_ + 3]);
            double vH = ParseHueDouble(tokens[base_ + 4]);

            // (opcional) autocorrección por coherencia en C* y h°
            bool corrected = false;
            if (ENABLE_COHERENCE_FIX)
            {
                double cHat, hHat;
                FromAB(vA, vB, out cHat, out hHat);

                if (Math.Abs(vC - cHat) > 0.6)
                {
                    bool okFixC; double fixedC;
                    TryFixOneDigit(tokens[base_ + 3], vC, delegate (double vv) { return Math.Abs(vv - cHat); }, out okFixC, out fixedC);
                    if (okFixC) { vC = fixedC; corrected = true; }
                }

                double hueErr = HueError(vH, hHat);
                if (hueErr > 1.2)
                {
                    bool okFixH; double fixedH;
                    TryFixOneDigit(tokens[base_ + 4], vH,
                        delegate (double vv) { return HueError(vv, hHat); }, out okFixH, out fixedH);
                    if (okFixH) { vH = fixedH; corrected = true; }
                }
            }

            bool ok = true;
            Action<string> Warn = msg =>
            {
                if (log != null) log.Add(string.Format("[WARN] {0}/{1} {2} -> Revisar OCR", illuminant, type, msg));
                ok = false;
            };

            if (!ColorimetryRanges.IsValidL(vL)) Warn(string.Format("L={0} fuera de [0,100]", vL));
            if (!ColorimetryRanges.IsValidAB(vA)) Warn(string.Format("a={0} fuera de [-200,200]", vA));
            if (!ColorimetryRanges.IsValidAB(vB)) Warn(string.Format("b={0} fuera de [-200,200]", vB));
            if (!ColorimetryRanges.IsValidChroma(vC)) Warn(string.Format("Chroma={0} fuera de [0,200]", vC));
            if (!ColorimetryRanges.IsValidHue(vH)) Warn(string.Format("Hue={0} fuera de [0,360]", vH));

            // Validación cruzada: Chroma debe ser aproximadamente sqrt(a² + b²).
            if (ok)
            {
                double chromaCalc = Math.Sqrt(vA * vA + vB * vB);
                double chromaErr = Math.Abs(chromaCalc - vC);
                // Tolerancia: mayor entre 5% de Chroma y 1.5 unidades
                double chromaTol = Math.Max(vC * 0.05, 1.5);
                if (chromaErr > chromaTol)
                {
                    if (log != null)
                        log.Add(string.Format(
                            "[WARN] {0}/{1} Chroma={2:F2} pero sqrt(a²+b²)={3:F2} (err={4:F2}) -> digito OCR incorrecto en a*, b* o Chroma",
                            illuminant, type, vC, chromaCalc, chromaErr));
                    ok = false;
                }
            }

            return new ColorimetricRow
            {
                Illuminant = illuminant,
                Type = type,
                L = vL,
                A = vA,
                B = vB,
                Chroma = vC,
                Hue = vH,
                NeedsReview = (!ok || corrected)
            };
        }

        /// Localiza el índice base del quintuplo [L, a*, b*, Chroma, Hue] en la lista de tokens.
        private static int FindMeasureBase(List<string> tokens)
        {
            for (int i = 0; i <= tokens.Count - 5; i++)
            {
                string tokL = tokens[i];
                string tokA = tokens[i + 1];
                string tokB = tokens[i + 2];
                string tokC = tokens[i + 3];
                string tokH = tokens[i + 4];

                double l = RestoreMeasureDecimal(tokL);
                double a = RestoreMeasureDecimal(tokA);
                double b = RestoreMeasureDecimal(tokB);
                double c = RestoreMeasureDecimal(tokC);
                double h;
                if (!double.TryParse(tokH.Replace(',', '.'),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out h))
                    continue;

                if (!ColorimetryRanges.IsValidL(l)) continue;
                if (!ColorimetryRanges.IsValidAB(a)) continue;
                if (!ColorimetryRanges.IsValidAB(b)) continue;
                if (!ColorimetryRanges.IsValidChroma(c)) continue;
                if (!ColorimetryRanges.IsValidHue(h)) continue;

                // Hue: rechazar tokens con 2+ decimales (son mediciones, no angulo Hue).
                if (!LooksLikeHue(tokH)) continue;

                // L* real siempre tiene decimales (43.70). Evita "65"/"84" del nombre del iluminante.
                if (!tokL.Contains(".") && tokL.TrimStart('-').Length <= 2) continue;

                // Evita aceptar un grupo desplazado donde Chroma es absurdamente diferente.
                double chromaCalc = Math.Sqrt(a * a + b * b);
                if (c > 0.001 && chromaCalc > 0.001)
                {
                    double ratio = c > chromaCalc ? c / chromaCalc : chromaCalc / c;
                    if (ratio > 3.0) continue;
                }

                return i;
            }
            return 0; // fallback prudente
        }

        private static bool LooksLikeHue(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return false;
            string s = token.Trim().Replace(',', '.');
            double v;
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                return false;
            if (v < 0 || v > 360) return false;

            int dotIdx = s.IndexOf('.');
            if (dotIdx >= 0)
            {
                // Con punto: solo aceptar exactamente 1 decimal (ej: "259.0", "241.6") // Rechazar "17.81", "27.15" (2 decimales → medición, no Hue)
                string decPart = s.Substring(dotIdx + 1);
                if (decPart.Length > 1) return false;
            }
            return true;
        }

        // Inserta punto decimal en mediciones si el OCR lo omitió (XX.XX)
        private static double RestoreMeasureDecimal(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            token = token.Trim().Replace(',', '.');
            if (token.Contains(".")) return SafeParse(token);

            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;
            if (!Regex.IsMatch(d, @"^\d+$")) return SafeParse(token);
            if (d.Length <= 2) return SafeParse(token);

            return SafeParse((neg ? "-" : "") + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2));
        }

        // Deltas
        private static double RestoreDeltaDecimal(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            token = token.Trim().Replace(',', '.');
            if (token.Contains(".")) return SafeParse(token);

            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;
            if (!Regex.IsMatch(d, @"^\d+$")) return SafeParse(token);

            string sign = neg ? "-" : "";
            if (d.Length == 1) return SafeParse(token);
            if (d.Length == 2) return SafeParse(sign + d[0] + "." + d[1]);
            return SafeParse(sign + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2));
        }

        private static double ParseHueDouble(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            double v = SafeParse(token.Trim());
            if (v < 0) v = 0;
            if (v >= 360) v %= 360;
            return v;
        }

        private static double SafeParse(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Trim().Replace(',', '.');
            double v;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v) ? v : 0;
        }

        // Corrección de signos vía Hue
        private static void CorrectSignsByHue(ref double a, ref double b, double hue)
        {
            const double MinCertainty = 0.05;
            double rad = hue * Math.PI / 180.0;
            double cosH = Math.Cos(rad);
            double sinH = Math.Sin(rad);

            if (Math.Abs(cosH) > MinCertainty)
            {
                bool shouldNeg = cosH < 0;
                if (shouldNeg && a > 0) a = -a;
                else if (!shouldNeg && a < 0) a = -a;
            }
            if (Math.Abs(sinH) > MinCertainty)
            {
                bool shouldNeg = sinH < 0;
                if (shouldNeg && b > 0) b = -b;
                else if (!shouldNeg && b < 0) b = -b;
            }
        }

        private static void FromAB(double a, double b, out double C, out double H)
        {
            C = Math.Sqrt(a * a + b * b);
            H = Math.Atan2(b, a) * 180.0 / Math.PI;
            if (H < 0) H += 360.0;
        }

        private static double HueError(double h, double refH)
        {
            double e = Math.Abs(h - refH);
            return e > 180 ? 360 - e : e;
        }

        // Autocorrección de un dígito (8↔3/5/6, 6↔0, 1↔7, etc.)
        private static void TryFixOneDigit(string originalText, double current, Func<double, double> scorer, out bool ok, out double fixedVal)
        {
            string s = (originalText ?? "").Trim();
            var set = new HashSet<string>();
            set.Add(s);

            var map = new Dictionary<char, char[]>();
            map['0'] = new[] { '9', '8' };
            map['1'] = new[] { '7' };
            map['3'] = new[] { '8' };
            map['5'] = new[] { '6' };
            map['6'] = new[] { '8', '0' };
            map['7'] = new[] { '1' };
            map['8'] = new[] { '3', '6', '0' };
            map['9'] = new[] { '0' };

            for (int i = 0; i < s.Length; i++)
            {
                char[] alts;
                if (map.TryGetValue(s[i], out alts))
                {
                    for (int k = 0; k < alts.Length; k++)
                    {
                        string cand = s.Substring(0, i) + alts[k] + s.Substring(i + 1);
                        set.Add(cand);
                    }
                }
            }

            // Insertar punto si no existe (ej.: 2486 -> 24.86)
            if (s.IndexOf('.') < 0 && Regex.IsMatch(s, @"^-?\d{3,}$"))
            {
                int at = s.StartsWith("-") ? s.Length - 2 : s.Length - 2;
                if (at > 0 && at < s.Length)
                {
                    string candDot = s.Substring(0, at) + "." + s.Substring(at);
                    set.Add(candDot);
                }
            }

            double best = double.PositiveInfinity;
            double bestVal = current;

            foreach (var cand in set)
            {
                double val;
                if (!double.TryParse(cand.Replace(',', '.'),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out val))
                    continue;

                double sc = scorer(val);
                if (sc < best) { best = sc; bestVal = val; }
            }

            ok = best < 0.9;
            fixedVal = bestVal;
        }

        // DELTAS CMC — búsqueda por rango
        private static CmcDifferenceRow ParseCmcFromStdPart(
            string stdRaw, string lotNorm, string illuminant, List<string> log)
        {
            stdRaw = StripRowPrefix(stdRaw);
            var tokens = ExtractNumericTokens(stdRaw);
            if (tokens.Count < 6) return null;

            int base_ = FindMeasureBase(tokens);
            if (base_ < 0) return null;

            int deltaStart = base_ + 5;
            if (deltaStart >= tokens.Count) return null;

            var deltaTokens = tokens.Skip(deltaStart).ToList();
            var cmc = new CmcDifferenceRow { Illuminant = illuminant };
            bool needsReview = false;

            int di = FindDeltaBase(deltaTokens, log, illuminant);
            if (di >= 0)
            {
                double dL = di < deltaTokens.Count ? RestoreDeltaDecimal(deltaTokens[di]) : 0;
                double dC = di + 1 < deltaTokens.Count ? RestoreDeltaDecimal(deltaTokens[di + 1]) : 0;
                double dH = di + 2 < deltaTokens.Count ? RestoreDeltaDecimal(deltaTokens[di + 2]) : 0;
                double dE = di + 3 < deltaTokens.Count ? RestoreDeltaDecimal(deltaTokens[di + 3]) : double.NaN;

                Action<string, double, string> WarnDelta = (name, v, range) =>
                {
                    if (log != null) log.Add(string.Format("[WARN] {0} {1}={2} fuera de {3} -> Revisar OCR", illuminant, name, v, range));
                    needsReview = true;
                };

                if (!ColorimetryRanges.IsValidDL(dL)) WarnDelta("ΔL", dL, "[-10,10]");
                if (!ColorimetryRanges.IsValidDC(dC)) WarnDelta("ΔC", dC, "[-10,10]");
                if (!ColorimetryRanges.IsValidDH(dH)) WarnDelta("ΔH", dH, "[-50,50]");
                if (!double.IsNaN(dE) && !ColorimetryRanges.IsValidDE(dE)) WarnDelta("ΔE CMC", dE, "[0,10]");

                cmc.DeltaLightness = dL;
                cmc.DeltaChroma = dC;
                cmc.DeltaHue = dH;
                cmc.DeltaCMC = double.IsNaN(dE) ? (double?)null : dE;
            }
            else if (deltaTokens.Count >= 3)
            {
                cmc.DeltaLightness = RestoreDeltaDecimal(deltaTokens[0]);
                cmc.DeltaChroma = RestoreDeltaDecimal(deltaTokens[1]);
                cmc.DeltaHue = RestoreDeltaDecimal(deltaTokens[2]);
                cmc.DeltaCMC = deltaTokens.Count >= 4 ? RestoreDeltaDecimal(deltaTokens[3]) : (double?)null;
                if (log != null) log.Add(string.Format("[INFO] {0}: deltas por fallback (sin cuarteto válido)", illuminant));
            }

            cmc.NeedsReview = needsReview;

            if (!string.IsNullOrWhiteSpace(lotNorm))
                ExtractFlagsOcr(lotNorm, cmc);

            return cmc;
        }

        private static int FindDeltaBase(List<string> dt, List<string> log, string ill)
        {
            // Intentar cuarteto completo
            for (int i = 0; i <= dt.Count - 4; i++)
            {
                double dL = RestoreDeltaDecimal(dt[i]);
                double dC = RestoreDeltaDecimal(dt[i + 1]);
                double dH = RestoreDeltaDecimal(dt[i + 2]);
                double dE = RestoreDeltaDecimal(dt[i + 3]);

                if (ColorimetryRanges.IsValidDL(dL) &&
                    ColorimetryRanges.IsValidDC(dC) &&
                    ColorimetryRanges.IsValidDH(dH) &&
                    ColorimetryRanges.IsValidDE(dE))
                    return i;
            }

            // Intentar trío (sin CMC)
            for (int i = 0; i <= dt.Count - 3; i++)
            {
                double dL = RestoreDeltaDecimal(dt[i]);
                double dC = RestoreDeltaDecimal(dt[i + 1]);
                double dH = RestoreDeltaDecimal(dt[i + 2]);

                if (ColorimetryRanges.IsValidDL(dL) &&
                    ColorimetryRanges.IsValidDC(dC) &&
                    ColorimetryRanges.IsValidDH(dH))
                {
                    if (log != null) log.Add(string.Format("[INFO] {0}: CMC(2:1) no encontrado en rango válido.", ill));
                    return i;
                }
            }
            return -1;
        }

        // EXTRACCIÓN DE TOKENS NUMÉRICOS
        private static List<string> ExtractNumericTokens(string rawLine)
        {
            var result = new List<string>();

            string line = (rawLine ?? string.Empty)
                .Replace('\u2212', '-')   // signo menos unicode → '-'
                .Replace(',', '.');       // coma decimal → punto

            // 1) Enmascarar números dentro del nombre del iluminante (D65, TL84, F11, ...)
            line = MaskIlluminantNumbers(line);

            // 2) Eliminar etiquetas STD/LOT
            line = Regex.Replace(line, @"\b(STD|5TD|LOT)\b", " ", RegexOptions.IgnoreCase);

            // 3) Única corrección textual segura: "CW/F" → "CWF"
            line = Regex.Replace(line, @"\bCW/F\b", "CWF", RegexOptions.IgnoreCase);

            // 4) Extraer tokens numéricos (enteros y decimales con signo)
            var matches = Regex.Matches(line, @"-?\d+(?:\.\d+)?");
            var rawTokens = new List<string>(matches.Count);
            foreach (Match m in matches) rawTokens.Add(m.Value);

            // 5) Dividir tokens largos (>6 dígitos) que el OCR pegó
            foreach (var tok in rawTokens)
            {
                bool neg = tok.StartsWith("-");
                string digs = neg ? tok.Substring(1) : tok;

                if (tok.IndexOf('.') >= 0 || digs.Length <= 6)
                {
                    result.Add(tok);
                    continue;
                }

                // Heurística de corte en bloques de 4
                var segments = new List<string>();
                string rem = digs;
                bool firstNeg = neg;
                while (rem.Length >= 4)
                {
                    string seg = rem.Substring(0, 4);
                    segments.Add(firstNeg ? "-" + seg : seg);
                    rem = rem.Substring(4);
                    firstNeg = false;
                }
                if (rem.Length > 0) segments.Add(rem);

                bool allValid = true;
                foreach (var s in segments)
                {
                    double v = RestoreMeasureDecimal(s);
                    if (!(v >= -200 && v <= 200)) { allValid = false; break; }
                }

                if (allValid && segments.Count > 1) result.AddRange(segments);
                else result.Add(tok);
            }

            return result;
        }

        // Enmascara los números dentro del nombre del iluminante para que no se confundan con L*, a*, b*
        private static string MaskIlluminantNumbers(string line)
        {
            line = Regex.Replace(line, @"\bD65\b", "DXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bD50\b", "DXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bD55\b", "DXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bD75\b", "DXX", RegexOptions.IgnoreCase);

            line = Regex.Replace(line, @"\bTL84\b", "TLXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bTL83\b", "TLXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bTL85\b", "TLXX", RegexOptions.IgnoreCase);

            line = Regex.Replace(line, @"\bCWF2\b", "CWFX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bF11\b", "FXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bF12\b", "FXX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bF2\b", "FX", RegexOptions.IgnoreCase);
            line = Regex.Replace(line, @"\bF7\b", "FX", RegexOptions.IgnoreCase);
            // A, CWF, UV no tienen dígitos → no enmascarar
            return line;
        }

        // LOCALIZAR INICIO DE SECCIÓN DE DATOS
        private static int FindSectionStart(List<string> normLines)
        {
            for (int i = 0; i < normLines.Count; i++)
            {
                string win = normLines[i] + " " + (i + 1 < normLines.Count ? normLines[i + 1] : "");
                bool hasCmc = Regex.IsMatch(win, @"\bCMC\b");
                bool hasDiff = Regex.IsMatch(win, @"\bDIFF(ERENCE)?\b") || Regex.IsMatch(win, @"\bILLUMINANT\b");
                if (hasCmc && hasDiff) return i + 1;
            }
            // Fallback
            for (int i = 0; i < normLines.Count; i++)
            {
                if (Regex.IsMatch(normLines[i], IlluminantPattern, RegexOptions.IgnoreCase) &&
                    Regex.IsMatch(normLines[i], @"\b(STD|5TD|LOT)\b"))
                    return i;
            }
            return 0;
        }

        // FLAGS (texto) — solo diagnóstico
        private static void ExtractFlagsOcr(string normLine, CmcDifferenceRow row)
        {
            if (row == null) return;

            // Según el formato del reporte: (Lightness) lleva el texto de Croma,// (Chroma) el de Lightness, (Hue) el de Hue.
            if (Regex.IsMatch(normLine, @"\bFULLER\b")) row.LightnessFlagOcr = "Fuller";
            else if (Regex.IsMatch(normLine, @"\bFULL\b")) row.LightnessFlagOcr = "Full";
            else if (Regex.IsMatch(normLine, @"\bDULLER\b")) row.LightnessFlagOcr = "Duller";
            else if (Regex.IsMatch(normLine, @"\bSAME\b")) row.LightnessFlagOcr = "Same";

            if (Regex.IsMatch(normLine, @"\bBRIGHTER\b")) row.ChromaFlagOcr = "Brighter";
            else if (Regex.IsMatch(normLine, @"\bDARKER\b")) row.ChromaFlagOcr = "Darker";

            if (Regex.IsMatch(normLine, @"\bYELLOWER\b")) row.HueFlagOcr = "Yellower";
            else if (Regex.IsMatch(normLine, @"\bBLUER\b")) row.HueFlagOcr = "Bluer/Greener";
            else if (Regex.IsMatch(normLine, @"\bGREENER\b")) row.HueFlagOcr = "Greener";
            else if (Regex.IsMatch(normLine, @"\bREDDER\b")) row.HueFlagOcr = "Redder";
        }

        // API COMPATIBLE
        public List<ColorimetricRow> ParseColorimetricData(string ocrText)
        {
            var report = new OcrReport();
            ParseCombinedTable(ocrText, report);
            return report.Measures;
        }

        // TOLERANCIAS
        private Tuple<double, double, double, double> ParseTolerances(string ocrText)
        {
            string t = NormalizeOCRLine(ocrText);
            // Busca "TOLERANCES ... DL ... DC ... DH ... DE ..."
            var m = Regex.Match(
                t,
                @"TOLERANCES.*?DL\s*(-?\d+(?:\.\d+)?).*?DC\s*(-?\d+(?:\.\d+)?).*?DH\s*(-?\d+(?:\.\d+)?).*?DE\s*(\d+(?:\.\d+)?)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            return m.Success
                ? Tuple.Create(SafeParse(m.Groups[1].Value), SafeParse(m.Groups[2].Value),
                               SafeParse(m.Groups[3].Value), SafeParse(m.Groups[4].Value))
                : Tuple.Create(0d, 0d, 0d, 0d);
        }

        // PRINT DATE
        private string ParsePrintDate(string ocrText)
        {
            string[] lines = ocrText.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                var line = NormalizeOCRLine(lines[i]);
                if (line.Contains("PRINT") && line.Contains("DATE")) return line;
            }
            return null;
        }

        // DEDUP / ORDEN
        private List<ColorimetricRow> DedupAndSort(List<ColorimetricRow> rows)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var deduped = new List<ColorimetricRow>();

            foreach (var r in rows)
            {
                string key = string.Format("{0}|{1}|{2:F2}|{3:F2}|{4:F2}|{5:F2}|{6:F2}",
                    r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue);
                if (seen.Add(key)) deduped.Add(r);
            }

            if (ENFORCE_ONE_PER_ILLUMINANT_TYPE)
            {
                var seenIT = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var final = new List<ColorimetricRow>();
                foreach (var r in deduped)
                {
                    string key = r.Illuminant + "|" + r.Type;
                    if (seenIT.Add(key)) final.Add(r);
                }
                deduped = final;
            }

            deduped.Sort((x, y) =>
            {
                int ox = GetIllumOrder(x.Illuminant);
                int oy = GetIllumOrder(y.Illuminant);
                if (ox != oy) return ox.CompareTo(oy);
                int tx = x.Type == "Std" ? 0 : 1;
                int ty = y.Type == "Std" ? 0 : 1;
                return tx.CompareTo(ty);
            });

            return deduped;
        }

        private static int GetIllumOrder(string illuminant)
        {
            string ill = (illuminant ?? "").ToUpperInvariant();
            switch (ill)
            {
                case "D65": return 0;
                case "TL84": return 1;
                case "TL83": return 2;
                case "TL85": return 3;
                case "CWF": return 4;
                case "A": return 5;
                default: return 6;
            }
        }
        private static List<CmcDifferenceRow> DedupCmc(List<CmcDifferenceRow> rows)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<CmcDifferenceRow>();
            foreach (var r in rows)
                if (seen.Add(r.Illuminant)) result.Add(r);
            return result;
        }

        // HELPERS
        private bool ShouldSkipIlluminant(string illum)
        {
            if (!KnownIlluminants.Contains(illum)) return true;
            return AllowedIlluminants != null && AllowedIlluminants.Count > 0 && !AllowedIlluminants.Contains(illum);
        }

        /// Normaliza línea OCR
        private static string NormalizeOCRLine(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            string t = s
                .Replace('\u2212', '-')  // menos unicode
                .Replace(',', '.')       // coma → punto
                .Replace('\u2018', '\'')
                .Replace('\u201C', '"')
                .Replace('\u201D', '"')
                .Replace('\u2013', '-')
                .Replace('\u2014', '-');

            t = t.ToUpperInvariant();

            // Correcciones en NOMBRES (no números)
            t = Regex.Replace(t, @"\b5TD\b", "STD");
            t = Regex.Replace(t, @"\bTLS4\b", "TL84");
            t = Regex.Replace(t, @"\bTLS3\b", "TL83");
            t = Regex.Replace(t, @"\bTLS5\b", "TL85");
            t = Regex.Replace(t, @"\bDO5\b", "D65");
            t = Regex.Replace(t, @"\bD6O\b", "D65");
            t = Regex.Replace(t, @"\bDSO\b", "D50");
            t = Regex.Replace(t, @"\bD5O\b", "D50");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*4\b", "TL84");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*3\b", "TL83");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*5\b", "TL85");
            t = Regex.Replace(t, @"\bD\s*6\s*5\b", "D65");
            t = Regex.Replace(t, @"\bD\s*[5S]\s*0\b", "D50");
            t = Regex.Replace(t, @"\bD\s*[5S]\s*5\b", "D55");
            t = Regex.Replace(t, @"\bD\s*7\s*5\b", "D75");
            t = Regex.Replace(t, @"\bC\s*W\s*F\b", "CWF");
            t = Regex.Replace(t, @"\bCW/F\b", "CWF");
            t = Regex.Replace(t, @"\bF\s*1\s*1\b", "F11");
            t = Regex.Replace(t, @"\bF\s*1\s*2\b", "F12");
            t = Regex.Replace(t, @"\[LLUMINANT\b", "ILLUMINANT");
            t = Regex.Replace(t, @"\bL0T\b", "LOT");

            // Limpieza conservadora (letras, dígitos, punto, guión, espacio)
            t = Regex.Replace(t, @"[^A-Z0-9\.\-\s]", " ");
            t = Regex.Replace(t, @"\s+", " ").Trim();
            return t;
        }
    }
}