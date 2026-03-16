// Dataextraxtor.cs — versión C# 7.3 para OCR de colorimetríca
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Http;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Tesseract;

namespace Color
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

    // Región de celda detectada en imagen original
    internal class CellRegion
    {
        public Rectangle Bounds { get; set; }
        public string ColumnHint { get; set; } // "a","b","C","h"
    }

    // EXTRACTOR PRINCIPAL
    public class ColorimetricDataExtractor : IDisposable
    {
        private readonly string _tessDataPath;
        private const int SCALE_FACTOR = 3;
        private const bool ENFORCE_ONE_PER_ILLUMINANT_TYPE = true;

        // Instancia reutilizable de Tesseract (evita recrearla en cada llamada → mucho más rápido)
        private TesseractEngine _sharedEngine;
        private readonly object _engineLock = new object();

        // Autocorrección por coherencia (1 dígito) en C* y h°
        private const bool ENABLE_COHERENCE_FIX = true;

        // A+C: Re-OCR dirigido cuando sqrt(a²+b²) no coincide con Chroma
        private const bool ENABLE_REOCR = true;
        private const double REOCR_CHROMA_THRESHOLD = 0.35; // error máximo aceptable
        private const int REOCR_SCALE = 6;    // escala extra para celdas individuales
        private const int REOCR_PADDING = 4;  // padding en px (imagen original) al recortar

        private TesseractEngine GetEngine()
        {
            if (_sharedEngine == null)
                _sharedEngine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default);
            return _sharedEngine;
        }

        public void Dispose()
        {
            _sharedEngine?.Dispose();
            _sharedEngine = null;
        }

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
                var report = new OcrReport();
                ParseCombinedTable(RunOCR(processed), report);
                if (ENABLE_REOCR)
                    ReOcrFailedCells(original, report);
                return report.Measures;
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
                var report = ParseFullReport(text);

                // A+C: Segunda pasada — re-OCR dirigido sobre celdas con error Chroma
                if (ENABLE_REOCR)
                    ReOcrFailedCells(original, report);

                // B: Corrección via API local (100% local, sin dependencias externas)
                try
                {
                    var apiClient = new ColorimetriaApiClient("http://localhost:5000");
                    if (apiClient.IsApiAvailable())
                    {
                        var corrections = apiClient.CorrectReport(report);
                        foreach (var c in corrections)
                            report.ParseLog.Add(string.Format(
                                "[API] {0}/{1} {2}: corregido ({3})",
                                c.Illuminant, c.Type, c.Field, c.Reason));
                    }
                    else
                    {
                        report.ParseLog.Add("[API] ColorimetriaAPI no disponible en localhost:5000");
                    }
                }
                catch (Exception ex)
                {
                    report.ParseLog.Add("[API] Error al llamar ColorimetriaAPI: " + ex.Message);
                }// B: Corrección via Claude API (solo tokens erróneos — sin datos industriales)
                /*  if (_claudeCorrector != null && _claudeCorrector.IsEnabled)
                  {
                      var corrections = _claudeCorrector.CorrectReport(report);
                      foreach (var c in corrections)
                          report.ParseLog.Add(string.Format(
                              "[CLAUDE] {0}/{1} {2}: {3:F4} → {4:F4} ({5})",
                              c.Illuminant, c.Type, c.Field,
                              c.CorrectedValue, c.NewCoherenceError, c.Reason));
                  }*/

                return report;
            }
        }

        // ══════════════════════════════════════════════════════════════
        // ESTRATEGIA A+C — Re-OCR dirigido sobre celdas con error Chroma
        // ══════════════════════════════════════════════════════════════

        /// Punto de entrada: recorre las filas ya parseadas, detecta cuáles
        /// tienen error Chroma > umbral y lanza re-OCR solo sobre esas celdas.
        private void ReOcrFailedCells(Bitmap original, OcrReport report)
        {
            if (report == null || report.Measures == null || report.Measures.Count == 0) return;

            // Verificar si hay filas que necesitan corrección antes de lanzar detección
            bool anyNeeds = false;
            foreach (var row in report.Measures)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                if (Math.Abs(chromaCalc - row.Chroma) > REOCR_CHROMA_THRESHOLD)
                { anyNeeds = true; break; }
            }
            if (!anyNeeds) return;

            // Obtener bounding boxes de todos los tokens usando Tesseract
            var wordBoxes = GetWordBoundingBoxes(original);
            if (wordBoxes == null || wordBoxes.Count == 0) return;

            foreach (var row in report.Measures)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                double chromaErr = Math.Abs(chromaCalc - row.Chroma);
                if (chromaErr <= REOCR_CHROMA_THRESHOLD) continue;

                report.ParseLog.Add(string.Format(
                    "[REOCR] {0}/{1} chromaErr={2:F3} → segunda pasada",
                    row.Illuminant, row.Type, chromaErr));

                // A+C: localizar la fila en la imagen usando bounding boxes
                var rowRect = FindRowRectByBoxes(wordBoxes, row.Illuminant, row.Type);
                if (rowRect == null)
                {
                    report.ParseLog.Add(string.Format(
                        "[REOCR] {0}/{1} no se encontró banda en imagen", row.Illuminant, row.Type));
                    continue;
                }

                TryReOcrRow(original, rowRect.Value, row, report.ParseLog);
            }
        }

        // ── Obtener bounding boxes de palabras via Tesseract ──────────────

        private List<Tuple<string, Rectangle>> GetWordBoundingBoxes(Bitmap original)
        {
            var result = new List<Tuple<string, Rectangle>>();
            string tmp = Path.Combine(Path.GetTempPath(),
                string.Format("bbox_{0:N}.png", Guid.NewGuid()));
            try
            {
                // Usar preprocesado estándar para las bboxes
                using (var processed = Preprocess(original))
                {
                    processed.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                    int sw = processed.Width, sh = processed.Height;

                    using (var img = Pix.LoadFromFile(tmp))
                    using (var page = GetEngine().Process(img, PageSegMode.SingleBlock))
                    using (var iter = page.GetIterator())
                    {
                        iter.Begin();
                        do
                        {
                            string word = iter.GetText(PageIteratorLevel.Word);
                            if (string.IsNullOrWhiteSpace(word)) continue;

                            Tesseract.Rect bbox;
                            if (!iter.TryGetBoundingBox(PageIteratorLevel.Word, out bbox)) continue;

                            // Convertir coordenadas de imagen procesada → imagen original
                            float scaleX = (float)original.Width / sw;
                            float scaleY = (float)original.Height / sh;
                            var origRect = new Rectangle(
                                (int)(bbox.X1 * scaleX),
                                (int)(bbox.Y1 * scaleY),
                                (int)((bbox.X2 - bbox.X1) * scaleX),
                                (int)((bbox.Y2 - bbox.Y1) * scaleY));

                            result.Add(Tuple.Create(word.Trim(), origRect));
                        }
                        while (iter.Next(PageIteratorLevel.Word));
                    }
                }
            }
            catch (Exception ex)
            {
                // Si falla la iteración de bboxes, continuar sin re-OCR
                System.Diagnostics.Debug.WriteLine("[REOCR] GetWordBoundingBoxes: " + ex.Message);
            }
            finally { if (File.Exists(tmp)) File.Delete(tmp); }

            return result;
        }

        // ── Localizar fila por bounding boxes ────────────────────────────

        /// Busca el bounding box de la palabra del iluminante en la fila correcta (Std/Lot).
        /// Devuelve el rectángulo completo de la fila en coordenadas de la imagen original.
        private Rectangle? FindRowRectByBoxes(
            List<Tuple<string, Rectangle>> boxes, string illuminant, string type)
        {
            string illUpper = illuminant.ToUpperInvariant();
            string typeUpper = type.ToUpperInvariant();

            // Buscar todas las ocurrencias del iluminante
            var illBoxes = new List<Rectangle>();
            foreach (var box in boxes)
            {
                string w = NormalizeOCRLine(box.Item1);
                if (w.Contains(illUpper)) illBoxes.Add(box.Item2);
            }
            if (illBoxes.Count == 0) return null;

            // Para cada ocurrencia del iluminante, buscar si en la misma línea (±Y) hay STD o LOT
            foreach (var illRect in illBoxes)
            {
                int rowY = illRect.Y;
                int rowH = illRect.Height;
                int rowTol = Math.Max(rowH, 12); // tolerancia vertical

                bool foundType = false;
                int minX = illRect.X, maxX = illRect.Right;
                int minY = rowY - rowTol, maxY = rowY + rowH + rowTol;

                foreach (var box in boxes)
                {
                    string w = NormalizeOCRLine(box.Item1);
                    bool sameRow = box.Item2.Y >= minY && box.Item2.Y <= maxY;
                    if (!sameRow) continue;

                    if (w.Contains(typeUpper) ||
                        (type == "Std" && (w.Contains("STD") || w.Contains("5TD"))) ||
                        (type == "Lot" && w.Contains("LOT")))
                    {
                        foundType = true;
                    }
                    // Expandir el rectángulo horizontal con todos los tokens de esta línea
                    if (box.Item2.X < minX) minX = box.Item2.X;
                    if (box.Item2.Right > maxX) maxX = box.Item2.Right;
                }

                if (foundType)
                {
                    // Construir rectángulo completo de la fila con margen
                    int margin = 3;
                    return new Rectangle(
                        Math.Max(0, minX - margin),
                        Math.Max(0, rowY - margin),
                        maxX - minX + margin * 2,
                        rowH + margin * 2);
                }
            }
            return null;
        }

        /// Re-OCR de todos los campos numéricos de una banda y actualiza la fila si mejora.
        private void TryReOcrRow(Bitmap original, Rectangle band, ColorimetricRow row, List<string> log)
        {
            int w = original.Width;

            // Dividir la banda en 5 zonas para L,a,b,C,h
            // Las columnas de datos empiezan aproximadamente al 32% del ancho
            int dataStart = (int)(w * 0.32);
            int dataWidth = w - dataStart;
            int colW = dataWidth / 5;

            var colNames = new[] { "L", "a", "b", "C", "h" };
            var results = new double[5];
            var success = new bool[5];

            for (int ci = 0; ci < 5; ci++)
            {
                var cellRect = new Rectangle(
                    dataStart + ci * colW + 2,
                    band.Y,
                    colW - 4,
                    band.Height);

                // A: OCR de celda individual con preproceso agresivo
                string cellText = OcrCellRegion(original, cellRect, PageSegMode.SingleWord);
                cellText = cellText.Trim()
                    .Replace(" ", "").Replace(",", ".")
                    .Replace("O", "0").Replace("l", "1").Replace("I", "1");

                var m = Regex.Match(cellText, @"-?\d+(?:\.\d+)?");
                if (m.Success)
                {
                    double v = RestoreMeasureDecimal(m.Value);
                    results[ci] = v;
                    success[ci] = true;
                    if (log != null)
                        log.Add(string.Format("[REOCR] {0}/{1} col={2} raw='{3}' → {4:F4}",
                            row.Illuminant, row.Type, colNames[ci], cellText.Trim(), v));
                }
            }

            double newA = success[1] ? results[1] : row.A;
            double newB = success[2] ? results[2] : row.B;
            double newC = success[3] ? results[3] : row.Chroma;
            double newH = success[4] ? results[4] : row.Hue;
            double newL = success[0] ? results[0] : row.L;

            // Aplicar correcciones de coherencia sobre los valores re-OCR
            newB = FixBviaChroma(newB, newB.ToString("F2"), newA, newC);
            newA = FixAviaChroma(newA, newA.ToString("F2"), newB, newC);

            double errNew = Math.Abs(Math.Sqrt(newA * newA + newB * newB) - newC);
            double errOld = Math.Abs(Math.Sqrt(row.A * row.A + row.B * row.B) - row.Chroma);

            if (errNew < errOld)
            {
                if (log != null)
                    log.Add(string.Format(
                        "[REOCR] {0}/{1} MEJORA err {2:F4}→{3:F4} | a:{4:F2}→{5:F2} b:{6:F2}→{7:F2}",
                        row.Illuminant, row.Type, errOld, errNew, row.A, newA, row.B, newB));

                if (ColorimetryRanges.IsValidL(newL)) row.L = newL;
                if (ColorimetryRanges.IsValidAB(newA)) row.A = newA;
                if (ColorimetryRanges.IsValidAB(newB)) row.B = newB;
                if (ColorimetryRanges.IsValidChroma(newC)) row.Chroma = newC;
                if (ColorimetryRanges.IsValidHue(newH)) row.Hue = newH;
                row.NeedsReview = errNew > REOCR_CHROMA_THRESHOLD;
            }
            else
            {
                if (log != null)
                    log.Add(string.Format("[REOCR] {0}/{1} sin mejora errNew={2:F4} >= errOld={3:F4}",
                        row.Illuminant, row.Type, errNew, errOld));
            }
        }

        /// A: OCR de una región específica de la imagen original con preproceso agresivo.
        private string OcrCellRegion(Bitmap original, Rectangle region, PageSegMode psm)
        {
            // Validar bounds
            region = Rectangle.Intersect(region, new Rectangle(0, 0, original.Width, original.Height));
            if (region.Width < 4 || region.Height < 4) return string.Empty;

            // Agregar padding
            int pad = REOCR_PADDING;
            var padded = new Rectangle(
                Math.Max(0, region.X - pad),
                Math.Max(0, region.Y - pad),
                Math.Min(original.Width - Math.Max(0, region.X - pad), region.Width + pad * 2),
                Math.Min(original.Height - Math.Max(0, region.Y - pad), region.Height + pad * 2));

            // Recortar
            Bitmap crop;
            try { crop = original.Clone(padded, original.PixelFormat); }
            catch { return string.Empty; }

            // A: Escalar más agresivamente que el preproceso normal
            int nW = crop.Width * REOCR_SCALE;
            int nH = crop.Height * REOCR_SCALE;
            var scaled = new Bitmap(nW, nH, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(crop, 0, 0, nW, nH);
            }
            crop.Dispose();

            // A: Binarización adaptativa — umbral Otsu simplificado
            var binary = BinarizeOtsu(scaled);
            scaled.Dispose();

            // OCR con PSM específico para celda
            string tmp = Path.Combine(Path.GetTempPath(),
                string.Format("reocr_{0:N}.png", Guid.NewGuid()));
            try
            {
                binary.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                binary.Dispose();

                lock (_engineLock)
                {
                    var eng = GetEngine();
                    eng.SetVariable("tessedit_char_whitelist", "0123456789.-");
                    using (var img = Pix.LoadFromFile(tmp))
                    using (var page = eng.Process(img, psm))
                    {
                        return page.GetText() ?? string.Empty;
                    }
                }
            }
            catch { return string.Empty; }
            finally { if (File.Exists(tmp)) File.Delete(tmp); }
        }

        /// Binarización por umbral Otsu — C# 7.3, sin unsafe, usa Marshal.Copy para velocidad.
        private static Bitmap BinarizeOtsu(Bitmap src)
        {
            int w = src.Width, h = src.Height;
            var hist = new int[256];

            // Leer todos los píxeles en un array byte[] via Marshal.Copy (sin unsafe)
            var srcData = src.LockBits(new Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            int stride = srcData.Stride;
            int byteCount = stride * h;
            var pixels = new byte[byteCount];
            System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, pixels, 0, byteCount);
            src.UnlockBits(srcData);

            // Calcular histograma de luminancia
            for (int y = 0; y < h; y++)
            {
                int rowOffset = y * stride;
                for (int x = 0; x < w; x++)
                {
                    int idx = rowOffset + x * 4;
                    int b2 = pixels[idx];
                    int g2 = pixels[idx + 1];
                    int r2 = pixels[idx + 2];
                    int lum = (int)(0.299 * r2 + 0.587 * g2 + 0.114 * b2);
                    hist[Math.Min(255, Math.Max(0, lum))]++;
                }
            }

            // Calcular umbral Otsu
            long total = w * h;
            long sumB = 0, wB = 0, sum1 = 0;
            for (int i = 0; i < 256; i++) sum1 += i * hist[i];
            double maxVar = 0;
            int threshold = 128;
            for (int t = 0; t < 256; t++)
            {
                wB += hist[t];
                if (wB == 0) continue;
                long wF = total - wB;
                if (wF == 0) break;
                sumB += t * hist[t];
                double mB = (double)sumB / wB;
                double mF = (double)(sum1 - sumB) / wF;
                double varT = wB * wF * (mB - mF) * (mB - mF);
                if (varT > maxVar) { maxVar = varT; threshold = t; }
            }

            // Aplicar umbral y escribir resultado via Marshal.Copy
            var result = new Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var dstData = result.LockBits(new Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.WriteOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            int dstStride = dstData.Stride;
            var dstPixels = new byte[dstStride * h];

            for (int y = 0; y < h; y++)
            {
                int srcRow = y * stride;
                int dstRow = y * dstStride;
                for (int x = 0; x < w; x++)
                {
                    int si = srcRow + x * 4;
                    int di = dstRow + x * 4;
                    int b2 = pixels[si];
                    int g2 = pixels[si + 1];
                    int r2 = pixels[si + 2];
                    int lum = (int)(0.299 * r2 + 0.587 * g2 + 0.114 * b2);
                    byte val = lum <= threshold ? (byte)0 : (byte)255;
                    dstPixels[di] = val;
                    dstPixels[di + 1] = val;
                    dstPixels[di + 2] = val;
                    dstPixels[di + 3] = 255;
                }
            }

            System.Runtime.InteropServices.Marshal.Copy(dstPixels, 0, dstData.Scan0, dstPixels.Length);
            result.UnlockBits(dstData);
            return result;
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
                bmp.Save(tmp, System.Drawing.Imaging.ImageFormat.Png);
                lock (_engineLock)
                {
                    using (var img = Pix.LoadFromFile(tmp))
                    using (var page = GetEngine().Process(img, PageSegMode.SingleBlock))
                    {
                        return page.GetText();
                    }
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
                    // Corregir signo de a*/b* del Lot usando Std como referencia
                    if (stdRow != null && lotRow != null)
                        FixSignByStd(stdRow, lotRow);
                    // Corregir L* del Lot si es outlier respecto a mediciones ya parseadas
                    if (lotRow != null)
                    {
                        string lTok = ExtractNumericTokens(lotR2).Count > 0 ? ExtractNumericTokens(lotR2)[0] : "";
                        lotRow.L = FixLviaNeighbors(lotRow.L, lTok, report.Measures);
                    }
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

            double vL = FixLabRange(RestoreMeasureDecimal(tokens[base_]), 0, 100);
            double vA = FixLabRange(RestoreMeasureDecimal(tokens[base_ + 1]), -100, 100);
            double vB = FixLabRange(RestoreMeasureDecimal(tokens[base_ + 2]), -100, 100);
            double vC = FixLabRange(RestoreMeasureDecimal(tokens[base_ + 3]), 0, 200);

            if (log != null)
                log.Add(string.Format("[TOKENS] {0}/{1} L={2} a={3} b={4} C={5} h={6}",
                    illuminant, type,
                    tokens[base_], tokens[base_ + 1], tokens[base_ + 2],
                    tokens[base_ + 3], tokens[base_ + 4]));

            // FIX UNIVERSAL...
            // FIX UNIVERSAL: deteccion de escala x10
            // Cuando OCR lee a*, b* y Chroma todos x10 (ej: 6.10->0.61, 17.89->1.78, 18.90->1.89)
            // la coherencia interna sqrt(a2+b2)~C es correcta pero los valores son 10x el real.
            // Senal: Chroma/L > 0.6 es anomalo en colorimetria textil (tipico Chroma < L*0.5)
            if (vL > 0 && vC / vL > 0.6 && vC > 5.0)
            {
                double aDiv = vA / 10.0;
                double bDiv = vB / 10.0;
                double cDiv = vC / 10.0;
                double chromaCheckDiv = Math.Sqrt(aDiv * aDiv + bDiv * bDiv);
                double errDiv = Math.Abs(chromaCheckDiv - cDiv);
                // Aplicar /10 si coherencia se mantiene (err<0.5) Y ratio C/L baja a rango normal
                if (errDiv < 0.5 && cDiv / vL < 0.5)
                {
                    if (log != null)
                        log.Add(string.Format(
                            "[FIX/10] {0}/{1} a:{2:F2}->{3:F2} b:{4:F2}->{5:F2} C:{6:F2}->{7:F2}",
                            illuminant, type, vA, aDiv, vB, bDiv, vC, cDiv));
                    vA = aDiv; vB = bDiv; vC = cDiv;
                }
            }

            // Corregir b* via coherencia con Chroma (OCR confunde 3↔8, 6↔8, dígito faltante)
            vB = FixBviaChroma(vB, tokens[base_ + 2], vA, vC);
            // FIX: Corregir a* también via coherencia con Chroma (mismo tipo de error)
            vA = FixAviaChroma(vA, tokens[base_ + 1], vB, vC);
            double vH = ParseHueDouble(tokens[base_ + 4]);

            // Corregir signos de a* y b* usando el ángulo Hue como referencia
            CorrectSignsByHue(ref vA, ref vB, vH);

            // FIX: Recalcular Hue desde a*/b* corregidos cuando el OCR entregó valor inválido
            // Cubre: 784→194, 249→195, 257→206, 0→206
            vH = CorrectHueFromAB(vH, vA, vB);

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

                double l = FixLabRange(RestoreMeasureDecimal(tokL), 0, 100);
                double a = FixLabRange(RestoreMeasureDecimal(tokA), -100, 100);
                double b = FixLabRange(RestoreMeasureDecimal(tokB), -100, 100);
                double c = FixLabRange(RestoreMeasureDecimal(tokC), 0, 200);
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
        // También detecta punto mal ubicado: "-36.0"→"-3.60", "-50.6"→"-5.06"
        private static double RestoreMeasureDecimal(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            token = token.Trim().Replace(',', '.');

            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;

            if (token.Contains("."))
            {
                double vDirect = SafeParse(token);
                // Si ya está en rango LAB razonable, confiar directamente
                if (Math.Abs(vDirect) <= 100) return vDirect;
                // Fuera de rango: OCR puso el punto en posición incorrecta
                // Ej: "-36.0" → allDigits="360" → insertar 2 del final → "-3.60"
                int dotIdx2 = d.IndexOf('.');
                if (dotIdx2 > 0)
                {
                    string allDigits = d.Replace(".", "");
                    if (allDigits.Length >= 3)
                    {
                        string rebuilt = (neg ? "-" : "")
                            + allDigits.Substring(0, allDigits.Length - 2)
                            + "." + allDigits.Substring(allDigits.Length - 2);
                        double cand;
                        if (double.TryParse(rebuilt,
                            System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out cand)
                            && Math.Abs(cand) <= 100)
                            return cand;
                    }
                }
                return vDirect;
            }

            if (!Regex.IsMatch(d, @"^\d+$")) return SafeParse(token);
            if (d.Length <= 2) return SafeParse(token);

            // Para 3 dígitos: preferir XX.X (ej: 247 → 24.7) sobre X.XX (ej: 247 → 2.47)
            if (d.Length == 3)
            {
                double candXX1;
                string rebuildXX1 = (neg ? "-" : "") + d.Substring(0, 2) + "." + d[2];
                double candX2X;
                string rebuildX2X = (neg ? "-" : "") + d[0] + "." + d.Substring(1);
                bool okXX1 = double.TryParse(rebuildXX1, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out candXX1) && Math.Abs(candXX1) <= 100;
                bool okX2X = double.TryParse(rebuildX2X, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out candX2X) && Math.Abs(candX2X) <= 100;

                if (okXX1) return candXX1;
                if (okX2X) return candX2X;
            }

            return SafeParse((neg ? "-" : "") + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2));
        }

        // Corrige signo de a*/b* del Lot usando el Std como referencia
        // Si Lot.a tiene signo opuesto a Std.a (físicamente imposible), invertir
        private static void FixSignByStd(ColorimetricRow std, ColorimetricRow lot)
        {
            // FIX: Solo corregir si Std tiene magnitud suficiente para ser referencia confiable
            // (evita cascada cuando Std también llegó con signo incorrecto del OCR)
            const double MIN_RELIABLE = 0.5;

            if (Math.Abs(std.A) >= MIN_RELIABLE && std.A * lot.A < 0) lot.A = -lot.A;
            if (Math.Abs(std.B) >= MIN_RELIABLE && std.B * lot.B < 0) lot.B = -lot.B;

            // Recalcular Hue del Lot desde a*/b* corregidos
            double newHue = Math.Atan2(lot.B, lot.A) * 180.0 / Math.PI;
            if (newHue < 0) newHue += 360.0;
            // Solo actualizar Hue si el cambio es significativo (>5°)
            double diff = Math.Abs(newHue - lot.Hue);
            if (diff > 180) diff = 360 - diff;
            if (diff > 5) lot.Hue = Math.Round(newHue);
        }

        // Corrige b* usando coherencia con Chroma: prueba variantes OCR (3↔8, 6↔8)
        private static double FixBviaChroma(double bRestored, string bToken, double a, double cOcr)
        {
            const double THRESHOLD = 0.5;
            double bOrig = bRestored;

            double errOrig = Math.Abs(Math.Sqrt(a * a + bOrig * bOrig) - cOcr);
            if (errOrig <= THRESHOLD) return bOrig;

            double bestErr = errOrig;
            double bestVal = bOrig;

            // ── FIX CRÍTICO: punto decimal desplazado ────────────────────────────
            // OCR lee "-36.0" pero real es "-3.60" | "-50.6"→"-5.06" | "-63.2"→"-6.32"
            // Método A: desde el token string (más preciso)
            {
                string bStr2 = (bToken ?? "").Trim().Replace(',', '.');
                bool bNeg2 = bStr2.StartsWith("-");
                string abs2 = bNeg2 ? bStr2.Substring(1) : bStr2;
                int dp2 = abs2.IndexOf('.');
                if (dp2 > 1) // "36.0", "50.6", "63.2" — 2+ dígitos antes del punto
                {
                    // Mover punto 1 izquierda: "36.0"→"3.60", "50.6"→"5.06"
                    string newAbs2 = abs2.Substring(0, dp2 - 1) + "." + abs2[dp2 - 1] + abs2.Substring(dp2 + 1);
                    double cs2;
                    if (double.TryParse((bNeg2 ? "-" : "") + newAbs2,
                        NumberStyles.Float, CultureInfo.InvariantCulture, out cs2))
                    {
                        double ep = Math.Abs(Math.Sqrt(a * a + cs2 * cs2) - cOcr);
                        double epn = Math.Abs(Math.Sqrt(a * a + (-cs2) * (-cs2)) - cOcr);
                        if (ep < bestErr && Math.Abs(cs2) <= 100) { bestErr = ep; bestVal = cs2; }
                        if (epn < bestErr && Math.Abs(cs2) <= 100) { bestErr = epn; bestVal = -cs2; }
                    }
                }
            }
            // Método B: desde el valor numérico (cubre casos sin token limpio)
            // bOrig=-36.0 → digits="360" → "-3.60"
            {
                string absNumStr = Math.Abs(bOrig).ToString("G10", CultureInfo.InvariantCulture);
                string digs2 = absNumStr.Replace(".", "").Replace(",", "");
                bool bOrigNeg = bOrig < 0;
                if (digs2.Length >= 3)
                {
                    // Insertar punto a 2 del final
                    string rb2 = (bOrigNeg ? "-" : "") + digs2.Substring(0, digs2.Length - 2) + "." + digs2.Substring(digs2.Length - 2);
                    double cv2;
                    if (double.TryParse(rb2, NumberStyles.Float, CultureInfo.InvariantCulture, out cv2)
                        && Math.Abs(cv2) <= 100)
                    {
                        double ep2 = Math.Abs(Math.Sqrt(a * a + cv2 * cv2) - cOcr);
                        double ep2n = Math.Abs(Math.Sqrt(a * a + (-cv2) * (-cv2)) - cOcr);
                        if (ep2 < bestErr) { bestErr = ep2; bestVal = cv2; }
                        if (ep2n < bestErr) { bestErr = ep2n; bestVal = -cv2; }
                    }
                    // Insertar punto a 1 del final
                    string rb1 = (bOrigNeg ? "-" : "") + digs2.Substring(0, digs2.Length - 1) + "." + digs2.Substring(digs2.Length - 1);
                    double cv1;
                    if (double.TryParse(rb1, NumberStyles.Float, CultureInfo.InvariantCulture, out cv1)
                        && Math.Abs(cv1) <= 100)
                    {
                        double ep1 = Math.Abs(Math.Sqrt(a * a + cv1 * cv1) - cOcr);
                        double ep1n = Math.Abs(Math.Sqrt(a * a + (-cv1) * (-cv1)) - cOcr);
                        if (ep1 < bestErr) { bestErr = ep1; bestVal = cv1; }
                        if (ep1n < bestErr) { bestErr = ep1n; bestVal = -cv1; }
                    }
                }
                // Retorno temprano si ya encontramos solución buena
                if (bestErr <= THRESHOLD) return bestVal;
            }

            // ── Si el positivo es incoherente con Chroma, probar negativo ────────
            if (errOrig > 2.0 && bOrig > 0)
            {
                double errNeg = Math.Abs(Math.Sqrt(a * a + (-bOrig) * (-bOrig)) - cOcr);
                if (errNeg < bestErr) { bestErr = errNeg; bestVal = -bOrig; }
            }

            // ── Confusiones de dígitos (3↔8, 6↔8, etc.) ────────────────────────
            var confusions = new Dictionary<char, char[]>
            {
                {'3', new[]{'8'}}, {'8', new[]{'3'}},
                {'6', new[]{'8'}}, {'0', new[]{'9'}},
                {'9', new[]{'0', '8'}},
                {'1', new[]{'7'}}, {'7', new[]{'1'}}
            };

            bool neg = bestVal < 0;
            string s = bToken.TrimStart().TrimStart('-').Replace(".", "").Replace(",", "");

            // Probar 1 dígito
            for (int i = 0; i < s.Length; i++)
            {
                char[] reps;
                if (!confusions.TryGetValue(s[i], out reps)) continue;
                foreach (char rep in reps)
                {
                    string ns = s.Substring(0, i) + rep + s.Substring(i + 1);
                    double v;
                    if (!double.TryParse((neg ? "-" : "") + ns,
                        NumberStyles.Float, CultureInfo.InvariantCulture, out v)) continue;
                    double e = Math.Abs(Math.Sqrt(a * a + v * v) - cOcr);
                    if (e < bestErr) { bestErr = e; bestVal = v; }
                }
            }

            // Probar 2 dígitos
            for (int i = 0; i < s.Length; i++)
            {
                char[] reps1;
                if (!confusions.TryGetValue(s[i], out reps1)) continue;
                for (int j = i + 1; j < s.Length; j++)
                {
                    char[] reps2;
                    if (!confusions.TryGetValue(s[j], out reps2)) continue;
                    foreach (char r1 in reps1)
                        foreach (char r2 in reps2)
                        {
                            var chars = s.ToCharArray();
                            chars[i] = r1; chars[j] = r2;
                            double v;
                            if (!double.TryParse((neg ? "-" : "") + new string(chars),
                                NumberStyles.Float, CultureInfo.InvariantCulture, out v)) continue;
                            double e = Math.Abs(Math.Sqrt(a * a + v * v) - cOcr);
                            if (e < bestErr) { bestErr = e; bestVal = v; }
                        }
                }
            }

            // ── Dígito faltante ──────────────────────────────────────────────────
            {
                string bStrFix = (bToken ?? "").Trim().Replace(',', '.');
                bool bNegFix = bStrFix.StartsWith("-");
                string absFix = bNegFix ? bStrFix.Substring(1) : bStrFix;
                int absDot = absFix.IndexOf('.');
                if (absDot > 0 && errOrig > 2.0)
                {
                    string absInt = absFix.Substring(0, absDot);
                    string decPart = absFix.Substring(absDot);
                    if (absInt.Length <= 2)
                    {
                        for (char ins = '0'; ins <= '9'; ins++)
                        {
                            string candidate = (neg ? "-" : "") + absInt + ins + decPart;
                            double v;
                            if (!double.TryParse(candidate, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out v)) continue;
                            if (Math.Abs(v) > 100) continue;
                            double e = Math.Abs(Math.Sqrt(a * a + v * v) - cOcr);
                            if (e < bestErr) { bestErr = e; bestVal = v; }
                        }
                        for (char ins = '1'; ins <= '9'; ins++)
                        {
                            string candidate = (neg ? "-" : "") + ins + absInt + decPart;
                            double v;
                            if (!double.TryParse(candidate, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out v)) continue;
                            if (Math.Abs(v) > 100) continue;
                            double e = Math.Abs(Math.Sqrt(a * a + v * v) - cOcr);
                            if (e < bestErr) { bestErr = e; bestVal = v; }
                        }
                    }
                }
            }

            return bestVal;
        }

        // Corrige a* via coherencia con Chroma (simétrico a FixBviaChroma)
        private static double FixAviaChroma(double aRestored, string aToken, double b, double cOcr)
        {
            const double THRESHOLD = 0.5;
            double aOrig = aRestored;

            double errOrig = Math.Abs(Math.Sqrt(aOrig * aOrig + b * b) - cOcr);
            if (errOrig <= THRESHOLD) return aOrig;

            double bestErr = errOrig;
            double bestVal = aOrig;

            // ── FIX CRÍTICO: punto decimal desplazado + signo perdido ────────────
            // OCR lee "50.01" pero real es "-15.03"
            // Método A: desde token string
            {
                string aStr = (aToken ?? "").Trim().Replace(',', '.');
                bool aStrNeg = aStr.StartsWith("-");
                string absStr2 = aStrNeg ? aStr.Substring(1) : aStr;
                int dotPos2 = absStr2.IndexOf('.');
                if (dotPos2 > 1)
                {
                    string newAbs = absStr2.Substring(0, dotPos2 - 1) + "." + absStr2[dotPos2 - 1] + absStr2.Substring(dotPos2 + 1);
                    double candShift;
                    if (double.TryParse((aStrNeg ? "-" : "") + newAbs,
                        NumberStyles.Float, CultureInfo.InvariantCulture, out candShift))
                    {
                        foreach (double cand in new[] { candShift, -candShift })
                        {
                            if (Math.Abs(cand) > 100) continue;
                            double e = Math.Abs(Math.Sqrt(cand * cand + b * b) - cOcr);
                            if (e < bestErr) { bestErr = e; bestVal = cand; }
                        }
                    }
                }
            }
            // Método B: desde valor numérico (cubre casos con token corrupto)
            // aOrig=50.01 → digits="5001" → "-5.00"... también probar negativo
            {
                string absNumStr = Math.Abs(aOrig).ToString("G10", CultureInfo.InvariantCulture);
                string digs2 = absNumStr.Replace(".", "").Replace(",", "");
                bool aOrigNeg = aOrig < 0;
                if (digs2.Length >= 3)
                {
                    string rb2 = (aOrigNeg ? "-" : "") + digs2.Substring(0, digs2.Length - 2) + "." + digs2.Substring(digs2.Length - 2);
                    double cv2;
                    if (double.TryParse(rb2, NumberStyles.Float, CultureInfo.InvariantCulture, out cv2)
                        && Math.Abs(cv2) <= 100)
                    {
                        foreach (double cand in new[] { cv2, -cv2 })
                        {
                            double ep = Math.Abs(Math.Sqrt(cand * cand + b * b) - cOcr);
                            if (ep < bestErr) { bestErr = ep; bestVal = cand; }
                        }
                    }
                }
                if (bestErr <= THRESHOLD) return bestVal;
            }

            // ── Confusiones de dígitos ───────────────────────────────────────────
            var confusions = new Dictionary<char, char[]>
            {
                {'3', new[]{'8'}}, {'8', new[]{'3'}},
                {'6', new[]{'8'}}, {'0', new[]{'9'}},
                {'9', new[]{'0', '8'}},
                {'1', new[]{'7'}}, {'7', new[]{'1'}}
            };

            bool neg = aRestored < 0;
            string s = aToken.TrimStart().TrimStart('-').Replace(".", "").Replace(",", "");

            // Probar 1 dígito
            for (int i = 0; i < s.Length; i++)
            {
                char[] reps;
                if (!confusions.TryGetValue(s[i], out reps)) continue;
                foreach (char rep in reps)
                {
                    string ns = s.Substring(0, i) + rep + s.Substring(i + 1);
                    double v;
                    if (!double.TryParse((neg ? "-" : "") + ns,
                        NumberStyles.Float, CultureInfo.InvariantCulture, out v)) continue;
                    double e = Math.Abs(Math.Sqrt(v * v + b * b) - cOcr);
                    if (e < bestErr) { bestErr = e; bestVal = v; }
                }
            }

            // ── Dígito faltante ──────────────────────────────────────────────────
            {
                string aStrFix = (aToken ?? "").Trim().Replace(',', '.');
                bool aNegFix = aStrFix.StartsWith("-");
                string absFix2 = aNegFix ? aStrFix.Substring(1) : aStrFix;
                int absDot = absFix2.IndexOf('.');
                if (absDot > 0 && errOrig > 2.0)
                {
                    string absInt = absFix2.Substring(0, absDot);
                    string decPart = absFix2.Substring(absDot);
                    if (absInt.Length <= 2)
                    {
                        for (char ins = '0'; ins <= '9'; ins++)
                        {
                            string candidate = (neg ? "-" : "") + absInt + ins + decPart;
                            double v;
                            if (!double.TryParse(candidate, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out v)) continue;
                            if (Math.Abs(v) > 100) continue;
                            double e = Math.Abs(Math.Sqrt(v * v + b * b) - cOcr);
                            if (e < bestErr) { bestErr = e; bestVal = v; }
                        }
                    }
                }
            }

            return bestVal;
        }

        // Corrige L* usando coherencia entre iluminantes: si difiere >2 del promedio, prueba 6↔8
        private static double FixLviaNeighbors(double lVal, string lToken, List<ColorimetricRow> existing)
        {
            const double OUTLIER_THRESHOLD = 2.0; // FIX: bajado de 3.0 a 2.0 (cubre 76.94 vs 78.94)
            if (existing == null || existing.Count < 2) return lVal;
            double sum = 0; int cnt = 0;
            foreach (var r in existing) { sum += r.L; cnt++; }
            double avg = sum / cnt;
            if (Math.Abs(lVal - avg) <= OUTLIER_THRESHOLD) return lVal;

            // Intentar reemplazar dígito confundido
            string s = lToken.Trim().Replace(',', '.');
            var confusions = new Dictionary<char, char[]>
            {
                {'6', new[]{'8','9'}},
                {'8', new[]{'6','9'}},
                {'3', new[]{'8'}},
                {'4', new[]{'8'}},   // FIX: 4 puede confundirse con 8 (76.94 → 78.94)
                {'7', new[]{'8'}},
                {'9', new[]{'8'}},
                {'0', new[]{'8'}}
            };
            double bestDiff = Math.Abs(lVal - avg);
            double bestVal = lVal;
            for (int i = 0; i < s.Length; i++)
            {
                char[] reps;
                if (!confusions.TryGetValue(s[i], out reps)) continue;
                foreach (char rep in reps)
                {
                    string ns = s.Substring(0, i) + rep + s.Substring(i + 1);
                    double v;
                    if (!double.TryParse(ns, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out v)) continue;
                    if (v < 0 || v > 100) continue;
                    double diff = Math.Abs(v - avg);
                    if (diff < bestDiff) { bestDiff = diff; bestVal = v; }
                }
            }
            return bestVal;
        }

        // Corrige valores LAB/Chroma fuera de rango por OCR (8283.00 → 82.83, 1468.00 → 14.68)
        private static double FixLabRange(double v, double minV, double maxV)
        {
            if (v >= minV && v <= maxV) return v;
            bool neg = v < 0;
            string s = ((long)Math.Abs(v)).ToString();

            // Insertar punto decimal desde el final (XX.XX)
            if (s.Length >= 3)
            {
                string rebuilt = (neg ? "-" : "") + s.Substring(0, s.Length - 2) + "." + s.Substring(s.Length - 2);
                double candidate;
                if (double.TryParse(rebuilt, NumberStyles.Float, CultureInfo.InvariantCulture, out candidate)
                    && candidate >= minV && candidate <= maxV)
                    return candidate;
            }

            // FIX: 3 digitos sin punto → X.XX (ej: 229 → 2.29, 289 → 2.89)
            if (s.Length == 3)
            {
                string rebuilt2 = (neg ? "-" : "") + s[0] + "." + s.Substring(1);
                double candidate2;
                if (double.TryParse(rebuilt2, NumberStyles.Float, CultureInfo.InvariantCulture, out candidate2)
                    && candidate2 >= minV && candidate2 <= maxV)
                    return candidate2;
            }

            return v;
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
            string s = token.Trim().Replace(',', '.');
            double v = SafeParse(s);
            if (v < 0) v = 0;

            // FIX: Hue con dígito extra al inicio (ej: 784 → 84 o 194)
            // Probar quitar primer dígito, segundo dígito y último dígito
            if (v > 360)
            {
                string digits = s.TrimStart('-');
                // Probar quitar primer dígito: 784 → 84
                if (digits.Length >= 3)
                {
                    string trimFirst = digits.Substring(1);
                    double cand1;
                    if (double.TryParse(trimFirst, NumberStyles.Float, CultureInfo.InvariantCulture, out cand1)
                        && cand1 >= 0 && cand1 <= 360)
                        return cand1;
                }
                // Probar quitar último dígito: 2490 → 249 (si aún >360, seguir)
                if (digits.Length >= 4)
                {
                    string trimLast = digits.Substring(0, digits.Length - 1);
                    double cand2;
                    if (double.TryParse(trimLast, NumberStyles.Float, CultureInfo.InvariantCulture, out cand2)
                        && cand2 >= 0 && cand2 <= 360)
                        return cand2;
                }
                // Fallback: módulo 360
                v = v % 360;
            }

            // FIX: Hue = 0 cuando debería ser ~200: el OCR perdió dígitos.
            // No podemos corregir aquí sin a*/b* — se corregirá en CorrectHueFromAB post-parse.
            return v;
        }

        /// <summary>
        /// Recalcula Hue desde a* y b* cuando el valor OCR es claramente incorrecto.
        /// Llamar después de que a* y b* estén ya corregidos.
        /// </summary>
        private static double CorrectHueFromAB(double hue, double a, double b)
        {
            // Hue calculado desde a*/b* corregidos
            double hueCalc = Math.Atan2(b, a) * 180.0 / Math.PI;
            if (hueCalc < 0) hueCalc += 360.0;

            // Siempre usar el hue calculado desde a*/b* cuando:
            // 1) Hue OCR es 0 (OCR perdió el valor)
            // 2) Error > 5° (OCR leyó dígito incorrecto: 248→194, 256→202)
            double hueErr = HueError(hue, hueCalc);
            if (hue < 1.0 || hueErr > 5.0)
                return Math.Round(hueCalc);

            return hue;
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
            map['8'] = new[] { '3', '6', '0', '9' }; // FIX: agregar 8→9
            map['9'] = new[] { '0', '8' };             // FIX: agregar 9→8

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

            // FIX: Eliminar dígito extra al inicio (ej: 784 → 84 o 194; probar quitar cada dígito)
            if (s.IndexOf('.') < 0 && s.TrimStart('-').Length > 3)
            {
                string digits = s.TrimStart('-');
                bool isNeg = s.StartsWith("-");
                // Quitar primer dígito
                set.Add((isNeg ? "-" : "") + digits.Substring(1));
                // Quitar último dígito
                set.Add((isNeg ? "-" : "") + digits.Substring(0, digits.Length - 1));
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