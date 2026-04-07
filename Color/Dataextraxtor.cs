using OpenCvSharp;
using OpenCvSharp.Extensions;
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

namespace Color
{
    // MODELOS
    public class ColorimetricRow
    {
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double L { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double Chroma { get; set; }
        public double Hue { get; set; }

        public int HueInt
        {
            get { return (int)Math.Round(Hue, MidpointRounding.AwayFromZero); }
        }

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
        public double DeltaLightness { get; set; }
        public double DeltaChroma { get; set; }
        public double DeltaHue { get; set; }
        public double? DeltaCMC { get; set; }

        // Texto crudo (del OCR) solo para diagnóstico
        public string LightnessFlagOcr { get; set; }
        public string ChromaFlagOcr { get; set; }
        public string HueFlagOcr { get; set; }

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

    public class BatchInfo
    {
        public string ShadeName { get; set; }
        public string LotNo { get; set; }
        public string BatchId { get; set; }
        public double dE { get; set; }
        public double dL { get; set; }
        public double dC { get; set; }
        public double dH { get; set; }
        public string PF { get; set; } // Pass / Fail
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

        // --- Estructura para Historial Detallado (Know-How) ---
        public BatchInfo Batch { get; set; } = new BatchInfo();

        public string DiagnosticoL { get; set; }
        public string Recomendacion { get; set; }
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
        public string ColumnHint { get; set; }
    }
    // EXTRACTOR PRINCIPAL
    public class ColorimetricDataExtractor : IDisposable
    {
        private readonly string _tessDataPath;
        private const int SCALE_FACTOR = 3;
        private const bool ENFORCE_ONE_PER_ILLUMINANT_TYPE = true;

        // Instancia reutilizable de Tesseract 
        private TesseractEngine _sharedEngine;
        private readonly object _engineLock = new object();

        // Autocorrección por coherencia (1 dígito) en C* y h°
        private const bool ENABLE_COHERENCE_FIX = true;

        // A+C: Re-OCR dirigido cuando sqrt(a²+b²) no coincide con Chroma
        private const bool ENABLE_REOCR = true;
        private const double REOCR_CHROMA_THRESHOLD = 0.35;
        private const int REOCR_SCALE = 6;
        private const int REOCR_PADDING = 4;

        // Preprocesado adaptativo: ajusta escala y contraste según calidad de imagen
        private static readonly bool ENABLE_ADAPTIVE_PREPROCESS = true;
        private const float SHARPNESS_LOW_THRESHOLD = 40f;
        private const float SHARPNESS_HIGH_THRESHOLD = 120f;
        private const int SCALE_FACTOR_LOW = 4;
        private const int SCALE_FACTOR_HIGH = 2;
        private const float CONTRAST_LOW = 2.2f;
        private const float CONTRAST_HIGH = 1.5f;

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
            // Delegar al pipeline completo (OpenCV + OCR por celda)
            return ExtractReportFromBitmap(original).Measures;
        }
        // ... dentro de la clase ColorimetricDataExtractor ...
        private double ExtractDoubleSafe(Mat gray, Rectangle rect)
        {
            lock (_engineLock)
            {
                return ExtractDouble(GetEngine(), gray, rect);
            }
        }

        private static double ExtractDouble(TesseractEngine engine, Mat grayImage, Rectangle rect)
        {
            try
            {
                // Asegurar que el recorte esté dentro de los límites de la imagen
                rect.X = Math.Max(0, rect.X);
                rect.Y = Math.Max(0, rect.Y);
                rect.Width = Math.Min(rect.Width, grayImage.Width - rect.X);
                rect.Height = Math.Min(rect.Height, grayImage.Height - rect.Y);

                using (Mat roi = new Mat(grayImage, new OpenCvSharp.Rect(rect.X, rect.Y, rect.Width, rect.Height)))
                using (Mat resized = new Mat())
                {
                    // Agrandar 3 veces para capturar puntos decimales pequeños
                    Cv2.Resize(roi, resized, new OpenCvSharp.Size(0, 0), 3.0, 3.0, InterpolationFlags.Cubic);
                    Cv2.Threshold(resized, resized, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

                    using (Bitmap bmp = OpenCvSharp.Extensions.BitmapConverter.ToBitmap(resized))
                    using (var page = engine.Process(bmp, PageSegMode.SingleLine))
                    {
                        string text = (page.GetText() ?? "").Trim().Replace(" ", "").Replace(",", ".");
                        text = System.Text.RegularExpressions.Regex.Replace(text, @"[^0-9.\-]", "");

                        if (double.TryParse(text, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double val))
                            return val;
                    }
                }
            }
            catch { /* Error de límites o formato */ }
            return 0.0;
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
            // ── INTENTO 1: OpenCV — detección de tabla + OCR por celda ──────────
            OcrReport report = null;
            bool openCvSuccess = false;

            try
            {
                report = ExtractReportFromBitmapOpenCV(original);
                if (report != null && report.Measures != null && report.Measures.Count >= 4)
                {
                    openCvSuccess = true;
                    report.ParseLog.Add(string.Format(
                        "[OPENCV] Extracción exitosa: {0} mediciones", report.Measures.Count));
                }
                else
                {
                    var prevLog = report != null ? new List<string>(report.ParseLog) : new List<string>();
                    report = new OcrReport();
                    report.ParseLog.AddRange(prevLog);
                    report.ParseLog.Add("[OPENCV] Detección insuficiente → fallback OCR global");
                }
            }
            catch (Exception ex)
            {
                report = new OcrReport();
                report.ParseLog.Add("[OPENCV] Error → fallback OCR global: " + ex.Message);
            }

            // ── INTENTO 2: OCR global (fallback cuando OpenCV falla) ─────────────
            if (!openCvSuccess)
            {
                try
                {
                    var prevLog = new List<string>(report.ParseLog);
                    using (var processed = Preprocess(original))
                    {
                        string text = RunOCR(processed);
                        var fallbackReport = new OcrReport();
                        ParseCombinedTable(text, fallbackReport);
                        var tFb = ParseTolerances(text);
                        fallbackReport.TolDL = tFb.Item1;
                        fallbackReport.TolDC = tFb.Item2;
                        fallbackReport.TolDH = tFb.Item3;
                        fallbackReport.TolDE = tFb.Item4;
                        fallbackReport.PrintDate = ParsePrintDate(text);
                        fallbackReport.ParseLog.InsertRange(0, prevLog);
                        fallbackReport.ParseLog.Add("[FALLBACK] OCR global ejecutado.");
                        report = fallbackReport;
                    }
                    if (ENABLE_REOCR)
                        ReOcrFailedCells(original, report);
                }
                catch (Exception exFb)
                {
                    if (report == null) report = new OcrReport();
                    report.ParseLog.Add("[FALLBACK] Error: " + exFb.Message);
                }
            }

            // ── POST-PROCESADO: LocalCorrector ───────────────────────────────────
            try
            {
                var corrections = LocalCorrector.CorrectReport(report);
                foreach (var c in corrections)
                    report.ParseLog.Add(string.Format(
                        "[LOCAL] {0}/{1} {2}: {3:F4}→{4:F4} ({5})",
                        c.Illuminant, c.Type, c.Field,
                        c.OriginalValue, c.CorrectedValue, c.Reason));
            }
            catch (Exception ex)
            {
                report.ParseLog.Add("[LOCAL] Error: " + ex.Message);
            }

            return report;
        }

        // ── OpenCV: extracción por detección de tabla ─────────────────────────

        private OcrReport ExtractReportFromBitmapOpenCV(Bitmap original)
        {
            var report = new OcrReport();

            // 1. Detectar tabla con OpenCV
            var detection = ColrTableDetector.Detect(original);
            if (!detection.Success)
            {
                report.ParseLog.Add("[OPENCV] " + (detection.FailReason ?? "Tabla no detectada"));
                return report;
            }

            report.ParseLog.Add(string.Format(
                "[OPENCV] Tabla detectada: {0} filas × {1} cols, {2} celdas",
                detection.RowCount, detection.ColCount, detection.Cells.Count));

            // 2. OCR por celda — usar ScaledImage si está disponible (mayor resolución)
            var cellTexts = new Dictionary<int, Dictionary<int, string>>();
            Bitmap ocrSource = (detection.ScaledImage != null) ? detection.ScaledImage : original;

            try
            {
                foreach (var cell in detection.Cells)
                {
                    bool isText = cell.Col <= 1;
                    string cellText = OcrCellRegion(ocrSource, cell.Bounds,
                        isText ? PageSegMode.SingleWord : PageSegMode.SingleWord,
                        numericOnly: !isText);
                    cellText = CleanCellText(cellText, cell.Col);

                    if (!cellTexts.ContainsKey(cell.Row))
                        cellTexts[cell.Row] = new Dictionary<int, string>();
                    cellTexts[cell.Row][cell.Col] = cellText;

                    if (!string.IsNullOrWhiteSpace(cellText))
                        report.ParseLog.Add(string.Format(
                            "[OPENCV] Celda ({0},{1})[{2}]: '{3}'",
                            cell.Row, cell.Col, cell.FieldName, cellText));
                }
            }
            finally
            {
                if (detection.ScaledImage != null)
                {
                    detection.ScaledImage.Dispose();
                    detection.ScaledImage = null;
                }
            }

            // 3. Encontrar fila de inicio de datos
            int _rowCount = 0;
            foreach (var _c in detection.Cells)
                if (_c.Row + 1 > _rowCount) _rowCount = _c.Row + 1;
            int dataStartRow = FindDataStartRow(cellTexts, _rowCount);
            if (dataStartRow < 0)
            {
                report.ParseLog.Add("[OPENCV] No se encontró fila de datos con iluminante conocido");
                return report;
            }

            report.ParseLog.Add(string.Format("[OPENCV] Datos desde fila {0}", dataStartRow));

            // 4. Construir OcrReport desde las celdas
            var builtReport = ColrTableParser.BuildReport(cellTexts, dataStartRow, report.ParseLog);

            // 5. Leer tolerancias si están presentes (OCR global en banda inferior)
            var tols = ParseTolerancesFromImage(original, detection.TableBounds);
            if (tols != null)
            {
                builtReport.TolDL = tols.Item1;
                builtReport.TolDC = tols.Item2;
                builtReport.TolDH = tols.Item3;
                builtReport.TolDE = tols.Item4;
            }

            // Transferir log
            builtReport.ParseLog.InsertRange(0, report.ParseLog);

            // 6. Aplicar validación por Deltas CMC — detectar ×10 automáticamente
            ApplyDeltaCmcValidation(builtReport);

            // 7. Dedup y orden
            builtReport.Measures = DedupAndSort(builtReport.Measures);
            builtReport.CmcDifferences = DedupCmc(builtReport.CmcDifferences);

            return builtReport;
        }

        private static string CleanCellText(string text, int col)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            text = text.Trim();

            if (col <= 1)
            {
                // Columnas de texto: iluminante y tipo
                return NormalizeOCRLine(text);
            }

            // Columnas numéricas: limpiar OCR manteniendo signo y punto
            text = text
                .Replace(',', '.')
                .Replace('O', '0')
                .Replace('o', '0')
                .Replace('l', '1')
                .Replace('I', '1')
                .Replace('|', '1')
                .Replace('S', '5')
                .Replace('s', '5')
                // MEJORA 1: sustituciones OCR adicionales
                .Replace('Z', '2')
                .Replace('z', '2')
                .Replace('`', ' ')
                .Replace('\'', ' ');

            // MEJORA 1: eliminar cualquier carácter que no sea dígito, punto, signo o espacio
            text = Regex.Replace(text, @"[^0-9\.\-\s]", "");

            // Quitar espacios internos
            text = Regex.Replace(text, @"\s+", "");

            // normalizar signo negativo — si hay un '-' no al inicio, moverlo al inicio
            {
                int midMinus = text.IndexOf('-', 1);
                if (midMinus > 0)
                    text = text.Substring(0, midMinus) + text.Substring(midMinus + 1);
            }

            // Si tiene múltiples puntos, quitar los extras
            int firstDot = text.IndexOf('.');
            if (firstDot >= 0)
            {
                string before = text.Substring(0, firstDot + 1);
                string after = text.Substring(firstDot + 1).Replace(".", "");
                text = before + after;
            }

            return text;
        }

        /// Detecta en qué fila empiezan los datos buscando la primera celda con iluminante conocido.
        private static int FindDataStartRow(
            Dictionary<int, Dictionary<int, string>> cellTexts, int totalRows)
        {
            for (int r = 0; r < totalRows; r++)
            {
                if (!cellTexts.ContainsKey(r)) continue;
                var row = cellTexts[r];

                // Buscar en col 0 y col 1 (por si hay desplazamiento de columnas en OCR)
                for (int c = 0; c <= 1; c++)
                {
                    string cellVal;
                    if (!row.TryGetValue(c, out cellVal)) continue;
                    string norm = NormalizeOCRLine(cellVal ?? "");
                    foreach (string ill in ColorimetricDataExtractor.KnownIlluminants)
                    {
                        if (norm == ill.ToUpperInvariant() || norm.Contains(ill.ToUpperInvariant()))
                            return r;
                    }
                }
            }
            return -1;
        }
        /// OCR sobre la banda inferior de la imagen (debajo de la tabla) para leer tolerancias.
        private Tuple<double, double, double, double> ParseTolerancesFromImage(
            Bitmap original, Rectangle tableBounds)
        {
            try
            {
                int belowY = tableBounds.Bottom + 5;
                if (belowY >= original.Height) return null;

                var band = new Rectangle(0, belowY, original.Width,
                    Math.Min(60, original.Height - belowY));
                band = Rectangle.Intersect(band,
                    new Rectangle(0, 0, original.Width, original.Height));
                if (band.IsEmpty) return null;

                string bandText = OcrCellRegion(original, band, PageSegMode.SingleBlock);
                return ParseTolerances(bandText ?? "");
            }
            catch { return null; }
        }

        /// Usa los Deltas CMC reportados en la imagen para detectar error ×10 automáticamente.
        private static void ApplyDeltaCmcValidation(OcrReport report)
        {
            if (report.CmcDifferences == null || report.CmcDifferences.Count == 0) return;
            if (report.Measures == null || report.Measures.Count < 2) return;

            foreach (var cmc in report.CmcDifferences)
            {
                double dCReported = Math.Abs(cmc.DeltaChroma);
                if (dCReported < 0.001) continue;

                // Buscar Std y Lot del mismo iluminante
                ColorimetricRow stdRow = null, lotRow = null;
                foreach (var r in report.Measures)
                {
                    if (!string.Equals(r.Illuminant, cmc.Illuminant,
                        StringComparison.OrdinalIgnoreCase)) continue;
                    if (r.Type == "Std") stdRow = r;
                    else if (r.Type == "Lot") lotRow = r;
                }

                if (stdRow == null || lotRow == null) continue;

                double dCCalculated = Math.Abs(lotRow.Chroma - stdRow.Chroma);
                if (dCCalculated < 0.001) continue;

                double ratio = dCCalculated / dCReported;

                // Si el ratio es ~10 → los valores están ×10
                if (ratio > 5.0 && ratio < 20.0)
                {
                    double dCDiv10 = dCCalculated / 10.0;
                    double errDiv = Math.Abs(dCDiv10 - dCReported);
                    double errOrig = Math.Abs(dCCalculated - dCReported);

                    if (errDiv < errOrig * 0.3) // ÷10 mejora al menos 70%
                    {
                        report.ParseLog.Add(string.Format(
                            "[CMC-VAL] {0}: ΔC calculado={1:F3} reportado={2:F3} ratio={3:F1} → ×10 detectado, dividiendo por 10",
                            cmc.Illuminant, dCCalculated, dCReported, ratio));

                        // Dividir a*, b*, C* por 10 en Std y Lot
                        DivideByTen(stdRow, report.ParseLog);
                        DivideByTen(lotRow, report.ParseLog);
                    }
                }
            }
        }
        private static void DivideByTen(ColorimetricRow row, List<string> log)
        {
            double newA = row.A / 10.0;
            double newB = row.B / 10.0;
            double newC = row.Chroma / 10.0;

            if (log != null)
                log.Add(string.Format(
                    "[CMC-VAL] {0}/{1} ÷10: a:{2:F2}→{3:F2} b:{4:F2}→{5:F2} C:{6:F2}→{7:F2}",
                    row.Illuminant, row.Type, row.A, newA, row.B, newB, row.Chroma, newC));

            row.A = newA;
            row.B = newB;
            row.Chroma = newC;

            // Recalcular Hue
            double newHue = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
            if (newHue < 0) newHue += 360.0;
            row.Hue = Math.Round(newHue);

            row.NeedsReview = false;
        }
        // ══════════════════════════════════════════════════════════════
        // ESTRATEGIA A+C — Re-OCR dirigido sobre celdas con error Chroma
        // ══════════════════════════════════════════════════════════════
        private void ReOcrFailedCells(Bitmap original, OcrReport report)
        {
            if (report == null || report.Measures == null || report.Measures.Count == 0)
                return;

            bool anyNeeds = false;
            foreach (var row in report.Measures)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                if (Math.Abs(chromaCalc - row.Chroma) > REOCR_CHROMA_THRESHOLD)
                {
                    anyNeeds = true;
                    break;
                }
            }
            if (!anyNeeds) return;

            // ✅ AQUÍ se crean mat y gray
            using (var mat = OpenCvSharp.Extensions.BitmapConverter.ToMat(original))
            using (var gray = new Mat())
            {
                Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);

                var detection = ColrTableDetector.Detect(original);
                if (!detection.Success || detection.Cells.Count == 0)
                {
                    if (detection.ScaledImage != null) detection.ScaledImage.Dispose();
                    return;
                }

                try
                {
                    foreach (var row in report.Measures)
                    {
                        double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                        double chromaErr = Math.Abs(chromaCalc - row.Chroma);

                        if (chromaErr <= REOCR_CHROMA_THRESHOLD)
                            continue;

                        var bRectScaled = FindRowRectByBoxes(detection, row.Illuminant, row.Type);
                        if (!bRectScaled.HasValue)
                            continue;

                        // Des-escalar el rectángulo a la imagen original
                        float invScale = 1f / detection.ScaleFactor;
                        var bRect = new Rectangle(
                            (int)(bRectScaled.Value.X * invScale),
                            (int)(bRectScaled.Value.Y * invScale),
                            (int)(bRectScaled.Value.Width * invScale),
                            (int)(bRectScaled.Value.Height * invScale)
                        );

                        // ✅ AQUÍ ES DONDE SE LLAMA CORRECTAMENTE
                        double newB = ExtractDoubleSafe(gray, bRect);

                        if (ColorimetryRanges.IsValidAB(newB))
                        {
                            double newChroma = Math.Sqrt(row.A * row.A + newB * newB);
                            if (Math.Abs(newChroma - row.Chroma) < chromaErr)
                            {
                                report.ParseLog.Add(
                                    $"[REOCR/NUM] {row.Illuminant}/{row.Type} b*: {row.B:F2} → {newB:F2}"
                                );

                                row.B = newB;
                                row.Chroma = newChroma;

                                double h = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                                if (h < 0) h += 360.0;
                                row.Hue = Math.Round(h);

                                row.NeedsReview = false;
                            }
                        }
                    }
                }

                finally
                {
                    if (detection.ScaledImage != null)
                        detection.ScaledImage.Dispose();
                }
            }
        }
        // ── Localizar celda B* usando celdas OpenCV ────────────────────────────

        private Rectangle? FindRowRectByBoxes(
            ColrTableDetector.CvDetectionResult detection, string illuminant, string type)
        {
            if (detection.Cells == null || detection.Cells.Count == 0) return null;

            string illUpper = illuminant.ToUpperInvariant();
            string typeUpper = type.ToUpperInvariant();

            int maxRow = 0;
            foreach (var c in detection.Cells)
                if (c.Row > maxRow) maxRow = c.Row;

            for (int r = 0; r <= maxRow; r++)
            {
                ColrTableDetector.CvTableCell illCell = null;
                ColrTableDetector.CvTableCell typeCell = null;
                ColrTableDetector.CvTableCell bCell = null;

                foreach (var c in detection.Cells)
                {
                    if (c.Row == r && c.Col == 0) illCell = c;
                    if (c.Row == r && c.Col == 1) typeCell = c;
                    if (c.Row == r && c.Col == 4) bCell = c;
                }

                if (illCell == null || typeCell == null || bCell == null) continue;

                string illText = NormalizeOCRLine(OcrCellRegion(detection.ScaledImage, illCell.Bounds, PageSegMode.SingleWord, false));
                if (!illText.Contains(illUpper)) continue;

                string typeText = NormalizeOCRLine(OcrCellRegion(detection.ScaledImage, typeCell.Bounds, PageSegMode.SingleWord, false));
                if (typeText.Contains(typeUpper) ||
                    (type == "Std" && (typeText.Contains("STD") || typeText.Contains("5TD") || typeText.Contains("ST0"))) ||
                    (type == "Lot" && (typeText.Contains("LOT") || typeText.Contains("L0T"))))
                {
                    // Retornar directamente la celda OpenCV de la columna 4 (b*)
                    return bCell.Bounds;
                }
            }
            return null;
        }
        /// A: OCR de una región específica de la imagen original con preproceso agresivo.
        private string OcrCellRegion(Bitmap original, Rectangle region, PageSegMode psm, bool numericOnly = true)
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

            // Escala dinámica: celdas grandes (≥30px) ×2, pequeñas ×4
            int cellScale = crop.Height >= 30 ? 2 : 4;
            int nW = crop.Width * cellScale;
            int nH = crop.Height * cellScale;
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
                    if (numericOnly) eng.SetVariable("tessedit_char_whitelist", "0123456789.-");
                    try
                    {
                        using (var img = Pix.LoadFromFile(tmp))
                        using (var page = eng.Process(img, psm))
                        {
                            return page.GetText() ?? string.Empty;
                        }
                    }
                    finally
                    {
                        // MEJORA 7: resetear whitelist para no contaminar llamadas OCR globales
                        eng.SetVariable("tessedit_char_whitelist", "");
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

        private static float MeasureSharpness(Bitmap src)
        {
            int w = src.Width, h = src.Height;
            // Kernel Laplaciano 3x3
            int[] kernel = { 0, -1, 0, -1, 4, -1, 0, -1, 0 };

            // Leer pixels via LockBits (evita GetPixel que es lento)
            var bmpData = src.LockBits(
                new Rectangle(0, 0, w, h),
                System.Drawing.Imaging.ImageLockMode.ReadOnly,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            int stride = bmpData.Stride;
            var pixels = new byte[stride * h];
            System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, pixels, 0, pixels.Length);
            src.UnlockBits(bmpData);

            // Zona de muestreo: 25%-75% de la imagen para evitar márgenes blancos
            int x0 = w / 4, x1 = w * 3 / 4;
            int y0 = h / 4, y1 = h * 3 / 4;

            double sum = 0, sumSq = 0;
            long count = 0;

            for (int y = Math.Max(1, y0); y < Math.Min(h - 1, y1); y++)
            {
                for (int x = Math.Max(1, x0); x < Math.Min(w - 1, x1); x++)
                {
                    int lap = 0;
                    int ki = 0;
                    for (int dy = -1; dy <= 1; dy++)
                    {
                        for (int dx = -1; dx <= 1; dx++)
                        {
                            int idx = (y + dy) * stride + (x + dx) * 4;
                            int b2 = pixels[idx];
                            int g2 = pixels[idx + 1];
                            int r2 = pixels[idx + 2];
                            int lum = (int)(0.299 * r2 + 0.587 * g2 + 0.114 * b2);
                            lap += kernel[ki] * lum;
                            ki++;
                        }
                    }
                    sum += lap;
                    sumSq += (double)lap * lap;
                    count++;
                }
            }

            if (count == 0) return 50f; // valor neutro si no hay píxeles
            double mean = sum / count;
            double variance = (sumSq / count) - (mean * mean);
            return (float)Math.Max(0, variance);
        }

        private Bitmap Preprocess(Bitmap src)
        {
            int scaleFactor;
            float contrast;

            if (ENABLE_ADAPTIVE_PREPROCESS)
            {
                float sharpness = MeasureSharpness(src);

                if (sharpness < SHARPNESS_LOW_THRESHOLD)
                {
                    // Imagen borrosa o baja resolución → escala agresiva + contraste alto
                    scaleFactor = SCALE_FACTOR_LOW;
                    contrast = CONTRAST_LOW;
                }
                else if (sharpness > SHARPNESS_HIGH_THRESHOLD)
                {
                    // Imagen ya nítida → escala moderada + contraste suave
                    scaleFactor = SCALE_FACTOR_HIGH;
                    contrast = CONTRAST_HIGH;
                }
                else
                {
                    // Rango normal → valores originales (sin cambio de comportamiento)
                    scaleFactor = SCALE_FACTOR;
                    contrast = 1.8f;
                }
            }
            else
            {
                // Comportamiento original preservado exactamente
                scaleFactor = SCALE_FACTOR;
                contrast = 1.8f;
            }

            int nW = src.Width * scaleFactor, nH = src.Height * scaleFactor;
            var scaled = new Bitmap(nW, nH, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.DrawImage(src, 0, 0, nW, nH);
            }
            var gray = ToGrayscaleContrast(scaled, contrast, -0.1f);
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
            // CAMBIO 3 — Resumen de extracción al final del log para diagnóstico rápido
            int totalMeasures = report.Measures.Count;
            int needsReview = 0;
            foreach (var r in report.Measures) { if (r.NeedsReview) needsReview++; }
            int warnCount = 0;
            foreach (var entry in report.ParseLog) { if (entry.StartsWith("[WARN]")) warnCount++; }

            report.ParseLog.Add(string.Format(
                "[RESUMEN] Filas extraídas: {0} | Requieren revisión: {1} | Advertencias log: {2}",
                totalMeasures, needsReview, warnCount));

            if (needsReview > 0)
                report.ParseLog.Add(string.Format(
                    "[RESUMEN] Iluminantes con NeedsReview=true: {0}",
                    string.Join(", ", report.Measures
                        .FindAll(delegate (ColorimetricRow r) { return r.NeedsReview; })
                        .ConvertAll(delegate (ColorimetricRow r) { return r.Illuminant + "/" + r.Type; }))));
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

            // Eliminar FIX UNIVERSAL (/10) y FIX UNIVERSAL INVERSO (x10): 
            {
                double vH_pre = ParseHueDouble(tokens[base_ + 4]);
                bool hueValid = vH_pre >= 1.0 && vH_pre <= 360.0;

                // Contar dígitos enteros (antes del punto) en cada token numérico relevante
                int digL = CountIntDigits(tokens[base_]);
                int digA = CountIntDigits(tokens[base_ + 1]);
                int digB = CountIntDigits(tokens[base_ + 2]);
                int digC = CountIntDigits(tokens[base_ + 3]);

                // Señal de truncamiento: L* tiene más dígitos que a*, b* y C*
                bool suspiciouslySmallChroma = vL > 10.0 && vC < vL * 0.15 && vC < 5.0;
                bool tokensMissingDigit = digL >= 2 && digA <= 1 && digB <= 1 && digC <= 1;

                // GUARDIA: si L* está fuera del rango textil confiable (20-100),
                bool lInTextileRange = vL >= 20.0 && vL <= 100.0;

                if (hueValid && lInTextileRange && (suspiciouslySmallChroma || tokensMissingDigit))
                {
                    // Calcular Chroma esperado desde Hue y la magnitud del L*
                    double hRad = vH_pre * Math.PI / 180.0;
                    double cosH = Math.Cos(hRad);
                    double sinH = Math.Sin(hRad);

                    double bestTripleErr = double.MaxValue;
                    double bestA = vA, bestB = vB, bestC = vC;

                    // Probar insertar cada dígito '1'-'9' como primer dígito de los tokens
                    for (char ins = '1'; ins <= '9'; ins++)
                    {
                        double tryA = TryInsertLeadingDigit(tokens[base_ + 1], ins);
                        double tryB = TryInsertLeadingDigit(tokens[base_ + 2], ins);
                        double tryC = TryInsertLeadingDigit(tokens[base_ + 3], ins);

                        if (double.IsNaN(tryA) || double.IsNaN(tryB) || double.IsNaN(tryC)) continue;
                        if (Math.Abs(tryA) > 100 || Math.Abs(tryB) > 100 || tryC > 200) continue;

                        // Coherencia interna con los nuevos valores
                        double chromaCalc = Math.Sqrt(tryA * tryA + tryB * tryB);
                        double chromaErr = Math.Abs(chromaCalc - tryC);
                        if (chromaErr > 1.5) continue;

                        // Coherencia con Hue: el ángulo atan2(b,a) debe coincidir con vH_pre
                        double hueCalc = Math.Atan2(tryB, tryA) * 180.0 / Math.PI;
                        if (hueCalc < 0) hueCalc += 360.0;
                        double hueErr = Math.Abs(hueCalc - vH_pre);
                        if (hueErr > 180) hueErr = 360 - hueErr;
                        if (hueErr > 15.0) continue;

                        double totalErr = chromaErr + hueErr * 0.1;
                        if (totalErr < bestTripleErr)
                        {
                            bestTripleErr = totalErr;
                            bestA = tryA; bestB = tryB; bestC = tryC;
                        }
                    }

                    if (bestTripleErr < double.MaxValue && (bestC > vC * 1.5))
                    {
                        if (log != null)
                            log.Add(string.Format(
                                "[FIX/DIGIT] {0}/{1} truncamiento simultaneo detectado: " +
                                "a:{2:F2}->{3:F2} b:{4:F2}->{5:F2} C:{6:F2}->{7:F2}",
                                illuminant, type, vA, bestA, vB, bestB, vC, bestC));
                        vA = bestA; vB = bestB; vC = bestC;
                    }
                }
            }
            // Corregir b* via coherencia con Chroma 
            vB = FixBviaChroma(vB, tokens[base_ + 2], vA, vC);
            // FIX: Corregir a* también via coherencia con Chroma 
            vA = FixAviaChroma(vA, tokens[base_ + 1], vB, vC);
            double vH = ParseHueDouble(tokens[base_ + 4]);

            // Corregir signos de a* y b* usando el ángulo Hue como referencia
            CorrectSignsByHue(ref vA, ref vB, vH);

            // FIX: Recalcular Hue desde a*/b* 
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

                // MEJORA 4: cuando la coherencia interna es excelente (err < 0.3),
                if (chromaErr < 0.3)
                {
                    vC = Math.Round(chromaCalc, 2);
                    double hCalc = Math.Atan2(vB, vA) * 180.0 / Math.PI;
                    if (hCalc < 0) hCalc += 360.0;
                    vH = Math.Round(hCalc, 2);
                }
                else if (chromaErr > chromaTol)
                {
                    if (log != null)
                        log.Add(string.Format(
                            "[WARN] {0}/{1} Chroma={2:F2} pero sqrt(a^2+b^2)={3:F2} (err={4:F2}) -> digito OCR incorrecto en a*, b* o Chroma",
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
                    if (ratio > 3.0)
                    {
                        // FIX: el token a* positivo de 3 dígitos puede haberse restaurado
                        double aAlt = RestoreMeasureDecimalAlt(tokA);
                        if (!double.IsNaN(aAlt))
                        {
                            double aAltFixed = FixLabRange(aAlt, -100, 100);
                            double chromaAlt = Math.Sqrt(aAltFixed * aAltFixed + b * b);
                            double ratioAlt = c > 0.001 && chromaAlt > 0.001
                                ? (c > chromaAlt ? c / chromaAlt : chromaAlt / c)
                                : 1.0;
                            if (ratioAlt <= 3.0)
                            {
                                // Usar a* alternativo — tokens[i+1] se corregirá
                                a = aAltFixed;
                                // Continuar con la validación usando a* corregido
                                goto AcceptPosition;
                            }
                        }
                        continue;
                    }
                }

            AcceptPosition:
                return i;
            }
            return 0;
        }

        private static double RestoreMeasureDecimalAlt(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return double.NaN;
            string t = token.Trim().Replace(',', '.');
            if (t.StartsWith("-")) return double.NaN;
            if (t.Contains(".")) return double.NaN;
            if (!Regex.IsMatch(t, @"^\d{3}$")) return double.NaN;
            double v;
            string candidate = t[0] + "." + t.Substring(1);
            return double.TryParse(candidate, NumberStyles.Float,
                CultureInfo.InvariantCulture, out v) && v <= 100 ? v : double.NaN;
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
                // Con punto: solo aceptar exactamente 1 decimal (ej: "259.0", "241.6") 
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

            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;

            if (token.Contains("."))
            {
                double vDirect = SafeParse(token);
                // Si ya está en rango LAB razonable, confiar directamente
                if (Math.Abs(vDirect) <= 100) return vDirect;
                // Fuera de rango: OCR puso el punto en posición incorrecta
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

            // Sin punto decimal y más de 2 caracteres: asumir 2 decimales fijos (ej. 189 -> 1.89)
            return SafeParse((neg ? "-" : "") + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2));
        }

        // Corrige signo de a*/b* del Lot usando el Std como referencia
        private static void FixSignByStd(ColorimetricRow std, ColorimetricRow lot)
        {
            // FIX: Solo corregir si Std tiene magnitud suficiente para ser referencia confiable
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
            {
                string bStr2 = (bToken ?? "").Trim().Replace(',', '.');
                bool bNeg2 = bStr2.StartsWith("-");
                string abs2 = bNeg2 ? bStr2.Substring(1) : bStr2;
                int dp2 = abs2.IndexOf('.');
                if (dp2 > 1)
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
            // Método B: desde el TOKEN original (más confiable que el valor numérico)
            {
                string bTokClean = (bToken ?? "").Trim().Replace(',', '.');
                bool bTokNeg = bTokClean.StartsWith("-");
                string bTokAbs = bTokNeg ? bTokClean.Substring(1) : bTokClean;
                string digs2 = bTokAbs.Replace(".", "").Replace(",", "");
                bool bOrigNeg = bOrig < 0;
                if (digs2.Length >= 3)
                {
                    // Insertar punto a 2 del final: "360"→"3.60", "506"→"5.06"
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
                    // Insertar punto a 1 del final: "360"→"36.0"
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
            // ── Punto decimal desplazado ×10 ─────────────────────────────────────
            if (errOrig > 2.0)
            {
                double bMul = bOrig * 10.0;
                if (Math.Abs(bMul) <= 100.0)
                {
                    double eMul = Math.Abs(Math.Sqrt(a * a + bMul * bMul) - cOcr);
                    if (eMul < bestErr) { bestErr = eMul; bestVal = bMul; }
                }
                double bMulNeg = -Math.Abs(bOrig * 10.0);
                if (Math.Abs(bMulNeg) <= 100.0)
                {
                    double eMulNeg = Math.Abs(Math.Sqrt(a * a + bMulNeg * bMulNeg) - cOcr);
                    if (eMulNeg < bestErr) { bestErr = eMulNeg; bestVal = bMulNeg; }
                }
                if (bestErr <= THRESHOLD) return bestVal;
            }
            // ── Confusiones de dígitos (3↔8, 6↔8, etc.) ────────────────────────
            var confusions = new Dictionary<char, char[]>
            {
                {'3', new[]{'8'}},         {'8', new[]{'3', '6', '0'}},
                {'6', new[]{'8', '5'}},    {'0', new[]{'9'}},
                {'9', new[]{'0', '8', '4'}},
                {'1', new[]{'7'}},         {'7', new[]{'1', '2'}},
                {'2', new[]{'7'}},         {'4', new[]{'9'}},
                {'5', new[]{'6', '8'}},
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
            // Método B: desde el TOKEN original (más confiable que el valor numérico)
            {
                string aTokClean = (aToken ?? "").Trim().Replace(',', '.');
                bool aTokNeg = aTokClean.StartsWith("-");
                string aTokAbs = aTokNeg ? aTokClean.Substring(1) : aTokClean;
                string digs2 = aTokAbs.Replace(".", "").Replace(",", "");
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

            // ── Punto decimal desplazado ×10 ─────────────────────────────────────
            if (errOrig > 2.0)
            {
                double aMul = aOrig * 10.0;
                if (Math.Abs(aMul) <= 100.0)
                {
                    double eMul = Math.Abs(Math.Sqrt(aMul * aMul + b * b) - cOcr);
                    if (eMul < bestErr) { bestErr = eMul; bestVal = aMul; }
                }
                double aMulNeg = -Math.Abs(aOrig * 10.0);
                if (Math.Abs(aMulNeg) <= 100.0)
                {
                    double eMulNeg = Math.Abs(Math.Sqrt(aMulNeg * aMulNeg + b * b) - cOcr);
                    if (eMulNeg < bestErr) { bestErr = eMulNeg; bestVal = aMulNeg; }
                }
                if (bestErr <= THRESHOLD) return bestVal;
            }

            // ── Confusiones de dígitos ───────────────────────────────────────────
            var confusions = new Dictionary<char, char[]>
            {
                {'3', new[]{'8'}},         {'8', new[]{'3', '6', '0'}},
                {'6', new[]{'8', '5'}},    {'0', new[]{'9'}},
                {'9', new[]{'0', '8', '4'}},
                {'1', new[]{'7'}},         {'7', new[]{'1', '2'}},
                {'2', new[]{'7'}},         {'4', new[]{'9'}},
                {'5', new[]{'6', '8'}},
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
            const double OUTLIER_THRESHOLD = 2.0;
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
                {'4', new[]{'8'}},
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

        // Corrige valores LAB/Chroma fuera de rango por OCR 
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

            // FIX: 3 digitos sin punto 
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

            // FIX: Hue con dígito extra al inicio 
            if (v > 360)
            {
                string digits = s.TrimStart('-');
                // Probar quitar primer dígito: 
                if (digits.Length >= 3)
                {
                    string trimFirst = digits.Substring(1);
                    double cand1;
                    if (double.TryParse(trimFirst, NumberStyles.Float, CultureInfo.InvariantCulture, out cand1)
                        && cand1 >= 0 && cand1 <= 360)
                        return cand1;
                }
                // Probar quitar último dígito: 
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
            return v;
        }

        private static double CorrectHueFromAB(double hue, double a, double b)
        {
            // Hue calculado desde a*/b* corregidos
            double hueCalc = Math.Atan2(b, a) * 180.0 / Math.PI;
            if (hueCalc < 0) hueCalc += 360.0;

            // Siempre usar el hue calculado desde a*/b* cuando:
            double hueErr = HueError(hue, hueCalc);
            if (hue < 1.0 || hueErr > 5.0)
                return Math.Round(hueCalc);

            return hue;
        }
        // ── Helpers para FIX/DIGIT (truncamiento simultáneo) ──────────────────
        private static int CountIntDigits(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            string s = token.Trim().TrimStart('-');
            int dot = s.IndexOf('.');
            string intPart = dot >= 0 ? s.Substring(0, dot) : s;
            return intPart.TrimStart('0').Length == 0 ? 1 : intPart.Length;
        }
        /// Intenta insertar 'ins' como primer dígito de la parte entera del token.
        private static double TryInsertLeadingDigit(string token, char ins)
        {
            if (string.IsNullOrWhiteSpace(token)) return double.NaN;
            string s = token.Trim().Replace(',', '.');
            bool neg = s.StartsWith("-");
            string abs = neg ? s.Substring(1) : s;

            // Construir nuevo token con dígito insertado al inicio de la parte entera
            string candidate = (neg ? "-" : "") + ins + abs;
            double v;
            if (!double.TryParse(candidate, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out v)) return double.NaN;
            if (Math.Abs(v) > 100) return double.NaN;
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
            map['8'] = new[] { '3', '6', '0', '9' };
            map['9'] = new[] { '0', '8' };

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
            // Insertar punto si no existe 
            if (s.IndexOf('.') < 0 && Regex.IsMatch(s, @"^-?\d{3,}$"))
            {
                int at = s.StartsWith("-") ? s.Length - 2 : s.Length - 2;
                if (at > 0 && at < s.Length)
                {
                    string candDot = s.Substring(0, at) + "." + s.Substring(at);
                    set.Add(candDot);
                }
            }
            // FIX: Eliminar dígito extra al inicio
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
                .Replace('\u2212', '-')
                .Replace(',', '.');

            // 1) Enmascarar números dentro del nombre del iluminante 
            line = MaskIlluminantNumbers(line);

            // 2) Eliminar etiquetas STD/LOT
            line = Regex.Replace(line, @"\b(STD|5TD|LOT)\b", " ", RegexOptions.IgnoreCase);

            // 3) Única corrección textual segura: "CW/F" → "CWF"
            line = Regex.Replace(line, @"\bCW/F\b", "CWF", RegexOptions.IgnoreCase);

            // 4) Extraer tokens numéricos 
            var matches = Regex.Matches(line, @"-?\d+(?:\.\d+)?");
            var rawTokens = new List<string>(matches.Count);
            foreach (Match m in matches) rawTokens.Add(m.Value);

            // 5) Dividir tokens largos 
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
                .Replace('\u2212', '-')
                .Replace(',', '.')
                .Replace('\u2018', '\'')
                .Replace('\u201C', '"')
                .Replace('\u201D', '"')
                .Replace('\u2013', '-')
                .Replace('\u2014', '-');

            t = t.ToUpperInvariant();

            // ── STD / LOT ─────────────────────────────────────────────────────────
            t = Regex.Replace(t, @"\b5TD\b", "STD");
            t = Regex.Replace(t, @"\b5T0\b", "STD");
            t = Regex.Replace(t, @"\bST0\b", "STD");
            t = Regex.Replace(t, @"\bSTB\b", "STD");
            t = Regex.Replace(t, @"\bSTO\b", "STD");
            t = Regex.Replace(t, @"\b5L0\b", "STD");
            t = Regex.Replace(t, @"\bL0T\b", "LOT");

            // ── TL84 / TL83 / TL85 ───────────────────────────────────────────────
            t = Regex.Replace(t, @"\bTLS4\b", "TL84");
            t = Regex.Replace(t, @"\bTLS3\b", "TL83");
            t = Regex.Replace(t, @"\bTLS5\b", "TL85");
            t = Regex.Replace(t, @"\bTL8A\b", "TL84");
            t = Regex.Replace(t, @"\bTL8B\b", "TL83");
            t = Regex.Replace(t, @"\bTLB4\b", "TL84");
            t = Regex.Replace(t, @"\bTLB3\b", "TL83");
            t = Regex.Replace(t, @"\bTI84\b", "TL84");
            t = Regex.Replace(t, @"\bTI83\b", "TL83");
            t = Regex.Replace(t, @"\bT1[8S]4\b", "TL84");
            t = Regex.Replace(t, @"\b1L84\b", "TL84");
            t = Regex.Replace(t, @"\b1L83\b", "TL83");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*4\b", "TL84");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*3\b", "TL83");
            t = Regex.Replace(t, @"\bT\s*L\s*[8S]\s*5\b", "TL85");

            // ── D65 / D50 / D55 / D75 ────────────────────────────────────────────
            t = Regex.Replace(t, @"\bD6S\b", "D65");
            t = Regex.Replace(t, @"\bDG5\b", "D65");
            t = Regex.Replace(t, @"\b065\b", "D65");
            t = Regex.Replace(t, @"\bDO5\b", "D65");
            t = Regex.Replace(t, @"\bD6O\b", "D65");
            t = Regex.Replace(t, @"\bDSO\b", "D50");
            t = Regex.Replace(t, @"\bD5O\b", "D50");
            t = Regex.Replace(t, @"\bDS0\b", "D50");
            t = Regex.Replace(t, @"\b0S0\b", "D50");
            t = Regex.Replace(t, @"\bD\s*6\s*5\b", "D65");
            t = Regex.Replace(t, @"\bD\s*[5S]\s*0\b", "D50");
            t = Regex.Replace(t, @"\bD\s*[5S]\s*5\b", "D55");
            t = Regex.Replace(t, @"\bD\s*7\s*5\b", "D75");

            // ── CWF / CWF2 ───────────────────────────────────────────────────────
            t = Regex.Replace(t, @"\bCVF\b", "CWF");
            t = Regex.Replace(t, @"\bGWF\b", "CWF");
            t = Regex.Replace(t, @"\bCW-F\b", "CWF");
            t = Regex.Replace(t, @"\bC\s*W\s*F\b", "CWF");
            t = Regex.Replace(t, @"\bCW/F\b", "CWF");

            // ── F11 / F12 ─────────────────────────────────────────────────────────
            t = Regex.Replace(t, @"\bF\s*1\s*1\b", "F11");
            t = Regex.Replace(t, @"\bF\s*1\s*2\b", "F12");
            t = Regex.Replace(t, @"\bFI1\b", "F11");
            t = Regex.Replace(t, @"\bFl1\b", "F11");
            t = Regex.Replace(t, @"\bF1I\b", "F11");

            // ── MEJORA 6: iluminantes adicionales no cubiertos antes ──────────────
            t = Regex.Replace(t, @"\bUY\b", "UV");
            t = Regex.Replace(t, @"\bCWF1\b", "CWF");
            t = Regex.Replace(t, @"\bD6S5\b", "D65");
            t = Regex.Replace(t, @"\bD\s+65\b", "D65");

            // ── Otros ────────────────────────────────────────────────────────────
            t = Regex.Replace(t, @"\[LLUMINANT\b", "ILLUMINANT");

            // Limpieza conservadora (letras, dígitos, punto, guión, espacio)
            t = Regex.Replace(t, @"[^A-Z0-9\.\-\s]", " ");
            t = Regex.Replace(t, @"\s+", " ").Trim();
            return t;

        }

        // NUEVO — LECTURA DESDE EXCEL (MEDICIONES)
        public OcrReport ExtractReportFromExcel(string excelPath)
        {
            var report = new OcrReport();
            report.Measures = OCR.ExcelReader.LoadMeasurements(excelPath);
            return report;
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // CV TABLE DETECTOR — Detección de tabla colorimétrica con OpenCV multi-pass
    // ══════════════════════════════════════════════════════════════════════════
    internal static class ColrTableDetector
    {
        public class CvDetectionResult
        {
            public bool Success;
            public string FailReason;
            public Rectangle TableBounds;
            public int RowCount;
            public int ColCount;
            public List<CvTableCell> Cells = new List<CvTableCell>();
            public Bitmap ScaledImage;
            public float ScaleFactor;
        }

        public class CvTableCell
        {
            public int Row;
            public int Col;
            public Rectangle Bounds;
            public string Field;
            public string FieldName { get { return Field; } }
        }

        private const int TARGET_WIDTH = 1200;

        private struct DetectConfig
        {
            public int BlockSize; public double C;
            public int KernelW; public int KernelH;
            public int MinIntersect;
            public double HoughMinLen; public int HoughGap; public int MergeTol;
            public string Label;
        }

        private static readonly DetectConfig[] Configs = new DetectConfig[]
        {
            new DetectConfig { BlockSize=15, C=6,  KernelW=40, KernelH=40, MinIntersect=12, HoughMinLen=0.40, HoughGap=20, MergeTol=8,  Label="Sharp"    },
            new DetectConfig { BlockSize=21, C=8,  KernelW=30, KernelH=30, MinIntersect=8,  HoughMinLen=0.30, HoughGap=25, MergeTol=10, Label="Normal"   },
            new DetectConfig { BlockSize=31, C=10, KernelW=20, KernelH=20, MinIntersect=4,  HoughMinLen=0.20, HoughGap=35, MergeTol=14, Label="LowRes"   },
            new DetectConfig { BlockSize=41, C=12, KernelW=12, KernelH=12, MinIntersect=4,  HoughMinLen=0.15, HoughGap=50, MergeTol=18, Label="Degraded" },
        };

        public static CvDetectionResult Detect(Bitmap original)
        {
            float scale = (float)TARGET_WIDTH / original.Width;
            if (scale > 4f) scale = 4f;
            int scaledW = (int)(original.Width * scale);
            int scaledH = (int)(original.Height * scale);

            Bitmap scaled = new Bitmap(scaledW, scaledH, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, 0, 0, scaledW, scaledH);
            }

            CvDetectionResult result = null;
            try { result = DetectInternal(scaled); }
            catch { scaled.Dispose(); throw; }

            if (result == null) { scaled.Dispose(); return new CvDetectionResult { Success = false, FailReason = "Error interno." }; }

            result.ScaledImage = scaled;
            result.ScaleFactor = scale;

            // TableBounds en espacio original para ParseTolerancesFromImage
            if (result.Success)
            {
                float inv = 1f / scale;
                result.TableBounds = ScaleRect(result.TableBounds, inv);
            }
            // Bounds de celdas permanecen en espacio escalado (ScaledImage)
            return result;
        }

        private static CvDetectionResult DetectInternal(Bitmap bmp)
        {
            Mat mat = BitmapConverter.ToMat(bmp);
            Mat gray = new Mat();
            try
            {
                Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                Cv2.GaussianBlur(gray, gray, new OpenCvSharp.Size(3, 3), 0);

                CvDetectionResult best = null;
                foreach (var cfg in Configs)
                {
                    var attempt = TryDetect(gray, cfg);
                    if (attempt.Success)
                    {
                        if (attempt.Cells.Count >= 14) return attempt;
                        if (best == null || attempt.Cells.Count > best.Cells.Count) best = attempt;
                    }
                    else if (best == null) best = attempt;
                }
                if (best != null && best.Cells.Count > 0) { best.Success = true; return best; }
                return best ?? new CvDetectionResult { Success = false, FailReason = "Ningún pass detectó rejilla." };
            }
            finally { mat.Dispose(); gray.Dispose(); }
        }
        private static CvDetectionResult TryDetect(Mat gray, DetectConfig cfg)
        {
            var result = new CvDetectionResult();
            Mat bin = new Mat(); Mat horiz = null; Mat vert = null; Mat inter = null;
            try
            {
                Cv2.AdaptiveThreshold(gray, bin, 255, AdaptiveThresholdTypes.GaussianC,
                    ThresholdTypes.BinaryInv, cfg.BlockSize, cfg.C);

                // ROI hint via Canny+contornos
                OpenCvSharp.Rect? roiHint = GetTableRegion(gray);
                Mat workBin = (roiHint.HasValue &&
                               roiHint.Value.Width > bin.Width / 4 &&
                               roiHint.Value.Height > bin.Height / 6)
                    ? new Mat(bin, roiHint.Value) : bin;

                horiz = workBin.Clone();
                Mat kH = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(cfg.KernelW, 1));
                Cv2.Erode(horiz, horiz, kH); Cv2.Dilate(horiz, horiz, kH);

                vert = workBin.Clone();
                Mat kV = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(1, cfg.KernelH));
                Cv2.Erode(vert, vert, kV); Cv2.Dilate(vert, vert, kV);

                inter = new Mat();
                Cv2.BitwiseAnd(horiz, vert, inter);

                Mat nzm = new Mat();
                Cv2.FindNonZero(inter, nzm);
                OpenCvSharp.Point[] pts = null;
                if (nzm.Rows > 0)
                {
                    pts = new OpenCvSharp.Point[nzm.Rows];
                    for (int i = 0; i < nzm.Rows; i++) pts[i] = nzm.At<OpenCvSharp.Point>(i, 0);
                }
                nzm.Dispose();

                if (pts == null || pts.Length < cfg.MinIntersect)
                {
                    result.FailReason = string.Format("[{0}] Intersecciones: {1}<{2}",
                        cfg.Label, pts == null ? 0 : pts.Length, cfg.MinIntersect);
                    return result;
                }

                // Offset ROI → coordenadas absolutas
                int offX = roiHint.HasValue ? roiHint.Value.X : 0;
                int offY = roiHint.HasValue ? roiHint.Value.Y : 0;

                OpenCvSharp.Rect tr = Cv2.BoundingRect(pts);
                tr = new OpenCvSharp.Rect(tr.X + offX, tr.Y + offY, tr.Width, tr.Height);
                result.TableBounds = new Rectangle(tr.X, tr.Y, tr.Width, tr.Height);

                // HoughLinesP — coords relativas a workBin (ROI), luego añadir offset
                double minH = tr.Width * cfg.HoughMinLen;
                double minV = tr.Height * cfg.HoughMinLen;

                // Recortar horiz y vert a tableRect relativo a workBin
                OpenCvSharp.Rect trLocal = new OpenCvSharp.Rect(tr.X - offX, tr.Y - offY, tr.Width, tr.Height);

                int wbW = workBin.Width, wbH = workBin.Height;
                trLocal = new OpenCvSharp.Rect(
                    Math.Max(0, trLocal.X), Math.Max(0, trLocal.Y),
                    Math.Min(trLocal.Width, wbW - Math.Max(0, trLocal.X)),
                    Math.Min(trLocal.Height, wbH - Math.Max(0, trLocal.Y)));

                if (trLocal.Width < 10 || trLocal.Height < 10)
                {
                    result.FailReason = string.Format("[{0}] ROI local inválida", cfg.Label);
                    return result;
                }

                LineSegmentPoint[] hLines = Cv2.HoughLinesP(horiz[trLocal], 1, Math.PI / 180, 40, minH, cfg.HoughGap);
                LineSegmentPoint[] vLines = Cv2.HoughLinesP(vert[trLocal], 1, Math.PI / 180, 40, minV, cfg.HoughGap);

                if (hLines == null || hLines.Length < 2 || vLines == null || vLines.Length < 2)
                {
                    result.FailReason = string.Format("[{0}] Líneas H={1} V={2}",
                        cfg.Label, hLines == null ? 0 : hLines.Length, vLines == null ? 0 : vLines.Length);
                    return result;
                }

                Array.Sort(hLines, delegate (LineSegmentPoint a, LineSegmentPoint b2) { return a.P1.Y.CompareTo(b2.P1.Y); });
                Array.Sort(vLines, delegate (LineSegmentPoint a, LineSegmentPoint b2) { return a.P1.X.CompareTo(b2.P1.X); });
                hLines = MergeLines(hLines, true, cfg.MergeTol);
                vLines = MergeLines(vLines, false, cfg.MergeTol);

                // Construir celdas — compensar offset absoluto (trLocal coords → image coords)
                int absOffX = tr.X;
                int absOffY = tr.Y;

                for (int r = 0; r < hLines.Length - 1; r++)
                    for (int c = 0; c < vLines.Length - 1; c++)
                    {
                        var cell = Rectangle.FromLTRB(
                            vLines[c].P1.X + absOffX,
                            hLines[r].P1.Y + absOffY,
                            vLines[c + 1].P1.X + absOffX,
                            hLines[r + 1].P1.Y + absOffY);

                        if (cell.Width < 10 || cell.Height < 8) continue;

                        result.Cells.Add(new CvTableCell
                        {
                            Row = r,
                            Col = c,
                            Bounds = cell,
                            Field = MapColumnToField(c)
                        });
                    }

                result.RowCount = hLines.Length - 1;
                result.ColCount = vLines.Length - 1;
                result.Success = result.Cells.Count > 0;
                if (!result.Success)
                    result.FailReason = string.Format("[{0}] Sin celdas válidas", cfg.Label);
                return result;
            }
            finally
            {
                bin.Dispose();
                if (horiz != null) horiz.Dispose();
                if (vert != null) vert.Dispose();
                if (inter != null) inter.Dispose();
            }
        }
        private static OpenCvSharp.Rect? GetTableRegion(Mat gray)
        {
            Mat edges = new Mat();
            try
            {
                Cv2.Canny(gray, edges, 50, 150);
                Mat k = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(3, 3));
                Cv2.Dilate(edges, edges, k);

                OpenCvSharp.Point[][] contours;
                HierarchyIndex[] hierarchy;
                Cv2.FindContours(edges, out contours, out hierarchy,
                    RetrievalModes.External, ContourApproximationModes.ApproxSimple);

                OpenCvSharp.Rect? best = null;
                foreach (var cnt in contours)
                {
                    var r = Cv2.BoundingRect(cnt);
                    if (r.Width * r.Height > 5000 && r.Width > r.Height)
                        if (!best.HasValue || r.Width * r.Height > best.Value.Width * best.Value.Height)
                            best = r;
                }
                return best;
            }
            finally { edges.Dispose(); }
        }
        private static LineSegmentPoint[] MergeLines(LineSegmentPoint[] lines, bool horizontal, int tol)
        {
            if (lines == null || lines.Length == 0) return lines;
            var groups = new List<List<LineSegmentPoint>>();
            var current = new List<LineSegmentPoint> { lines[0] };
            for (int i = 1; i < lines.Length; i++)
            {
                int pos = horizontal ? lines[i].P1.Y : lines[i].P1.X;
                int prevPos = horizontal ? current[current.Count - 1].P1.Y : current[current.Count - 1].P1.X;
                if (Math.Abs(pos - prevPos) <= tol) current.Add(lines[i]);
                else { groups.Add(current); current = new List<LineSegmentPoint> { lines[i] }; }
            }
            groups.Add(current);
            var merged = new LineSegmentPoint[groups.Count];
            for (int g = 0; g < groups.Count; g++)
            {
                var grp = groups[g];
                LineSegmentPoint best2 = grp[0]; int bestLen = 0;
                foreach (var ln in grp)
                {
                    int dx = ln.P2.X - ln.P1.X, dy = ln.P2.Y - ln.P1.Y, len = dx * dx + dy * dy;
                    if (len > bestLen) { bestLen = len; best2 = ln; }
                }
                int med = horizontal ? grp[grp.Count / 2].P1.Y : grp[grp.Count / 2].P1.X;
                OpenCvSharp.Point p1 = best2.P1, p2 = best2.P2;
                if (horizontal) { p1.Y = med; p2.Y = med; } else { p1.X = med; p2.X = med; }
                merged[g] = new LineSegmentPoint(p1, p2);
            }
            return merged;
        }
        private static string MapColumnToField(int col)
        {
            switch (col)
            {
                case 0: return "Illuminant";
                case 1: return "Type";
                case 2: return "L";
                case 3: return "a";
                case 4: return "b";
                case 5: return "Chroma";
                case 6: return "Hue";
                case 7: return "DeltaL";
                case 8: return "DeltaC";
                case 9: return "DeltaH";
                case 10: return "DeltaE";
                default: return "Other";
            }
        }
        private static Rectangle ScaleRect(Rectangle r, float s)
        {
            return new Rectangle((int)(r.X * s), (int)(r.Y * s), (int)(r.Width * s), (int)(r.Height * s));
        }
    }
    // ══════════════════════════════════════════════════════════════════════════
    // CV TABLE PARSER — Construye OcrReport desde diccionario de textos por celda
    // ══════════════════════════════════════════════════════════════════════════
    internal static class ColrTableParser
    {
        public static OcrReport BuildReport(
            Dictionary<int, Dictionary<int, string>> cellTexts,
            int dataStartRow, List<string> log)
        {
            var report = new OcrReport();
            int maxRow = 0;
            foreach (var kvp in cellTexts) if (kvp.Key > maxRow) maxRow = kvp.Key;

            string lastIlluminant = string.Empty;

            for (int r = dataStartRow; r <= maxRow; r++)
            {
                Dictionary<int, string> row;
                if (!cellTexts.TryGetValue(r, out row)) continue;

                // Col 0: Illuminant
                string illRaw = GetCell(row, 0);
                string illuminant = NormalizeIlluminant(illRaw);
                if (!string.IsNullOrWhiteSpace(illuminant)) lastIlluminant = illuminant;
                else if (!string.IsNullOrWhiteSpace(lastIlluminant)) illuminant = lastIlluminant;
                else { if (log != null) log.Add(string.Format("[PARSER] Fila {0}: sin iluminante", r)); continue; }

                // Col 1: Type
                string typeRaw = GetCell(row, 1);
                string type = NormalizeType(typeRaw);
                if (string.IsNullOrWhiteSpace(type))
                {
                    string combined = NormUpper(illRaw);
                    if (combined.Contains("STD")) type = "Std";
                    else if (combined.Contains("LOT")) type = "Lot";
                }
                if (string.IsNullOrWhiteSpace(type))
                { if (log != null) log.Add(string.Format("[PARSER] Fila {0}: tipo desconocido '{1}'", r, typeRaw)); continue; }

                // Cols 2-6: L a b Chroma Hue
                double vL = SafeParse(GetCell(row, 2)), vA = SafeParse(GetCell(row, 3));
                double vB = SafeParse(GetCell(row, 4)), vC = SafeParse(GetCell(row, 5));
                double vH = SafeParse(GetCell(row, 6));

                bool valid = ColorimetryRanges.IsValidL(vL) && ColorimetryRanges.IsValidAB(vA) &&
                             ColorimetryRanges.IsValidAB(vB) && ColorimetryRanges.IsValidChroma(vC) &&
                             ColorimetryRanges.IsValidHue(vH);

                // Recalcular Hue si el OCR lo tiene mal
                double hCalc = Math.Atan2(vB, vA) * 180.0 / Math.PI;
                if (hCalc < 0) hCalc += 360.0;
                double hErr = Math.Abs(vH - hCalc); if (hErr > 180) hErr = 360 - hErr;
                if (vH < 1.0 || hErr > 10.0) vH = Math.Round(hCalc, 2);

                report.Measures.Add(new ColorimetricRow
                {
                    Illuminant = illuminant,
                    Type = type,
                    L = vL,
                    A = vA,
                    B = vB,
                    Chroma = vC,
                    Hue = vH,
                    NeedsReview = !valid
                });

                if (log != null) log.Add(string.Format(
                    "[PARSER] {0}/{1} L={2:F2} a={3:F2} b={4:F2} C={5:F2} h={6:F2}{7}",
                    illuminant, type, vL, vA, vB, vC, vH, valid ? "" : " [REVISAR]"));

                // Cols 7-10: Deltas CMC
                string dLr = GetCell(row, 7), dCr = GetCell(row, 8), dHr = GetCell(row, 9), dEr = GetCell(row, 10);
                bool hasDelta = !string.IsNullOrWhiteSpace(dLr) || !string.IsNullOrWhiteSpace(dCr) || !string.IsNullOrWhiteSpace(dHr);
                if (hasDelta)
                {
                    bool already = false;
                    foreach (var ex2 in report.CmcDifferences)
                        if (string.Equals(ex2.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase)) { already = true; break; }
                    if (!already)
                    {
                        double dL = SafeParse(dLr), dC = SafeParse(dCr), dH = SafeParse(dHr);
                        double dE = string.IsNullOrWhiteSpace(dEr) ? double.NaN : SafeParse(dEr);
                        report.CmcDifferences.Add(new CmcDifferenceRow
                        {
                            Illuminant = illuminant,
                            DeltaLightness = dL,
                            DeltaChroma = dC,
                            DeltaHue = dH,
                            DeltaCMC = double.IsNaN(dE) ? (double?)null : dE,
                            NeedsReview = !ColorimetryRanges.IsValidDL(dL) || !ColorimetryRanges.IsValidDC(dC) || !ColorimetryRanges.IsValidDH(dH)
                        });
                    }
                }
            }
            return report;
        }
        private static string GetCell(Dictionary<int, string> row, int col)
        { if (col < 0) return string.Empty; string v; return row.TryGetValue(col, out v) ? (v ?? string.Empty) : string.Empty; }

        private static string NormUpper(string s)
        { return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim().ToUpperInvariant(); }
        private static string NormalizeIlluminant(string raw)
        {
            string u = NormUpper(raw);
            foreach (string known in ColorimetricDataExtractor.KnownIlluminants)
                if (u == known.ToUpperInvariant() || u.Contains(known.ToUpperInvariant())) return known.ToUpperInvariant();
            return string.Empty;
        }

        private static string NormalizeType(string raw)
        {
            string u = NormUpper(raw);
            if (u.Contains("STD") || u.Contains("5TD") || u.Contains("ST0")) return "Std";
            if (u.Contains("LOT") || u.Contains("L0T")) return "Lot";
            return string.Empty;
        }

        private static double SafeParse(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0.0;
            s = s.Trim().Replace(',', '.');
            double v; return double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out v) ? v : 0.0;
        }
    }
    // ══════════════════════════════════════════════════════════════════════════
    // LOCAL CORRECTOR — Portado de ClaudeService (ColorimetriaAPI)
    // ══════════════════════════════════════════════════════════════════════════
    public class LocalCorrectionResult
    {
        public string Field { get; set; }
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double OriginalValue { get; set; }
        public double CorrectedValue { get; set; }
        public double OriginalCoherenceError { get; set; }
        public double NewCoherenceError { get; set; }
        public string Reason { get; set; }
    }
    internal static class LocalCorrector
    {
        private const double IMPROVEMENT_THRESHOLD = 0.3;
        private const double CHROMA_THRESHOLD = 0.35;

        // ── Punto de entrada ─────────────────────────────────────────────────
        public static List<LocalCorrectionResult> CorrectReport(OcrReport report)
        {
            var results = new List<LocalCorrectionResult>();
            if (report == null || report.Measures == null || report.Measures.Count == 0)
                return results;

            // Trabajar sobre copia para no interferir con los valores ya corregidos
            foreach (var row in report.Measures)
            {
                // Detectar error de Chroma
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                double chromaErr = Math.Abs(chromaCalc - row.Chroma);

                if (chromaErr > CHROMA_THRESHOLD)
                {
                    // Intentar corregir b* primero
                    var corrB = TryCorrectField("b", row.B, row.A, row.B, row.Chroma, chromaErr);
                    if (corrB != null)
                    {
                        corrB.Illuminant = row.Illuminant;
                        corrB.Type = row.Type;
                        ApplyToRow(row, corrB);
                        results.Add(corrB);
                        // Recalcular error tras corrección de b*
                        chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                        chromaErr = Math.Abs(chromaCalc - row.Chroma);
                    }

                    // Si sigue con error, intentar a*
                    if (chromaErr > CHROMA_THRESHOLD)
                    {
                        var corrA = TryCorrectField("a", row.A, row.A, row.B, row.Chroma, chromaErr);
                        if (corrA != null)
                        {
                            corrA.Illuminant = row.Illuminant;
                            corrA.Type = row.Type;
                            ApplyToRow(row, corrA);
                            results.Add(corrA);
                        }
                    }
                }

                // Detectar error de Hue
                double hueCalc = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                double hueErr = Math.Abs(row.Hue - hueCalc);
                if (hueErr > 180) hueErr = 360 - hueErr;

                if (hueErr > 5.0)
                {
                    double correctedHue = Math.Round(hueCalc, 2);
                    results.Add(new LocalCorrectionResult
                    {
                        Field = "Hue",
                        Illuminant = row.Illuminant,
                        Type = row.Type,
                        OriginalValue = row.Hue,
                        CorrectedValue = correctedHue,
                        OriginalCoherenceError = hueErr,
                        NewCoherenceError = 0,
                        Reason = "Recalculado desde atan2(b,a)"
                    });
                    row.Hue = correctedHue;
                    row.NeedsReview = false;
                }
            }

            return results;
        }

        // ── Corrección de un campo ───────────────────────────────────────────

        private static LocalCorrectionResult TryCorrectField(
            string field, double ocrValue, double a, double b, double chroma, double errOrig)
        {
            double bestErr = errOrig;
            double bestVal = ocrValue;
            string bestReason = null;

            // Candidatos: desplazamientos de punto decimal + cambio de signo
            var candidates = new List<KeyValuePair<double, string>>
            {
                new KeyValuePair<double, string>(ocrValue / 10.0,   "Punto decimal ÷10"),
                new KeyValuePair<double, string>(ocrValue * 10.0,   "Punto decimal ×10"),
                new KeyValuePair<double, string>(-ocrValue,         "Signo invertido"),
                new KeyValuePair<double, string>(-ocrValue / 10.0,  "Signo invertido + ÷10"),
                new KeyValuePair<double, string>(-ocrValue * 10.0,  "Signo invertido + ×10"),
            };

            // Resolución directa desde Chroma = sqrt(a²+b²)
            double chroma2 = chroma * chroma;
            if (field == "b")
            {
                double a2 = a * a;
                if (chroma2 >= a2)
                {
                    double bSolved = Math.Sqrt(chroma2 - a2);
                    candidates.Add(new KeyValuePair<double, string>(bSolved, "Resuelto sqrt(C²-a²)"));
                    candidates.Add(new KeyValuePair<double, string>(-bSolved, "Resuelto -sqrt(C²-a²)"));
                }
            }
            else if (field == "a")
            {
                double b2 = b * b;
                if (chroma2 >= b2)
                {
                    double aSolved = Math.Sqrt(chroma2 - b2);
                    candidates.Add(new KeyValuePair<double, string>(aSolved, "Resuelto sqrt(C²-b²)"));
                    candidates.Add(new KeyValuePair<double, string>(-aSolved, "Resuelto -sqrt(C²-b²)"));
                }
            }

            foreach (var kv in candidates)
            {
                double cand = kv.Key;
                string reason = kv.Value;

                if (!IsPhysicallyValid(field, cand)) continue;

                double newA = field == "a" ? cand : a;
                double newB = field == "b" ? cand : b;
                double newErr = Math.Abs(Math.Sqrt(newA * newA + newB * newB) - chroma);

                if (newErr < bestErr - IMPROVEMENT_THRESHOLD)
                {
                    bestErr = newErr;
                    bestVal = cand;
                    bestReason = reason;
                }
            }

            if (bestReason == null) return null;

            return new LocalCorrectionResult
            {
                Field = field,
                OriginalValue = ocrValue,
                CorrectedValue = Math.Round(bestVal, 4),
                OriginalCoherenceError = errOrig,
                NewCoherenceError = Math.Round(bestErr, 4),
                Reason = bestReason
            };
        }

        private static bool IsPhysicallyValid(string field, double value)
        {
            if (field == "L" && (value < 0 || value > 100)) return false;
            if ((field == "a" || field == "b") && Math.Abs(value) > 150) return false;
            if (field == "Chroma" && (value < 0 || value > 200)) return false;
            if (field == "Hue" && (value < 0 || value > 360)) return false;
            return true;
        }
        private static void ApplyToRow(ColorimetricRow row, LocalCorrectionResult c)
        {
            if (c == null) return;

            switch (c.Field.ToLower())
            {
                case "l": row.L = c.CorrectedValue; break;
                case "a": row.A = c.CorrectedValue; break;
                case "b": row.B = c.CorrectedValue; break;
                case "chroma": row.Chroma = c.CorrectedValue; break;
                case "hue": row.Hue = (double)c.CorrectedValue; break;
            }
        }

    }
}