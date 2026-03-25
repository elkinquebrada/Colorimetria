using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace Color
{
    /// Celda detectada en la tabla CMC con su posición y etiquetas.
    public class DetectedCell
    {
        public Rectangle Bounds { get; set; }
        public int Row { get; set; }  // índice de fila en la tabla (0 = encabezado)
        public int Col { get; set; }  // índice de columna
        public string Illuminant { get; set; }  // "D65", "TL84", etc. (null si no determinado)
        public string RowType { get; set; }  // "Std" | "Lot" | "Header" | null
        public string FieldName { get; set; }  // "L", "A", "B", "Chroma", "Hue", "dL"...
    }

    /// Resultado de la detección de tabla.
    public class TableDetectionResult
    {
        public bool Success { get; set; }
        public List<DetectedCell> Cells { get; set; } = new List<DetectedCell>();
        public Rectangle TableBounds { get; set; }
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public string FailReason { get; set; }
    }

    /// Detecta la estructura de la tabla CMC en una imagen usando OpenCV.
    /// Estrategia: morfología + detección de líneas + clustering de intersecciones.
    public static class OpenCvTableDetector
    {
        // ── Constantes de configuración ───────────────────────────────────────
        private const int MIN_H_LINE_LENGTH_PCT = 20;  // % del ancho mínimo para línea H
        private const int MIN_V_LINE_LENGTH_PCT = 5;   // % del alto mínimo para línea V
        private const int LINE_GAP = 8;   // tolerancia de gap en HoughLinesP
        private const int LINE_THICKNESS = 2;   // grosor para dilatar líneas
        private const int CELL_PADDING = 3;   // padding dentro de cada celda para OCR
        private const int MIN_ROWS = 4;   // mínimo filas esperadas (D65+TL84+A = 6 filas de datos)
        private const int MIN_COLS = 5;   // mínimo columnas (L,a,b,C,H)
        private const int CLUSTER_TOLERANCE = 12;  // px para agrupar líneas paralelas cercanas
        private const double MIN_TABLE_WIDTH_PCT = 0.4; // la tabla debe ocupar al menos 40% del ancho

        // Nombres de campos por índice de columna (0-based, columna 0 = iluminante)
        private static readonly string[] FIELD_NAMES =
            { "Illuminant", "Type", "L", "A", "B", "Chroma", "Hue", "dL", "dC", "dH", "CMC" };

        // ── API pública ───────────────────────────────────────────────────────

        /// Detecta la tabla CMC en la imagen y retorna las celdas con sus coordenadas.
        public static TableDetectionResult Detect(Bitmap original)
        {
            var result = new TableDetectionResult();

            try
            {
                using (var src = BitmapConverter.ToMat(original))
                {
                    // 1. Preprocesar
                    using (var gray = Preprocess(src))
                    using (var hLines = new Mat())
                    using (var vLines = new Mat())
                    {
                        // 2. Extraer líneas horizontales y verticales por morfología
                        ExtractLines(gray, hLines, vLines, src.Width, src.Height);

                        // 3. Encontrar posiciones Y de líneas H y X de líneas V
                        var yPositions = FindLinePositions(hLines, true, src.Height, CLUSTER_TOLERANCE);
                        var xPositions = FindLinePositions(vLines, false, src.Width, CLUSTER_TOLERANCE);

                        if (yPositions.Count < MIN_ROWS + 1 || xPositions.Count < MIN_COLS + 1)
                        {
                            result.FailReason = string.Format(
                                "Pocas líneas detectadas: {0} H, {1} V (mínimo {2}H, {3}V)",
                                yPositions.Count, xPositions.Count, MIN_ROWS + 1, MIN_COLS + 1);
                            return result;
                        }

                        // 4. Verificar que la tabla ocupa área razonable
                        int tableWidth = xPositions[xPositions.Count - 1] - xPositions[0];
                        if (tableWidth < src.Width * MIN_TABLE_WIDTH_PCT)
                        {
                            result.FailReason = string.Format(
                                "Tabla demasiado estrecha: {0}px ({1}% del ancho)",
                                tableWidth, (int)(tableWidth * 100.0 / src.Width));
                            return result;
                        }

                        // 5. Construir celdas desde la cuadrícula
                        result.RowCount = yPositions.Count - 1;
                        result.ColCount = xPositions.Count - 1;
                        result.TableBounds = new Rectangle(
                            xPositions[0], yPositions[0],
                            tableWidth,
                            yPositions[yPositions.Count - 1] - yPositions[0]);

                        for (int r = 0; r < result.RowCount; r++)
                        {
                            for (int c = 0; c < result.ColCount; c++)
                            {
                                int x = xPositions[c] + CELL_PADDING;
                                int y = yPositions[r] + CELL_PADDING;
                                int w = xPositions[c + 1] - xPositions[c] - CELL_PADDING * 2;
                                int h = yPositions[r + 1] - yPositions[r] - CELL_PADDING * 2;

                                if (w < 4 || h < 4) continue;

                                // Asegurar que el rectángulo está dentro de la imagen
                                var bounds = Rectangle.Intersect(
                                    new Rectangle(x, y, w, h),
                                    new Rectangle(0, 0, original.Width, original.Height));

                                if (bounds.IsEmpty) continue;

                                result.Cells.Add(new DetectedCell
                                {
                                    Bounds = bounds,
                                    Row = r,
                                    Col = c,
                                    FieldName = c < FIELD_NAMES.Length ? FIELD_NAMES[c] : string.Format("Col{0}", c)
                                });
                            }
                        }

                        result.Success = result.Cells.Count > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                result.FailReason = "OpenCV error: " + ex.Message;
            }

            return result;
        }

        // ── Preprocesamiento ──────────────────────────────────────────────────

        private static Mat Preprocess(Mat src)
        {
            var gray = new Mat();

            // Convertir a gris si es color
            if (src.Channels() > 1)
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
            else
                src.CopyTo(gray);

            // Invertir si el fondo es oscuro (texto blanco sobre fondo oscuro)
            double meanVal = Cv2.Mean(gray).Val0;
            if (meanVal < 128)
                Cv2.BitwiseNot(gray, gray);

            // Suavizar levemente para eliminar ruido
            Cv2.GaussianBlur(gray, gray, new OpenCvSharp.Size(3, 3), 0);

            return gray;
        }

        // ── Extracción de líneas por morfología ───────────────────────────────

        private static void ExtractLines(Mat gray, Mat hDst, Mat vDst, int width, int height)
        {
            // Umbral adaptativo para binarizar
            var binary = new Mat();
            Cv2.AdaptiveThreshold(gray, binary, 255,
                AdaptiveThresholdTypes.MeanC, ThresholdTypes.BinaryInv, 15, 5);

            // Kernel horizontal: ancho = 20% del ancho de imagen
            int hKernelW = Math.Max(20, width / MIN_H_LINE_LENGTH_PCT);
            using (var hKernel = Cv2.GetStructuringElement(
                MorphShapes.Rect, new OpenCvSharp.Size(hKernelW, 1)))
            {
                Cv2.Erode(binary, hDst, hKernel);
                Cv2.Dilate(hDst, hDst, hKernel);
            }

            // Kernel vertical: alto = 5% del alto de imagen
            int vKernelH = Math.Max(10, height / MIN_V_LINE_LENGTH_PCT);
            using (var vKernel = Cv2.GetStructuringElement(
                MorphShapes.Rect, new OpenCvSharp.Size(1, vKernelH)))
            {
                Cv2.Erode(binary, vDst, vKernel);
                Cv2.Dilate(vDst, vDst, vKernel);
            }

            binary.Dispose();
        }

        // ── Encontrar posiciones de líneas ────────────────────────────────────

        private static List<int> FindLinePositions(Mat lineMat, bool horizontal, int size, int tolerance)
        {
            var positions = new List<int>();

            int rows = lineMat.Rows;
            int cols = lineMat.Cols;

            // Calcular proyección manualmente: sumar píxeles por fila o por columna
            // Convertir Mat → Bitmap para lectura via LockBits (compatible .NET 4.8)
            int threshold = (int)(size * 0.20 * 255);

            if (horizontal)
            {
                // Proyección Y: para cada fila, sumar todos los píxeles
                for (int r = 0; r < rows; r++)
                {
                    long rowSum = 0;
                    for (int c = 0; c < cols; c++)
                        rowSum += lineMat.At<byte>(r, c);

                    if (rowSum > threshold)
                        positions.Add(r);
                }
            }
            else
            {
                // Proyección X: para cada columna, sumar todos los píxeles
                for (int c = 0; c < cols; c++)
                {
                    long colSum = 0;
                    for (int r = 0; r < rows; r++)
                        colSum += lineMat.At<byte>(r, c);

                    if (colSum > threshold)
                        positions.Add(c);
                }
            }

            // Agrupar posiciones cercanas (cluster) → tomar el centro de cada grupo
            return ClusterPositions(positions, tolerance);
        }

        private static List<int> ClusterPositions(List<int> positions, int tolerance)
        {
            if (positions.Count == 0) return positions;

            var clusters = new List<List<int>>();
            var current = new List<int> { positions[0] };

            for (int i = 1; i < positions.Count; i++)
            {
                if (positions[i] - positions[i - 1] <= tolerance)
                    current.Add(positions[i]);
                else
                {
                    clusters.Add(current);
                    current = new List<int> { positions[i] };
                }
            }
            clusters.Add(current);

            // Retornar el valor medio de cada cluster
            var result = new List<int>(clusters.Count);
            foreach (var cluster in clusters)
                result.Add((int)cluster.Average());

            return result;
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // PARSER DE TABLA DETECTADA
    // ══════════════════════════════════════════════════════════════════════════

    internal static class TableParser
    {
        private const double SAFE_PARSE_DEFAULT = double.NaN;

        /// Construye un OcrReport a partir de un diccionario (fila,col) → texto OCR de celda.
   
        public static OcrReport BuildReport(
            Dictionary<int, Dictionary<int, string>> cellTexts,
            int dataStartRow,
            List<string> log)
        {
            var report = new OcrReport();

            // Agrupar filas de datos por iluminante
            // Estructura esperada: cada iluminante ocupa 2 filas (Std + Lot)
            // Col 0 = Illuminant (solo en fila Std), Col 1 = Type, Col 2-6 = L,a,b,C,H
            string currentIlluminant = null;
            ColorimetricRow stdRow = null;

            for (int r = dataStartRow; r < dataStartRow + 20; r++)
            {
                if (!cellTexts.ContainsKey(r)) break;
                var row = cellTexts[r];

                // Leer iluminante y tipo
                string illText = GetCell(row, 0);
                string typeText = GetCell(row, 1);

                string illuminant = NormalizeIlluminant(illText);
                string rowType = NormalizeType(typeText);

                // Si hay un iluminante en col 0, es fila Std
                if (!string.IsNullOrEmpty(illuminant))
                {
                    currentIlluminant = illuminant;
                    rowType = "Std";
                }
                else if (!string.IsNullOrEmpty(rowType))
                {
                    // rowType viene de col 1 (Lot/Std)
                }
                else
                {
                    // Determinar por posición: filas impares = Lot si no hay iluminante
                    if (currentIlluminant != null)
                        rowType = string.IsNullOrEmpty(rowType) ? "Lot" : rowType;
                }

                if (string.IsNullOrEmpty(currentIlluminant)) continue;
                if (string.IsNullOrEmpty(rowType)) continue;

                // Leer valores numéricos L, a, b, Chroma, Hue
                double vL = ParseCell(GetCell(row, 2));
                double vA = ParseCell(GetCell(row, 3));
                double vB = ParseCell(GetCell(row, 4));
                double vC = ParseCell(GetCell(row, 5));
                double vH = ParseCell(GetCell(row, 6));

                if (double.IsNaN(vL) || double.IsNaN(vA)) continue;
                if (!ColorimetryRanges.IsValidL(vL)) continue;
                if (!ColorimetryRanges.IsValidAB(vA)) continue;

                var measRow = new ColorimetricRow
                {
                    Illuminant = currentIlluminant,
                    Type = rowType,
                    L = double.IsNaN(vL) ? 0 : vL,
                    A = double.IsNaN(vA) ? 0 : vA,
                    B = double.IsNaN(vB) ? 0 : vB,
                    Chroma = double.IsNaN(vC) ? 0 : vC,
                    Hue = double.IsNaN(vH) ? 0 : vH,
                    NeedsReview = false
                };

                // Recalcular Hue si el OCR lo perdió
                if (measRow.Hue < 1.0)
                {
                    double hCalc = Math.Atan2(measRow.B, measRow.A) * 180.0 / Math.PI;
                    if (hCalc < 0) hCalc += 360.0;
                    measRow.Hue = Math.Round(hCalc);
                }

                if (log != null)
                    log.Add(string.Format("[OPENCV] {0}/{1} L={2:F2} a={3:F2} b={4:F2} C={5:F2} H={6:F0}",
                        measRow.Illuminant, measRow.Type,
                        measRow.L, measRow.A, measRow.B, measRow.Chroma, measRow.Hue));

                report.Measures.Add(measRow);

                // Leer deltas CMC si están presentes (col 7-10)
                double dL = ParseCell(GetCell(row, 7));
                double dC = ParseCell(GetCell(row, 8));
                double dH = ParseCell(GetCell(row, 9));
                double dE = ParseCell(GetCell(row, 10));

                if (rowType == "Std" && !double.IsNaN(dL) && !double.IsNaN(dC))
                {
                    report.CmcDifferences.Add(new CmcDifferenceRow
                    {
                        Illuminant = currentIlluminant,
                        DeltaLightness = dL,
                        DeltaChroma = dC,
                        DeltaHue = double.IsNaN(dH) ? 0 : dH,
                        DeltaCMC = double.IsNaN(dE) ? (double?)null : dE
                    });
                }

                if (rowType == "Std") stdRow = measRow;
                else if (rowType == "Lot" && stdRow != null)
                {
                    // Corregir signo usando Std como referencia
                    FixSignByStdSimple(stdRow, measRow);
                    stdRow = null;
                }
            }

            return report;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static string GetCell(Dictionary<int, string> row, int col)
        {
            string val;
            return row.TryGetValue(col, out val) ? (val ?? "").Trim() : "";
        }

        private static double ParseCell(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return double.NaN;

            // Limpiar caracteres OCR comunes
            text = text.Trim()
                .Replace(',', '.')
                .Replace('O', '0')
                .Replace('l', '1')
                .Replace('I', '1');

            // Quitar caracteres no numéricos excepto punto y signo
            var cleaned = System.Text.RegularExpressions.Regex.Replace(text, @"[^0-9\.\-\+]", "");
            if (string.IsNullOrEmpty(cleaned)) return double.NaN;

            double v;
            if (double.TryParse(cleaned,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out v))
                return v;

            return double.NaN;
        }

        private static string NormalizeIlluminant(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.ToUpperInvariant().Trim();

            // Aplicar correcciones básicas de OCR para iluminantes
            if (text.Contains("D65") || text.Contains("D6S") || text.Contains("DG5")) return "D65";
            if (text.Contains("D50") || text.Contains("DSO")) return "D50";
            if (text.Contains("TL84") || text.Contains("TL8A") || text.Contains("TLB4")) return "TL84";
            if (text.Contains("TL83") || text.Contains("TLS3")) return "TL83";
            if (text.Contains("CWF") || text.Contains("CVF") || text.Contains("GWF")) return "CWF";
            if (text == "A") return "A";
            if (text.Contains("F11")) return "F11";
            if (text.Contains("F12")) return "F12";
            if (text.Contains("UV")) return "UV";

            return null;
        }

        private static string NormalizeType(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.ToUpperInvariant().Trim();
            if (text.Contains("STD") || text.Contains("5TD")) return "Std";
            if (text.Contains("LOT") || text.Contains("L0T")) return "Lot";
            return null;
        }

        private static void FixSignByStdSimple(ColorimetricRow std, ColorimetricRow lot)
        {
            const double MIN_RELIABLE = 0.5;
            if (Math.Abs(std.A) >= MIN_RELIABLE && std.A * lot.A < 0) lot.A = -lot.A;
            if (Math.Abs(std.B) >= MIN_RELIABLE && std.B * lot.B < 0) lot.B = -lot.B;

            double newHue = Math.Atan2(lot.B, lot.A) * 180.0 / Math.PI;
            if (newHue < 0) newHue += 360.0;
            double diff = Math.Abs(newHue - lot.Hue);
            if (diff > 180) diff = 360 - diff;
            if (diff > 5) lot.Hue = Math.Round(newHue);
        }
    }
}