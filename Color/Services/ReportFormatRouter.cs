using System;
using System.Drawing;
using System.IO;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Tesseract;
namespace Color.Services
{
    // ─────────────────────────────────────────────────────────────────────────
    // Clasificación geométrica del formato de imagen de reporte
    // ─────────────────────────────────────────────────────────────────────────

    public enum ReportFormatType
    {
        LegacyCombinedFormat,
        DynamicSplitGridFormat,
        UnknownFallback
    }

    public static class ReportFormatRouter
    {
        // ── Configuración de detección ────────────────────────────────────────

        private const double MIN_LINE_WIDTH_FRACTION = 0.25;
        private const int STRUCTURAL_LINES_THRESHOLD = 4;

        // ── API pública ───────────────────────────────────────────────────────

        public static ReportFormatType DetermineFormat(string imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return ReportFormatType.UnknownFallback;

            string tessdataPath = @".\tessdata";
            try
            {
                using (Bitmap bmp = LoadUniversalImage24bpp(imagePath))
                using (Mat src = BitmapConverter.ToMat(bmp))
                {
                    // Recortar SOLO el 15% superior de la imagen para escaneo rápido
                    int topH = (int)(src.Height * 0.15);
                    var topRoi = new OpenCvSharp.Rect(0, 0, src.Width, topH);
                    
                    using (Mat topMat = new Mat(src, topRoi))
                    using (Mat gray = new Mat())
                    using (Mat scaled = new Mat())
                    {
                        Cv2.CvtColor(topMat, gray, ColorConversionCodes.BGR2GRAY);

                        // Escalar 2x para lectura rápida y nítida
                        Cv2.Resize(gray, scaled, new OpenCvSharp.Size(gray.Width * 2, gray.Height * 2), 0, 0, InterpolationFlags.Cubic);
                        
                        using (Bitmap topBmp = BitmapConverter.ToBitmap(scaled))
                        using (var engine = new TesseractEngine(tessdataPath, "eng+spa", EngineMode.Default))
                        using (var page = engine.Process(topBmp, PageSegMode.SingleBlock))
                        {
                            string ocrText = page.GetText()?.ToUpper() ?? "";
                            
                            // Evaluación determinística rápida de palabras clave exclusivas
                            if (ocrText.Contains("BULK") || ocrText.Contains("CHEESES") || ocrText.Contains("COL GROUP"))
                            {
                                return ReportFormatType.DynamicSplitGridFormat;
                            }
                            
                            if (ocrText.Contains("PASS / FAIL") || ocrText.Contains("SHADE HISTORY") || ocrText.Contains("EQUATION"))
                            {
                                return ReportFormatType.LegacyCombinedFormat;
                            }
                        }
                    }
                    
                    // Si el OCR no detecta palabras clave contundentes, aplicamos heurística geométrica de fallback (Hough Lines)
                    using (Mat gray = new Mat())
                    using (Mat binary = new Mat())
                    {
                        Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
                        Cv2.AdaptiveThreshold(gray, binary, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.BinaryInv, 15, 4);
                        
                        int lowerY = src.Height / 2;
                        int lowerH = src.Height - lowerY;
                        using (Mat lower = new Mat(binary, new OpenCvSharp.Rect(0, lowerY, src.Width, lowerH)))
                        {
                            int minLineLength = (int)(src.Width * MIN_LINE_WIDTH_FRACTION);
                            var lines = Cv2.HoughLinesP(lower, 1, Cv2.PI / 180, 50, minLineLength, 10);
                            int horiz = 0;
                            foreach (var line in lines) { if (Math.Abs(line.P1.Y - line.P2.Y) <= 3) horiz++; }
                            
                            if (horiz >= STRUCTURAL_LINES_THRESHOLD) return ReportFormatType.LegacyCombinedFormat;
                        }
                    }
                }
            }
            catch
            {
                return ReportFormatType.UnknownFallback;
            }

            return ReportFormatType.DynamicSplitGridFormat;
        }

        // ── Utilidad de carga ────────────────────────────────────────────────

        public static Bitmap LoadUniversalImage24bpp(string imagePath)
        {
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            using (var original = new Bitmap(stream))
            {
                var flat = new Bitmap(
                    original.Width,
                    original.Height,
                    System.Drawing.Imaging.PixelFormat.Format24bppRgb);

                flat.SetResolution(original.HorizontalResolution, original.VerticalResolution);

                using (Graphics g = Graphics.FromImage(flat))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(original, new System.Drawing.Rectangle(0, 0, flat.Width, flat.Height));
                }

                return flat;
            }
        }
    }
}
