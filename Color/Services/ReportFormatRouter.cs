using System;
using System.Drawing;
using System.IO;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace Color.Services
{
    // ─────────────────────────────────────────────────────────────────────────
    // Clasificación geométrica del formato de imagen de reporte
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Tipo de formato detectado por análisis estructural de líneas.
    /// </summary>
    public enum ReportFormatType
    {
        /// <summary>
        /// Reporte completo: encabezado + receta + tabla CMC inferior.
        /// </summary>
        LegacyCombinedFormat,

        /// <summary>
        /// Ticket plano de receta sin tabla CMC (solo matriz de puntos).
        /// </summary>
        DynamicSplitGridFormat,

        /// <summary>
        /// No se pudo determinar el formato; se aplica la ruta de ticket plano.
        /// </summary>
        UnknownFallback
    }

    /// <summary>
    /// Analizador estructural autónomo que clasifica el tipo de imagen
    /// utilizando detección morfológica de líneas horizontales (OpenCV).
    /// Opera sin referencias a palabras clave del contenido textual.
    /// </summary>
    public static class ReportFormatRouter
    {
        // ── Configuración de detección ────────────────────────────────────────

        /// <summary>
        /// Fracción mínima del ancho de imagen que debe tener una línea
        /// para considerarse estructural (25 %).
        /// </summary>
        private const double MIN_LINE_WIDTH_FRACTION = 0.25;

        /// <summary>
        /// Número mínimo de líneas horizontales estructurales que deben
        /// detectarse en la mitad inferior para clasificar como LegacyCombinedFormat.
        /// </summary>
        private const int STRUCTURAL_LINES_THRESHOLD = 4;

        // ── API pública ───────────────────────────────────────────────────────

        /// <summary>
        /// Determina el formato del reporte mediante análisis geométrico de líneas.
        /// </summary>
        /// <param name="imagePath">Ruta absoluta a la imagen PNG.</param>
        /// <returns>El tipo de formato detectado.</returns>
        public static ReportFormatType DetermineFormat(string imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                return ReportFormatType.UnknownFallback;

            try
            {
                using (Bitmap bmp = LoadUniversalImage24bpp(imagePath))
                using (Mat src = BitmapConverter.ToMat(bmp))
                using (Mat gray = new Mat())
                using (Mat binary = new Mat())
                {
                    // 1. Convertir a escala de grises
                    Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);

                    // 2. Umbral adaptativo para binarizar el texto/líneas
                    Cv2.AdaptiveThreshold(gray, binary, 255,
                        AdaptiveThresholdTypes.MeanC,
                        ThresholdTypes.BinaryInv,
                        blockSize: 15,
                        c: 4);

                    // 3. Analizar SOLO la mitad inferior (donde aparece la tabla CMC)
                    int lowerY   = src.Height / 2;
                    int lowerH   = src.Height - lowerY;
                    var lowerRoi = new OpenCvSharp.Rect(0, lowerY, src.Width, lowerH);

                    using (Mat lower = new Mat(binary, lowerRoi))
                    {
                        // 4. Detectar líneas con Probabilistic Hough Transform
                        int minLineLength = (int)(src.Width * MIN_LINE_WIDTH_FRACTION);

                        LineSegmentPoint[] lines = Cv2.HoughLinesP(
                            lower,
                            rho:          1,
                            theta:        Cv2.PI / 180,
                            threshold:    50,
                            minLineLength: minLineLength,
                            maxLineGap:   10);

                        // 5. Contar únicamente las líneas casi horizontales (Δy ≤ 3 px)
                        int horizontalCount = 0;
                        foreach (var line in lines)
                        {
                            if (Math.Abs(line.P1.Y - line.P2.Y) <= 3)
                                horizontalCount++;
                        }

                        // 6. Si hay suficientes líneas estructurales → formato combinado completo
                        if (horizontalCount >= STRUCTURAL_LINES_THRESHOLD)
                            return ReportFormatType.LegacyCombinedFormat;
                    }
                }
            }
            catch
            {
                // Ante cualquier error de procesamiento, usar la ruta más segura
                return ReportFormatType.UnknownFallback;
            }

            return ReportFormatType.DynamicSplitGridFormat;
        }

        // ── Utilidad de carga ────────────────────────────────────────────────

        /// <summary>
        /// Carga una imagen en formato 24 bpp RGB con fondo blanco,
        /// resolviendo transparencias y formatos de color no estándar.
        /// </summary>
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
