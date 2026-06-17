using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using Tesseract;
using Color.Models;

namespace Color.Services
{
    /// Extractor aislado de metadatos texiles del encabezado del reporte de color.
    public class TextileMetadataExtractor
    {
        private readonly string _tessDataPath;
        private readonly object _engineLock = new object();

        // Umbrales heredados de ColorimetricDataExtractor (Dataextraxtor.cs)
        private const float SHARPNESS_LOW_THRESHOLD  = 40f;
        private const float SHARPNESS_HIGH_THRESHOLD = 120f;
        private const int   SCALE_FACTOR_MAX = 4;
        private const int   SCALE_FACTOR_MIN = 2;

        // Anchors de busqueda (tolerantes a OCR: lower-case, sin acentos)
        private static readonly string[] SHADE_KEYS   = { "shade name", "shade", "sombra" };
        private static readonly string[] CLASS_KEYS   = { "dyeing class", "clase", "class" };
        private static readonly string[] SUBST_KEYS   = { "substrate", "sustrato" };
        private static readonly string[] COUNT_KEYS   = { "count/ply", "count / ply", "count", "ply", "count/pl" };
        private static readonly string[] FIBER_KEYS   = { "fibre type", "fiber type", "fibra", "fibre" };

        public TextileMetadataExtractor(string tessDataPath = @".\tessdata")
        {
            _tessDataPath = tessDataPath;
        }

        // =====================================================================
        // API PUBLICA
        // =====================================================================
        /// Extrae los metadatos del encabezado superior izquierdo de la imagen.
        public TextileMetadata Extract(string imagePath)
        {
            var meta = new TextileMetadata();
            if (!File.Exists(imagePath)) return meta;

            try
            {
                using (var bmp = new Bitmap(imagePath))
                {
                    return ExtractFromBitmap(bmp);
                }
            }
            catch { return meta; }
        }

        /// Sobrecarga que acepta directamente un Bitmap en memoria
        public TextileMetadata ExtractFromBitmap(Bitmap original)
        {
            var meta = new TextileMetadata();
            if (original == null) return meta;

            try
            {
                // ── 1. ROI ampliado: 100% ancho × 50% alto (captura más filas de cabecera) ──
                int roiW = original.Width;
                int roiH = (int)(original.Height * 0.50);
                var roiRect = Rectangle.Intersect(
                    new Rectangle(0, 0, roiW, roiH),
                    new Rectangle(0, 0, original.Width, original.Height));

                if (roiRect.Width < 20 || roiRect.Height < 20) return meta;

                using (var roi = original.Clone(roiRect, original.PixelFormat))
                {

                    // ── 2. Escalado adaptativo basado en nitidez medida ─────────────────────
                    float sharpness = MeasureSharpness(roi);
                    int scaleFactor;
                    if      (sharpness < SHARPNESS_LOW_THRESHOLD)  scaleFactor = SCALE_FACTOR_MAX; 
                    else if (sharpness < SHARPNESS_HIGH_THRESHOLD) scaleFactor = 3;                
                    else                                            scaleFactor = SCALE_FACTOR_MIN; 

                    int nW = roi.Width  * scaleFactor;
                    int nH = roi.Height * scaleFactor;

                    using (var scaled = new Bitmap(nW, nH, PixelFormat.Format32bppArgb))
                    {
                        using (var g = Graphics.FromImage(scaled))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode     = SmoothingMode.HighQuality;
                            g.DrawImage(roi, 0, 0, nW, nH);
                        }

                        // ── 3. PASE 1: OCR estándar con segmentación Auto ──────────────────
                        string text1 = RunOcrWithMode(scaled, PageSegMode.Auto);
                        ParseKeyValue(text1, meta);

                        // ── 4. PASE 2: Binarización Otsu + SparseText si faltan campos ─────
                        if (HasMissingFields(meta))
                        {
                            using (var binarized = BinarizeOtsu(scaled))
                            {
                                string text2 = RunOcrWithMode(binarized, PageSegMode.SparseText);
                                ParseKeyValue(text2, meta);
                            }
                        }

                        // ── 5. PASE 3: SingleBlock como último recurso ─────────────────────
                        if (HasMissingFields(meta))
                        {
                            string text3 = RunOcrWithMode(scaled, PageSegMode.SingleBlock);
                            ParseKeyValue(text3, meta);
                        }

                        // ── 6. PASE 4: Binarizado + SingleBlock (imágenes muy degradadas) ──
                        if (HasMissingFields(meta))
                        {
                            using (var binarized = BinarizeOtsu(scaled))
                            {
                                string text4 = RunOcrWithMode(binarized, PageSegMode.SingleBlock);
                                ParseKeyValue(text4, meta);
                            }
                        }
                    }
                }
            }
            catch { /* Devolver meta parcial */ }

            return meta;
        }

        private bool HasMissingFields(TextileMetadata meta)
        {
            return (string.IsNullOrEmpty(meta.ShadeName)   || meta.ShadeName   == "-") ||
                   (string.IsNullOrEmpty(meta.DyeingClass) || meta.DyeingClass == "-") ||
                   (string.IsNullOrEmpty(meta.Substrate)   || meta.Substrate   == "-") ||
                   (string.IsNullOrEmpty(meta.CountPly)    || meta.CountPly    == "-") ||
                   (string.IsNullOrEmpty(meta.FiberType)   || meta.FiberType   == "-");
        }

        // =====================================================================
        // OCR INTERNO — MULTI-MODO
        // =====================================================================

        private string RunOcrWithMode(Bitmap bmp, PageSegMode mode)
        {
            string tmpFile = Path.Combine(Path.GetTempPath(),
                $"textile_hdr_{Guid.NewGuid():N}.png");
            try
            {
                bmp.Save(tmpFile, System.Drawing.Imaging.ImageFormat.Png);

                lock (_engineLock)
                {
                    using (var engine = new TesseractEngine(_tessDataPath, "eng", EngineMode.Default))
                    {
                        engine.DefaultPageSegMode = mode;

                        using (var pix  = Pix.LoadFromFile(tmpFile))
                        using (var page = engine.Process(pix))
                        {
                            return page.GetText() ?? string.Empty;
                        }
                    }
                }
            }
            catch { return string.Empty; }
            finally
            {
                if (File.Exists(tmpFile)) File.Delete(tmpFile);
            }
        }

        // =====================================================================
        // PARSER DE CLAVE-VALOR
        // =====================================================================
        /// Mapea las lineas del texto OCR a los campos del modelo usando
        private static void ParseKeyValue(string text, TextileMetadata meta)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            string[] allAnchorsList = new string[] {
                "shade name", "sombra", "dyeing class", "clase", "substrate", "sustrato",
                "count/ply", "count / ply", "count", "ply", "fibre type", "fiber type", "fibra",
                "component", "std type", "stdtype", "std", "redye flag", "redye",
                "new thread flag", "new thread", "newthread",
                "batch id", "batchid", "batch", "recipe type", "recipetype", "recipe",
                "lot no", "lotno", "lot"
            };

            meta.ShadeName   = ExtractWithRegex(text, SHADE_KEYS, allAnchorsList, meta.ShadeName);
            meta.DyeingClass = ExtractWithRegex(text, CLASS_KEYS, allAnchorsList, meta.DyeingClass);
            meta.Substrate   = ExtractWithRegex(text, SUBST_KEYS, allAnchorsList, meta.Substrate);
            meta.CountPly    = ExtractWithRegex(text, COUNT_KEYS, allAnchorsList, meta.CountPly);
            meta.FiberType   = ExtractWithRegex(text, FIBER_KEYS, allAnchorsList, meta.FiberType);
        }

        private static string ExtractWithRegex(string text, string[] anchors, string[] stopWords, string currentValue)
        {
            if (currentValue != null && currentValue != "-") return currentValue;

            // Construir regex tolerante a OCR: Permite multiples espacios dentro del anchor (ej: "shade   name")
            string anchorsPattern = string.Join("|", System.Linq.Enumerable.Select(anchors, a => string.Join(@"\s+", a.Split(' '))));
            
            // Busca el anchor, luego ignora cualquier espacio, salto de linea, dos puntos, guiones, igual
            string pattern = $@"(?:{anchorsPattern})[\s=:;.-]*(?<val>[^\r\n]+)";
            
            var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                string val = match.Groups["val"].Value.Trim();

                // Separar si en la misma linea vienen otras columnas (reducir umbral a 2+ espacios)
                var parts = Regex.Split(val, @"\s{2,}|\t");
                string firstPart = parts.Length > 0 ? parts[0].Trim() : val.Trim();

                // Truncar usando StopWords contra el texto COMPLETO (para capturar 'stdtype' pegado)
                string valFull = val.ToLowerInvariant();
                string firstLower = firstPart.ToLowerInvariant();
                int cutInFull = val.Length;

                foreach (var sw in stopWords)
                {
                    bool isOwnAnchor = false;
                    foreach (var own in anchors) if (own == sw) { isOwnAnchor = true; break; }
                    if (isOwnAnchor) continue;

                    // Buscar en el valor completo (no solo en la primera parte) para detectar tokens pegados
                    var matchSw = Regex.Match(valFull, @"(^|\s|\b)" + Regex.Escape(sw) + @"(\b|\s|:)");
                    if (matchSw.Success && matchSw.Index > 0 && matchSw.Index < cutInFull)
                        cutInFull = matchSw.Index;
                }

                // Aplicar corte al valor original preservando mayúsculas
                string finalVal = val.Substring(0, cutInFull).Trim();

                // Si el corte quedó en la primera columna, preferirla
                if (firstPart.Length < finalVal.Length) finalVal = firstPart;

                // Guardia máxima: evitar capturas multi-columna (máx 70 caracteres)
                if (finalVal.Length > 70) finalVal = finalVal.Substring(0, 70);

                finalVal = finalVal.TrimEnd(':', '-', ' ', '.', ',');
                if (!string.IsNullOrWhiteSpace(finalVal)) return finalVal;
            }
            return "-";
        }

        private static bool MatchesAny(string key, string[] anchors)
        {
            foreach (var anchor in anchors)
                if (key.Contains(anchor)) return true;
            return false;
        }

        // =====================================================================
        // PROCESAMIENTO DE IMAGEN — HEREDADO DE Dataextraxtor.cs
        // =====================================================================
        /// Identica a MeasureSharpness de ColorimetricDataExtractor.
        private static float MeasureSharpness(Bitmap src)
        {
            int w = src.Width, h = src.Height;
            int[] kernel = { 0, -1, 0, -1, 4, -1, 0, -1, 0 };

            var bmpData = src.LockBits(
                new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            int    stride = bmpData.Stride;
            byte[] pixels = new byte[stride * h];
            System.Runtime.InteropServices.Marshal.Copy(bmpData.Scan0, pixels, 0, pixels.Length);
            src.UnlockBits(bmpData);

            int x0 = w / 4, x1 = w * 3 / 4;
            int y0 = h / 4, y1 = h * 3 / 4;
            double sum = 0, sumSq = 0;
            long count = 0;

            for (int y = Math.Max(1, y0); y < Math.Min(h - 1, y1); y++)
            {
                for (int x = Math.Max(1, x0); x < Math.Min(w - 1, x1); x++)
                {
                    int lap = 0, ki = 0;
                    for (int dy = -1; dy <= 1; dy++)
                    {
                        for (int dx = -1; dx <= 1; dx++)
                        {
                            int idx = (y + dy) * stride + (x + dx) * 4;
                            int lum = (int)(0.299 * pixels[idx + 2]
                                          + 0.587 * pixels[idx + 1]
                                          + 0.114 * pixels[idx]);
                            lap += kernel[ki] * lum;
                            ki++;
                        }
                    }
                    sum   += lap;
                    sumSq += (double)lap * lap;
                    count++;
                }
            }

            if (count == 0) return 50f;
            double mean     = sum / count;
            double variance = (sumSq / count) - (mean * mean);
            return (float)Math.Max(0, variance);
        }

        /// Umbralizado Otsu (varianza inter-clase) sobre mapa de memoria.
        private static Bitmap BinarizeOtsu(Bitmap src)
        {
            int w = src.Width, h = src.Height;
            int[] hist = new int[256];

            var srcData = src.LockBits(
                new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            int    stride = srcData.Stride;
            byte[] pixels = new byte[stride * h];
            System.Runtime.InteropServices.Marshal.Copy(srcData.Scan0, pixels, 0, pixels.Length);
            src.UnlockBits(srcData);

            for (int y = 0; y < h; y++)
            {
                int offset = y * stride;
                for (int x = 0; x < w; x++)
                {
                    int idx = offset + x * 4;
                    int lum = (int)(0.299 * pixels[idx + 2]
                                  + 0.587 * pixels[idx + 1]
                                  + 0.114 * pixels[idx]);
                    hist[Math.Min(255, Math.Max(0, lum))]++;
                }
            }

            long total = (long)w * h;
            long sumB = 0, wB = 0, sum1 = 0;
            for (int i = 0; i < 256; i++) sum1 += i * hist[i];

            double maxVar    = 0;
            int    threshold = 128;

            for (int t = 0; t < 256; t++)
            {
                wB += hist[t];
                if (wB == 0) continue;
                long   wF   = total - wB;
                if (wF == 0) break;
                sumB += t * hist[t];
                double mB   = (double)sumB / wB;
                double mF   = (double)(sum1 - sumB) / wF;
                double varT = wB * wF * (mB - mF) * (mB - mF);
                if (varT > maxVar) { maxVar = varT; threshold = t; }
            }

            var result  = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            var dstData = result.LockBits(
                new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
            int    dstStride = dstData.Stride;
            byte[] dstPixels = new byte[dstStride * h];

            for (int y = 0; y < h; y++)
            {
                int srcRow = y * stride, dstRow = y * dstStride;
                for (int x = 0; x < w; x++)
                {
                    int si = srcRow + x * 4, di = dstRow + x * 4;
                    int lum = (int)(0.299 * pixels[si + 2]
                                  + 0.587 * pixels[si + 1]
                                  + 0.114 * pixels[si]);
                    byte v = lum <= threshold ? (byte)0 : (byte)255;
                    dstPixels[di]     = v;
                    dstPixels[di + 1] = v;
                    dstPixels[di + 2] = v;
                    dstPixels[di + 3] = 255;
                }
            }

            System.Runtime.InteropServices.Marshal.Copy(dstPixels, 0, dstData.Scan0, dstPixels.Length);
            result.UnlockBits(dstData);
            return result;
        }
    }
}
