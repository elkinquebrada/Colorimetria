using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Color
{
    public class CielabChartControl : UserControl
    {
        private double _dL;
        private double _dA;
        private double _dB;
        private double _dE;
        private double _toleranceDE = 1.2;

        public enum ViewMode { Relative, Absolute }

        public double DeltaL { get => _dL; set { _dL = value; InvalidateSafer(); } }
        public double DeltaA { get => _dA; set { _dA = value; InvalidateSafer(); } }
        public double DeltaB { get => _dB; set { _dB = value; InvalidateSafer(); } }
        public double DeltaE { get => _dE; set { _dE = value; InvalidateSafer(); } }
        public double ToleranceDE { get => _toleranceDE; set { _toleranceDE = value; InvalidateSafer(); } }
        public string Title { get; set; } = "Análisis CIELAB";
        public ViewMode Mode { get; set; } = ViewMode.Relative;
        public string InstructionMessage { get; set; } = "";

        public double AbsoluteL { get; set; } = 50.0;
        public double AbsoluteA { get; set; } = 0.0;
        public double AbsoluteB { get; set; } = 0.0;

        public double LotL { get; set; } = 50.0;
        public double LotA { get; set; } = 0.0;
        public double LotB { get; set; } = 0.0;

        public CielabChartControl()
        {
            // CORRECCIÓN 1: DoubleBuffered y ResizeRedraw para evitar duplicidad visual
            this.DoubleBuffered = true;
            this.ResizeRedraw = true;
            this.Size = new Size(400, 400);
        }

        private void InvalidateSafer()
        {
            if (this.IsHandleCreated)
            {
                this.BeginInvoke(new Action(() => this.Invalidate()));
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Fondo con degradado profesional
            using (LinearGradientBrush backBrush = new LinearGradientBrush(this.ClientRectangle,
                System.Drawing.Color.FromArgb(252, 252, 254),
                System.Drawing.Color.FromArgb(240, 242, 248), 45f))
            {
                g.FillRectangle(backBrush, this.ClientRectangle);
            }

            // CORRECCIÓN 2: Cálculo de área dinámico para Pantalla Completa
            int margin = 50;
            // Calculamos un tamaño cuadrado basado en el lado más corto disponible
            int size = Math.Min(this.Width - (margin * 2), this.Height - (margin * 2) - 20);
            
            Rectangle chartArea = new Rectangle(
                (this.Width - size) / 2,
                (this.Height - size) / 2 + 10,
                size,
                size
            );

            Point center = new Point(chartArea.X + chartArea.Width / 2, chartArea.Y + chartArea.Height / 2);

            if (!string.IsNullOrEmpty(Title))
            {
                using (Font f = new Font("Segoe UI", 12, FontStyle.Bold))
                {
                    g.DrawString(Title, f, Brushes.MidnightBlue, margin, 10);
                }
            }

            // --- Determinar Escala y Rango Dinámico ---
            double maxLabValue = 120.0; // Rango visual completo
            if (Mode == ViewMode.Relative)
            {
                double maxD = Math.Max(Math.Abs(DeltaA), Math.Abs(DeltaB));
                maxLabValue = Math.Max(ToleranceDE * 2.5, maxD * 1.5);
                if (maxLabValue < 3) maxLabValue = 3; 
            }
            float scale = (chartArea.Width / 2f) / (float)maxLabValue; 

            // Dibujar Círculo Cromático de fondo
            DrawChromaticCircle(g, center, chartArea.Width / 2, maxLabValue);

            // Dibujar Ejes
            using (Pen axisPen = new Pen(System.Drawing.Color.FromArgb(100, System.Drawing.Color.White), 2))
            {
                g.DrawLine(axisPen, chartArea.Left, center.Y, chartArea.Right, center.Y);
                g.DrawLine(axisPen, center.X, chartArea.Top, center.X, chartArea.Bottom);
            }

            // Etiquetas de los Ejes
            using (Font axisFont = new Font("Segoe UI", 9, FontStyle.Bold))
            {
                g.DrawString("Amarillo (+b*)", axisFont, Brushes.Gold, center.X - 40, chartArea.Top - 20);
                g.DrawString("Azul (-b*)", axisFont, Brushes.RoyalBlue, center.X - 30, chartArea.Bottom + 5);
                g.DrawString("Verde (-a*)", axisFont, Brushes.ForestGreen, chartArea.Left - 80, center.Y - 10);
                g.DrawString("Rojo (+a*)", axisFont, Brushes.Crimson, chartArea.Right + 5, center.Y - 10);
            }

            // --- Lógica de Dibujo de Puntos (Inmersión) ---
            double plotStdA = Mode == ViewMode.Absolute ? AbsoluteA : 0;
            double plotStdB = Mode == ViewMode.Absolute ? AbsoluteB : 0;
            double plotLotA = Mode == ViewMode.Absolute ? LotA : DeltaA;
            double plotLotB = Mode == ViewMode.Absolute ? LotB : DeltaB;

            // Punto Estándar (Verde)
            PointF pStd = new PointF(
                center.X + (float)plotStdA * scale,
                center.Y - (float)plotStdB * scale
            );

            // Punto Lote (Rojo)
            PointF pLot = new PointF(
                center.X + (float)plotLotA * scale,
                center.Y - (float)plotLotB * scale
            );

            // Tolerancia — CMC 2:1 Elipse Rotada
            if (Mode == ViewMode.Relative)
            {
                double C1 = Math.Sqrt(AbsoluteA * AbsoluteA + AbsoluteB * AbsoluteB);
                double h1_rad = Math.Atan2(AbsoluteB, AbsoluteA);
                double h1_deg = h1_rad * 180.0 / Math.PI;
                if (h1_deg < 0) h1_deg += 360.0;

                var axes = Color.ColorimetricCalculator.CalculateCmcSemiAxes(AbsoluteL, C1, h1_deg);
                float wC = (float)(axes.sc * ToleranceDE * scale);
                float hH = (float)(axes.sh * ToleranceDE * scale);

                using (Pen tolPen = new Pen(System.Drawing.Color.FromArgb(200, 255, 255, 255), 1.8f))
                {
                    tolPen.DashStyle = DashStyle.Dash;
                    
                    GraphicsState state = g.Save();
                    g.TranslateTransform(center.X, center.Y);
                    g.RotateTransform((float)(-h1_deg));
                    g.DrawEllipse(tolPen, -wC, -hH, wC * 2, hH * 2);
                    g.Restore(state);
                }
            }

            // Vector Tendencial
            using (Pen vectorPen = new Pen(System.Drawing.Color.FromArgb(60, 255, 255, 255), 2.5f))
            {
                vectorPen.EndCap = LineCap.ArrowAnchor;
                g.DrawLine(vectorPen, pStd, pLot);
            }

            // Dibujar Puntos (Diseño anidado para evitar ocultamientos cuando se superponen)
            // Estándar es más grande y traslúcido
            g.FillEllipse(Brushes.LimeGreen, pStd.X - 7, pStd.Y - 7, 14, 14);
            g.DrawEllipse(Pens.White, pStd.X - 7, pStd.Y - 7, 14, 14);
            g.DrawString("Est.", new Font("Segoe UI", 7, FontStyle.Bold), Brushes.DarkGreen, pStd.X + 8, pStd.Y - 6);

            // Lote es más pequeño y concéntrico
            g.FillEllipse(Brushes.Red, pLot.X - 4, pLot.Y - 4, 8, 8);
            g.DrawEllipse(Pens.White, pLot.X - 4, pLot.Y - 4, 8, 8);
            
            // Tooltip flotante de datos
            string info = Mode == ViewMode.Relative 
                ? $"Δa: {DeltaA:F2}\nΔb: {DeltaB:F2}\nΔE: {DeltaE:F2}"
                : $"Lote: a*={LotA:F1}, b*={LotB:F1}";
            Size box = TextRenderer.MeasureText(info, this.Font);
            Rectangle rectInfo = new Rectangle((int)pLot.X + 10, (int)pLot.Y - 20, box.Width + 10, box.Height + 5);
            
            g.FillRectangle(new SolidBrush(System.Drawing.Color.FromArgb(200, 0, 0, 0)), rectInfo);
            g.DrawString(info, this.Font, Brushes.White, rectInfo.X + 5, rectInfo.Y + 2);

            // Renderizar componentes adicionales perdidos
            int lWidth = 60;
            DrawComparisonSamples(g, this.Width - lWidth - 140, margin);
            DrawLightnessAxis(g, this.Width - lWidth + 10, chartArea.Top, lWidth - 30, chartArea.Height);
        }

        private void DrawChromaticCircle(Graphics g, Point center, int radius, double maxLabValue)
        {
            int size = radius * 2;
            using (Bitmap bmp = new Bitmap(size, size))
            {
                var bmpData = bmp.LockBits(new Rectangle(0, 0, size, size), System.Drawing.Imaging.ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                int stride = bmpData.Stride;
                byte[] pixels = new byte[stride * size];

                for (int py = 0; py < size; py++)
                {
                    for (int px = 0; px < size; px++)
                    {
                        double dx = px - radius;
                        double dy = py - radius;
                        double dist = Math.Sqrt(dx * dx + dy * dy);

                        if (dist <= radius)
                        {
                            // La rueda cromática siempre mostrará el espectro completo (independiente del zoom)
                            double a = (dx / radius) * 128.0;
                            double b = (-dy / radius) * 128.0;
                            
                            System.Drawing.Color c = LabToRgb(70, a, b);
                            double alpha = 255;
                            if (dist > radius - 3) alpha = 255 * (radius - dist) / 3.0;

                            int idx = py * stride + px * 4;
                            pixels[idx + 0] = c.B;
                            pixels[idx + 1] = c.G;
                            pixels[idx + 2] = c.R;
                            pixels[idx + 3] = (byte)alpha;
                        }
                    }
                }
                System.Runtime.InteropServices.Marshal.Copy(pixels, 0, bmpData.Scan0, pixels.Length);
                bmp.UnlockBits(bmpData);
                g.DrawImage(bmp, center.X - radius, center.Y - radius, size, size);
            }
        }

        private System.Drawing.Color LabToRgb(double L, double a, double b)
        {
            double y = (L + 16) / 116.0;
            double x = a / 500.0 + y;
            double z = y - b / 200.0;
            double ToX(double v) => v > 0.206893 ? Math.Pow(v, 3) : (v - 16 / 116.0) / 7.787;
            x = 0.95047 * ToX(x); y = 1.000 * ToX(y); z = 1.08883 * ToX(z);
            double r = x * 3.2406 + y * -1.5372 + z * -0.4986;
            double g = x * -0.9689 + y * 1.8758 + z * 0.0415;
            double bl = x * 0.0557 + y * -0.2040 + z * 1.0570;
            double FromR(double v) => v <= 0.0031308 ? 12.92 * v : 1.055 * Math.Pow(v, 1 / 2.4) - 0.055;
            return System.Drawing.Color.FromArgb(
                (int)Math.Max(0, Math.Min(255, FromR(r) * 255)),
                (int)Math.Max(0, Math.Min(255, FromR(g) * 255)),
                (int)Math.Max(0, Math.Min(255, FromR(bl) * 255)));
        }

        private void DrawComparisonSamples(Graphics g, int x, int y)
        {
            int sw = 60, sh = 60;
            Rectangle rStd = new Rectangle(x, y, sw, sh);
            Rectangle rLot = new Rectangle(x + sw + 10, y, sw, sh);

            System.Drawing.Color cStd = LabToRgb(AbsoluteL, AbsoluteA, AbsoluteB);
            System.Drawing.Color cLot = LabToRgb(LotL, LotA, LotB);

            // Sombra
            using (Brush shadow = new SolidBrush(System.Drawing.Color.FromArgb(40, 0, 0, 0)))
                g.FillRectangle(shadow, x + 3, y + 3, (sw * 2) + 10 + 3, sh);

            using (SolidBrush bStd = new SolidBrush(cStd)) g.FillRectangle(bStd, rStd);
            using (SolidBrush bLot = new SolidBrush(cLot)) g.FillRectangle(bLot, rLot);

            using (Pen borderPen = new Pen(System.Drawing.Color.White, 2f))
            {
                g.DrawRectangle(borderPen, rStd);
                g.DrawRectangle(borderPen, rLot);
            }

            using (Font f = new Font("Segoe UI", 8.5f, FontStyle.Bold))
            {
                string[] labels = { "STD", "LOT" };
                int[] xs = { x, x + sw + 10 };
                for (int i = 0; i < 2; i++)
                {
                    SizeF sz = g.MeasureString(labels[i], f);
                    float lx = xs[i] + (sw - sz.Width) / 2f;
                    float ly = y + sh + 5;
                    using (Brush bg = new SolidBrush(System.Drawing.Color.FromArgb(160, 0, 0, 0)))
                        g.FillRectangle(bg, lx - 2, ly, sz.Width + 4, sz.Height + 2);
                    g.DrawString(labels[i], f, Brushes.White, lx, ly);
                }
            }
        }

        private void DrawLightnessAxis(Graphics g, int x, int y, int w, int h)
        {
            using (LinearGradientBrush lBrush = new LinearGradientBrush(new Rectangle(x, y, w, h), System.Drawing.Color.White, System.Drawing.Color.Black, 90f))
            {
                g.FillRectangle(lBrush, x, y, w, h);
            }
            g.DrawRectangle(Pens.Gray, x, y, w, h);

            float stdY = (float)(y + h * (1.0 - AbsoluteL / 100.0));
            float lotY = (float)(y + h * (1.0 - LotL / 100.0));
            stdY = Math.Max(y, Math.Min(y + h, stdY));
            lotY = Math.Max(y, Math.Min(y + h, lotY));

            g.DrawLine(new Pen(System.Drawing.Color.LimeGreen, 3f), x - 5, stdY, x + w + 5, stdY);
            
            Point[] arrow = { new Point(x - 5, (int)lotY), new Point(x - 15, (int)lotY - 6), new Point(x - 15, (int)lotY + 6) };
            g.FillPolygon(Brushes.Crimson, arrow);
            
            using (Font bf = new Font(this.Font, FontStyle.Bold))
                g.DrawString("Lote", bf, Brushes.Crimson, x - 45, lotY - 7);
        }
    }
}