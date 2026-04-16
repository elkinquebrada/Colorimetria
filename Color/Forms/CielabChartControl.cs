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

        // ---- Propiedades de Coordenadas Absolutas (Inmersión de Datos) ----
        public double AbsoluteL { get; set; } = 50.0;
        public double AbsoluteA { get; set; } = 0.0;
        public double AbsoluteB { get; set; } = 0.0;

        public CielabChartControl()
        {
            this.DoubleBuffered = true;
            this.BackColor = System.Drawing.Color.White;
            this.Size = new Size(500, 450);
            this.Font = new Font("Segoe UI", 9f);
        }

        private void InvalidateSafer()
        {
            if (this.InvokeRequired) this.Invoke(new Action(() => this.Invalidate()));
            else this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Fondo profesional
            using (LinearGradientBrush backBrush = new LinearGradientBrush(this.ClientRectangle,
                System.Drawing.Color.FromArgb(252, 252, 254),
                System.Drawing.Color.FromArgb(240, 242, 248), 45f))
            {
                g.FillRectangle(backBrush, this.ClientRectangle);
            }

            int margin = 50;
            int lWidth = 60;
            Rectangle chartArea = new Rectangle(margin, margin + 25, this.Width - margin * 2 - lWidth, this.Height - margin * 2 - 25);
            Point center = new Point(chartArea.X + chartArea.Width / 2, chartArea.Y + chartArea.Height / 2);

            if (!string.IsNullOrEmpty(Title))
            {
                using (Font f = new Font("Segoe UI", 12, FontStyle.Bold))
                {
                    g.DrawString(Title, f, Brushes.MidnightBlue, margin, 10);
                }
            }

            // 1. DETERMINAR LÓGICA DE ESCALA SEGÚN MODO
            double plotStdA = 0, plotStdB = 0;
            double plotLotA = DeltaA, plotLotB = DeltaB;

            if (Mode == ViewMode.Absolute)
            {
                plotStdA = AbsoluteA;
                plotStdB = AbsoluteB;
                plotLotA = AbsoluteA + DeltaA;
                plotLotB = AbsoluteB + DeltaB;
            }

            double maxCoord = Math.Max(
                Math.Max(Math.Abs(plotStdA), Math.Abs(plotStdB)),
                Math.Max(Math.Abs(plotLotA), Math.Abs(plotLotB))
            );
            double maxVal = Math.Max(ToleranceDE * 2.0, maxCoord) + 1.0;
            float scale = (float)(Math.Min(chartArea.Width, chartArea.Height) / (2.0 * maxVal));

            // Rueda de colores CIELAB (rango fijo ±128 para colores reales)
            int wheelRadius = (int)(Math.Min(chartArea.Width, chartArea.Height) / 2.0);
            DrawCielabColorWheel(g, center, wheelRadius);

            // Ejes (Neutro 0,0) — blancos para contrastar con la rueda
            using (Pen axisPen = new Pen(System.Drawing.Color.FromArgb(200, 255, 255, 255), 1.5f))
            {
                g.DrawLine(axisPen, chartArea.Left, center.Y, chartArea.Right, center.Y);
                g.DrawLine(axisPen, center.X, chartArea.Top, center.X, chartArea.Bottom);
            }

            // Etiquetas de Ejes
            DrawAxisLabel(g, "Rojo (+a*)", chartArea.Right - 40, center.Y + 6, StringAlignment.Near, System.Drawing.Color.FromArgb(220, 0, 0));
            DrawAxisLabel(g, "Verde (-a*)", chartArea.Left, center.Y + 6, StringAlignment.Near, System.Drawing.Color.FromArgb(0, 160, 0));
            DrawAxisLabel(g, "Amarillo (+b*)", center.X + 6, chartArea.Top, StringAlignment.Near, System.Drawing.Color.FromArgb(180, 140, 0));
            DrawAxisLabel(g, "Azul (-b*)", center.X + 6, chartArea.Bottom - 18, StringAlignment.Near, System.Drawing.Color.FromArgb(30, 30, 200));

            // Tolerancia — círculo blanco para contrastar con la rueda
            if (Mode == ViewMode.Relative)
            {
                float tolRadius = (float)(ToleranceDE * scale);
                using (Pen tolPen = new Pen(System.Drawing.Color.FromArgb(220, 255, 255, 255), 1.8f))
                {
                    tolPen.DashStyle = DashStyle.Dash;
                    g.DrawEllipse(tolPen, center.X - tolRadius, center.Y - tolRadius, tolRadius * 2, tolRadius * 2);
                }
                // Etiqueta de tolerancia con fondo semitransparente
                using (Font tolFont = new Font(this.Font.FontFamily, 8f))
                {
                    string tolTxt = $"Tol ΔE: {ToleranceDE:F2}";
                    SizeF ts = g.MeasureString(tolTxt, tolFont);
                    float tx = center.X - tolRadius;
                    float ty = center.Y - tolRadius - 17;
                    using (Brush bgBrush = new SolidBrush(System.Drawing.Color.FromArgb(160, 0, 0, 0)))
                        g.FillRectangle(bgBrush, tx - 2, ty - 1, ts.Width + 4, ts.Height + 2);
                    g.DrawString(tolTxt, tolFont, Brushes.White, tx, ty);
                }
            }

            float sX = (float)(center.X + plotStdA * scale);
            float sY = (float)(center.Y - plotStdB * scale);
            float lX = (float)(center.X + plotLotA * scale);
            float lY = (float)(center.Y - plotLotB * scale);

            // Estándar
            g.FillEllipse(Brushes.LimeGreen, sX - 5, sY - 5, 10, 10);
            g.DrawEllipse(Pens.White, sX - 6, sY - 6, 12, 12);
            if (Mode == ViewMode.Absolute)
            {
                SizeF es = g.MeasureString("Est.", this.Font);
                using (Brush bg = new SolidBrush(System.Drawing.Color.FromArgb(160, 0, 0, 0)))
                    g.FillRectangle(bg, sX + 7, sY - 9, es.Width + 2, es.Height);
                g.DrawString("Est.", this.Font, Brushes.LimeGreen, sX + 8, sY - 8);
            }

            // Vector
            using (Pen vectorPen = new Pen(System.Drawing.Color.FromArgb(30, 144, 255), 2.5f))
            {
                vectorPen.EndCap = LineCap.ArrowAnchor;
                g.DrawLine(vectorPen, sX, sY, lX, lY);
            }

            // Lote
            using (Brush lotB = new SolidBrush(System.Drawing.Color.FromArgb(220, 20, 60)))
            {
                g.FillEllipse(lotB, lX - 5, lY - 5, 10, 10);
                g.DrawEllipse(Pens.White, lX - 6, lY - 6, 12, 12);
            }
            // Etiqueta "Lote" con fondo
            {
                SizeF ls = g.MeasureString("Lote", this.Font);
                using (Brush bg = new SolidBrush(System.Drawing.Color.FromArgb(160, 0, 0, 0)))
                    g.FillRectangle(bg, lX + 9, lY + 5, ls.Width + 2, ls.Height);
                g.DrawString("Lote", this.Font, Brushes.White, lX + 10, lY + 6);
            }

            // Datos Δa Δb ΔE con fondo semitransparente
            string dataTxt = $"Δa: {DeltaA:F2}\nΔb: {DeltaB:F2}\nΔE: {DeltaE:F2}";
            using (Font boldF = new Font("Segoe UI", 9.5f, FontStyle.Bold))
            {
                SizeF ds = g.MeasureString(dataTxt, boldF);
                float dx = lX + 12;
                float dy = lY - 25;
                using (Brush bgBrush = new SolidBrush(System.Drawing.Color.FromArgb(170, 0, 0, 0)))
                    g.FillRectangle(bgBrush, dx - 3, dy - 2, ds.Width + 6, ds.Height + 4);
                g.DrawString(dataTxt, boldF, Brushes.White, dx, dy);
            }

            // Info final
            string modeName = Mode == ViewMode.Absolute ? "Vista Real" : "Vista Relativa";
            string spatialInfo = $"[{modeName}] Std: L={AbsoluteL:F1}, a={AbsoluteA:F1}, b={AbsoluteB:F1}";
            g.DrawString(spatialInfo, new Font("Segoe UI", 8.5f, FontStyle.Italic), Brushes.DarkSlateGray, margin, this.Height - 30);

            DrawComparisonSamples(g, this.Width - lWidth - 110, 20);

            // lateral
            DrawLightnessAxis(g, this.Width - lWidth + 10, chartArea.Top, lWidth - 25, chartArea.Height);
        }

        private void DrawComparisonSamples(Graphics g, int x, int y)
        {
            int sw = 50, sh = 50;
            Rectangle rStd = new Rectangle(x, y, sw, sh);
            Rectangle rLot = new Rectangle(x + sw + 5, y, sw, sh);

            System.Drawing.Color cStd = LabToRgb(AbsoluteL, AbsoluteA, AbsoluteB);
            System.Drawing.Color cLot = LabToRgb(AbsoluteL + DeltaL, AbsoluteA + DeltaA, AbsoluteB + DeltaB);

            // Sombra
            using (Brush shadow = new SolidBrush(System.Drawing.Color.FromArgb(40, 0, 0, 0)))
                g.FillRectangle(shadow, x + 3, y + 3, sw * 2 + 5 + 3, sh);

            // Relleno con gradiente sutil para dar volumen
            using (LinearGradientBrush bStd = new LinearGradientBrush(rStd,
                System.Drawing.Color.FromArgb(Math.Min(255, cStd.R + 30), Math.Min(255, cStd.G + 30), Math.Min(255, cStd.B + 30)),
                cStd, 135f))
                g.FillRectangle(bStd, rStd);

            using (LinearGradientBrush bLot = new LinearGradientBrush(rLot,
                System.Drawing.Color.FromArgb(Math.Min(255, cLot.R + 30), Math.Min(255, cLot.G + 30), Math.Min(255, cLot.B + 30)),
                cLot, 135f))
                g.FillRectangle(bLot, rLot);

            // Borde blanco para resaltar
            using (Pen borderPen = new Pen(System.Drawing.Color.White, 2f))
            {
                g.DrawRectangle(borderPen, rStd);
                g.DrawRectangle(borderPen, rLot);
            }

            // Etiquetas STD / LOT con fondo oscuro
            using (Font f = new Font("Segoe UI", 7.5f, FontStyle.Bold))
            {
                string[] labels = { "STD", "LOT" };
                int[] xs = { x, x + sw + 5 };
                for (int i = 0; i < 2; i++)
                {
                    SizeF sz = g.MeasureString(labels[i], f);
                    float lx = xs[i] + (sw - sz.Width) / 2f;
                    float ly = y + sh + 3;
                    using (Brush bg = new SolidBrush(System.Drawing.Color.FromArgb(160, 0, 0, 0)))
                        g.FillRectangle(bg, lx - 1, ly, sz.Width + 2, sz.Height);
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
            float lotY = (float)(y + h * (1.0 - (AbsoluteL + DeltaL) / 100.0));
            stdY = Math.Max(y, Math.Min(y + h, stdY));
            lotY = Math.Max(y, Math.Min(y + h, lotY));

            g.DrawLine(new Pen(System.Drawing.Color.LimeGreen, 2.5f), x - 5, stdY, x + w + 5, stdY);
            Point[] arrow = { new Point(x - 5, (int)lotY), new Point(x - 15, (int)lotY - 6), new Point(x - 15, (int)lotY + 6) };
            g.FillPolygon(Brushes.Crimson, arrow);
            g.DrawString("Lote", new Font(this.Font, FontStyle.Bold), Brushes.Crimson, x - 45, lotY - 7);
        }

        private void DrawAxisLabel(Graphics g, string text, float x, float y, StringAlignment align, System.Drawing.Color color)
        {
            using (Brush b = new SolidBrush(color))
            using (Font f = new Font("Segoe UI Semibold", 9))
            {
                StringFormat sf = new StringFormat { Alignment = align };
                g.DrawString(text, f, b, x, y, sf);
            }
        }

        private void DrawCielabColorWheel(Graphics g, Point center, int radius)
        {
            if (radius <= 0) return;
            int size = radius * 2;
            const double LAB_RANGE = 128.0;
            const double L_FIXED = 65.0;

            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(size, size))
            {
                System.Drawing.Imaging.BitmapData bmpData = bmp.LockBits(
                    new Rectangle(0, 0, size, size),
                    System.Drawing.Imaging.ImageLockMode.WriteOnly,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb);

                int stride = bmpData.Stride;
                byte[] pixels = new byte[stride * size];

                for (int py = 0; py < size; py++)
                {
                    for (int px = 0; px < size; px++)
                    {
                        double dx = px - radius;
                        double dy = radius - py;
                        double dist = Math.Sqrt(dx * dx + dy * dy);

                        if (dist <= radius)
                        {
                            double aVal = (dx / radius) * LAB_RANGE;
                            double bVal = (dy / radius) * LAB_RANGE;
                            double L = L_FIXED + (dist / radius) * 12.0;

                            System.Drawing.Color c = LabToRgb(L, aVal, bVal);

                            int alpha = 255;
                            if (radius - dist < 2.0)
                                alpha = (int)((radius - dist) / 2.0 * 255);

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
            double gr = x * -0.9689 + y * 1.8758 + z * 0.0415;
            double bl = x * 0.0557 + y * -0.2040 + z * 1.0570;
            double ToR(double v)
            {
                v = v > 0.0031308 ? 1.055 * Math.Pow(v, 1 / 2.4) - 0.055 : 12.92 * v;
                return Math.Max(0, Math.Min(255, v * 255.0));
            }
            return System.Drawing.Color.FromArgb((int)ToR(r), (int)ToR(gr), (int)ToR(bl));
        }
    }
}