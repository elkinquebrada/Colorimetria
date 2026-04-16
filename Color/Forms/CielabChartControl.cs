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

        public double DeltaL { get => _dL; set { _dL = value; InvalidateSafer(); } }
        public double DeltaA { get => _dA; set { _dA = value; InvalidateSafer(); } }
        public double DeltaB { get => _dB; set { _dB = value; InvalidateSafer(); } }
        public double DeltaE { get => _dE; set { _dE = value; InvalidateSafer(); } }
        public double ToleranceDE { get => _toleranceDE; set { _toleranceDE = value; InvalidateSafer(); } }
        public string Title { get; set; } = "";

        public CielabChartControl()
        {
            this.DoubleBuffered = true;
            this.BackColor = System.Drawing.Color.White;
            this.Size = new Size(300, 250);
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

            int margin = 40; // Aumentado para evitar corte de etiquetas
            int lWidth = 40; // Espacio para el eje L* a la derecha
            Rectangle chartArea = new Rectangle(margin, margin + 20, this.Width - margin * 2 - lWidth, this.Height - margin * 2 - 20);
            Point center = new Point(chartArea.Left + chartArea.Width / 2, chartArea.Top + chartArea.Height / 2);

            // ---- Título Opcional ----
            if (!string.IsNullOrEmpty(Title))
            {
                using (Font titleFont = new Font(this.Font, FontStyle.Bold))
                {
                    g.DrawString(Title, titleFont, Brushes.DarkSlateGray, margin, 5);
                }
            }

            // Determinar escala (unidades CIELAB por pixel)
            double maxVal = Math.Max(ToleranceDE * 1.5, Math.Max(Math.Abs(DeltaA), Math.Abs(DeltaB))) + 0.5;
            float scale = (float)(Math.Min(chartArea.Width, chartArea.Height) / (2.0 * maxVal));

            // ---- Fondo Gradiente a*-b* ----
            PointF[] gradientPoints = new PointF[]
            {
                new PointF(chartArea.Left, chartArea.Top),      
                new PointF(chartArea.Right, chartArea.Top),     
                new PointF(chartArea.Right, chartArea.Bottom),  
                new PointF(chartArea.Left, chartArea.Bottom)    
            };

            using (PathGradientBrush pgb = new PathGradientBrush(gradientPoints))
            {
                // el centro (a = 0 ; b = 0 ) es neutro/blanco
                pgb.CenterColor = System.Drawing.Color.White; 
                pgb.SurroundColors = new System.Drawing.Color[]
                {
                    System.Drawing.Color.FromArgb(80, 150, 255, 150), 
                    System.Drawing.Color.FromArgb(80, 255, 150, 150), 
                    System.Drawing.Color.FromArgb(80, 255, 150, 255), 
                    System.Drawing.Color.FromArgb(80, 150, 255, 255)  
                };
                g.FillRectangle(pgb, chartArea);
            }

            // ---- Dibujar Ejes a*-b* ----
            using (Pen axisPen = new Pen(System.Drawing.Color.DarkGray, 1f))
            {
                axisPen.DashStyle = DashStyle.Dash;
                g.DrawLine(axisPen, chartArea.Left, center.Y, chartArea.Right, center.Y); 
                g.DrawLine(axisPen, center.X, chartArea.Top, center.X, chartArea.Bottom); 
            }

            // ---- Etiquetas de los Ejes ----
            DrawAxisLabel(g, "Rojo (+a*)", chartArea.Right - 40, center.Y + 5, StringAlignment.Near, System.Drawing.Color.DarkRed);
            DrawAxisLabel(g, "Verde (-a*)", chartArea.Left, center.Y + 5, StringAlignment.Near, System.Drawing.Color.DarkGreen);
            DrawAxisLabel(g, "Amarillo (+b*)", center.X + 5, chartArea.Top, StringAlignment.Near, System.Drawing.Color.DarkGoldenrod);
            DrawAxisLabel(g, "Azul (-b*)", center.X + 5, chartArea.Bottom - 15, StringAlignment.Near, System.Drawing.Color.DarkBlue);

            // ---- Círculo de Tolerancia ----
            float tolRadius = (float)(ToleranceDE * scale);
            using (Pen tolPen = new Pen(System.Drawing.Color.FromArgb(100, 200, 200, 200), 1.5f))
            {
                tolPen.DashStyle = DashStyle.Dot;
                g.DrawEllipse(tolPen, center.X - tolRadius, center.Y - tolRadius, tolRadius * 2, tolRadius * 2);
                
                // Texto de tolerancia
                string tolText = $"Tol ΔE: {ToleranceDE:F2}";
                g.DrawString(tolText, this.Font, Brushes.Gray, center.X - tolRadius, center.Y - tolRadius - 15);
            }

            // ---- Dibujar Punto de Medición (Lote) ----
            float ptX = (float)(center.X + DeltaA * scale);
            float ptY = (float)(center.Y - DeltaB * scale); 

            // Vector de corrección
            using (Pen vectorPen = new Pen(System.Drawing.Color.FromArgb(0, 102, 204), 2f))
            {
                vectorPen.EndCap = LineCap.ArrowAnchor;
                g.DrawLine(vectorPen, center.X, center.Y, ptX, ptY);
            }

            // Punto (Standard es el centro)
            g.FillEllipse(Brushes.LimeGreen, center.X - 4, center.Y - 4, 8, 8); 
            
            // Punto (Lote)
            using (Brush pointBrush = new SolidBrush(System.Drawing.Color.Red))
            {
                g.FillEllipse(pointBrush, ptX - 4, ptY - 4, 8, 8);
                g.DrawEllipse(Pens.White, ptX - 5, ptY - 5, 10, 10);
            }

            // Etiqueta del punto
            string batchText = $"Δa={DeltaA:F2}, Δb={DeltaB:F2}\nΔE={DeltaE:F2}";
            g.DrawString(batchText, new Font(this.Font, FontStyle.Bold), Brushes.Black, ptX + 8, ptY - 10);

            // ---- Eje L* (Luminosidad) ----
            int lX = this.Width - lWidth + 5;
            int lYStart = chartArea.Top;
            int lYEnd = chartArea.Bottom;
            int lHeight = lYEnd - lYStart;

            // Gradiente Blanco-Negro
            using (LinearGradientBrush lBrush = new LinearGradientBrush(new Point(lX, lYStart), new Point(lX, lYEnd), System.Drawing.Color.White, System.Drawing.Color.Black))
            {
                g.FillRectangle(lBrush, lX, lYStart, 12, lHeight);
            }
            g.DrawRectangle(Pens.Gray, lX, lYStart, 12, lHeight);

            // Indicadores L*
            g.DrawString("L+", this.Font, Brushes.Gray, lX + 15, lYStart);
            g.DrawString("L-", this.Font, Brushes.Gray, lX + 15, lYEnd - 15);

            // Posición de la medición en L*
            double maxValL = Math.Max(Math.Abs(DeltaL), 2.0) + 0.5;
            float lCenterY = lYStart + lHeight / 2f;
            float lPtY = (float)(lCenterY - (DeltaL / maxValL) * (lHeight / 2f));
            
            // Limitar a los bordes
            lPtY = Math.Max(lYStart, Math.Min(lYEnd, lPtY));

            using (Pen markerPen = new Pen(System.Drawing.Color.DarkBlue, 2f))
            {
                g.DrawLine(markerPen, lX - 5, lPtY, lX + 17, lPtY);
                g.DrawString($"ΔL={DeltaL:F2}", this.Font, Brushes.DarkBlue, lX - 45, lPtY - 7);
            }

            // ---- Leyenda (Legend) ----
            int legendY = chartArea.Bottom + 5;
            int legendX = chartArea.Left;
            
            // Punto Referencia
            g.FillEllipse(Brushes.LimeGreen, legendX, legendY + 5, 8, 8);
            g.DrawString("Punto referencia", this.Font, Brushes.Black, legendX + 12, legendY + 1);
            
            // Punto Resultado
            using (Brush pointBrush = new SolidBrush(System.Drawing.Color.Red))
            {
                g.FillEllipse(pointBrush, legendX + 130, legendY + 5, 8, 8);
                g.DrawEllipse(Pens.DarkRed, legendX + 130, legendY + 5, 8, 8);
            }
            g.DrawString("Punto resultado", this.Font, Brushes.Black, legendX + 142, legendY + 1);
        }

        private void DrawAxisLabel(Graphics g, string text, int x, int y, StringAlignment align, System.Drawing.Color color)
        {
            using (StringFormat sf = new StringFormat { Alignment = align })
            using (Brush brush = new SolidBrush(color))
            {
                g.DrawString(text, this.Font, brush, x, y, sf);
            }
        }
    }
}
