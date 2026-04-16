using System;
using System.Drawing;
using System.Windows.Forms;

namespace Color
{
    // Renombrado a FormDetalleCielab para evitar colisiones con metadatos antiguos
    public class FormDetalleCielab : Form
    {
        private CielabChartControl chartFull;
        private Label lblTitle;
        private RichTextBox txtAdvice;
        private Button btnClose;

        public FormDetalleCielab(double dL, double dA, double dB, double dE, double tolerance, string advice)
        {
            InitializeComponents();
            
            // Asignación de datos
            chartFull.DeltaL = dL;
            chartFull.DeltaA = dA;
            chartFull.DeltaB = dB;
            chartFull.DeltaE = dE;
            chartFull.ToleranceDE = tolerance;
            chartFull.Title = "Proyección Espacial CIELAB";
            
            txtAdvice.Text = advice;
            lblTitle.Text = $"Análisis Gráfico CIELAB — ΔE: {dE:F2}";
        }

        private void InitializeComponents()
        {
            this.Text = "Colorimetría Detallada";
            this.Size = new Size(1150, 850);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.BackColor = System.Drawing.Color.White;

            // 1. Cabecera (Título)
            lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 60,
                Font = new Font("Segoe UI", 15, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = System.Drawing.Color.FromArgb(0, 80, 160),
                ForeColor = System.Drawing.Color.White,
                Text = "Análisis del Reporte"
            };

            // 2. Pie (Contenedor de botón)
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 80, BackColor = System.Drawing.Color.WhiteSmoke };
            btnClose = new Button
            {
                Text = "Cerrar Ventana de Análisis",
                Size = new Size(240, 45),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.IndianRed,
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => this.Close();
            pnlBottom.Controls.Add(btnClose);
            pnlBottom.Resize += (s, e) => {
                btnClose.Left = (pnlBottom.Width - btnClose.Width) / 2;
                btnClose.Top = (pnlBottom.Height - btnClose.Height) / 2;
            };

            // 3. Área Central con SplitContainer (Lo más robusto para 2 áreas)
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.None,
                SplitterDistance = 750, 
                IsSplitterFixed = false,
                BorderStyle = BorderStyle.None
            };

            chartFull = new CielabChartControl 
            { 
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.Color.White 
            };
            
            txtAdvice = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = System.Drawing.Color.FromArgb(250, 250, 252),
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI Semibold", 11.5f),
                ForeColor = System.Drawing.Color.FromArgb(30, 30, 60),
                Padding = new Padding(20)
            };

            split.Panel1.Controls.Add(chartFull);
            split.Panel2.Controls.Add(txtAdvice);
            split.Panel2.Padding = new Padding(10);

            // Agregar al formulario cuidando el orden de Dock
            this.Controls.Add(split);      
            this.Controls.Add(lblTitle);    
            this.Controls.Add(pnlBottom);   

            // Forzar z-order correcto
            split.BringToFront();
            lblTitle.SendToBack();
            pnlBottom.SendToBack();
        }
    }
}
