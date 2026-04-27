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

        public FormDetalleCielab(double dL, double dA, double dB, double dE, double cmc, double tolerance, string advice, 
            double absL = 50, double absA = 0, double absB = 0)
        {
            InitializeComponents();
            
            // Asignación de datos (Deltas)
            chartFull.DeltaL = dL;
            chartFull.DeltaA = dA;
            chartFull.DeltaB = dB;
            chartFull.DeltaE = dE;
            chartFull.ToleranceDE = tolerance;

            chartFull.Title = "Grafico de Desviación de Color (CIELAB)";
            chartFull.InstructionMessage = advice;

            // Inmersión de Datos (Valores Absolutos)
            chartFull.AbsoluteL = absL;
            chartFull.AbsoluteA = absA;
            chartFull.AbsoluteB = absB;
            
            // Inmersión de Datos Lote (Requerido por nueva UI)
            chartFull.LotL = absL + dL;
            chartFull.LotA = absA + dA;
            chartFull.LotB = absB + dB;
            
            txtAdvice.Text = advice;
            lblTitle.Text = $"Análisis Grafico — Lote vs Estándar";
        }

        private void InitializeComponents()
        {
            this.Text = "TINT COATS CADENA";
            this.Size = new Size(1250, 920);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.BackColor = System.Drawing.Color.FromArgb(245, 245, 250);

            // 1. Cabecera (Título Premium)
            lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 60,
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = System.Drawing.Color.FromArgb(20, 40, 80),
                ForeColor = System.Drawing.Color.White,
                Text = "Análisis de Colorimetría"
            };

            // 1.1 Barra de controles superior
            Panel pnlControls = new Panel 
            { 
                Dock = DockStyle.Top, 
                Height = 45, 
                BackColor = System.Drawing.Color.FromArgb(30, 60, 110),
                Padding = new Padding(10, 0, 10, 0)
            };
            
            CheckBox chkViewMode = new CheckBox
            {
                Text = "Activar Vista Espacial Real (Neutro al centro)",
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI Semibold", 10.5f),
                AutoSize = true,
                Location = new Point(15, 10),
                Cursor = Cursors.Hand
            };
            chkViewMode.CheckedChanged += (s, e) => {
                chartFull.Mode = chkViewMode.Checked ? CielabChartControl.ViewMode.Absolute : CielabChartControl.ViewMode.Relative;
                chartFull.Invalidate();
            };
            pnlControls.Controls.Add(chkViewMode);

            // 2. Pie (Contenedor de botón)
            var pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 70, BackColor = System.Drawing.Color.FromArgb(230, 230, 235) };
            btnClose = new Button
            {
                Text = "Regresar a Resultados",
                Size = new Size(200, 40),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(180, 50, 50),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => this.Close();

            Button btnSave = new Button
            {
                Text = "Guardar Imagen",
                Size = new Size(200, 40),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(40, 120, 80),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            pnlBottom.Controls.Add(btnClose);
            pnlBottom.Controls.Add(btnSave);
            
            pnlBottom.Resize += (s, e) => {
                int totalWidth = btnClose.Width + btnSave.Width + 20;
                int startX = (pnlBottom.Width - totalWidth) / 2;
                btnClose.Left = startX;
                btnClose.Top = (pnlBottom.Height - btnClose.Height) / 2;
                btnSave.Left = startX + btnClose.Width + 20;
                btnSave.Top = (pnlBottom.Height - btnSave.Height) / 2;
            };

            // 3. Área Central con SplitContainer
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.None,
                SplitterDistance = 820, 
                IsSplitterFixed = false,
                BorderStyle = BorderStyle.None,
                BackColor = System.Drawing.Color.FromArgb(220, 220, 230)
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
                BackColor = System.Drawing.Color.FromArgb(252, 252, 255),
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 11f),
                ForeColor = System.Drawing.Color.FromArgb(10, 10, 30),
                Padding = new Padding(25)
            };

            split.Panel1.Controls.Add(chartFull);
            split.Panel2.Controls.Add(txtAdvice);
            split.Panel2.Padding = new Padding(15);
            split.Panel2.BackColor = System.Drawing.Color.White;

            this.Controls.Add(split);      
            this.Controls.Add(pnlControls);
            this.Controls.Add(lblTitle);    
            this.Controls.Add(pnlBottom);   

            split.BringToFront();
            pnlControls.SendToBack();
            lblTitle.SendToBack();
            pnlBottom.SendToBack();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Imagen PNG (*.png)|*.png|Imagen JPEG (*.jpg)|*.jpg";
                sfd.Title = "Guardar Diagnóstico CIELAB";
                sfd.FileName = "Diagnostico_CIELAB.png";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (Bitmap bmp = new Bitmap(this.Width, this.Height))
                        {
                            // Captura toda la ventana con la gráfica completa e información.
                            this.DrawToBitmap(bmp, new Rectangle(0, 0, this.Width, this.Height));
                            
                            var format = sfd.FileName.EndsWith(".jpg") ? 
                                System.Drawing.Imaging.ImageFormat.Jpeg : 
                                System.Drawing.Imaging.ImageFormat.Png;
                                
                            bmp.Save(sfd.FileName, format);
                        }
                        MessageBox.Show("Imagen guardada exitosamente.", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error al guardar la imagen:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
