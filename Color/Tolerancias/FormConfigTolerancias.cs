using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Color;

namespace Color.Tolerancias
{
    /// Formulario de configuración para establecer los límites de tolerancia colorimétrica.
    public partial class FormConfigTolerancias : Form
    {
        private List<ToleranceResult> _profiles;
        private ToleranceResult _selectedProfile = null;
        private Panel _selectedPanel = null;

        public FormConfigTolerancias()
        {
            InitializeComponent();
            this.Load += FormConfigTolerancias_Load;
        }

        private void FormConfigTolerancias_Load(object sender, EventArgs e)
        {
            LoadProfiles();
            RenderCards();
        }

        private void LoadProfiles()
        {
            try
            {
                // Ruta relativa al excel de configuración dentro del proyecto
                string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"LogicDocs\calculos a realizar con el programa.xlsx");
                
                if (!File.Exists(excelPath))
                {
                    // Fallback para depuración en VS
                    excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\LogicDocs\calculos a realizar con el programa.xlsx");
                }

                if (File.Exists(excelPath))
                {
                    _profiles = OCR.ExcelReader.LoadTolerances(excelPath);
                }
                else
                {
                    MessageBox.Show("No se encontró el archivo Excel de configuración:\n" + excelPath, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    _profiles = new List<ToleranceResult>();
                }

                // Asegurar que siempre tengamos al menos 4 perfiles (el 4to es para entrada manual)
                while (_profiles.Count < 4)
                {
                    _profiles.Add(new ToleranceResult { DE = 0, DL = 0, DC = 0, DH = 0 });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _profiles = new List<ToleranceResult>();
                
                // Fallback de emergencia
                while (_profiles.Count < 4)
                {
                    _profiles.Add(new ToleranceResult { DE = 0, DL = 0, DC = 0, DH = 0 });
                }
            }
        }

        private void RenderCards()
        {
            flowCards.Controls.Clear();
            
            if (_profiles.Count > 0)
            {
                int cardWidth = 110;

                // Reducido para asegurar que quepan 4
                int gap = 10; 
                int totalWidth = _profiles.Count * cardWidth + (_profiles.Count - 1) * (gap + 20); 

                int leftPadding = (flowCards.Width - totalWidth) / 2;
                if (leftPadding < 0) leftPadding = 0;
                flowCards.Padding = new Padding(leftPadding, 20, 0, 0);
            }

            double currentDE = Math.Round(Properties.Settings.Default.ToleranciaDE, 2);

            for (int i = 0; i < _profiles.Count; i++)
            {
                var profile = _profiles[i];
                var pnlCard = new Panel
                {
                    Width = 110,
                    Height = 170,
                    Margin = new Padding(10),
                    BorderStyle = BorderStyle.FixedSingle,
                    BackColor = System.Drawing.Color.White,
                    Cursor = Cursors.Hand,
                    Tag = profile
                };

                // Encabezado Azul
                var lblHeader = new Label
                {
                    Text = (i == 3) ? "Ingresa el DE" : $"DE {profile.DE.ToString("0.00", CultureInfo.InvariantCulture)}",
                    Dock = DockStyle.Top,
                    Height = 35,
                    BackColor = System.Drawing.Color.FromArgb(43, 142, 227), 
                    ForeColor = System.Drawing.Color.White,
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                
                // Cuerpo con DL, DC, DH
                var lblBody = new Label
                {
                    Text = $"\nDL  {profile.DL.ToString("0.000", CultureInfo.InvariantCulture)}\nDC  {profile.DC.ToString("0.000", CultureInfo.InvariantCulture)}\nDH  {profile.DH.ToString("0.000", CultureInfo.InvariantCulture)}",
                    Dock = DockStyle.Fill,
                    Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                    TextAlign = ContentAlignment.TopCenter,
                    ForeColor = System.Drawing.Color.Black,
                    BackColor = System.Drawing.Color.Transparent
                };

                // Si es el perfil manual (4to), añadimos un TextBox para el DE
                if (i == 3)
                {
                    // Ajustamos el body para dejar espacio al textbox
                    lblBody.Padding = new Padding(0, 32, 0, 0);

                    var txtDE = new TextBox
                    {
                        Width = 80,
                        Location = new Point(15, 45),
                        TextAlign = HorizontalAlignment.Center,
                        Text = profile.DE > 0 ? profile.DE.ToString("0.00", CultureInfo.InvariantCulture) : "",
                        Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                        BorderStyle = BorderStyle.FixedSingle
                    };

                    txtDE.TextChanged += (s, ev) =>
                    {
                        string valStr = txtDE.Text.Replace(',', '.');
                        if (double.TryParse(valStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                        {
                            var newTol = ColorimetricCalculator.CalculateTolerance(val);
                            profile.DE = newTol.DE;
                            profile.DL = newTol.DL;
                            profile.DC = newTol.DC;
                            profile.DH = newTol.DH;
                            
                            lblBody.Text = $"\nDL  {profile.DL.ToString("0.000", CultureInfo.InvariantCulture)}\nDC  {profile.DC.ToString("0.000", CultureInfo.InvariantCulture)}\nDH  {profile.DH.ToString("0.000", CultureInfo.InvariantCulture)}";
                        }
                    };

                    // Asegurar que el click en el textbox no rompa la selección pero permita escribir
                    txtDE.GotFocus += (s, ev) => { SelectCard(pnlCard, profile); };

                    pnlCard.Controls.Add(txtDE);
                    txtDE.BringToFront();
                }

                // Suscribir eventos de click
                pnlCard.Click += Card_Click;
                lblHeader.Click += (s, e) => Card_Click(pnlCard, e);
                lblBody.Click += (s, e) => Card_Click(pnlCard, e);

                pnlCard.Controls.Add(lblBody);
                pnlCard.Controls.Add(lblHeader);

                flowCards.Controls.Add(pnlCard);

                // Preseleccionar si coincide con el perfil global guardado
                if (Math.Abs(profile.DE - currentDE) < 0.01)
                {
                    SelectCard(pnlCard, profile);
                }
            }

            // asignacion por defecto del escenario DE 1.20 
            if (_selectedProfile == null && _profiles.Count > 0)
            {
                var defaultProfile = _profiles.FirstOrDefault(p => Math.Abs(p.DE - 1.20) < 0.01) ?? _profiles.FirstOrDefault();
                if (defaultProfile != null)
                {
                    var pnl = flowCards.Controls.Cast<Panel>().FirstOrDefault(p => p.Tag == defaultProfile);
                    if (pnl != null) SelectCard(pnl, defaultProfile);
                }
            }
        }

        private void Card_Click(object sender, EventArgs e)
        {
            var pnl = sender as Panel;
            if (pnl != null && pnl.Tag is ToleranceResult profile)
            {
                SelectCard(pnl, profile);
            }
        }

        private void SelectCard(Panel pnl, ToleranceResult profile)
        {
            // Reset previous selection
            if (_selectedPanel != null)
            {
                _selectedPanel.BorderStyle = BorderStyle.FixedSingle;
                _selectedPanel.BackColor = System.Drawing.Color.White;
                
                // Reset header color
                var prevHeader = _selectedPanel.Controls.Cast<Control>().FirstOrDefault(c => c is Label && c.Dock == DockStyle.Top);
                if (prevHeader != null)
                    prevHeader.BackColor = System.Drawing.Color.FromArgb(43, 142, 227);
            }

            // Apply selected state
            _selectedPanel = pnl;
            _selectedProfile = profile;
            
            pnl.BorderStyle = BorderStyle.Fixed3D; 
            pnl.BackColor = System.Drawing.Color.AliceBlue;
            
            // Highlight current header
            var currentHeader = pnl.Controls.Cast<Control>().FirstOrDefault(c => c is Label && c.Dock == DockStyle.Top);
            if (currentHeader != null)
                currentHeader.BackColor = System.Drawing.Color.FromArgb(20, 100, 180); 
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if (_selectedProfile == null)
            {

                // forzamos la búsqueda de DE 1.20
                _selectedProfile = _profiles.FirstOrDefault(p => Math.Abs(p.DE - 1.20) < 0.01) ?? _profiles.FirstOrDefault();
                if (_selectedProfile == null)
                {
                    MessageBox.Show("No se encontraron perfiles de tolerancia disponibles.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Mostrar el cuadro de confirmación
            var confirmResult = MessageBox.Show(
                $"¿Confirmas que deseas aplicar el perfil de tolerancia: DE = {_selectedProfile.DE.ToString("0.00", CultureInfo.InvariantCulture)}?",
                "Confirmar Cambio de Tolerancia",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirmResult == DialogResult.Yes)
            {
                Properties.Settings.Default.ToleranciaDL = _selectedProfile.DL;
                Properties.Settings.Default.ToleranciaDC = _selectedProfile.DC;
                Properties.Settings.Default.ToleranciaDH = _selectedProfile.DH;
                Properties.Settings.Default.ToleranciaDE = _selectedProfile.DE;
                Properties.Settings.Default.Save();

                MessageBox.Show("Configuración guardada y aplicada con éxito a todas las evaluaciones.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}