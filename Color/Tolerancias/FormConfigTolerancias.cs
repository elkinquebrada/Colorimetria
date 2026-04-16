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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al leer el archivo Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _profiles = new List<ToleranceResult>();
            }
        }

        private void RenderCards()
        {
            flowCards.Controls.Clear();
            
            if (_profiles.Count > 0)
            {
                int cardWidth = 110;
                int gap = 20;
                int totalWidth = _profiles.Count * cardWidth + (_profiles.Count - 1) * gap;
                int leftPadding = (flowCards.Width - totalWidth) / 2;
                if (leftPadding < 0) leftPadding = 0;
                flowCards.Padding = new Padding(leftPadding, 20, 0, 0);
            }

            double currentDE = Math.Round(Properties.Settings.Default.ToleranciaDE, 2);

            foreach (var profile in _profiles)
            {
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
                    Text = $"DE {profile.DE.ToString("0.00", CultureInfo.InvariantCulture)}",
                    Dock = DockStyle.Top,
                    Height = 35,
                    BackColor = System.Drawing.Color.FromArgb(43, 142, 227), 
                    ForeColor = System.Drawing.Color.White,
                    Font = new Font("Segoe UI", 11F, FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleCenter
                };
                
                // Cuerpo con DL, DC, DH
                var lblBody = new Label
                {
                    Text = $"\nDL  {profile.DL.ToString("0.000", CultureInfo.InvariantCulture)}\nDC  {profile.DC.ToString("0.000", CultureInfo.InvariantCulture)}\nDH  {profile.DH.ToString("0.000", CultureInfo.InvariantCulture)}",
                    Dock = DockStyle.Fill,
                    Font = new Font("Segoe UI", 11F, FontStyle.Regular),
                    TextAlign = ContentAlignment.TopCenter,
                    ForeColor = System.Drawing.Color.Black
                };

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
            // Reset state
            if (_selectedPanel != null)
            {
                _selectedPanel.BorderStyle = BorderStyle.FixedSingle;
                _selectedPanel.BackColor = System.Drawing.Color.White;
                if (_selectedPanel.Controls.Count > 1 && _selectedPanel.Controls[1] is Label lblHeader)
                {
                    lblHeader.BackColor = System.Drawing.Color.FromArgb(43, 142, 227);
                }
            }

            // Apply selected state
            _selectedPanel = pnl;
            _selectedProfile = profile;
            
            pnl.BorderStyle = BorderStyle.Fixed3D; 
            pnl.BackColor = System.Drawing.Color.AliceBlue;
            
            if (_selectedPanel.Controls.Count > 1 && _selectedPanel.Controls[1] is Label header)
            {
                header.BackColor = System.Drawing.Color.FromArgb(20, 100, 180); 
            }
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