using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Color;

namespace Color.Tolerancias
{
    public partial class FormConfigTolerancias : Form
    {
        private List<ToleranceResult> _profiles;
        private ToleranceResult _selectedProfile = null;
        private Panel _selectedPanel = null;

        // Perfil manual independiente que inicia en ceros
        private ToleranceResult _manualProfile = new ToleranceResult { DE = 0, DL = 0, DC = 0, DH = 0 };

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
            // NUEVO ORDEN SOLICITADO: 0.60 -> 1.20 -> 1.80
            _profiles = new List<ToleranceResult>
            {
                new ToleranceResult { DE = 0.60, DL = 0.346, DC = 0.346, DH = 0.346 },
                new ToleranceResult { DE = 1.20, DL = 0.693, DC = 0.693, DH = 0.693 },
                new ToleranceResult { DE = 1.80, DL = 1.039, DC = 1.039, DH = 1.039 }
            };
            // La tarjeta dinámica siempre al final (Índice 3)
            _profiles.Add(_manualProfile);
        }

        private void RenderCards()
        {
            flowCards.Controls.Clear();
            // Valor actualmente activo en los Settings del programa
            double activeDE = Math.Round(Properties.Settings.Default.ToleranciaDE, 2);

            for (int i = 0; i < _profiles.Count; i++)
            {
                var profile = _profiles[i];
                bool isManualCard = (i == 3);

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

                var lblHeader = new Label
                {
                    Text = isManualCard ? "Ingresa el DE" : $"DE {profile.DE:0.00}",
                    Dock = DockStyle.Top,
                    Height = 35,
                    BackColor = System.Drawing.Color.FromArgb(43, 142, 227),
                    ForeColor = System.Drawing.Color.White,
                    Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Enabled = false
                };

                var lblBody = new Label
                {
                    Text = $"\nDL  {profile.DL:0.000}\nDC  {profile.DC:0.000}\nDH  {profile.DH:0.000}",
                    Dock = DockStyle.Fill,
                    TextAlign = ContentAlignment.TopCenter,
                    Font = new Font("Segoe UI", 10F),
                    ForeColor = System.Drawing.Color.Black,
                    Enabled = false
                };

                if (isManualCard)
                {
                    lblBody.Padding = new Padding(0, 35, 0, 0);
                    var txtDE = new TextBox
                    {
                        Width = 70,
                        Location = new Point(20, 45),
                        TextAlign = HorizontalAlignment.Center,
                        Text = "", // Mantenemos vacío para evitar confusiones con las estáticas
                        Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                    };

                    txtDE.TextChanged += (s, ev) => {
                        string input = txtDE.Text.Replace(',', '.');
                        if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                        {
                            // Cálculo matemático basado en el ingreso
                            double factor = val / 1.732;
                            _manualProfile.DE = val;
                            _manualProfile.DL = Math.Round(factor, 3);
                            _manualProfile.DC = Math.Round(factor, 3);
                            _manualProfile.DH = Math.Round(factor, 3);

                            // SELECCIÓN AUTOMÁTICA: Al escribir, esta se vuelve la tarjeta activa
                            SelectCard(pnlCard, _manualProfile);
                        }
                        else
                        {
                            _manualProfile.DE = 0; _manualProfile.DL = 0; _manualProfile.DC = 0; _manualProfile.DH = 0;
                        }
                        lblBody.Text = $"\nDL  {_manualProfile.DL:0.000}\nDC  {_manualProfile.DC:0.000}\nDH  {_manualProfile.DH:0.000}";
                    };
                    pnlCard.Controls.Add(txtDE);
                    txtDE.BringToFront();
                }

                pnlCard.Controls.Add(lblBody);
                pnlCard.Controls.Add(lblHeader);
                flowCards.Controls.Add(pnlCard);

                // Evento de clic para seleccionar cualquier tarjeta
                pnlCard.Click += (s, e) => SelectCard(pnlCard, profile);

                // Resaltar la tarjeta que coincide con la configuración actual del programa
                if (Math.Abs(profile.DE - activeDE) < 0.01 && profile.DE > 0)
                    SelectCard(pnlCard, profile);
            }
        }

        private void SelectCard(Panel pnl, ToleranceResult profile)
        {
            if (_selectedPanel != null)
            {
                _selectedPanel.BackColor = System.Drawing.Color.White;
                _selectedPanel.BorderStyle = BorderStyle.FixedSingle;
            }

            _selectedPanel = pnl;
            _selectedProfile = profile; // Vinculamos el perfil actual a la selección
            _selectedPanel.BackColor = System.Drawing.Color.AliceBlue;
            _selectedPanel.BorderStyle = BorderStyle.Fixed3D;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            // Validación de seguridad
            if (_selectedProfile == null || _selectedProfile.DE <= 0)
            {
                MessageBox.Show("Por favor seleccione una tarjeta de tolerancia válida.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ENVÍO DE DATOS AL PROGRAMA:
            // Tomamos los valores exactos del objeto seleccionado (sea estático o el manual modificado)
            Properties.Settings.Default.ToleranciaDE = _selectedProfile.DE;
            Properties.Settings.Default.ToleranciaDL = _selectedProfile.DL;
            Properties.Settings.Default.ToleranciaDC = _selectedProfile.DC;
            Properties.Settings.Default.ToleranciaDH = _selectedProfile.DH;
            Properties.Settings.Default.Save();

            // Mensaje de confirmación con los valores reales enviados
            MessageBox.Show($"Tolerancia Enviada:\nDE: {_selectedProfile.DE:0.00}\nDL: {_selectedProfile.DL:0.000}\nDC: {_selectedProfile.DC:0.000}\nDH: {_selectedProfile.DH:0.000}",
                            "Confirmación de Envío", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e) => this.Close();
    }
}