using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.IO;
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
            this.TopMost = false;
            this.ShowInTaskbar = true;
            this.MinimizeBox = true;
            this.MaximizeBox = true;
            this.Load += FormConfigTolerancias_Load;
            AddBrandingLogo();
        }

        private void AddBrandingLogo()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logicDocs", "Coats_logo.svg.png");
                if (!File.Exists(path)) return;

                var logo = new PictureBox
                {
                    Image = Image.FromFile(path),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Width = 40,
                    Height = 40,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    BackColor = System.Drawing.Color.Transparent
                };
                logo.Location = new Point(this.Width - logo.Width - 15, 10);
                this.Controls.Add(logo);
                logo.BringToFront();
            }
            catch { }
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
            // La tarjeta dinamica siempre al final (Ãndice 3)
            _profiles.Add(_manualProfile);
        }

        private void RenderCards()
        {
            flowCards.Controls.Clear();
            // Valor actualmente activo en los Settings del programa
            double activeDE = Math.Round(Color.Properties.Settings.Default.ToleranciaDE, 2);

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
                    Text = $"\nL  {profile.DL:0.000}\nC  {profile.DC:0.000}\nHue  {profile.DH:0.000}",
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
                        Text = "",
                        Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                    };

                    txtDE.TextChanged += (s, ev) => {
                        string input = txtDE.Text.Replace(',', '.');
                        if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                        {
                            // 1. Asignamos el valor ingresado al Delta E global
                            _manualProfile.DE = val;

                            // 2. IMPLEMENTACION OBLIGATORIA DE LA FORMULA DE EXCEL 
                            double ejeCalculado = Math.Sqrt(Math.Pow(val, 2) / 3);

                            _manualProfile.DL = ejeCalculado;
                            _manualProfile.DC = ejeCalculado;
                            _manualProfile.DH = ejeCalculado;

                            // Seleccionamos la tarjeta automaticamente al escribir
                            SelectCard(pnlCard, _manualProfile);
                        }
                        else
                        {
                            _manualProfile.DE = 0; _manualProfile.DL = 0; _manualProfile.DC = 0; _manualProfile.DH = 0;
                        }

                        // 3. valores calculados con 3 decimales
                        lblBody.Text = $"\nL  {_manualProfile.DL:0.000}\nC  {_manualProfile.DC:0.000}\nHue  {_manualProfile.DH:0.000}";
                    };

                    pnlCard.Controls.Add(txtDE);
                    txtDE.BringToFront();
                }

                pnlCard.Controls.Add(lblBody);
                pnlCard.Controls.Add(lblHeader);
                flowCards.Controls.Add(pnlCard);

                // Evento de clic para seleccionar cualquier tarjeta
                pnlCard.Click += (s, e) => SelectCard(pnlCard, profile);

                // Resaltar la tarjeta que coincide con la configuracion actual del programa
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
            _selectedProfile = profile; 
            _selectedPanel.BackColor = System.Drawing.Color.AliceBlue;
            _selectedPanel.BorderStyle = BorderStyle.Fixed3D;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            // Validacion de seguridad
            if (_selectedProfile == null || _selectedProfile.DE <= 0)
            {
                MessageBox.Show("Por favor seleccione una tarjeta de tolerancia valida.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ENVIO DE DATOS AL PROGRAMA:
            Color.Properties.Settings.Default.ToleranciaDE = _selectedProfile.DE;
            Color.Properties.Settings.Default.ToleranciaDL = _selectedProfile.DL;
            Color.Properties.Settings.Default.ToleranciaDC = _selectedProfile.DC;
            Color.Properties.Settings.Default.ToleranciaDH = _selectedProfile.DH;
            Color.Properties.Settings.Default.Save();

            // Mensaje de confirmacion con los valores reales enviados
            MessageBox.Show($"Tolerancia Enviada:\nDE: {_selectedProfile.DE:0.00}\nL: {_selectedProfile.DL:0.000}\nC: {_selectedProfile.DC:0.000}\nHue: {_selectedProfile.DH:0.000}",
                            "Confirmacion de EnvÃ­o", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();
        }

        private void btnCancelar_Click(object sender, EventArgs e) => this.Close();
    }
}