using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace Color
{
    public partial class FormMainShell : Form
    {
        private Form _activeModule = null;
        private Button _activeButton = null;
        private System.Collections.Generic.Dictionary<string, Form> _modulos = new System.Collections.Generic.Dictionary<string, Form>();

        public FormMainShell()
        {
            InitializeComponent();
            this.Text = "COATS CADENA - TINT COATS UNIFICADO";
            this.DoubleBuffered = true;
            
            // Inicializar módulos base
            _modulos["SCAN"] = new Form1();
            _modulos["HISTORY"] = new FormHistorial();

            // Inicio por defecto
            NavigateTo(_modulos["SCAN"], btnNavScan);

            // Arquitectura de Eventos: El Cerebro (Shell) escucha al Corazón (Session)
            AnalysisSession.Instance.DataUpdated += (s, e) => {
                this.Invoke((MethodInvoker)delegate {
                    SetAnalysisEnabled(AnalysisSession.Instance.HasActiveData);
                });
            };

            // Estado inicial
            SetAnalysisEnabled(AnalysisSession.Instance.HasActiveData);
        }

        public void SetAnalysisEnabled(bool enabled)
        {
            btnNavDashboard.Enabled = enabled;
            btnNavCielab.Enabled = enabled;
            
            // Estilo visual para deshabilitado
            System.Drawing.Color activeColor = enabled ? System.Drawing.Color.White : System.Drawing.Color.FromArgb(100, 100, 120);
            btnNavDashboard.ForeColor = activeColor;
            btnNavCielab.ForeColor = activeColor;
        }

        public void NavigateTo(Form module, Button navButton)
        {
            if (module == null) return;

            // 1. UI Navigation State
            if (_activeButton != null) _activeButton.BackColor = System.Drawing.Color.Transparent;
            _activeButton = navButton;
            if (_activeButton != null) _activeButton.BackColor = System.Drawing.Color.FromArgb(45, 126, 247);

            // 2. Module Management (Sin .Clear() para persistencia)
            if (_activeModule != null) _activeModule.Hide();

            _activeModule = module;

            if (!pnlContent.Controls.Contains(module))
            {
                module.TopLevel = false;
                module.FormBorderStyle = FormBorderStyle.None;
                module.Dock = DockStyle.Fill;
                pnlContent.Controls.Add(module);
            }
            
            module.BringToFront();
            module.Show();

            lblStatusInfo.Text = $"Entorno Unificado | {navButton?.Text.Trim() ?? "Análisis Activo"}";
        }

        // Método para registrar/actualizar el Dashboard de resultados
        public void UpdateDashboard(FormResultados frm = null)
        {
            if (frm == null) frm = new FormResultados();

            if (_modulos.ContainsKey("DASHBOARD"))
            {
                _modulos["DASHBOARD"].Dispose();
            }
            _modulos["DASHBOARD"] = frm;
            NavigateTo(frm, btnNavDashboard);
        }

        // Acciones de la barra lateral
        private void btnNavScan_Click(object sender, EventArgs e) => NavigateTo(_modulos["SCAN"], btnNavScan);
        private void btnNavHistory_Click(object sender, EventArgs e) => NavigateTo(_modulos["HISTORY"], btnNavHistory);
        
        private void btnNavCielab_Click(object sender, EventArgs e)
        {
            if (!AnalysisSession.Instance.HasActiveData)
            {
                MessageBox.Show("No hay datos de análisis disponibles.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try 
            {
                // Encontrar la corrección D65 para el gráfico
                dynamic d65 = null;
                foreach (var item in AnalysisSession.Instance.CurrentCorrections)
                {
                    dynamic d = item;
                    if (d.Illuminant == "D65") { d65 = d; break; }
                }
                
                if (d65 == null) 
                {
                    MessageBox.Show("No se encontraron datos del iluminante D65 para generar el gráfico.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var frmCielab = new FormDetalleCielab(d65.DL, d65.Da, d65.Db, d65.DE, d65.CmcValue, 1.0, "Análisis unificado", 50, 0, 0);
                NavigateTo(frmCielab, btnNavCielab);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al generar el gráfico CIELAB: " + ex.Message, "Error de Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void btnNavDashboard_Click(object sender, EventArgs e)
        {
            if (_modulos.ContainsKey("DASHBOARD"))
            {
                NavigateTo(_modulos["DASHBOARD"], btnNavDashboard);
            }
            else
            {
                MessageBox.Show("No hay un análisis activo. Realice una lectura OCR primero.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnNavConfig_Click(object sender, EventArgs e)
        {
            using (var config = new Color.Tolerancias.FormConfigTolerancias())
            {
                config.ShowDialog();
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Finalizar sesión de trabajo?", "Coats Cadena", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
