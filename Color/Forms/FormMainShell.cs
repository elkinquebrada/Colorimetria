using System;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace Color
{
    public partial class FormMainShell : Form
    {
        private Form _formularioActivo = null;
        private Button _botonActivo = null;

        // Instancias persistentes para no perder datos al navegar
        private Form1 _formOcr;
        private FormHistorial _formHistorial;

        public FormMainShell()
        {
            InitializeComponent();
            SetupDashboard();
        }

        private void SetupDashboard()
        {
            // Inicializar formularios hijos
            _formOcr = new Form1();
            _formHistorial = new FormHistorial();

            // Cargar la pantalla de inicio (OCR) por defecto
            AbrirFormularioEnWorkspace(_formOcr, btnNavOcr);
        }

        public void AbrirFormularioEnWorkspace(Form formularioHijo, Button btnNav)
        {
            if (formularioHijo == null) return;

            // 1. Resaltar botón activo
            if (_botonActivo != null)
            {
                _botonActivo.BackColor = System.Drawing.Color.FromArgb(0, 51, 153); // Azul oscuro
            }
            _botonActivo = btnNav;
            if (_botonActivo != null)
            {
                _botonActivo.BackColor = System.Drawing.Color.FromArgb(45, 126, 247); // Azul claro resaltado
            }

            // 2. Si ya está activo, no hacer nada
            if (_formularioActivo == formularioHijo) return;

            // 3. Limpiar panel y preparar formulario
            pnlWorkspace.Controls.Clear();
            _formularioActivo = formularioHijo;

            formularioHijo.TopLevel = false;
            formularioHijo.FormBorderStyle = FormBorderStyle.None;
            formularioHijo.Dock = DockStyle.Fill;
            
            pnlWorkspace.Controls.Add(formularioHijo);
            pnlWorkspace.Tag = formularioHijo;
            formularioHijo.BringToFront();
            formularioHijo.Show();

            lblSeccionActiva.Text = btnNav?.Text.Replace("  ", "") ?? "Dashboard";
        }

        // Método público para ser llamado desde Form1 cuando se obtienen resultados
        public void MostrarResultados(FormResultados frmRes)
        {
            AbrirFormularioEnWorkspace(frmRes, btnNavAnalisis);
        }

        private void btnNavOcr_Click(object sender, EventArgs e) => AbrirFormularioEnWorkspace(_formOcr, btnNavOcr);
        private void btnNavHistorial_Click(object sender, EventArgs e) => AbrirFormularioEnWorkspace(_formHistorial, btnNavHistorial);
        private void btnNavAnalisis_Click(object sender, EventArgs e) 
        {
            // El análisis solo se habilita si hay datos; si no, avisamos
            if (pnlWorkspace.Controls.Count > 0 && pnlWorkspace.Controls[0] is FormResultados) return;
            MessageBox.Show("Por favor, realice una lectura OCR o cargue un análisis primero.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea salir de la aplicación?", "Confirmar Salida", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}
