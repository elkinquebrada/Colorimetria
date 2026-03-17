using System;
using System.Globalization;
using System.Windows.Forms;

namespace Color.Tolerancias
{
    public partial class FormConfigTolerancias : Form
    {
        public FormConfigTolerancias()
        {
            InitializeComponent();
            CargarValoresActuales();
        }

        private void CargarValoresActuales()
        {
            txtDL.Text = Properties.Settings.Default.ToleranciaDL.ToString(CultureInfo.InvariantCulture);
            txtDC.Text = Properties.Settings.Default.ToleranciaDC.ToString(CultureInfo.InvariantCulture);
            txtDH.Text = Properties.Settings.Default.ToleranciaDH.ToString(CultureInfo.InvariantCulture);
            txtDE.Text = Properties.Settings.Default.ToleranciaDE.ToString(CultureInfo.InvariantCulture);
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                // Aceptar tanto punto como coma decimal
                string dlStr = txtDL.Text.Replace(',', '.');
                string dhStr = txtDH.Text.Replace(',', '.');
                string deStr = txtDE.Text.Replace(',', '.');

                string dcStr = txtDC.Text.Replace(',', '.');
                double dl = double.Parse(dlStr, CultureInfo.InvariantCulture);
                double dc = double.Parse(dcStr, CultureInfo.InvariantCulture);
                double dh = double.Parse(dhStr, CultureInfo.InvariantCulture);
                double de = double.Parse(deStr, CultureInfo.InvariantCulture);

                Properties.Settings.Default.ToleranciaDL = dl;
                Properties.Settings.Default.ToleranciaDC = dc;
                Properties.Settings.Default.ToleranciaDH = dh;
                Properties.Settings.Default.ToleranciaDE = de;
                Properties.Settings.Default.Save();

                MessageBox.Show("Configuración guardada con éxito.", "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (FormatException)
            {
                MessageBox.Show("Por favor ingresa solo números válidos (ej: 1.20).",
                    "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}