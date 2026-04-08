using System;
using System.Globalization;
using System.Windows.Forms;

namespace Color.Tolerancias
{
    /// Formulario de configuración para establecer los límites de tolerancia colorimétrica.
    public partial class FormConfigTolerancias : Form
    {
        /// Inicializa una nueva instancia de la clase.
        public FormConfigTolerancias()
        {
            InitializeComponent();
            CargarValoresActuales();
        }

        /// Recupera las tolerancias almacenadas en la configuración persistente del usuario 
        private void CargarValoresActuales()
        {
            txtDL.Text = Properties.Settings.Default.ToleranciaDL.ToString(CultureInfo.InvariantCulture);
            txtDC.Text = Properties.Settings.Default.ToleranciaDC.ToString(CultureInfo.InvariantCulture);
            txtDH.Text = Properties.Settings.Default.ToleranciaDH.ToString(CultureInfo.InvariantCulture);
            txtDE.Text = Properties.Settings.Default.ToleranciaDE.ToString(CultureInfo.InvariantCulture);
        }

        /// Procesa y guarda los nuevos valores de tolerancia ingresados por el usuario.
        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                // Aceptar tanto punto como coma decimal
                string dlStr = txtDL.Text.Replace(',', '.');
                string dhStr = txtDH.Text.Replace(',', '.');
                string deStr = txtDE.Text.Replace(',', '.');

                // Conversión de strings a valores numéricos (double)
                string dcStr = txtDC.Text.Replace(',', '.');
                double dl = double.Parse(dlStr, CultureInfo.InvariantCulture);
                double dc = double.Parse(dcStr, CultureInfo.InvariantCulture);
                double dh = double.Parse(dhStr, CultureInfo.InvariantCulture);
                double de = double.Parse(deStr, CultureInfo.InvariantCulture);

                // Asignación de valores a la configuración persistente de la aplicación
                Properties.Settings.Default.ToleranciaDL = dl;
                Properties.Settings.Default.ToleranciaDC = dc;
                Properties.Settings.Default.ToleranciaDH = dh;
                Properties.Settings.Default.ToleranciaDE = de;

                // Persistencia de los datos en el archivo de configuración del usuario
                Properties.Settings.Default.Save();

                MessageBox.Show("Configuración guardada con éxito.", "Éxito",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (FormatException)
            {
                // Manejo de error en caso de que el usuario ingrese caracteres no numéricos
                MessageBox.Show("Por favor ingresa solo números válidos (ej: 1.20).",
                    "Error de formato", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// Cierra el formulario sin aplicar ni guardar ningún cambio.
        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}