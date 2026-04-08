using System;
using System.Windows.Forms;

namespace Color
{
    /// Clase contenedora del punto de entrada principal de la aplicación.
    static class Program
    {
        /// Punto de entrada principal para la aplicación.
        [STAThread]
        static void Main()
        {
            // Habilita los estilos visuales modernos (temas de Windows) para la aplicación.
            Application.EnableVisualStyles();
            // Define la resolución de renderizado de texto de forma estándar para mayor compatibilidad.
            Application.SetCompatibleTextRenderingDefault(false);
            // Inicia el bucle de mensajes de la aplicación en el hilo actual 
            // y abre el formulario principal (Form1).
            Application.Run(new Form1());
        }
    }
}