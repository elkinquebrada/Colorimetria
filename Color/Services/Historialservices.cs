using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace Color.Services
{
    public static class HistorialService
    {
        private static string rutaArchivo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Colorimetria.csv");

        public static void GuardarMedicionDetallada(
            DateTime fecha, string shadeName, string iluminante,
            string lightnessPct, string chromaPct,
            string diagL, string corrL,
            string diagA, string corrA,
            string diagB, string corrB)
        {
            try
            {
                if (!File.Exists(rutaArchivo))
                {
                    string headers = "FechaHora;ShadeName;Iluminante;Lightness;Chroma;Diagnostico_L;Correccion_L;Diagnostico_a;Correccion_a;Diagnostico_b;Correccion_b" + Environment.NewLine;
                    File.WriteAllText(rutaArchivo, headers, Encoding.UTF8);
                }

                string linea = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}",
                    fecha.ToString("yyyy-MM-dd HH:mm"),
                    shadeName ?? "N/A",
                    iluminante ?? "N/A",
                    lightnessPct ?? "N/A",
                    chromaPct ?? "N/A",
                    diagL ?? "N/A",
                    corrL ?? "N/A",
                    diagA ?? "N/A",
                    corrA ?? "N/A",
                    diagB ?? "N/A",
                    corrB ?? "N/A");

                File.AppendAllText(rutaArchivo, linea + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        public static DataTable ObtenerHistorial()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("FechaHora");
            dt.Columns.Add("ShadeName");
            dt.Columns.Add("Iluminante");
            dt.Columns.Add("Lightness");
            dt.Columns.Add("Chroma");
            dt.Columns.Add("Diagnostico_L");
            dt.Columns.Add("Correccion_L");
            dt.Columns.Add("Diagnostico_a");
            dt.Columns.Add("Correccion_a");
            dt.Columns.Add("Diagnostico_b");
            dt.Columns.Add("Correccion_b");

            try
            {
                if (File.Exists(rutaArchivo))
                {
                    string[] lineas = File.ReadAllLines(rutaArchivo, Encoding.UTF8);
                    for (int i = 1; i < lineas.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(lineas[i])) continue;
                        string[] celdas = lineas[i].Split(';');
                        if (celdas.Length == 11) dt.Rows.Add(celdas);
                    }
                }
            }
            catch { }

            return dt;
        }

        public static void GuardarHistorialCompleto(DataTable dt)
        {
            try
            {
                string headers = "FechaHora;ShadeName;Iluminante;Lightness;Chroma;Diagnostico_L;Correccion_L;Diagnostico_a;Correccion_a;Diagnostico_b;Correccion_b" + Environment.NewLine;
                File.WriteAllText(rutaArchivo, headers, Encoding.UTF8);

                foreach (DataRow row in dt.Rows)
                {
                    string linea = string.Format("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}",
                        row["FechaHora"] ?? "N/A", row["ShadeName"] ?? "N/A", row["Iluminante"] ?? "N/A",
                        row["Lightness"] ?? "N/A", row["Chroma"] ?? "N/A", row["Diagnostico_L"] ?? "N/A",
                        row["Correccion_L"] ?? "N/A", row["Diagnostico_a"] ?? "N/A", row["Correccion_a"] ?? "N/A",
                        row["Diagnostico_b"] ?? "N/A", row["Correccion_b"] ?? "N/A");

                    File.AppendAllText(rutaArchivo, linea + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}