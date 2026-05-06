using System;
using System.Data;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Globalization;

namespace Color.Services
{
    public static class HistorialService
    {
        private static string rutaArchivo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Coats_Consolidado.csv");

        // PK: ShadeName + DyeCode (evita duplicados en el historial)
        public static void GuardarRegistroMaestro(
            string shadeName, 
            DateTime fecha, 
            string iluminante,
            double dlEje, double dcEje, double dhEje,
            string dyeCode, string dyeName, 
            decimal concOriginal, 
            decimal ajusteDL, decimal ajusteDC, decimal ajusteDH, 
            decimal nuevaReceta)
        {
            try
            {
                var ci = CultureInfo.InvariantCulture;
                string nuevaLinea = string.Format(ci, "{0};{1};{2};{3:F5};{4:F5};{5:F5};{6};{7};{8:F5};{9:F5};{10:F5};{11:F5};{12:F5}",
                    shadeName ?? "N/A",
                    fecha.ToString("dd/MM/yyyy HH:mm"),
                    iluminante ?? "D65",
                    dlEje, dcEje, dhEje,
                    dyeCode ?? "0",
                    dyeName ?? "Unknown",
                    concOriginal,
                    ajusteDL,
                    ajusteDC,
                    ajusteDH,
                    nuevaReceta);

                // Si el archivo no existe, crearlo con encabezado y la nueva línea
                if (!File.Exists(rutaArchivo))
                {
                    string header = "ShadeName;FechaHora;Iluminante;DLEje;DCEje;DHEje;DyeCode;DyeName;ConcOriginal;AjusteDL;AjusteDC;AjusteDH;NuevaReceta" + Environment.NewLine;
                    File.WriteAllText(rutaArchivo, header + nuevaLinea + Environment.NewLine, Encoding.UTF8);
                    return;
                }

                // Leer todas las líneas existentes
                string[] lineasExistentes = File.ReadAllLines(rutaArchivo, Encoding.UTF8);
                bool registroActualizado = false;
                var nuevasLineas = new List<string>();

                // Preservar el encabezado
                if (lineasExistentes.Length > 0)
                    nuevasLineas.Add(lineasExistentes[0]);

                // Clave de unicidad: ShadeName + DyeCode
                string claveNueva = $"{(shadeName ?? "N/A").Trim().ToUpper()};{(dyeCode ?? "0").Trim().ToUpper()}";

                for (int i = 1; i < lineasExistentes.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lineasExistentes[i])) continue;
                    string[] celdas = lineasExistentes[i].Split(';');
                    if (celdas.Length >= 7)
                    {
                        string claveExistente = $"{celdas[0].Trim().ToUpper()};{celdas[6].Trim().ToUpper()}";
                        if (claveExistente == claveNueva)
                        {
                            // Reemplazar con los datos más recientes
                            nuevasLineas.Add(nuevaLinea);
                            registroActualizado = true;
                            continue;
                        }
                    }
                    nuevasLineas.Add(lineasExistentes[i]);
                }

                // Si no existía, agregar como registro nuevo
                if (!registroActualizado)
                    nuevasLineas.Add(nuevaLinea);

                File.WriteAllLines(rutaArchivo, nuevasLineas, Encoding.UTF8);
            }
            catch { }
        }

        public static DataTable ObtenerHistorial()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ShadeName");
            dt.Columns.Add("FechaHora");
            dt.Columns.Add("Iluminante");
            dt.Columns.Add("DLEje");
            dt.Columns.Add("DCEje");
            dt.Columns.Add("DHEje");
            dt.Columns.Add("DyeCode");
            dt.Columns.Add("DyeName");
            dt.Columns.Add("ConcOriginal");
            dt.Columns.Add("AjusteDL");
            dt.Columns.Add("AjusteDC");
            dt.Columns.Add("AjusteDH");
            dt.Columns.Add("NuevaReceta");

            try
            {
                if (File.Exists(rutaArchivo))
                {
                    string[] lineas = File.ReadAllLines(rutaArchivo, Encoding.UTF8);
                    for (int i = 1; i < lineas.Length; i++) 
                    {
                        if (string.IsNullOrWhiteSpace(lineas[i])) continue;
                        string[] celdas = lineas[i].Split(';');
                        
                        // Compatibilidad con versiones anteriores (12 columnas) o actual (13 columnas)
                        if (celdas.Length == 12)
                        {
                            // Insertar "0" en la posición de AjusteDC (índice 10)
                            var lista = new List<string>(celdas);
                            lista.Insert(10, "0");
                            dt.Rows.Add(lista.ToArray());
                        }
                        else if (celdas.Length >= 13)
                        {
                            dt.Rows.Add(celdas);
                        }
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
                string headers = "ShadeName;FechaHora;Iluminante;DLEje;DCEje;DHEje;DyeCode;DyeName;ConcOriginal;AjusteDL;AjusteDC;AjusteDH;NuevaReceta" + Environment.NewLine;
                File.WriteAllText(rutaArchivo, headers, Encoding.UTF8);

                var ci = CultureInfo.InvariantCulture;
                foreach (DataRow row in dt.Rows)
                {
                    string linea = string.Format(ci, "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12}",
                        row["ShadeName"], row["FechaHora"], row["Iluminante"],
                        row["DLEje"], row["DCEje"], row["DHEje"],
                        row["DyeCode"], row["DyeName"], 
                        row["ConcOriginal"], row["AjusteDL"], row["AjusteDC"], row["AjusteDH"], row["NuevaReceta"]);

                    File.AppendAllText(rutaArchivo, linea + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}
