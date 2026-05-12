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

        // PERSISTENCIA INDUSTRIAL: Guardado por anexado (Append-Only) para trazabilidad total
        public static void GuardarRegistroMaestro(
            string shadeName, 
            DateTime fecha, 
            string iluminante,
            string dyeName, 
            decimal concOriginal, 
            string r1, string r2, string r3,
            string impL = "", string diagL = "", string recL = "",
            string impC = "", string diagC = "", string recC = "",
            string impH = "", string diagH = "", string recH = "",
            string factorA = "0", string factorB = "0", string deltaE = "0")
        {
            try
            {
                var ci = CultureInfo.InvariantCulture;
                string nuevaLinea = string.Format(ci, "{0};{1};{2};{3};{4:F5};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};{18};{19}",
                    shadeName ?? "N/A",
                    fecha.ToString("dd/MM/yyyy HH:mm"),
                    iluminante ?? "D65",
                    dyeName ?? "Unknown",
                    concOriginal,
                    r1 ?? "---", r2 ?? "---", r3 ?? "---",
                    impL ?? "", diagL ?? "", recL ?? "",
                    impC ?? "", diagC ?? "", recC ?? "",
                    impH ?? "", diagH ?? "", recH ?? "",
                    factorA ?? "0", factorB ?? "0", deltaE ?? "0");

                // Si el archivo no existe, crearlo con encabezado
                if (!File.Exists(rutaArchivo))
                {
                    string header = "ShadeName;FechaHora;Iluminante;DyeName;ConcOriginal;Receta1;Receta2;Receta3;Impactodl;Diagdl;Recdl;Impactoda;Diagda;Recda;Impactodb;Diagdb;Recdb;FactorA;FactorB;DeltaE" + Environment.NewLine;
                    File.WriteAllText(rutaArchivo, header, Encoding.UTF8);
                }

                // Anexar directamente (Permite múltiples registros por ShadeName - Paridad Industrial)
                File.AppendAllText(rutaArchivo, nuevaLinea + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        public static DataTable ObtenerHistorial()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ShadeName");
            dt.Columns.Add("FechaHora");
            dt.Columns.Add("Iluminante");
            dt.Columns.Add("DyeName");
            dt.Columns.Add("ConcOriginal");
            dt.Columns.Add("Receta1");
            dt.Columns.Add("Receta2");
            dt.Columns.Add("Receta3");
            dt.Columns.Add("Impactodl");
            dt.Columns.Add("Diagdl");
            dt.Columns.Add("Recdl");
            dt.Columns.Add("Impactoda");
            dt.Columns.Add("Diagda");
            dt.Columns.Add("Recda");
            dt.Columns.Add("Impactodb");
            dt.Columns.Add("Diagdb");
            dt.Columns.Add("Recdb");
            dt.Columns.Add("FactorA");
            dt.Columns.Add("FactorB");
            dt.Columns.Add("DeltaE");

            try
            {
                if (File.Exists(rutaArchivo))
                {
                    string[] lineas = File.ReadAllLines(rutaArchivo, Encoding.UTF8);
                    for (int i = 1; i < lineas.Length; i++) 
                    {
                        if (string.IsNullOrWhiteSpace(lineas[i])) continue;
                        string[] celdas = lineas[i].Split(';');
                        
                        var lista = new List<string>();
                        
                        if (celdas.Length == 20)
                        {
                            dt.Rows.Add(celdas);
                        }
                        else if (celdas.Length == 21)
                        {
                            // Migración: Omitir DyeCode (índice 3)
                            for (int j = 0; j <= 2; j++) lista.Add(celdas[j]);
                            for (int j = 4; j < celdas.Length; j++) lista.Add(celdas[j]);
                            dt.Rows.Add(lista.ToArray());
                        }
                        else if (celdas.Length == 18)
                        {
                            // Formato antiguo de 18: [0:Shade, 1:Fecha, 2:Ilu, 3:Code, 4:Name, 5:Conc, 6:R1, 7:R2, 8:R3, ...]
                            for (int j = 0; j <= 2; j++) lista.Add(celdas[j]);
                            lista.Add(celdas[4]); // Name
                            for (int j = 5; j < 18; j++) lista.Add(celdas[j]);
                            lista.Add("0"); lista.Add("0"); lista.Add("0"); 
                            dt.Rows.Add(lista.ToArray());
                        }
                        else
                        {
                            // Fallback genérico
                            for (int j = 0; j < Math.Min(celdas.Length, 20); j++) lista.Add(celdas[j]);
                            while (lista.Count < 20) lista.Add("0");
                            dt.Rows.Add(lista.ToArray());
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
                string headers = "ShadeName;FechaHora;Iluminante;DyeName;ConcOriginal;Receta1;Receta2;Receta3;Impactodl;Diagdl;Recdl;Impactoda;Diagda;Recda;Impactodb;Diagdb;Recdb;FactorA;FactorB;DeltaE" + Environment.NewLine;
                File.WriteAllText(rutaArchivo, headers, Encoding.UTF8);

                var ci = CultureInfo.InvariantCulture;
                foreach (DataRow row in dt.Rows)
                {
                    string linea = string.Format(ci, "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};{18};{19}",
                        row["ShadeName"], row["FechaHora"], row["Iluminante"],
                        row["DyeName"], 
                        row["ConcOriginal"], row["Receta1"], row["Receta2"], row["Receta3"],
                        row["Impactodl"], row["Diagdl"], row["Recdl"],
                        row["Impactoda"], row["Diagda"], row["Recda"],
                        row["Impactodb"], row["Diagdb"], row["Recdb"],
                        row["FactorA"], row["FactorB"], row["DeltaE"]);

                    File.AppendAllText(rutaArchivo, linea + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}
