using System;
using System.Data;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Color;
using EngineRes = Color.ColorCorrectionResult;

namespace Color.Services
{
    public static class HistorialService
    {
        private static string rutaArchivo = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DB_Coats_Consolidado.csv");

        // Cadena de conexión - Ajustada tras validación exitosa en (localdb)\MSSQLLocalDB
        private static string connectionString = @"Server=(localdb)\MSSQLLocalDB;Database=ColorimetriaDB;Trusted_Connection=True;Connect Timeout=5;";
        // Alternativas:
        // private static string connectionString = @"Server=.;Database=ColorimetriaDB;Trusted_Connection=True;";
        // private static string connectionString = @"Server=(localdb)\MSSQLLocalDB;Database=ColorimetriaDB;Trusted_Connection=True;";

        // PERSISTENCIA INDUSTRIAL (SQL SERVER V4)
        public static bool GuardarAnalisisCompleto(
            string shadeName,
            string lotNo,
            EngineRes resTL84,
            EngineRes resA,
            List<RecipeItem> recetaOriginal,
            List<double> conOriginales)
        {
            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var trans = conn.BeginTransaction())
                    {
                        try
                        {
                            // 1. Insertar Cabecera
                            string sqlCabecera = @"
                                INSERT INTO tbl_analisis_cabecera 
                                (ShadeName, LotNo, FechaRegistro, DeltaE_TL84, CMC_TL84, Status_TL84, DeltaE_A, CMC_A, Status_A)
                                OUTPUT INSERTED.Id_Lote
                                VALUES 
                                (@Shade, @Lot, GETDATE(), @DeTL, @CmcTL, @StTL, @DeA, @CmcA, @StA)";

                            int idLote;
                            using (var cmd = new System.Data.SqlClient.SqlCommand(sqlCabecera, conn, trans))
                            {
                                cmd.Parameters.AddWithValue("@Shade", shadeName ?? (object)DBNull.Value);
                                cmd.Parameters.AddWithValue("@Lot", lotNo ?? (object)DBNull.Value);
                                
                                cmd.Parameters.AddWithValue("@DeTL", resTL84 != null ? (object)resTL84.DeltaE : DBNull.Value);
                                cmd.Parameters.AddWithValue("@CmcTL", resTL84 != null ? (object)resTL84.CmcValue : DBNull.Value);
                                cmd.Parameters.AddWithValue("@StTL", resTL84 != null ? (object)(resTL84.Pass ? "PASS" : "FAIL") : DBNull.Value);

                                cmd.Parameters.AddWithValue("@DeA", resA != null ? (object)resA.DeltaE : DBNull.Value);
                                cmd.Parameters.AddWithValue("@CmcA", resA != null ? (object)resA.CmcValue : DBNull.Value);
                                cmd.Parameters.AddWithValue("@StA", resA != null ? (object)(resA.Pass ? "PASS" : "FAIL") : DBNull.Value);

                                idLote = (int)cmd.ExecuteScalar();
                            }

                            // 2. Insertar Detalle (Colorantes y sus 3 recetas)
                            // Usamos un solo resultado (normalmente D65 o el principal) para las recetas correctivas
                            // En FormResultados se usa 'res' que suele ser el principal. 
                            // Aquí asumiremos que el llamador ya calculó las recetas en uno de los objetos EngineRes.
                            
                            // Usaremos resTL84 como referencia si resA es nulo para las recetas, 
                            // pero lo ideal es que el llamador pase el objeto que tiene las recetas calculadas.
                            var resRecetas = resTL84 ?? resA;

                            if (recetaOriginal != null && resRecetas != null && resRecetas.RecetaR1_Luminosidad != null)
                            {
                                double totalOri = conOriginales.Sum();
                                double totalR1 = resRecetas.RecetaR1_Luminosidad.Sum();
                                double totalR2 = resRecetas.RecetaR2_Croma.Sum();
                                double totalR3 = resRecetas.RecetaR3_Tono.Sum();

                                string sqlDetalle = @"
                                    INSERT INTO tbl_analisis_detalle 
                                    (Id_Lote, DyeCode, DyeName, Concentration_Original, Proportion_Original, 
                                     R1_Con_Percentage, R1_Part_Percentage, R1_Ajuste_Percentage,
                                     R2_Con_Percentage, R2_Part_Percentage, R2_Ajuste_Percentage,
                                     R3_Con_Percentage, R3_Part_Percentage, R3_Ajuste_Percentage)
                                    VALUES 
                                    (@Id, @Code, @Name, @ConcOri, @PropOri, 
                                     @R1C, @R1P, @R1A, 
                                     @R2C, @R2P, @R2A, 
                                     @R3C, @R3P, @R3A)";

                                for (int i = 0; i < recetaOriginal.Count; i++)
                                {
                                    var ing = recetaOriginal[i];
                                    double cOri = conOriginales[i];
                                    double r1 = resRecetas.RecetaR1_Luminosidad[i];
                                    double r2 = resRecetas.RecetaR2_Croma[i];
                                    double r3 = resRecetas.RecetaR3_Tono[i];

                                    using (var cmdD = new System.Data.SqlClient.SqlCommand(sqlDetalle, conn, trans))
                                    {
                                        cmdD.Parameters.AddWithValue("@Id", idLote);
                                        cmdD.Parameters.AddWithValue("@Code", ing.Code ?? "");
                                        cmdD.Parameters.AddWithValue("@Name", ing.Name ?? "");
                                        cmdD.Parameters.AddWithValue("@ConcOri", (decimal)cOri);
                                        cmdD.Parameters.AddWithValue("@PropOri", (decimal)(totalOri > 0 ? (cOri / totalOri * 100) : 0));

                                        // R1
                                        cmdD.Parameters.AddWithValue("@R1C", (decimal)r1);
                                        cmdD.Parameters.AddWithValue("@R1P", (decimal)(totalR1 > 0 ? (r1 / totalR1 * 100) : 0));
                                        cmdD.Parameters.AddWithValue("@R1A", (decimal)(Math.Abs(cOri > 0 ? (r1 / cOri - 1.0) * 100.0 : 0)));

                                        // R2
                                        cmdD.Parameters.AddWithValue("@R2C", (decimal)r2);
                                        cmdD.Parameters.AddWithValue("@R2P", (decimal)(totalR2 > 0 ? (r2 / totalR2 * 100) : 0));
                                        cmdD.Parameters.AddWithValue("@R2A", (decimal)(Math.Abs(cOri > 0 ? (r2 / cOri - 1.0) * 100.0 : 0)));

                                        // R3
                                        cmdD.Parameters.AddWithValue("@R3C", (decimal)r3);
                                        cmdD.Parameters.AddWithValue("@R3P", (decimal)(totalR3 > 0 ? (r3 / totalR3 * 100) : 0));
                                        cmdD.Parameters.AddWithValue("@R3A", (decimal)(Math.Abs(cOri > 0 ? (r3 / cOri - 1.0) * 100.0 : 0)));

                                        cmdD.ExecuteNonQuery();
                                    }
                                }
                            }

                            trans.Commit();
                            return true;
                        }
                        catch
                        {
                            trans.Rollback();
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Podrías loguear el error aquí
                throw new Exception("Error al guardar en SQL Server: " + ex.Message);
            }
        }

        // PERSISTENCIA LEGACY: Guardado por anexado (Append-Only) para trazabilidad total
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
                    string header = "ShadeName;FechaHora;Iluminante;DyeName;ConcOriginal;Receta1;Receta2;Receta3;Impactodl;Acciondl;Recdl;Impactoda;Accionda;Recda;Impactodb;Acciondb;Recdb;FactorA;FactorB;DeltaE" + Environment.NewLine;
                    File.WriteAllText(rutaArchivo, header, Encoding.UTF8);
                }

                File.AppendAllText(rutaArchivo, nuevaLinea + Environment.NewLine, Encoding.UTF8);
            }
            catch { }
        }

        public static DataTable ObtenerHistorialSQL()
        {
            DataTable dt = new DataTable();
            // Definimos columnas para la estructura V4
            dt.Columns.Add("Id_Lote");
            dt.Columns.Add("ShadeName");
            dt.Columns.Add("LotNo");
            dt.Columns.Add("FechaRegistro");
            dt.Columns.Add("DyeCode");
            dt.Columns.Add("DyeName");
            dt.Columns.Add("Concentration_Original");
            dt.Columns.Add("Proportion_Original");
            dt.Columns.Add("R1_Con");
            dt.Columns.Add("R1_Part");
            dt.Columns.Add("R1_Ajuste");
            dt.Columns.Add("R2_Con");
            dt.Columns.Add("R2_Part");
            dt.Columns.Add("R2_Ajuste");
            dt.Columns.Add("R3_Con");
            dt.Columns.Add("R3_Part");
            dt.Columns.Add("R3_Ajuste");
            dt.Columns.Add("DeltaE_TL84");
            dt.Columns.Add("Status_TL84");

            try
            {
                using (var conn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    string sql = @"
                        SELECT c.*, d.*
                        FROM tbl_analisis_cabecera c
                        INNER JOIN tbl_analisis_detalle d ON c.Id_Lote = d.Id_Lote
                        ORDER BY c.FechaRegistro DESC, c.Id_Lote DESC";

                    using (var cmd = new System.Data.SqlClient.SqlCommand(sql, conn))
                    {
                        conn.Open();
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var row = dt.NewRow();
                                row["Id_Lote"] = reader["Id_Lote"];
                                row["ShadeName"] = reader["ShadeName"];
                                row["LotNo"] = reader["LotNo"];
                                row["FechaRegistro"] = reader["FechaRegistro"];
                                row["DyeCode"] = reader["DyeCode"];
                                row["DyeName"] = reader["DyeName"];
                                row["Concentration_Original"] = string.Format("{0:F5}%", reader["Concentration_Original"]);
                                row["Proportion_Original"] = string.Format("{0:F1}%", reader["Proportion_Original"]);
                                
                                row["R1_Con"] = string.Format("{0:F5}%", reader["R1_Con_Percentage"]);
                                row["R1_Part"] = string.Format("{0:F1}%", reader["R1_Part_Percentage"]);
                                row["R1_Ajuste"] = string.Format("{0:F1}%", reader["R1_Ajuste_Percentage"]);

                                row["R2_Con"] = string.Format("{0:F5}%", reader["R2_Con_Percentage"]);
                                row["R2_Part"] = string.Format("{0:F1}%", reader["R2_Part_Percentage"]);
                                row["R2_Ajuste"] = string.Format("{0:F1}%", reader["R2_Ajuste_Percentage"]);

                                row["R3_Con"] = string.Format("{0:F5}%", reader["R3_Con_Percentage"]);
                                row["R3_Part"] = string.Format("{0:F1}%", reader["R3_Part_Percentage"]);
                                row["R3_Ajuste"] = string.Format("{0:F1}%", reader["R3_Ajuste_Percentage"]);

                                row["DeltaE_TL84"] = reader["DeltaE_TL84"];
                                row["Status_TL84"] = reader["Status_TL84"];
                                
                                dt.Rows.Add(row);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // En un entorno real se loguearía el error
                Console.WriteLine("Error SQL: " + ex.Message);
            }

            return dt;
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
            dt.Columns.Add("Acciondl");
            dt.Columns.Add("Recdl");
            dt.Columns.Add("Impactoda");
            dt.Columns.Add("Accionda");
            dt.Columns.Add("Recda");
            dt.Columns.Add("Impactodb");
            dt.Columns.Add("Acciondb");
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
                            // Migración: Omitir DyeCode 
                            for (int j = 0; j <= 2; j++) lista.Add(celdas[j]);
                            for (int j = 4; j < celdas.Length; j++) lista.Add(celdas[j]);
                            dt.Rows.Add(lista.ToArray());
                        }
                        else if (celdas.Length == 18)
                        {
                            // [0:Shade, 1:Fecha, 2:Ilu, 3:Code, 4:Name, 5:Conc, 6:R1, 7:R2, 8:R3, ...]
                            for (int j = 0; j <= 2; j++) lista.Add(celdas[j]);
                            lista.Add(celdas[4]);
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
                string headers = "ShadeName;FechaHora;Iluminante;DyeName;ConcOriginal;Receta1;Receta2;Receta3;Impactodl;Acciondl;Recdl;Impactoda;Accionda;Recda;Impactodb;Acciondb;Recdb;FactorA;FactorB;DeltaE" + Environment.NewLine;
                File.WriteAllText(rutaArchivo, headers, Encoding.UTF8);

                var ci = CultureInfo.InvariantCulture;
                foreach (DataRow row in dt.Rows)
                {
                    string linea = string.Format(ci, "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12};{13};{14};{15};{16};{17};{18};{19}",
                        row["ShadeName"], row["FechaHora"], row["Iluminante"],
                        row["DyeName"], 
                        row["ConcOriginal"], row["Receta1"], row["Receta2"], row["Receta3"],
                        row["Impactodl"], row["Acciondl"], row["Recdl"],
                        row["Impactoda"], row["Accionda"], row["Recda"],
                        row["Impactodb"], row["Acciondb"], row["Recdb"],
                        row["FactorA"], row["FactorB"], row["DeltaE"]);

                    File.AppendAllText(rutaArchivo, linea + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { }
        }
    }
}
