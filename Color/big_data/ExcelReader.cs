using System;
using System.Collections.Generic;
using System.Globalization; // Necesario para manejar puntos y comas
using ClosedXML.Excel;
using Color;

namespace OCR
{
    /// Proporciona métodos estáticos para la lectura y mapeo de datos del archivos Excel (.xlsx)
    public static class ExcelReader
    {
        /// Lee la primera hoja del archivo Excel y la convierte en una estructura de listas.
        public static List<List<string>> LoadTable(string excelPath)
        {
            var table = new List<List<string>>();

            using (var wb = new XLWorkbook(excelPath))
            {
                var ws = wb.Worksheet(1);

                foreach (var row in ws.RowsUsed())
                {
                    var list = new List<string>();
                    foreach (var cell in row.CellsUsed())
                        list.Add(GetCellValueSafe(cell));

                    table.Add(list);
                }
            }

            return table;
        }

        /// Carga y mapea los datos de una receta desde archivo Excel.
        public static List<RecipeItem> LoadRecipe(string excelPath)
        {
            var table = LoadTable(excelPath);
            var list = new List<RecipeItem>();

            foreach (var row in table)
            {
                if (row.Count < 3) continue;

                list.Add(new RecipeItem
                {
                    Code = row[0],
                    Name = row[1],
                    Percentage = row[2]
                });
            }

            return list;
        }

        /// Carga y valida datos colorimétricos con manejo seguro de decimales.
        public static List<ColorimetricRow> LoadMeasurements(string excelPath)
        {
            var table = LoadTable(excelPath);
            var list = new List<ColorimetricRow>();

            foreach (var row in table)
            {
                if (row.Count < 7) continue;

                list.Add(new ColorimetricRow
                {
                    Illuminant = row[0],
                    Type = row[1],
                    L = ParseDoubleSafe(row[2]),
                    A = ParseDoubleSafe(row[3]),
                    B = ParseDoubleSafe(row[4]),
                    Chroma = ParseDoubleSafe(row[5]),
                    Hue = ParseDoubleSafe(row[6])
                });
            }

            return list;
        }

        /// Lee los perfiles de tolerancia corrigiendo el error de conversión Double.
        public static List<ToleranceResult> LoadTolerances(string excelPath)
        {
            var list = new List<ToleranceResult>();

            try
            {
                using (var wb = new XLWorkbook(excelPath))
                {
                    // Buscamos la hoja "TOLERANCIA" de forma robusta (ignorando mayúsculas/minúsculas y espacios)
                    IXLWorksheet ws = null;
                    foreach (var sheet in wb.Worksheets)
                    {
                        if (sheet.Name.Trim().Equals("TOLERANCIA", StringComparison.OrdinalIgnoreCase))
                        {
                            ws = sheet;
                            break;
                        }
                    }

                    // Si no se encuentra por nombre, usamos el índice 3 como respaldo
                    if (ws == null) ws = wb.Worksheet(3);

                    // Columnas de interés según el esquema (C, F, I, L)
                    int[] cols = { 3, 6, 9, 12 };

                    foreach (var col in cols)
                    {
                        var cellDE = ws.Cell(2, col);
                        if (cellDE.IsEmpty()) continue;

                        // CORRECCIÓN: Usamos GetCellValueSafe para manejar fórmulas con referencias externas (evita error "References from other files")
                        double de = GetCellValueAsDouble(cellDE);
                        double dl = GetCellValueAsDouble(ws.Cell(3, col));
                        double dc = GetCellValueAsDouble(ws.Cell(4, col));
                        double dh = GetCellValueAsDouble(ws.Cell(5, col));

                        list.Add(new ToleranceResult
                        {
                            DE = Math.Round(de, 2),
                            DL = Math.Round(dl, 3),
                            DC = Math.Round(dc, 3),
                            DH = Math.Round(dh, 3)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error crítico al procesar tolerancias:\n" + ex.Message,
                    "Error de Lectura",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }

            return list;
        }

        /// Obtiene el valor de una celda como double de forma robusta.
        private static double GetCellValueAsDouble(IXLCell cell)
        {
            try
            {
                // Si ClosedXML ya identifica la celda como número, lo tomamos directamente
                if (cell.Value.IsNumber) return cell.Value.GetNumber();
                
                // Si es texto o fórmula, usamos el parseo seguro
                return ParseDoubleSafe(GetCellValueSafe(cell));
            }
            catch
            {
                // Fallback en caso de cualquier error
                return ParseDoubleSafe(GetCellValueSafe(cell));
            }
        }

        /// Obtiene el valor de una celda de forma segura, manejando fórmulas con referencias externas.
        private static string GetCellValueSafe(IXLCell cell)
        {
            try
            {
                // Intentamos obtener el valor (esto disparará la evaluación de fórmulas internas)
                return cell.Value.ToString();
            }
            catch (Exception ex) when (ex.Message.Contains("References from other files"))
            {
                // Si hay referencias externas que ClosedXML no puede resolver, usamos el último valor guardado (caché)
                return cell.CachedValue.ToString();
            }
            catch
            {
                // Fallback general para cualquier otro error de evaluación
                try { return cell.HasFormula ? cell.CachedValue.ToString() : cell.Value.ToString(); }
                catch { return ""; }
            }
        }

        /// MÉTODO AUXILIAR CLAVE: Convierte texto a número sin importar el separador decimal (punto o coma).
        private static double ParseDoubleSafe(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return 0;

            // Limpiamos el texto: convertimos comas en puntos para estandarizar
            string cleanValue = value.Replace(',', '.').Trim();

            // Intentamos convertir usando la cultura invariante (estándar internacional)
            if (double.TryParse(cleanValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }

            return 0;
        }
    }
}