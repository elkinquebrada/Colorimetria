using System;
using System.Collections.Generic;
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
                // Se procesa únicamente la primera hoja del libro
                var ws = wb.Worksheet(1);

                foreach (var row in ws.RowsUsed())
                {
                    var list = new List<string>();
                    foreach (var cell in row.CellsUsed())
                        // Se obtiene el valor de la celda convertido a string
                        list.Add(cell.GetValue<string>());

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
                // Validación: Se requieren al menos 3 columnas (Código, Nombre, Porcentaje)
                if (row.Count < 3)
                    continue;

                list.Add(new RecipeItem
                {
                    Code = row[0],
                    Name = row[1],
                    Percentage = row[2]
                });
            }

            return list;
        }

        /// Carga y valida datos colorimétricos (L, A, B, Chroma, Hue) 
        public static List<ColorimetricRow> LoadMeasurements(string excelPath)
        {
            var table = LoadTable(excelPath);

            var list = new List<ColorimetricRow>();

            foreach (var row in table)
            { 
                // Validación: Se requieren al menos 7 columnas para una medición completa
                if (row.Count < 7) continue;

                double L, A, B, Chroma, Hue;

                if (double.TryParse(row[2], out L) &&
                    double.TryParse(row[3], out A) &&
                    double.TryParse(row[4], out B) &&
                    double.TryParse(row[5], out Chroma) &&
                    double.TryParse(row[6], out Hue))
                {
                    list.Add(new ColorimetricRow
                    {
                        Illuminant = row[0],
                        Type = row[1],
                        L = L,
                        A = A,
                        B = B,
                        Chroma = Chroma,
                        Hue = Hue
                    });
                }
            }

            return list;
        }

        /// Lee los perfiles estáticos de tolerancia de la hoja de cálculo.
        public static List<ToleranceResult> LoadTolerances(string excelPath)
        {
            var list = new List<ToleranceResult>();

            try
            {
                using (var wb = new XLWorkbook(excelPath))
                {
                    var ws = wb.Worksheet(3);

                    // De acuerdo al esquema detectado, las tolerancias (DE, DL, DC, DH) se encuentran en:
                    int[] cols = { 3, 6, 9 };

                    foreach (var col in cols)
                    {
                        var cellDE = ws.Cell(2, col);
                        if (cellDE.IsEmpty()) continue;

                        if (double.TryParse(cellDE.GetValue<string>().Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double de))
                        {
                            // Leer los componentes
                            double dl = ws.Cell(3, col).GetValue<double>();
                            double dc = ws.Cell(4, col).GetValue<double>();
                            double dh = ws.Cell(5, col).GetValue<double>();

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
            }
            catch
            {
                // Devolvemos vacío si falla la lectura
            }

            return list;
        }
    }
}
