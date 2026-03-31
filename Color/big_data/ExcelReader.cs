using System.Collections.Generic;
using ClosedXML.Excel;
using Color;

namespace OCR
{
    public static class ExcelReader
    {
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
                        list.Add(cell.GetValue<string>());

                    table.Add(list);
                }
            }

            return table;
        }

        public static List<RecipeItem> LoadRecipe(string excelPath)
        {
            var table = LoadTable(excelPath);
            var list = new List<RecipeItem>();

            foreach (var row in table)
            {
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

        public static List<ColorimetricRow> LoadMeasurements(string excelPath)
        {
            var table = LoadTable(excelPath);

            var list = new List<ColorimetricRow>();

            foreach (var row in table)
            {
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
    }
}
