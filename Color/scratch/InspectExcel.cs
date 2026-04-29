using System;
using System.Linq;
using ClosedXML.Excel;

namespace Inspector
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"c:\Users\COPEEQGuapacha\OneDrive - Coats\Escritorio\Colorimetria\Color\LogicDocs\calculos a realizar con el programa.xlsx";
            try
            {
                using (var wb = new XLWorkbook(path))
                {
                    var ws = wb.Worksheet(1);
                    var row = ws.Row(1);
                    foreach (var cell in row.CellsUsed())
                    {
                        Console.WriteLine(cell.Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
