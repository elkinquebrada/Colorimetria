using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Tesseract;
using ClosedXML.Excel;

namespace OCR
{
    /// <summary>
    /// Extrae texto desde imágenes usando OCR, reconstruye una tabla
    /// agrupando palabras por coordenadas y exporta un archivo Excel
    /// temporal para ser leído por los extractores internos.
    /// </summary>
    public class ReportExtractor
    {
        private readonly string _tessdata;

        public ReportExtractor(string tessdataPath)
        {
            _tessdata = tessdataPath;
        }

        /// <summary>
        /// Convierte una imagen -> OCR -> tabla -> Excel temporal.
        /// </summary>
        public string ProcessImageToExcel(string imagePath)
        {
            var words = ExtractWordsWithBoxes(imagePath);
            var rows = BuildTable(words);
            return SaveToTempExcel(rows);
        }

        // ======================================================
        // =============== 1. OCR + Bounding Boxes ==============
        // ======================================================
        private List<OcrWord> ExtractWordsWithBoxes(string imagePath)
        {
            var result = new List<OcrWord>();

            using (var engine = new TesseractEngine(_tessdata, "eng", EngineMode.Default))
            using (var img = Pix.LoadFromFile(imagePath))
            using (var page = engine.Process(img))
            using (var it = page.GetIterator())
            {
                it.Begin();

                do
                {
                    if (it.TryGetBoundingBox(PageIteratorLevel.Word, out var rect))
                    {
                        string text = it.GetText(PageIteratorLevel.Word);

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            result.Add(new OcrWord
                            {
                                Text = text.Trim(),
                                X = rect.X1,
                                Y = rect.Y1,
                                Width = rect.Width,
                                Height = rect.Height
                            });
                        }
                    }
                }
                while (it.Next(PageIteratorLevel.Word));
            }

            return result;
        }

        // ======================================================
        // =============== 2. RECONSTRUCCIÓN TABLA ==============
        // ======================================================

        private List<List<string>> BuildTable(List<OcrWord> words)
        {
            var table = new List<List<string>>();

            if (words == null || words.Count == 0)
                return table;

            // Ordenar por Y ascendente (de arriba hacia abajo)
            words.Sort((a, b) => a.Y.CompareTo(b.Y));

            int tolerance = 10;  // margen vertical para determinar filas
            int currentRowY = -9999;
            List<OcrWord> currentRow = null;

            foreach (var w in words)
            {
                // ¿Nueva fila?
                if (Math.Abs(w.Y - currentRowY) > tolerance)
                {
                    if (currentRow != null)
                        table.Add(ConvertRow(currentRow));

                    currentRow = new List<OcrWord>();
                    currentRowY = w.Y;
                }

                currentRow.Add(w);
            }

            // Agregar última fila
            if (currentRow != null)
                table.Add(ConvertRow(currentRow));

            return table;
        }

        private List<string> ConvertRow(List<OcrWord> row)
        {
            // Ordenar de izquierda a derecha
            row.Sort((a, b) => a.X.CompareTo(b.X));

            var list = new List<string>();
            foreach (var w in row)
                list.Add(w.Text);

            return list;
        }

        // ======================================================
        // ================== 3. EXCEL TEMPORAL =================
        // ======================================================

        private string SaveToTempExcel(List<List<string>> table)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "OCR_" + Guid.NewGuid().ToString("N") + ".xlsx"
            );

            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Extract");

                for (int r = 0; r < table.Count; r++)
                {
                    for (int c = 0; c < table[r].Count; c++)
                    {
                        ws.Cell(r + 1, c + 1).Value = table[r][c];
                    }
                }

                wb.SaveAs(path);
            }

            return path;
        }
    }

    /// <summary>
    /// Contenedor de palabra OCR + posición en la imagen.
    /// </summary>
    public class OcrWord
    {
        public string Text { get; set; }

        public int X { get; set; }
        public int Y { get; set; }

        public int Width { get; set; }
        public int Height { get; set; }
    }
}