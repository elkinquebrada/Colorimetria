using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Color;

namespace OCR
{
    
    /// REGLA DE ORO: El Excel recibe EXCLUSIVAMENTE los valores numéricos del OCR.
    ///               Las fórmulas derivadas (deltas, %, diagnóstico) las calcula Excel solo.
    public static class ExcelWriter
    {
        // ──────────────────────────────────────────────────────────────────────
        // PUNTO DE ENTRADA PRINCIPAL — llamado automáticamente desde FormResultados
        // ──────────────────────────────────────────────────────────────────────

        /// Escribe todos los datos OCR en el Excel de referencia de forma silenciosa.
        public static bool WriteAll(
            string excelPath,
            List<ColorimetricRow> ocrRows,
            List<CmcResult> cmcResults,
            List<IlluminantCorrectionResult> recipeResults)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
                return false;

            try
            {
                using (var wb = new XLWorkbook(excelPath))
                {
                    bool ok1 = WriteHoja1_CalculoMedicion(wb, ocrRows);
                    bool ok2 = WriteHoja2_CalculoReceta(wb, cmcResults, recipeResults, ocrRows);

                    if (ok1 || ok2)
                    {
                        wb.Save();
                        return true;
                    }
                }
            }
            catch
            {
                // Silencioso — nunca interrumpir el flujo del usuario
            }

            return false;
        }

        // ──────────────────────────────────────────────────────────────────────
        // HOJA 1: CALCULO SAMPLE COMPARISON
        // Celdas azules: L/A/B del Estándar y Lote por iluminante
        // ──────────────────────────────────────────────────────────────────────

        private static bool WriteHoja1_CalculoMedicion(IXLWorkbook wb, List<ColorimetricRow> rows)
        {
            if (rows == null || rows.Count == 0) return false;

            IXLWorksheet ws;
            try { ws = wb.Worksheet(1); }
            catch { return false; }

            // --- Limpieza previa de celdas azules (Regla de oro) ---
            ClearCell(ws, 3, 2); ClearCell(ws, 3, 3); 
            ClearCell(ws, 4, 2); ClearCell(ws, 4, 3); 
            ClearCell(ws, 5, 2); ClearCell(ws, 5, 3); 

            // Limpieza tabla expandida (filas 15-20, cols D-F)
            for (int r = 15; r <= 20; r++)
            {
                for (int c = 4; c <= 6; c++) ClearCell(ws, r, c);
            }

            // --- Buscar D65 (fila resumen R03–R05) ---
            var stdD65 = rows.FirstOrDefault(r =>
                string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.Type, "Std", StringComparison.OrdinalIgnoreCase));

            var lotD65 = rows.FirstOrDefault(r =>
                string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.Type, "Lot", StringComparison.OrdinalIgnoreCase));

            if (stdD65 != null && lotD65 != null)
            {
                // R03: L* Std (B3) y L* Lot (C3)
                SetNumericCell(ws, 3, 2, stdD65.L);
                SetNumericCell(ws, 3, 3, lotD65.L);

                // R04: a* Std (B4) y a* Lot (C4)
                SetNumericCell(ws, 4, 2, stdD65.A);
                SetNumericCell(ws, 4, 3, lotD65.A);

                // R05: b* Std (B5) y b* Lot (C5)
                SetNumericCell(ws, 5, 2, stdD65.B);
                SetNumericCell(ws, 5, 3, lotD65.B);
            }

            // --- Tabla expandida por iluminante (R15–R20, columnas D=4, E=5, F=6) ---
            var iluminantOrder = new[] { "D65", "TL84", "A" };
            int baseRow = 15;

            foreach (var illuminant in iluminantOrder)
            {
                var std = rows.FirstOrDefault(r =>
                    string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(r.Type, "Std", StringComparison.OrdinalIgnoreCase));

                var lot = rows.FirstOrDefault(r =>
                    string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(r.Type, "Lot", StringComparison.OrdinalIgnoreCase));

                if (std != null)
                {
                    SetNumericCell(ws, baseRow, 4, std.L);     
                    SetNumericCell(ws, baseRow, 5, std.A);     
                    SetNumericCell(ws, baseRow, 6, std.B);     
                }

                if (lot != null)
                {
                    SetNumericCell(ws, baseRow + 1, 4, lot.L); 
                    SetNumericCell(ws, baseRow + 1, 5, lot.A); 
                    SetNumericCell(ws, baseRow + 1, 6, lot.B); 
                }

                baseRow += 2; // Siguiente par Std/Lot
            }

            return true;
        }

        // ──────────────────────────────────────────────────────────────────────
        // HOJA 2: CALCULO SHADE HISTORY REPORT
        // Celdas azules: CMC por iluminante + tintes y concentraciones
        // ──────────────────────────────────────────────────────────────────────

        private static bool WriteHoja2_CalculoReceta(
            IXLWorkbook wb,
            List<CmcResult> cmcResults,
            List<IlluminantCorrectionResult> recipeResults,
            List<ColorimetricRow> ocrRows)
        {
            IXLWorksheet ws;
            try { ws = wb.Worksheet(2); }
            catch { return false; }

            bool anyWrite = false;

            // ─── Sección CMC(2:1) por iluminante ─────────────────────────────
            // Estructura del Excel:

            if (cmcResults != null && cmcResults.Count > 0)
            {
                // Limpieza celdas CMC
                foreach (int r in new[] { 5, 7, 9 })
                {
                    for (int c = 3; c <= 6; c++) ClearCell(ws, r, c);
                }

                var cmcRows = new[] {
                    ("D65",  5),
                    ("TL84", 7),
                    ("A",    9)
                };

                foreach (var (illuminant, row) in cmcRows)
                {
                    var cmc = cmcResults.FirstOrDefault(x =>
                        string.Equals(x.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase));

                    if (cmc == null) continue;

                    SetNumericCell(ws, row, 3, cmc.Lightness);  
                    SetNumericCell(ws, row, 4, cmc.Chroma);     
                    SetNumericCell(ws, row, 5, cmc.Hue);        
                    SetNumericCell(ws, row, 6, cmc.CmcValue);   

                    // HUE Std / Lot (cols I y J en filas 5, 7, 9)
                    var stdRow = ocrRows?.FirstOrDefault(r =>
                        string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(r.Type, "Std", StringComparison.OrdinalIgnoreCase));
                    var lotRow = ocrRows?.FirstOrDefault(r =>
                        string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(r.Type, "Lot", StringComparison.OrdinalIgnoreCase));

                    if (stdRow != null) SetNumericCell(ws, row, 9, stdRow.Hue);   
                    if (lotRow != null) SetNumericCell(ws, row, 10, lotRow.Hue);  

                    anyWrite = true;
                }
            }

            // ─── Sección Receta/Tintes por iluminante ────────────────────────
            // Estructura del Excel:

            if (recipeResults != null && recipeResults.Count > 0)
            {
                var recipeRows = new[] {
                    ("D65",  15),
                    ("TL84", 23),
                    ("A",    31)
                };

                foreach (var (illuminant, startRow) in recipeRows)
                {
                    var result = recipeResults.FirstOrDefault(r =>
                        string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase));

                    if (result == null || result.Ingredients == null || result.Ingredients.Count == 0)
                        continue;

                    // Limpiar las filas de tintes antes de escribir
                    for (int r = startRow; r < startRow + result.Ingredients.Count; r++)
                    {
                        ClearCell(ws, r, 2); 
                        ClearCell(ws, r, 3); 
                    }

                    int rowIdx = startRow;
                    foreach (var ing in result.Ingredients)
                    {
                        SetTextCell(ws, rowIdx, 2, ing.Name);            
                        SetNumericCell(ws, rowIdx, 3, ing.Original);     

                        rowIdx++;
                        if (rowIdx > startRow + 10) break; 
                    }

                    anyWrite = true;
                }
            }

            return anyWrite;
        }

        // ──────────────────────────────────────────────────────────────────────
        // HELPERS
        // ──────────────────────────────────────────────────────────────────────

        private static void SetNumericCell(IXLWorksheet ws, int row, int col, double value)
        {
            try
            {
                var cell = ws.Cell(row, col);
                cell.SetValue(Math.Round(value, 4));
            }
            catch { }
        }

        private static void SetTextCell(IXLWorksheet ws, int row, int col, string value)
        {
            try
            {
                ws.Cell(row, col).SetValue(value ?? string.Empty);
            }
            catch { }
        }

        private static void ClearCell(IXLWorksheet ws, int row, int col)
        {
            try
            {
                ws.Cell(row, col).SetValue(string.Empty);
            }
            catch { }
        }

        // ──────────────────────────────────────────────────────────────────────
        // RUTA DEL EXCEL DE REFERENCIA (Busca automáticamente)
        // ──────────────────────────────────────────────────────────────────────

        public static string FindReferenceExcelPath()
        {
            // 1. Ruta estándar del ejecutable
            string path1 = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"LogicDocs\calculos a realizar con el programa.xlsx");

            if (File.Exists(path1)) return path1;

            // 2. Fallback para depuración en Visual Studio
            string path2 = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                @"..\..\LogicDocs\calculos a realizar con el programa.xlsx");

            if (File.Exists(path2)) return path2;

            // 3. Directorio de trabajo
            string path3 = Path.Combine(
                Directory.GetCurrentDirectory(),
                @"LogicDocs\calculos a realizar con el programa.xlsx");

            if (File.Exists(path3)) return path3;

            return null;
        }
    }
}
