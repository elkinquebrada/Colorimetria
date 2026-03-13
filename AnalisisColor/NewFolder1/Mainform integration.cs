// ════════════════════════════════════════════════════════════════════════════
//  INTEGRACIÓN EN MainForm.cs
//  Agrega este código a tu MainForm existente.
//  Requiere: Tesseract NuGet package   →  Install-Package Tesseract
// ════════════════════════════════════════════════════════════════════════════

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Colorimetria
{
    public partial class MainForm : Form
    {
        // ── Instancia del extractor (reutilizable) ───────────────────────────
        private readonly ColorimetricDataExtractor _extractor = new ColorimetricDataExtractor(@".\tessdata");

        // ── Guarda los últimos resultados para uso posterior ─────────────────
        private List<ColorimetricRow> _lastConfirmedRows;
        private List<CorrectionResult> _lastCorrectionResults;

        // ════════════════════════════════════════════════════════════════════
        //  PUNTO DE ENTRADA: llamar desde el botón "Iniciar escaneo"
        // ════════════════════════════════════════════════════════════════════
        private void BtnIniciarEscaneo_Click(object sender, EventArgs e)
        {
            // 1) Obtener la imagen cargada (ajusta según tu lógica actual)
            Bitmap imagenCargada = ObtenerImagenCargada();
            if (imagenCargada == null)
            {
                MessageBox.Show("No hay imagen cargada.", "Aviso",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2) Ejecutar OCR y extraer datos
            List<ColorimetricRow> rows;
            try
            {
                rows = _extractor.ExtractFromBitmap(imagenCargada);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en OCR:\n{ex.Message}", "Error OCR",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (rows.Count == 0)
            {
                MessageBox.Show(
                    "No se encontraron datos colorimétricos en la imagen.\n" +
                    "Verifica que la imagen contenga una tabla con D65, TL84 o CWF.",
                    "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 3) Mostrar diálogo de confirmación
            using (var dlg = new FormConfirmacionOCR(rows))
            {
                var result = dlg.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    // 4) Usuario confirmó → guardar filas y calcular correcciones
                    _lastConfirmedRows = dlg.RowsConfirmed;
                    EjecutarCalculosCorreccion(_lastConfirmedRows);
                }
                else
                {
                    // Usuario canceló → no hacer nada
                    MessageBox.Show("Operación cancelada.", "Cancelado",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        // ════════════════════════════════════════════════════════════════════
        //  CÁLCULOS DE CORRECCIÓN COLORIMÉTRICA
        // ════════════════════════════════════════════════════════════════════
        private void EjecutarCalculosCorreccion(List<ColorimetricRow> rows)
        {
            try
            {
                // Calcular
                _lastCorrectionResults = ColorimetricCalculator.Calculate(rows);

                // Obtener tolerancias (puedes leerlas de tu BD o config)
                double tolDL = 1.60, tolDC = 1.60, tolDH = 1.20, tolDE = 1.60;

                // Generar reporte en texto
                string resumen = ColorimetricCalculator.BuildSummary(
                    rows, _lastCorrectionResults, tolDL, tolDC, tolDH, tolDE);

                // Mostrar resultado
                MostrarResultados(resumen, _lastCorrectionResults);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en cálculos:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════════
        //  MOSTRAR RESULTADOS (adaptar a tu UI)
        // ════════════════════════════════════════════════════════════════════
        private void MostrarResultados(string resumen, List<CorrectionResult> results)
        {
            // Opción A: Abrir formulario de resultados
            using (var frmResultados = new FormResultados(resumen, results))
            {
                frmResultados.ShowDialog(this);
            }

            // Opción B (alternativa simple): guardar en archivo
            // string path = System.IO.Path.Combine(Application.StartupPath, "resultado.txt");
            // System.IO.File.WriteAllText(path, resumen);
            // System.Diagnostics.Process.Start("notepad.exe", path);
        }

        // ════════════════════════════════════════════════════════════════════
        //  HELPER: obtener Bitmap de la imagen cargada actualmente
        //  → Adapta este método a cómo guardas la imagen en tu form
        // ════════════════════════════════════════════════════════════════════
        private Bitmap ObtenerImagenCargada()
        {
            // OPCIÓN 1: si tienes un PictureBox
            // return pictureBoxPreview.Image as Bitmap;

            // OPCIÓN 2: si guardas la ruta
            // return string.IsNullOrEmpty(_imagenPath) ? null : new Bitmap(_imagenPath);

            // OPCIÓN 3: si tienes el bitmap en una variable
            // return _currentBitmap;

            // ── Placeholder para compilar ─────────────────────────────────
            // Reemplaza con la lógica real de tu aplicación:
            throw new NotImplementedException(
                "Implementa ObtenerImagenCargada() según tu lógica actual de carga de imagen.");
        }
    }
}