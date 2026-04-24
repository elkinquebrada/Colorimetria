using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Color.Services;
using Color.Forms;

namespace Color
{
    public partial class Form1 : Form
    {
        private Label lblLeftLoaded;
        private Label lblRightLoaded;
        private Button btnCambiarLeft;
        private Button btnCambiarRight;

        // Extractor de receta e instancia de resultado
        private readonly ShadeReportExtractor _shadeExtractor = new ShadeReportExtractor(@".\tessdata");
        private ShadeExtractionResult _lastShadeResult;

        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = true;

            WireEvents();
            UpdateHints();
            LayoutBottomArea();
            PositionExitButtonAtBottom();
            MinimizarNavegador();
        }

        #region Utilidades de Ventana
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        private const int SW_SHOWMINNOACTIVE = 7;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;

        private void MinimizarNavegador()
        {
            try
            {
                this.TopMost = true;
                var timer = new System.Windows.Forms.Timer { Interval = 100 };
                int ticks = 0;
                timer.Tick += (s, e) => {
                    ticks++;
                    string[] navs = { "chrome", "msedge", "edge", "firefox", "iexplore" };
                    foreach (var proc in Process.GetProcesses())
                    {
                        if (navs.Any(n => proc.ProcessName.ToLower() == n) && proc.MainWindowHandle != IntPtr.Zero)
                        {
                            ShowWindow(proc.MainWindowHandle, SW_SHOWMINNOACTIVE);
                            SetWindowPos(proc.MainWindowHandle, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
                        }
                    }
                    if (ticks >= 30) { timer.Stop(); this.TopMost = false; }
                };
                timer.Start();
            }
            catch { }
        }
        #endregion

        private void WireEvents()
        {
            btnCargarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "Sample Comparison");
            btnCargarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Shade History Report");

            btnCambiarLeft = CrearBotonCambiar();
            btnCambiarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "Sample Comparison");
            contentBorder.Controls.Add(btnCambiarLeft);

            btnCambiarRight = CrearBotonCambiar();
            btnCambiarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Shade History Report");
            contentBorder.Controls.Add(btnCambiarRight);

            EnableDragDrop(pnlLeftFrame, picLeft, lblLeftHint, "Sample Comparison");
            EnableDragDrop(pnlRightFrame, picRight, lblRightHint, "Shade History Report");

            btnTolerancias.Click += (s, e) => { using (var f = new Color.Tolerancias.FormConfigTolerancias()) f.ShowDialog(this); };
            btnBaseDatos.Click += BtnBaseDatos_Click;
            btnSalir.Click += (s, e) => { if (MessageBox.Show("¿Salir?", "Confirme", MessageBoxButtons.YesNo) == DialogResult.Yes) Application.Exit(); };

            btnIniciar.Click += BtnIniciar_Click;
            btnCancelarAccion.Click += BtnCancelarAccion_Click;

            mainArea.Resize += (s, e) => LayoutBottomArea();
            leftNav.Resize += (s, e) => PositionExitButtonAtBottom();
        }

        private Button CrearBotonCambiar() => new Button
        {
            Text = "🔄 Cambiar imagen",
            Size = new Size(160, 32),
            BackColor = System.Drawing.Color.FromArgb(30, 90, 180),
            ForeColor = System.Drawing.Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Visible = false,
            Cursor = Cursors.Hand
        };

        private void LoadInto(PictureBox target, Label hint, string path, string etiqueta)
        {
            if (!Path.GetExtension(path).Equals(".png", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("Solo archivos PNG.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Bitmap tempBmp;
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                tempBmp = new Bitmap(Image.FromStream(fs));
            }

            // --- VALIDACIÓN ESTRICTA DE FORMATO DE SHADE HISTORY REPORT ---
            if (etiqueta == "Shade History Report")
            {
                lblStatus.Text = "Validando formato 'Shade History Report'...";
                lblStatus.ForeColor = System.Drawing.Color.Blue;
                Application.DoEvents();

                // 1. Extraer datos
                var validacion = _shadeExtractor.ExtractFromBitmap(tempBmp);
                _lastShadeResult = validacion;
            }

            // CARGA EXITOSA
            if (target.Image != null) target.Image.Dispose();
            target.Image = tempBmp;
            lblStatus.ForeColor = System.Drawing.Color.Black;
            if (hint != null) hint.Visible = false;

            string info = $"{Path.GetFileName(path)} ({target.Image.Width}x{target.Image.Height})";
            lblStatus.Text = $"{etiqueta} cargada.";

            if (target == picLeft)
            {
                btnCargarLeft.Visible = false;
                ActualizarLabelCarga(ref lblLeftLoaded, info, btnCargarLeft, btnCambiarLeft);
            }
            else
            {
                btnCargarRight.Visible = false;
                ActualizarLabelCarga(ref lblRightLoaded, info, btnCargarRight, btnCambiarRight);
            }

            CheckIfBothImagesLoaded();
        }

        private void ActualizarLabelCarga(ref Label lbl, string text, Button btnBase, Button btnCambiar)
        {
            if (lbl == null)
            {
                lbl = new Label
                {
                    AutoSize = true,
                    Font = new Font("Segoe UI", 9),
                    ForeColor = System.Drawing.Color.FromArgb(60, 64, 70)
                };
                contentBorder.Controls.Add(lbl);
            }
            lbl.Text = text;
            lbl.Location = new Point(btnBase.Left, btnBase.Top + 6);
            lbl.Visible = true;
            btnCambiar.Location = new Point(btnBase.Left, btnBase.Top + 34);
            btnCambiar.Visible = true;
            btnCambiar.BringToFront();
        }

        private void BtnIniciar_Click(object sender, EventArgs e)
        {
            if (picLeft.Image == null || picRight.Image == null || _lastShadeResult == null) return;

            lblStatus.Text = "Procesando Datos...";
            Cursor = Cursors.WaitCursor;

            try
            {
                OcrReport ocrMediciones = null;
                using (var bmpLeft = new Bitmap(picLeft.Image))
                {
                    var medExtractor = new ColorimetricDataExtractor(@".\tessdata");
                    ocrMediciones = medExtractor.ExtractReportFromBitmap(bmpLeft);
                }

                if (ocrMediciones == null || ocrMediciones.Measures.Count == 0)
                {
                    MessageBox.Show("No se detectaron datos en la imagen de Sample Comparison.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                OcrReport.SetLastReport(ocrMediciones);

                var dlgConfirm = new Colorimetria.FormConfirmacionOCR(ocrMediciones, _lastShadeResult);
                dlgConfirm.MainFormOwner = this;
                this.Hide();

                if (dlgConfirm.ShowDialog() == DialogResult.OK)
                {
                    // 1. Calcular correcciones estándar (CIE76)
                    var correcciones = ColorimetricCalculator.Calculate(dlgConfirm.RowsConfirmed);
                    
                    // 2. Calcular CMC(2:1) y sincronizar (¡NUEVO!)
                    var cmcRes = ColorimetricCalculator.CalculateCmc(correcciones, dlgConfirm.RowsConfirmed);
                    foreach (var c in correcciones)
                    {
                        var m = cmcRes.FirstOrDefault(x => string.Equals(x.Illuminant, c.Illuminant, StringComparison.OrdinalIgnoreCase));
                        if (m != null) c.CmcValue = m.CmcValue;
                    }

                    var ingredientes = RecipeCorrector.IngredientsFromShade(_lastShadeResult);
                    var deltas = RecipeCorrector.DeltasFromReport(ocrMediciones);
                    var corrReceta = RecipeCorrector.Calculate(ingredientes, deltas);

                    // Ciclo: el usuario puede regresar al OCR y volver a Resultados
                    bool seguir = true;
                    while (seguir)
                    {
                        var frmRes = new FormResultados(BuildResumenReceta(_lastShadeResult), correcciones, corrReceta as List<IlluminantCorrectionResult>);

                        // Crear nuevo OCR fresco para el botón Regresar
                        var dlgOcr = new Colorimetria.FormConfirmacionOCR(ocrMediciones, _lastShadeResult);
                        dlgOcr.MainFormOwner = this;
                        frmRes.FormOcrOrigen = dlgOcr;

                        frmRes.ShowDialog();

                        // Si el usuario regresó al OCR y volvió a confirmar
                        if (!dlgOcr.IsDisposed && dlgOcr.RowsConfirmed != null && dlgOcr.RowsConfirmed.Count > 0)
                        {
                            // Recalcular con los nuevos datos (Incluyendo CMC)
                            correcciones = ColorimetricCalculator.Calculate(dlgOcr.RowsConfirmed);
                            
                            var cmcResLoop = ColorimetricCalculator.CalculateCmc(correcciones, dlgOcr.RowsConfirmed);
                            foreach (var c in correcciones)
                            {
                                var m = cmcResLoop.FirstOrDefault(x => string.Equals(x.Illuminant, c.Illuminant, StringComparison.OrdinalIgnoreCase));
                                if (m != null) c.CmcValue = m.CmcValue;
                            }

                            ingredientes = RecipeCorrector.IngredientsFromShade(_lastShadeResult);
                            deltas = RecipeCorrector.DeltasFromReport(ocrMediciones);
                            corrReceta = RecipeCorrector.Calculate(ingredientes, deltas);
                        }
                        else
                        {
                            seguir = false;
                        }

                        if (!dlgOcr.IsDisposed) dlgOcr.Dispose();
                        frmRes.Dispose();
                    }
                }
                this.Show();
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            finally { Cursor = Cursors.Default; }
        }

        private void BtnCancelarAccion_Click(object sender, EventArgs e)
        {
            ClearPicture(picLeft); ClearPicture(picRight);
            _lastShadeResult = null;
            btnCargarLeft.Visible = btnCargarRight.Visible = true;
            if (lblLeftLoaded != null) lblLeftLoaded.Visible = false;
            if (lblRightLoaded != null) lblRightLoaded.Visible = false;
            btnCambiarLeft.Visible = btnCambiarRight.Visible = false;
            lblStatus.Text = "Cargue imágenes";
            lblStatus.ForeColor = System.Drawing.Color.Black;
            UpdateHints();
            ShowActionButtons(false);
        }

        #region Helpers UI
        private void CheckIfBothImagesLoaded()
        {
            // Solo permite el botón Iniciar si picRight tiene una receta validada en _lastShadeResult
            bool listo = picLeft.Image != null && picRight.Image != null && _lastShadeResult != null;
            ShowActionButtons(listo);
        }
        private void ShowActionButtons(bool visible) { btnIniciar.Visible = visible; btnCancelarAccion.Visible = visible; }
        private void UpdateHints() { if (lblLeftHint != null) lblLeftHint.Visible = picLeft.Image == null; if (lblRightHint != null) lblRightHint.Visible = picRight.Image == null; }
        private void ClearPicture(PictureBox pb) { if (pb?.Image != null) { pb.Image.Dispose(); pb.Image = null; } }

        private void SelectAndLoadPng(PictureBox target, Label hint, string etiqueta)
        {
            using (var ofd = new OpenFileDialog { Filter = "PNG|*.png" })
                if (ofd.ShowDialog() == DialogResult.OK) LoadInto(target, hint, ofd.FileName, etiqueta);
        }

        private void EnableDragDrop(Control surf, PictureBox target, Label hint, string etiqueta)
        {
            surf.AllowDrop = target.AllowDrop = true;
            surf.DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            surf.DragDrop += (s, e) => {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0) LoadInto(target, hint, files[0], etiqueta);
            };
        }

        private void LayoutBottomArea()
        {
            int centerX = mainArea.ClientSize.Width / 2;
            lblStatus.Location = new Point(centerX - lblStatus.Width / 2, mainArea.ClientSize.Height - 110);
            btnIniciar.Location = new Point(centerX - 168, mainArea.ClientSize.Height - 70);
            btnCancelarAccion.Location = new Point(centerX + 8, mainArea.ClientSize.Height - 70);
        }

        private void PositionExitButtonAtBottom() => btnSalir.Location = new Point(20, Math.Max(210, leftNav.Height - btnSalir.Height - 20));

        private void BtnBaseDatos_Click(object sender, EventArgs e)
        {
            try
            {
                var tabla = HistorialService.ObtenerHistorial();

                if (tabla == null || tabla.Rows.Count == 0)
                {
                    MessageBox.Show(
                        "La base de datos no contiene registros todavía.\n\n" +
                        "Los registros se generan al guardar los resultados desde el botón 'Historial' dentro de un análisis.",
                        "Base de datos vacía",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var frm = new FormHistorial();
                frm.CargarHistorial(tabla);
                frm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error al abrir la base de datos: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private string BuildResumenReceta(ShadeExtractionResult result)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Shade History Report EXTRAÍDA");
            if (result?.Recipe != null)
                foreach (var i in result.Recipe) sb.AppendLine($"{i.Code} - {i.Name}: {i.Percentage}%");
            return sb.ToString();
        }
        #endregion
    }
}