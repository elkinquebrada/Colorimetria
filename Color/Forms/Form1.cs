using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
            btnCargarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "Mediciones");
            btnCargarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Receta");

            btnCambiarLeft = CrearBotonCambiar();
            btnCambiarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "Mediciones");
            contentBorder.Controls.Add(btnCambiarLeft);

            btnCambiarRight = CrearBotonCambiar();
            btnCambiarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Receta");
            contentBorder.Controls.Add(btnCambiarRight);

            EnableDragDrop(pnlLeftFrame, picLeft, lblLeftHint, "Mediciones");
            EnableDragDrop(pnlRightFrame, picRight, lblRightHint, "Receta");

            btnTolerancias.Click += (s, e) => { using (var f = new Color.Tolerancias.FormConfigTolerancias()) f.ShowDialog(this); };
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

            // --- VALIDACIÓN ESTRICTA DE FORMATO DE RECETA ---
            if (etiqueta == "Receta")
            {
                lblStatus.Text = "Validando formato 'Shade History Report'...";
                lblStatus.ForeColor = System.Drawing.Color.Blue;
                Application.DoEvents();

                // 1. Extraer datos
                var validacion = _shadeExtractor.ExtractFromBitmap(tempBmp);

                // 2. VERIFICACIÓN DE CABECERA (Requisito indispensable según imagen cargada)
                bool esFormatoValido = false;
                if (validacion != null)
                {
                    // El extractor debe detectar "Shade History Report" o al menos tener ingredientes y LAB coherentes
                    if (validacion.Recipe.Count > 0 || validacion.Lab != null)
                    {
                        esFormatoValido = true;
                    }
                }

                if (!esFormatoValido)
                {
                    tempBmp.Dispose();
                    target.Image = null;
                    lblStatus.Text = "ERROR: Formato de receta no reconocido.";
                    lblStatus.ForeColor = System.Drawing.Color.Red;
                    MessageBox.Show("La imagen no corresponde al formato 'Shade History Report'.\n\nSolo se permite el formato oficial de Coats Cadena.",
                                    "Error de Formato", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    CheckIfBothImagesLoaded();
                    return;
                }
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

            lblStatus.Text = "Procesando mediciones...";
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
                    MessageBox.Show("No se detectaron datos en la imagen de mediciones.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var dlgConfirm = new Colorimetria.FormConfirmacionOCR(ocrMediciones, _lastShadeResult);
                dlgConfirm.MainFormOwner = this;
                this.Hide();

                if (dlgConfirm.ShowDialog() == DialogResult.OK)
                {
                    var correcciones = ColorimetricCalculator.Calculate(dlgConfirm.RowsConfirmed);
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
                            // Recalcular con los nuevos datos
                            correcciones = ColorimetricCalculator.Calculate(dlgOcr.RowsConfirmed);
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

        private string BuildResumenReceta(ShadeExtractionResult result)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("RECETA EXTRAÍDA");
            if (result?.Recipe != null)
                foreach (var i in result.Recipe) sb.AppendLine($"{i.Code} - {i.Name}: {i.Percentage}%");
            return sb.ToString();
        }
        #endregion
    }
}