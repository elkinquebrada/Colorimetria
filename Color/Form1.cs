using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

// ► AGREGAR: referencia al extractor de recetas


namespace Color
{
    public partial class Form1 : Form
    {
        // Labels que muestran el nombre de archivo debajo de cada cuadro
        private Label lblLeftLoaded;
        private Label lblRightLoaded;

        // ► AGREGAR: extractor de receta e instancia de resultado
        private readonly ShadeReportExtractor _shadeExtractor =
            new ShadeReportExtractor(@".\tessdata");
        private ShadeExtractionResult _lastShadeResult;


        public Form1()
        {
            InitializeComponent();

            WireEvents();
            UpdateHints();
            LayoutBottomArea();
            PositionExitButtonAtBottom();
        }

        // -------------------- Eventos y cableado --------------------
        private void WireEvents()
        {
            btnCargarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "Mediciones");
            btnCargarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Receta");

            EnableDragDrop(pnlLeftFrame, picLeft, lblLeftHint, "Mediciones");
            EnableDragDrop(pnlRightFrame, picRight, lblRightHint, "Receta");

            btnTolerancias.Click += (s, e) => MessageBox.Show("Abrir Config. de tolerancias", "Info");
            btnBaseDatos.Click += (s, e) => MessageBox.Show("Abrir Base de datos", "Info");

            btnSalir.Click += (s, e) =>
            {
                var r = MessageBox.Show("¿Seguro que deseas salir?",
                                        "Confirmar salida",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question);
                if (r == DialogResult.Yes) Application.Exit();
            };

            // ► MODIFICAR: btnIniciar ahora llama al flujo real de OCR
            btnIniciar.Click += BtnIniciar_Click;

            btnCancelarAccion.Click += (s, e) =>
            {
                ClearPicture(picLeft);
                ClearPicture(picRight);

                btnCargarLeft.Visible = true;
                btnCargarRight.Visible = true;

                if (lblLeftLoaded != null) lblLeftLoaded.Visible = false;
                if (lblRightLoaded != null) lblRightLoaded.Visible = false;

                UpdateHints();
                lblStatus.Text = "Cargue dos imágenes para comparación";
                ShowActionButtons(false);
            };

            mainArea.Resize += (s, e) => LayoutBottomArea();
            leftNav.Resize += (s, e) => PositionExitButtonAtBottom();
        }

        // -------------------- ► NUEVO: flujo de escaneo --------------------
        private void BtnIniciar_Click(object sender, EventArgs e)
        {
            if (picLeft.Image == null || picRight.Image == null)
            {
                MessageBox.Show(
                    "Debes cargar las dos imágenes antes de iniciar el escaneo.\n\n" +
                    "• Izquierda: Mediciones\n• Derecha: Receta (Shade History Report)",
                    "Imágenes faltantes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            lblStatus.Text = "Procesando OCR…";
            Cursor = Cursors.WaitCursor;

            try
            {
                // 1) Extraer RECETA (imagen derecha)
                _lastShadeResult = _shadeExtractor.ExtractFromBitmap(
                    new Bitmap(picRight.Image));

                // 2) Extraer OcrReport COMPLETO de imagen izquierda (incluye CMC)
                var medExtractor = new ColorimetricDataExtractor(@".\tessdata");
                OcrReport ocrReport = medExtractor.ExtractReportFromBitmap(
                    new Bitmap(picLeft.Image));

                if (ocrReport == null || ocrReport.Measures == null || ocrReport.Measures.Count == 0)
                {
                    MessageBox.Show(
                        "No se encontraron datos colorimétricos en la imagen de mediciones.\n" +
                        "Verifica que la imagen contenga una tabla con D65, TL84 o CWF.",
                        "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text = "Sin datos de mediciones.";
                    return;
                }

                // 3) Mostrar formulario de CONFIRMACIÓN con OcrReport completo (muestra CMC)
                List<ColorimetricRow> rowsConfirmados;
                using (var dlgConfirm = new FormConfirmacionOCR(ocrReport, _lastShadeResult))
                {
                    dlgConfirm.MainFormOwner = this; // minimiza Form1 al aparecer
                    Cursor = Cursors.Default;
                    var dlgResult = dlgConfirm.ShowDialog(this);

                    // Restaurar Form1 siempre al cerrar el dialog
                    this.WindowState = FormWindowState.Normal;
                    this.BringToFront();

                    if (dlgResult != DialogResult.OK)
                    {
                        lblStatus.Text = "Operación cancelada.";
                        return;
                    }

                    rowsConfirmados = dlgConfirm.RowsConfirmed;
                }

                Cursor = Cursors.WaitCursor;

                // 4) Calcular correcciones colorimetricas (existente - no se modifica)
                var correcciones = ColorimetricCalculator.Calculate(rowsConfirmados);

                // 4b) Calcular correcciones de receta (NUEVO - complementario)
                var ingredientes = RecipeCorrector.IngredientsFromShade(_lastShadeResult);
                var deltasReceta = RecipeCorrector.DeltasFromReport(ocrReport);

                var correccionesReceta = RecipeCorrector.Calculate(ingredientes, deltasReceta);

                // 5) Construir texto combinado (receta + mediciones)
                string textoFinal = BuildResumenReceta(_lastShadeResult)
                    + System.Environment.NewLine + System.Environment.NewLine
                    + "Mediciones procesadas correctamente.";

                // 6) Mostrar resultados (con correcciones de receta adicionales)
                using (var frmRes = new FormResultados(textoFinal, correcciones,
                    correccionesReceta as List<Color.IlluminantCorrectionResult>))
                {
                    this.WindowState = FormWindowState.Minimized;
                    frmRes.ShowDialog(this);
                    this.WindowState = FormWindowState.Normal;
                    this.BringToFront();
                }

                lblStatus.Text = "Escaneo completado.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error durante el escaneo:\n\n" + ex.Message,
                    "Error OCR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Error en el escaneo.";
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // -------------------- ► NUEVO: mostrar resultados --------------------
        private void MostrarResultados(string texto)
        {
            var frm = new Form
            {
                Text = "Resultados — Colorimetría",
                Size = new Size(900, 600),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                MinimizeBox = true,
                MaximizeBox = true
            };

            var txt = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9.5f),
                BackColor = System.Drawing.Color.FromArgb(20, 20, 20),
                ForeColor = System.Drawing.Color.LightGreen,
                ReadOnly = true,
                WordWrap = false,
                Text = texto
            };

            var btnCerrar = new Button
            {
                Text = "Cerrar",
                Dock = DockStyle.Bottom,
                Height = 36,
                BackColor = System.Drawing.Color.FromArgb(60, 60, 60),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCerrar.Click += (s, e) => frm.Close();

            frm.Controls.Add(txt);
            frm.Controls.Add(btnCerrar);
            frm.ShowDialog(this);
        }

        // -------------------- ► NUEVO: resumen de mediciones ----------------
        private static string BuildResumenMediciones(
            List<ColorimetricRow> rows, List<ColorCorrectionResult> correcciones)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine(" MEDICIONES COLORIMÉTRICAS");
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine(string.Format(" {0,-6} {1,-4} {2,7} {3,7} {4,7} {5,7} {6,6}",
                "Illum", "Tipo", "L*", "a*", "b*", "C*", "h°"));
            sb.AppendLine(" " + new string('─', 52));

            foreach (var r in rows)
                sb.AppendLine(string.Format(" {0,-6} {1,-4} {2,7:F2} {3,7:F2} {4,7:F2} {5,7:F2} {6,6:F1}",
                    r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue));

            if (correcciones != null && correcciones.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine(" CORRECCIONES (Lot - Std):");
                sb.AppendLine(string.Format(" {0,-6} {1,7} {2,7} {3,7} {4,7}",
                    "Illum", "ΔL", "Δa", "Δb", "ΔE"));
                sb.AppendLine(" " + new string('─', 42));

                foreach (var c in correcciones)
                    sb.AppendLine(string.Format(" {0,-6} {1,7:+0.00;-0.00} {2,7:+0.00;-0.00} {3,7:+0.00;-0.00} {4,7:F2}",
                        c.Illuminant, c.DeltaL, c.DeltaA, c.DeltaB, c.DeltaE));
            }

            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        // -------------------- ► NUEVO: resumen de receta --------------------
        private static string BuildResumenReceta(ShadeExtractionResult result)
        {
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine(" RECETA EXTRAÍDA  (Shade History Report)");
            sb.AppendLine(" Fecha: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            sb.AppendLine("═══════════════════════════════════════════════════════════════");

            if (result == null || !result.Success)
            {
                sb.AppendLine(" No se encontraron datos de receta en la imagen.");
                return sb.ToString();
            }

            // Ingredientes
            if (result.Recipe != null && result.Recipe.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine(" INGREDIENTES:");
                sb.AppendLine(" " + new string('─', 58));
                sb.AppendLine(string.Format(" {0,-10} {1,-32} {2,8}", "CÓDIGO", "NOMBRE", "%"));
                sb.AppendLine(" " + new string('─', 58));

                foreach (var item in result.Recipe)
                    sb.AppendLine(string.Format(" {0,-10} {1,-32} {2,8}",
                        item.Code, item.Name, item.Percentage));
            }
            else
            {
                sb.AppendLine(" No se detectaron ingredientes.");
            }

            // Valores LAB
            sb.AppendLine();
            if (result.Lab != null)
            {
                sb.AppendLine(" VALORES LAB:");
                sb.AppendLine(" " + new string('─', 58));
                sb.AppendLine(string.Format("  L={0}  A={1}  B={2}",
                    result.Lab.L, result.Lab.A, result.Lab.B));
                sb.AppendLine(string.Format("  dL={0}  da={1}  dB={2}  CDE={3}  P/F={4}",
                    result.Lab.dL, result.Lab.da, result.Lab.dB,
                    result.Lab.cde, result.Lab.PF));
            }
            else
            {
                sb.AppendLine(" No se detectaron valores LAB.");
            }

            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        // -------------------- Drag & Drop --------------------
        private void EnableDragDrop(Control dropSurface, PictureBox target, Label hint, string etiqueta)
        {
            dropSurface.AllowDrop = true;
            target.AllowDrop = true;

            DragEventHandler onDragEnter = (s, e) =>
            {
                if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    e.Effect = ContainsPng(files) ? DragDropEffects.Copy : DragDropEffects.None;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            };

            DragEventHandler onDragDrop = (s, e) =>
            {
                try
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    string png = FirstPng(files);
                    if (!string.IsNullOrEmpty(png))
                        LoadInto(target, hint, png, etiqueta);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("No se pudo cargar la imagen.\n\n" + ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            dropSurface.DragEnter += onDragEnter;
            target.DragEnter += onDragEnter;
            dropSurface.DragDrop += onDragDrop;
            target.DragDrop += onDragDrop;
        }

        private static bool ContainsPng(string[] files)
        {
            if (files == null) return false;
            return files.Any(f => string.Equals(Path.GetExtension(f), ".png",
                StringComparison.OrdinalIgnoreCase));
        }

        private static string FirstPng(string[] files)
        {
            if (files == null) return null;
            return files.FirstOrDefault(f => string.Equals(Path.GetExtension(f), ".png",
                StringComparison.OrdinalIgnoreCase));
        }

        // -------------------- Carga de imágenes --------------------
        private void SelectAndLoadPng(PictureBox target, Label hint, string etiqueta)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = "Selecciona una imagen PNG";
                ofd.Filter = "PNG (*.png)|*.png|Todos los archivos (*.*)|*.*";
                ofd.Multiselect = false;

                if (ofd.ShowDialog(this) == DialogResult.OK)
                    LoadInto(target, hint, ofd.FileName, etiqueta);
            }
        }

        private void LoadInto(PictureBox target, Label hint, string path, string etiqueta)
        {
            var ext = Path.GetExtension(path);
            if (!string.Equals(ext, ".png", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Formato no permitido. Seleccione un archivo PNG.");

            if (target.Image != null)
            {
                var old = target.Image;
                target.Image = null;
                old.Dispose();
            }

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var temp = Image.FromStream(fs))
            {
                target.Image = new Bitmap(temp);
            }

            if (hint != null) hint.Visible = (target.Image == null);

            var fi = new FileInfo(path);
            string sizeText = fi.Length < 1024 * 1024
                ? (fi.Length / 1024.0).ToString("0.#") + " KB"
                : (fi.Length / (1024.0 * 1024.0)).ToString("0.##") + " MB";

            lblStatus.Text = etiqueta + ": " + Path.GetFileName(path) +
                             "  " + target.Image.Width + "×" + target.Image.Height + " px  " + sizeText;

            string fileLabel = Path.GetFileName(path) +
                               $"   {target.Image.Width}×{target.Image.Height} px   {sizeText}";

            if (ReferenceEquals(target, picLeft))
            {
                btnCargarLeft.Visible = false;

                if (lblLeftLoaded == null)
                {
                    lblLeftLoaded = new Label
                    {
                        AutoSize = true,
                        Font = new Font("Segoe UI", 9F),
                        ForeColor = System.Drawing.Color.FromArgb(60, 64, 70)
                    };
                    lblLeftLoaded.Location = new Point(btnCargarLeft.Left, btnCargarLeft.Top + 6);
                    contentBorder.Controls.Add(lblLeftLoaded);
                }

                lblLeftLoaded.Text = fileLabel;
                lblLeftLoaded.Visible = true;
            }
            else
            {
                btnCargarRight.Visible = false;

                if (lblRightLoaded == null)
                {
                    lblRightLoaded = new Label
                    {
                        AutoSize = true,
                        Font = new Font("Segoe UI", 9F),
                        ForeColor = System.Drawing.Color.FromArgb(60, 64, 70)
                    };
                    lblRightLoaded.Location = new Point(btnCargarRight.Left, btnCargarRight.Top + 6);
                    contentBorder.Controls.Add(lblRightLoaded);
                }

                lblRightLoaded.Text = fileLabel;
                lblRightLoaded.Visible = true;
            }

            CheckIfBothImagesLoaded();
        }

        private void ClearPicture(PictureBox pb)
        {
            if (pb != null && pb.Image != null)
            {
                var img = pb.Image;
                pb.Image = null;
                img.Dispose();
            }
        }

        // -------------------- Estado inferior --------------------
        private void CheckIfBothImagesLoaded()
        {
            bool leftLoaded = picLeft.Image != null;
            bool rightLoaded = picRight.Image != null;
            ShowActionButtons(leftLoaded && rightLoaded);
        }

        private void ShowActionButtons(bool visible)
        {
            btnIniciar.Visible = visible;
            btnCancelarAccion.Visible = visible;
        }

        private void LayoutBottomArea()
        {
            int centerX = mainArea.ClientSize.Width / 2;
            int statusWidth = 520;

            lblStatus.Size = new Size(statusWidth, 24);
            lblStatus.Location = new Point(centerX - statusWidth / 2,
                                           mainArea.ClientSize.Height - 110);

            int btnW = 160, btnH = 36, gap = 16;
            int totalW = btnW * 2 + gap;
            int startX = centerX - totalW / 2;
            int y = mainArea.ClientSize.Height - 70;

            btnIniciar.Size = new Size(btnW, btnH);
            btnCancelarAccion.Size = new Size(btnW, btnH);

            btnIniciar.Location = new Point(startX, y);
            btnCancelarAccion.Location = new Point(btnIniciar.Right + gap, y);
        }

        private void PositionExitButtonAtBottom()
        {
            if (btnSalir == null || leftNav == null) return;

            int margin = 20;
            btnSalir.Location = new Point(
                20,
                Math.Max(210, leftNav.Height - btnSalir.Height - margin));
        }

        private void UpdateHints()
        {
            if (lblLeftHint != null) lblLeftHint.Visible = (picLeft.Image == null);
            if (lblRightHint != null) lblRightHint.Visible = (picRight.Image == null);
        }

        // -------------------- Limpieza --------------------
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            ClearPicture(picLeft);
            ClearPicture(picRight);
            base.OnFormClosing(e);
        }
    }
}