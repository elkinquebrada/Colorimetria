using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Color.Services;
using System.Threading.Tasks;
namespace Color
{
    public partial class Form1 : Form
    {
        private Label lblLeftLoaded;
        private Label lblRightLoaded;
        private Button btnCambiarLeft;
        private Button btnCambiarRight;

        // Extractor de receta e instancia de resultado
        private readonly ShadeReportExtractor       _shadeExtractor    = new ShadeReportExtractor(@".\tessdata");
        private readonly TextileMetadataExtractor    _textileExtractor  = new TextileMetadataExtractor(@".\tessdata");
        private readonly DynamicSplitGridExtractor   _splitExtractor    = new DynamicSplitGridExtractor(@".\tessdata");
        // Ruta de la imagen del Shade History Report actualmente cargada
        private string _lastShadeImagePath;
        private ShadeExtractionResult _lastShadeResult;
        private Color.Models.TextileMetadata _lastTextileMetadata;

        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            this.TopMost = false;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            WireEvents();
            UpdateHints();
            LayoutBottomArea();
            PositionExitButtonAtBottom();
            MinimizarNavegador();
            AddBrandingLogo();
        }

        private void AddBrandingLogo()
        {
            try
            {
                string finalPath = null;
                string currentDir = AppDomain.CurrentDomain.BaseDirectory;
                for (int i = 0; i < 5; i++)
                {
                    string candidate = Path.Combine(currentDir, "logicDocs", "Coats_logo.svg.png");
                    if (File.Exists(candidate)) { finalPath = candidate; break; }
                    currentDir = Path.GetDirectoryName(currentDir);
                    if (string.IsNullOrEmpty(currentDir)) break;
                }

                if (string.IsNullOrEmpty(finalPath)) return;

                var logo = new PictureBox
                {
                    Image = Image.FromFile(finalPath),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Width = 80,
                    Height = 80,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    BackColor = System.Drawing.Color.Transparent
                };
                
                // Agregar al mainArea para que no lo cubra el panel lateral
                logo.Location = new Point(this.mainArea.Width - logo.Width - 30, 20);
                this.mainArea.Controls.Add(logo);
                logo.BringToFront();
            }
            catch { }
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
            // Desactivado para permitir multitasking según requerimiento del cliente.
        }
        #endregion

        private void WireEvents()
        {
            btnCargarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "PASS / FAIL");
            btnCargarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Shade History Report");

            btnCambiarLeft = CrearBotonCambiar();
            btnCambiarLeft.Click += (s, e) => SelectAndLoadPng(picLeft, lblLeftHint, "PASS / FAIL");
            contentBorder.Controls.Add(btnCambiarLeft);

            btnCambiarRight = CrearBotonCambiar();
            btnCambiarRight.Click += (s, e) => SelectAndLoadPng(picRight, lblRightHint, "Shade History Report");
            contentBorder.Controls.Add(btnCambiarRight);

            EnableDragDrop(pnlLeftFrame, picLeft, lblLeftHint, "PASS / FAIL");
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
            Text = " Cambiar imagen",
            Size = new Size(160, 32),
            BackColor = System.Drawing.Color.FromArgb(30, 90, 180),
            ForeColor = System.Drawing.Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Visible = false,
            Cursor = Cursors.Hand
        };

        private async Task LoadIntoAsync(PictureBox target, Label hint, string path, string etiqueta)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            string[] allowedExts = { ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif" };
            if (!allowedExts.Contains(ext))
            {
                MessageBox.Show("Formato de imagen no soportado. Seleccione .png, .jpg, .jpeg, .bmp, .tiff o .gif.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Cursor = Cursors.WaitCursor;
            lblStatus.Text = $"Analizando y extrayendo datos de {etiqueta}...";
            lblStatus.ForeColor = System.Drawing.Color.Blue;
            btnIniciar.Enabled = false;

            Bitmap tempBmp = null;
            ShadeExtractionResult shadeRes = null;
            Color.Models.TextileMetadata textRes = null;

            try
            {
                await Task.Run(() =>
                {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                    {
                        tempBmp = new Bitmap(Image.FromStream(fs));
                    }

                    if (etiqueta == "Shade History Report")
                    {
                        shadeRes = OnShadeHistoryImageLoaded(path, tempBmp);
                    }
                    else if (etiqueta == "PASS / FAIL")
                    {
                        textRes = _textileExtractor.ExtractFromBitmap(tempBmp);
                    }
                });

                if (etiqueta == "Shade History Report")
                {
                    _lastShadeImagePath = path;
                    _lastShadeResult = shadeRes;
                }
                else if (etiqueta == "PASS / FAIL")
                {
                    _lastTextileMetadata = textRes;
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
            finally
            {
                Cursor = Cursors.Default;
                btnIniciar.Enabled = true;
            }
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

                // --- TRANSICIÓN FLUIDA ---
                this.TopMost = false;
                this.WindowState = FormWindowState.Minimized;

                OcrReport.SetLastReport(ocrMediciones);
                using (var dlgConfirm = new Colorimetria.FormConfirmacionOCR(ocrMediciones, _lastShadeResult))
                {
                    dlgConfirm.MainFormOwner = this;

                    bool volverAConfirmar = true;
                    while (volverAConfirmar)
                    {
                        volverAConfirmar = false;
                        if (dlgConfirm.ShowDialog() == DialogResult.OK)
                        {
                            // Sincronizar reporte con ediciones del usuario
                            dlgConfirm.Report.Measures = dlgConfirm.RowsConfirmed;

                            // Motores Industriales (D65, TL84, A)
                            var correcciones = ColorimetricCalculator.CalculateAllIlluminants(dlgConfirm.Report);
                            var mainResult = correcciones.FirstOrDefault(r => r.Illuminant == "D65") ?? correcciones.FirstOrDefault();

                            // Motor de Receta (Basado en D65)
                            var ingredientes = RecipeCorrector.IngredientsFromShade(_lastShadeResult);
                            var corrReceta = new List<CorrectiveRecipeResult> { 
                                RecipeCorrector.CalculateCorrectiveRecipe(ingredientes, mainResult) 
                            };

                            using (var frmRes = new FormResultados(BuildResumenReceta(_lastShadeResult), correcciones, corrReceta, _lastShadeResult))
                            {
                                if (_lastTextileMetadata != null)
                                {
                                    frmRes.UpdateTextileMetadataPanel(_lastTextileMetadata);
                                }

                                if (frmRes.ShowDialog() == DialogResult.Retry)
                                {
                                    volverAConfirmar = true; 
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
            finally 
            { 
                Cursor = Cursors.Default;
                lblStatus.Text = "";

                // Restaurar estado normal al finalizar sin bloquear el sistema
                this.WindowState = FormWindowState.Maximized;
            }
        }

        private void BtnCancelarAccion_Click(object sender, EventArgs e)
        {
            ClearPicture(picLeft); ClearPicture(picRight);
            _lastShadeResult = null;
            _lastTextileMetadata = null;
            btnCargarLeft.Visible = btnCargarRight.Visible = true;
            if (lblLeftLoaded != null) lblLeftLoaded.Visible = false;
            if (lblRightLoaded != null) lblRightLoaded.Visible = false;
            btnCambiarLeft.Visible = btnCambiarRight.Visible = false;

            lblStatus.Text = "Cargue Imágenes";
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

        private async void SelectAndLoadPng(PictureBox target, Label hint, string etiqueta)
        {
            using (var ofd = new OpenFileDialog { Filter = "Archivos de imagen|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|Todos los archivos|*.*" })
                if (ofd.ShowDialog() == DialogResult.OK) await LoadIntoAsync(target, hint, ofd.FileName, etiqueta);
        }

        private void EnableDragDrop(Control surf, PictureBox target, Label hint, string etiqueta)
        {
            surf.AllowDrop = target.AllowDrop = true;
            surf.DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            surf.DragDrop += async (s, e) => {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0) await LoadIntoAsync(target, hint, files[0], etiqueta);
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

        // ── Orquestador del pipeline geométrico (DESACOPLADO) ─────────────────

        /// Clasifica el reporte y extrae la receta con lógica limpia y desacoplada.
        private ShadeExtractionResult OnShadeHistoryImageLoaded(string imagePath, Bitmap bmp)
        {
            // 1. El Clasificador identifica qué tipo de reporte envió el usuario
            Color.Services.ReportFormatType format =
                Color.Services.ReportFormatRouter.DetermineFormat(imagePath);

            ShadeExtractionResult shadeResult;

            if (format == Color.Services.ReportFormatType.LegacyCombinedFormat)
            {
                // ══════════════════════════════════════════════════════════════
                // REQUERIMIENTO 1: Reporte Completo (Extractor Tradicional)
                // ══════════════════════════════════════════════════════════════
                shadeResult = _shadeExtractor.ExtractFromBitmap(bmp);

                // Limpiar porcentajes a valores numéricos puros
                if (shadeResult.Recipe != null)
                {
                    shadeResult.Recipe = CleanRecipeItems(shadeResult.Recipe);
                }

                // Diagnóstico y fallback si el tradicional no encontró receta
                if (shadeResult.Recipe == null || shadeResult.Recipe.Count == 0)
                {
                    WriteDiag(imagePath, bmp, format, shadeResult.Recipe, "LegacyCombined→FALLBACK a Split");
                    var fallbackItems = _splitExtractor.ExtractFromBitmap(bmp, null);
                    if (fallbackItems != null && fallbackItems.Count > 0)
                    {
                        shadeResult.Recipe = CleanRecipeItems(fallbackItems);
                        WriteDiag(imagePath, bmp, format, shadeResult.Recipe, "Fallback Split → OK");
                    }
                    else
                    {
                        WriteDiag(imagePath, bmp, format, shadeResult.Recipe, "LegacyCombined → VACÍO (ambos fallaron)");
                    }
                }
                else
                {
                    WriteDiag(imagePath, bmp, format, shadeResult.Recipe, "LegacyCombined → OK");
                }
            }
            else
            {
                // ══════════════════════════════════════════════════════════════
                // REQUERIMIENTO 2: Tickets planos de matriz de puntos
                // ══════════════════════════════════════════════════════════════
                var cleanRecipeItems = _splitExtractor.ExtractFromBitmap(bmp, null);

                // ── FALLBACK: si el split no encontró nada, probar el extractor tradicional ──
                if (cleanRecipeItems == null || cleanRecipeItems.Count == 0)
                {
                    WriteDiag(imagePath, bmp, format, cleanRecipeItems, "Split→FALLBACK a Traditional");
                    var fallback = _shadeExtractor.ExtractFromBitmap(bmp);
                    if (fallback?.Recipe != null && fallback.Recipe.Count > 0)
                    {
                        cleanRecipeItems = fallback.Recipe;
                        WriteDiag(imagePath, bmp, format, cleanRecipeItems, "Fallback Traditional → OK");
                    }
                }
                else
                {
                    WriteDiag(imagePath, bmp, format, cleanRecipeItems, "Split → OK");
                }

                var finalRecipe = CleanRecipeItems(cleanRecipeItems ?? new List<RecipeItem>());

                // Inicializa el resultado solo con la receta pura para la grilla visual
                shadeResult = new ShadeExtractionResult
                {
                    Recipe = finalRecipe
                };
            }

            return shadeResult;
        }

        /// Limpia porcentajes a valores numéricos puros.
        private static List<RecipeItem> CleanRecipeItems(List<RecipeItem> items)
        {
            var clean = new List<RecipeItem>();
            foreach (var item in items)
            {
                string pctDigits = System.Text.RegularExpressions.Regex.Replace(
                    item.Percentage ?? "", @"[^0-9\.]", "");
                double cleanPct = 0;
                double.TryParse(pctDigits,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out cleanPct);
                clean.Add(new RecipeItem
                {
                    Code = item.Code,
                    Name = item.Name,
                    Percentage = cleanPct.ToString(System.Globalization.CultureInfo.InvariantCulture)
                });
            }
            return clean;
        }

        /// Escribe diagnóstico completo en C:\Temp\shade_diag.txt para debugging.
        private static void WriteDiag(string imagePath, Bitmap bmp,
            Color.Services.ReportFormatType format,
            List<RecipeItem> items, string etapa)
        {
            try
            {
                System.IO.Directory.CreateDirectory(@"C:\Temp");
                string logPath = @"C:\Temp\shade_diag.txt";
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("══════════════════════════════════════════════");
                sb.AppendLine($"[{DateTime.Now:HH:mm:ss}] ETAPA: {etapa}");
                sb.AppendLine($"Imagen : {System.IO.Path.GetFileName(imagePath)}");
                sb.AppendLine($"Tamaño : {bmp?.Width}x{bmp?.Height} px");
                sb.AppendLine($"Formato: {format}");

                // Calcular zona que usa DynamicSplitGridExtractor para ticket plano
                if (bmp != null)
                {
                    int zoneTop    = (int)(bmp.Height * 0.18);
                    int zoneHeight = (int)(bmp.Height * 0.25);
                    sb.AppendLine($"Zona Split: top={zoneTop}px, h={zoneHeight}px, bot={zoneTop + zoneHeight}px");
                    sb.AppendLine($"Col CODE : 0..{(int)(bmp.Width * 0.15)}px");
                    sb.AppendLine($"Col NAME : {(int)(bmp.Width * 0.15)}..{(int)(bmp.Width * 0.53)}px");
                    sb.AppendLine($"Col PCT  : {(int)(bmp.Width * 0.53)}..{(int)(bmp.Width * 0.68)}px");
                }

                int count = items?.Count ?? 0;
                sb.AppendLine($"Items encontrados: {count}");
                if (items != null)
                    foreach (var it in items)
                        sb.AppendLine($"  → [{it.Code}] {it.Name} | {it.Percentage}");

                sb.AppendLine();
                System.IO.File.AppendAllText(logPath, sb.ToString());
            }
            catch { /* Diagnóstico no crítico */ }
        }

        #endregion
    }
}
