using System.IO;
using Color.Services;
using Color.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using EngineCalc = Color.ColorimetricCalculator;
using EngineRes = Color.ColorCorrectionResult;


namespace Color
{
    public class FormResultados : Form
    {
        private readonly OcrReport _report;
        private readonly List<EngineRes> _resultsLegacy;
        private readonly List<CorrectiveRecipeResult> _recipeResults;
        private readonly ShadeExtractionResult _shadeData;
        private FlowLayoutPanel pnlReportFlow;
        private IluminantReportBlock blockD65;
        private IluminantReportBlock blockTL84;
        private IluminantReportBlock blockCWF;
        private Panel pnlWhitePaper;
        private Label lblRightShadeValue;
        private DataGridView dgvShadeHistory;
        private DataGridView dgvCorrectiveRecipe;
        private Button btnGuardar;
        private Button btnCerrar;
        private Button btnRegresar;
        private CielabChartControl _cielabChart;
        private EngineRes _lastMainResult;
        private TableLayoutPanel pnlTolerances;
        private Label lblTolDe, lblTolDl, lblTolDc, lblTolDh;
        private Label lblValTolDe;
        private Label lblValTolDl;
        private Label lblValTolDc;
        private Label lblValTolDh;
        private Label lblTypeTolDe;
        private Label lblTypeTolDl;
        private Label lblTypeTolDc;
        private Label lblTypeTolDh;

        // Etiquetas del panel de metadatos textiles (superior izquierdo)
        private Label lblValueShadeName;
        private Label lblValueDyeingClass;
        private Label lblValueSubstrate;
        private Label lblValueCountPly;
        private Label lblValueFiberType;

        public FormResultados(OcrReport report)
        {
            _report = report ?? new OcrReport();
            InitializeComponents();
            PopulateFromReport(_report);
            AddBrandingLogo();
        }

        public FormResultados(string _, List<EngineRes> results, List<CorrectiveRecipeResult> recipeResults = null, ShadeExtractionResult shadeData = null)
        {
            _resultsLegacy = results ?? new List<EngineRes>();
            _recipeResults = recipeResults;
            _shadeData = shadeData;
            InitializeComponents();
            PopulateFromObjects(_shadeData, _resultsLegacy);
            AddBrandingLogo();
        }

        private void InitializeComponents()
        {
            this.Text = "TINT COATS CADENA";
            this.Size = new Size(1100, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = System.Drawing.Color.FromArgb(230, 230, 230);

            pnlWhitePaper = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = System.Drawing.Color.FromArgb(230, 230, 230), Padding = new Padding(20) };
            pnlReportFlow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                BackColor = System.Drawing.Color.White,
                Padding = new Padding(30),
                Dock = DockStyle.Top
            };
            pnlWhitePaper.Controls.Add(pnlReportFlow);

            // Cuando la ventana cambia de tamaño, el FlowPanel y los bloques se adaptan al nuevo ancho
            this.Resize += (s, e) =>
            {
                int nuevoAncho = Math.Max(900, pnlWhitePaper.ClientSize.Width - 40);
                pnlReportFlow.Width = nuevoAncho;
                if (blockD65 != null) { blockD65.Width = nuevoAncho - 60; blockD65.Height = 240; }
                if (blockTL84 != null) { blockTL84.Width = nuevoAncho - 60; blockTL84.Height = 240; }
                if (blockCWF != null) { blockCWF.Width = nuevoAncho - 60; blockCWF.Height = 240; }
            };

            // CABECERA
            var pnlHeader = new Panel { Width = 940, Height = 100, Margin = new Padding(0, 0, 0, 20) };
            pnlHeader.Controls.Add(new Label
            {
                Text = "Análisis de Color",
                BackColor = System.Drawing.Color.FromArgb(0, 102, 204),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Size = new Size(940, 40),
                Location = new Point(0, 30)
            });
            pnlHeader.Controls.Add(new Label { Text = "Shade History Report", Font = new Font("Segoe UI", 11, FontStyle.Bold), Location = new Point(0, 75), AutoSize = true });
            pnlReportFlow.Controls.Add(pnlHeader);

            // INFO GENERAL
            var pnlTopInfo = new TableLayoutPanel { Width = 940, Height = 180, ColumnCount = 3, Margin = new Padding(0, 0, 0, 20) };
            pnlTopInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            pnlTopInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnlTopInfo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            dgvShadeHistory = CreateStyledGrid();
            dgvShadeHistory.ColumnHeadersVisible = true;
            dgvShadeHistory.ColumnCount = 4;
            dgvShadeHistory.Columns[0].HeaderText = "Dye code";
            dgvShadeHistory.Columns[1].HeaderText = "Dye name";
            dgvShadeHistory.Columns[2].HeaderText = "Concentration ";
            dgvShadeHistory.Columns[3].HeaderText = "Proportion ";
            SetupShadeHistoryPainting();

            lblRightShadeValue = new Label { Text = "Shade Name:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Regular) };
            lblValueShadeName = new Label { Text = "-", AutoSize = true, Font = new Font("Segoe UI", 9) };
            lblValueDyeingClass = new Label { Text = "Dyeing Class: -", AutoSize = true };
            lblValueSubstrate = new Label { Text = "Substrate: -", AutoSize = true };
            lblValueCountPly = new Label { Text = "Count/Ply: -", AutoSize = true };
            lblValueFiberType = new Label { Text = "Fibre Type: -", AutoSize = true };

            var pnlGeneralInfo = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
            var pnlShadeName = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            pnlShadeName.Controls.Add(lblRightShadeValue);
            pnlGeneralInfo.Controls.Add(pnlShadeName);
            pnlGeneralInfo.Controls.Add(lblValueDyeingClass);
            pnlGeneralInfo.Controls.Add(lblValueSubstrate);
            pnlGeneralInfo.Controls.Add(lblValueCountPly);
            pnlGeneralInfo.Controls.Add(lblValueFiberType);

            pnlTolerances = CreateTolerancesTable();

            pnlTopInfo.Controls.Add(dgvShadeHistory, 0, 0);
            pnlTopInfo.Controls.Add(pnlGeneralInfo, 1, 0);
            pnlTopInfo.Controls.Add(pnlTolerances, 2, 0);
            pnlReportFlow.Controls.Add(pnlTopInfo);

            pnlTopInfo.Controls.Add(dgvShadeHistory, 0, 0);
            pnlTopInfo.Controls.Add(pnlGeneralInfo, 1, 0);
            pnlTopInfo.Controls.Add(pnlTolerances, 2, 0);
            pnlReportFlow.Controls.Add(pnlTopInfo);

            // =================================================================
            // --- DECLARACIÓN DE TÍTULOS (Cobertura Bloque 1 y 2) ---
            // =================================================================
            Panel pnlTitlesHeader = new Panel();
            pnlTitlesHeader.Width = 940;
            pnlTitlesHeader.Height = 35;

            pnlTitlesHeader.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            Padding marginTitles = new Padding(0, 15, 0, 5);
            pnlTitlesHeader.Margin = marginTitles;

            Font fontTitulos = new Font("Segoe UI", 10, FontStyle.Bold);
            System.Drawing.Color negroPuro = System.Drawing.Color.Black;

            // --- CALCULO DE POSICIÓN PARA EL ACOPLE DE BLOQUES --
            int posXDerecha = (int)(pnlTitlesHeader.Width * 0.58);

            // 1. TÍTULO IZQUIERDO: L, a, b (lot - std)
            Label lblLabTitle = new Label();
            lblLabTitle.Text = "L, a, b (Lot - Std)";
            lblLabTitle.Font = fontTitulos;
            lblLabTitle.ForeColor = negroPuro;
            Point locLabTitle = new Point(0, 5);
            lblLabTitle.Location = locLabTitle;
            lblLabTitle.AutoSize = true;

            // Línea negra izquierda
            Label lineLab = new Label();
            lineLab.BackColor = negroPuro;
            Point locLineLab = new Point(0, 28);
            lineLab.Location = locLineLab;
            Size sizeLineLab = new Size(posXDerecha - 15, 2);
            lineLab.Size = sizeLineLab;
            lineLab.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // 2. TÍTULO CENTRAL: 
            Label lblLchTitle = new Label();
            lblLchTitle.Text = "L, C, H ";
            lblLchTitle.Font = fontTitulos;
            lblLchTitle.ForeColor = negroPuro;
            Point locLchTitle = new Point(posXDerecha, 5);
            lblLchTitle.Location = locLchTitle;
            lblLchTitle.AutoSize = true;
            lblLchTitle.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // Línea negra derecha
            Label lineLch = new Label();
            lineLch.BackColor = negroPuro;
            Point locLineLch = new Point(posXDerecha, 28);
            lineLch.Location = locLineLch;
            Size sizeLineLch = new Size(395, 2);
            lineLch.Size = sizeLineLch;
            lineLch.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // Agregamos todos los elementos al panel contenedor
            pnlTitlesHeader.Controls.Add(lblLabTitle);
            pnlTitlesHeader.Controls.Add(lineLab);
            pnlTitlesHeader.Controls.Add(lblLchTitle);
            pnlTitlesHeader.Controls.Add(lineLch);

            pnlReportFlow.Controls.Add(pnlTitlesHeader);
            pnlReportFlow.PerformLayout();

            // Agregamos todos los elementos al panel contenedor
            pnlTitlesHeader.Controls.Add(lblLabTitle);
            pnlTitlesHeader.Controls.Add(lineLab);
            pnlTitlesHeader.Controls.Add(lblLchTitle);
            pnlTitlesHeader.Controls.Add(lineLch);

            // Inyectamos la barra de títulos completa al flujo del reporte
            pnlReportFlow.Controls.Add(pnlTitlesHeader);

            // BLOQUES - Ancho responsivo controlado por el evento Resize
            int anchoInicial = Math.Max(900, this.ClientSize.Width - 80);
            blockD65 = new IluminantReportBlock { Width = anchoInicial, Height = 240, Margin = new Padding(0, 0, 0, 10) };
            blockTL84 = new IluminantReportBlock { Width = anchoInicial, Height = 240, Margin = new Padding(0, 0, 0, 10) };
            blockCWF = new IluminantReportBlock { Width = anchoInicial, Height = 240, Margin = new Padding(0, 0, 0, 10) };

            pnlReportFlow.Controls.Add(blockD65);
            pnlReportFlow.Controls.Add(blockTL84);
            pnlReportFlow.Controls.Add(blockCWF);

            // RECETA Y GRAFICO 
            dgvCorrectiveRecipe = CreateCorrectiveGrid();
            _cielabChart = new CielabChartControl();

            // BOTONES
            var pnlButtons = new Panel { Dock = DockStyle.Bottom, Height = 60, BackColor = System.Drawing.Color.White };
            btnGuardar = CreateStyledButton("Guardar", System.Drawing.Color.FromArgb(45, 126, 247));
            btnCerrar = CreateStyledButton("Finalizar", System.Drawing.Color.FromArgb(90, 90, 90));
            btnRegresar = CreateStyledButton("← Regresar", System.Drawing.Color.FromArgb(180, 100, 30));

            btnGuardar.Location = new Point(pnlButtons.Width - 150, 12); btnGuardar.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            btnCerrar.Location = new Point(15, 12);
            btnRegresar.Location = new Point(130, 12);

            pnlButtons.Controls.Add(btnGuardar); pnlButtons.Controls.Add(btnCerrar); pnlButtons.Controls.Add(btnRegresar);
            this.Controls.Add(pnlWhitePaper); this.Controls.Add(pnlButtons);

            btnGuardar.Click += BtnGuardar_Click;
            btnCerrar.Click += (s, e) => this.Close();
            btnRegresar.Click += (s, e) => { this.DialogResult = DialogResult.Retry; this.Close(); };
        }

        /// Inyecta los metadatos textiles extraídos por TextileMetadataExtractor
        public void UpdateTextileMetadataPanel(TextileMetadata metadata)
        {

            if (metadata == null) return;
            {
                lblValueShadeName.Text = !string.IsNullOrEmpty(metadata.ShadeName) && metadata.ShadeName != "-" ? metadata.ShadeName.ToUpper() : "-";
                lblValueDyeingClass.Text = "Dyeing Class:       " + (!string.IsNullOrEmpty(metadata.DyeingClass) && metadata.DyeingClass != "-" ? metadata.DyeingClass.ToUpper() : "-");
                lblValueSubstrate.Text = "Substrate:            " + (!string.IsNullOrEmpty(metadata.Substrate) && metadata.Substrate != "-" ? metadata.Substrate.ToUpper() : "-");
                lblValueCountPly.Text = "Count/Ply:             " + (!string.IsNullOrEmpty(metadata.CountPly) && metadata.CountPly != "-" ? metadata.CountPly : "-");
                lblValueFiberType.Text = "Fibre Type:           " + (!string.IsNullOrEmpty(metadata.FiberType) && metadata.FiberType != "-" ? metadata.FiberType.ToUpper() : "-");
            }
            // 3. CORRECCIÓN PARA EVITAR LA DUPLICACIÓN: Unificamos el texto en el control compuesto derecho
            if (lblRightShadeValue != null && !string.IsNullOrEmpty(metadata.ShadeName) && metadata.ShadeName != "-")
            {
                lblRightShadeValue.Text = "Shade Name:      " + metadata.ShadeName.ToUpper();
            }
        }
        public void ActualizarTablaTolerancias(double de, double dl, double dc, double dh)
        {
            if (lblValTolDe != null) lblValTolDe.Text = de.ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblValTolDl != null) lblValTolDl.Text = dl.ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblValTolDc != null) lblValTolDc.Text = dc.ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblValTolDh != null) lblValTolDh.Text = dh.ToString("0.00#", CultureInfo.InvariantCulture);

            if (lblTypeTolDe != null) lblTypeTolDe.Text = (de * -1).ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblTypeTolDl != null) lblTypeTolDl.Text = (dl * -1).ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblTypeTolDc != null) lblTypeTolDc.Text = (dc * -1).ToString("0.00#", CultureInfo.InvariantCulture);
            if (lblTypeTolDh != null) lblTypeTolDh.Text = (dh * -1).ToString("0.00#", CultureInfo.InvariantCulture);

            // Pasamos estas tolerancias también a los bloques visuales de los iluminantes
            if (blockD65 != null) blockD65.UpdateTolerances(de, dl, dc, dh);
            if (blockTL84 != null) blockTL84.UpdateTolerances(de, dl, dc, dh);
            if (blockCWF != null) blockCWF.UpdateTolerances(de, dl, dc, dh);
        }

        private TableLayoutPanel CreateTolerancesTable()
        {
            var pnl = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 4,
                RowCount = 4,
                Margin = new Padding(10, 0, 0, 0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                BackColor = System.Drawing.Color.White
            };
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.Height = 105;

            var lblTitle = new Label
            {
                Text = "Tolerancias",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                Margin = new Padding(0)
            };
            pnl.Controls.Add(lblTitle, 0, 0);
            pnl.SetColumnSpan(lblTitle, 4);

            pnl.Controls.Add(CreateGridLabel("DE = ", true), 0, 1);
            pnl.Controls.Add(CreateGridLabel("L", true), 1, 1);
            pnl.Controls.Add(CreateGridLabel("C", true), 2, 1);
            pnl.Controls.Add(CreateGridLabel("Hue", true), 3, 1);

            lblValTolDe = CreateGridLabel("-", false);
            lblValTolDl = CreateGridLabel("-", false);
            lblValTolDc = CreateGridLabel("-", false);
            lblValTolDh = CreateGridLabel("-", false);

            Font fontTipoTol = new Font("Segoe UI", 9, FontStyle.Bold);
            lblTypeTolDe = CreateGridLabel("-", false); lblTypeTolDe.Font = fontTipoTol;
            lblTypeTolDl = CreateGridLabel("-", false); lblTypeTolDl.Font = fontTipoTol;
            lblTypeTolDc = CreateGridLabel("-", false); lblTypeTolDc.Font = fontTipoTol;
            lblTypeTolDh = CreateGridLabel("-", false); lblTypeTolDh.Font = fontTipoTol;

            pnl.Controls.Add(lblValTolDe, 0, 2);
            pnl.Controls.Add(lblValTolDl, 1, 2);
            pnl.Controls.Add(lblValTolDc, 2, 2);
            pnl.Controls.Add(lblValTolDh, 3, 2);

            pnl.Controls.Add(lblTypeTolDe, 0, 3);
            pnl.Controls.Add(lblTypeTolDl, 1, 3);
            pnl.Controls.Add(lblTypeTolDc, 2, 3);
            pnl.Controls.Add(lblTypeTolDh, 3, 3);

            return pnl;
        }

        private Label CreateGridLabel(string text, bool isHeader)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = isHeader ? ContentAlignment.MiddleLeft : ContentAlignment.MiddleRight,
                BackColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 9, FontStyle.Regular),
                Margin = new Padding(0)
            };
        }

        private void SetupShadeHistoryPainting()
        {
            dgvShadeHistory.CellPainting += (s, e) => {
                if (e.ColumnIndex == 3 && e.RowIndex >= 0 && e.Value != null)
                {
                    e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
                    string valStr = e.Value.ToString().Replace("%", "");
                    string rowVal = dgvShadeHistory.Rows[e.RowIndex].Cells[0].Value?.ToString();
                    if (rowVal != "¨[Dyes]" && float.TryParse(valStr, out float porcentaje))
                    {
                        int barWidth = (int)((e.CellBounds.Width - 10) * (porcentaje / 100f));
                        using (var brush = new SolidBrush(System.Drawing.Color.FromArgb(220, 220, 220)))
                            e.Graphics.FillRectangle(brush, e.CellBounds.X + 5, e.CellBounds.Y + 4, barWidth, e.CellBounds.Height - 9);
                    }
                    e.PaintContent(e.CellBounds);
                    e.Handled = true;
                }
            };
        }

        private DataGridView CreateStyledGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 8.2f)
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.White;
            dgv.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            return dgv;
        }

        private DataGridView CreateCorrectiveGrid()
        {
            var dgv = CreateStyledGrid(); dgv.ColumnCount = 8;
            dgv.Columns[0].Name = "Colorante"; dgv.Columns[2].Name = "Receta 1"; dgv.Columns[3].Name = "Part %";
            dgv.Columns[4].Name = "Receta 2"; dgv.Columns[5].Name = "Part %"; dgv.Columns[6].Name = "Receta 3"; dgv.Columns[7].Name = "Part %";
            dgv.Columns[1].Visible = false;
            return dgv;
        }

        private Button CreateStyledButton(string text, System.Drawing.Color color)
        {
            return new Button { Text = text, Size = new Size(130, 35), BackColor = color, ForeColor = System.Drawing.Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        }

        private void AddBrandingLogo()
        {
            try
            {
                string finalPath = null;
                string currentDir = AppDomain.CurrentDomain.BaseDirectory;

                // Búsqueda dinámica en bucle ascendente de directorios
                for (int i = 0; i < 5; i++)
                {
                    string candidate = Path.Combine(currentDir, "logicDocs", "Coats_logo.svg.png");
                    if (File.Exists(candidate)) { finalPath = candidate; break; }
                    currentDir = Path.GetDirectoryName(currentDir);
                    if (string.IsNullOrEmpty(currentDir)) break;
                }

                // Si se encontró el archivo, creamos el PictureBox y lo acomodamos en el reporte
                if (!string.IsNullOrEmpty(finalPath))
                {
                    var logo = new PictureBox
                    {
                        Image = Image.FromFile(finalPath),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Width = 140,
                        Height = 40,
                        Anchor = AnchorStyles.Top | AnchorStyles.Right,
                        BackColor = System.Drawing.Color.Transparent
                    };

                    // Colocamos el logo flotando en la esquina superior derecha del formulario
                    logo.Location = new Point(this.Width - logo.Width - 40, 15);
                    this.Controls.Add(logo);
                    logo.BringToFront();
                }
            }
            catch (Exception)
            {

            }
        }

        private void PopulateFromObjects(ShadeExtractionResult shadeData, List<EngineRes> results)
        {
            // LEER LAS TOLERANCIAS ELEGIDAS POR EL USUARIO DESDE EL ORIGEN DE CONFIGURACIÓN
            double tolDE = Properties.Settings.Default.ToleranciaDE;
            double tolDL = Properties.Settings.Default.ToleranciaDL;
            double tolDC = Properties.Settings.Default.ToleranciaDC;
            double tolDH = Properties.Settings.Default.ToleranciaDH;

            if (lblValTolDe != null)
            {
                lblValTolDe.Text = tolDE.ToString("0.00#", CultureInfo.InvariantCulture);
                lblValTolDl.Text = tolDL.ToString("0.00#", CultureInfo.InvariantCulture);
                lblValTolDc.Text = tolDC.ToString("0.00#", CultureInfo.InvariantCulture);
                lblValTolDh.Text = tolDH.ToString("0.00#", CultureInfo.InvariantCulture);
            }
            if (lblTypeTolDe != null)
            {
                lblTypeTolDe.Text = (tolDE * -1).ToString("0.00#", CultureInfo.InvariantCulture);
                lblTypeTolDl.Text = (tolDL * -1).ToString("0.00#", CultureInfo.InvariantCulture);
                lblTypeTolDc.Text = (tolDC * -1).ToString("0.00#", CultureInfo.InvariantCulture);
                lblTypeTolDh.Text = (tolDH * -1).ToString("0.00#", CultureInfo.InvariantCulture);
            }

            if (shadeData != null)
            {
                lblRightShadeValue.Text = "Shade: " + (shadeData.ShadeName ?? "N/A");
                dgvShadeHistory.Rows.Clear();
                if (shadeData.Recipe != null)
                {
                    double total = shadeData.Recipe.Sum(ing => ParsePercentageValue(ing.Percentage));
                    foreach (var ing in shadeData.Recipe)
                    {
                        double p = total > 0 ? (ParsePercentageValue(ing.Percentage) / total * 100) : 0;
                        dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage.Replace("%", ""), ((int)Math.Round(p)).ToString() + "%");
                    }

                    int totalRowIdx = dgvShadeHistory.Rows.Add("[Dyes]", "", total.ToString("F5"), "100%");
                    dgvShadeHistory.Rows[totalRowIdx].DefaultCellStyle.Font = new Font(dgvShadeHistory.Font, FontStyle.Bold);
                    dgvShadeHistory.Rows[totalRowIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
                }
            }

            if (results != null && results.Count > 0)
            {
                var d65 = results.FirstOrDefault(r => r.Illuminant.Contains("D65")) ?? results[0];
                var tl84 = results.FirstOrDefault(r => r.Illuminant.Contains("TL84"));
                var cwf = results.FirstOrDefault(r => r.Illuminant.Contains("CWF")) ?? results.FirstOrDefault(r => r.Illuminant.Contains("A"));

                // FASE 2 del pliego: inyección limpia con refresco explícito de interfaz
                blockD65.UpdateData(d65);
                blockD65.Invalidate();

                if (tl84 != null) { blockTL84.UpdateData(tl84); blockTL84.Invalidate(); }

                // === CALCULO CRUZADO TL84: CMC(D65) vs CMC(TL84) ===
                if (tl84 != null && d65 != null)
                    blockTL84.SetSpecialCrossCmcResult(d65.CmcValue, tl84.CmcValue);
                if (cwf != null) { blockCWF.UpdateData(cwf); blockCWF.Invalidate(); }

                UpdateChart(d65); _lastMainResult = d65;

                if (shadeData != null)
                {
                    LoadCorrectiveRecipeGrid(d65);
                }
            }

            if (!pnlReportFlow.Controls.Contains(dgvCorrectiveRecipe))
            {
                var pnlNewRecipesHeader = new Panel { Width = 940, Height = 35, Margin = new Padding(0, 20, 0, 10) };
                var lblNewRecipesTitle = new Label
                {
                    Text = "New Recipes",
                    Font = new Font("Segoe UI", 10, FontStyle.Bold),
                    ForeColor = System.Drawing.Color.Black,
                    Location = new Point(0, 5),
                    AutoSize = true
                };
                var lineNewRecipes = new Label { BackColor = System.Drawing.Color.Black, Location = new Point(0, 28), Size = new Size(940, 2) };

                pnlNewRecipesHeader.Controls.Add(lblNewRecipesTitle);
                pnlNewRecipesHeader.Controls.Add(lineNewRecipes);

                // PRIMERO: Inyectamos el título al contenedor de flujo
                pnlReportFlow.Controls.Add(pnlNewRecipesHeader);

                //  LA TABLA DE RECETAS SE AGREGA INMEDIATAMENTE DESPUÉS
                dgvCorrectiveRecipe.Width = 940;
                dgvCorrectiveRecipe.Height = 150;
                pnlReportFlow.Controls.Add(dgvCorrectiveRecipe);

                //  CONDICIONES Y GRÁFICO LADO A LADO
                if (_cielabChart != null)
                {
                    var pnlBottomSplit = new TableLayoutPanel { Width = 940, Height = 320, ColumnCount = 2 };
                    pnlBottomSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
                    pnlBottomSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));

                    var pnlCondiciones = CreateConditionsPanel();
                    pnlCondiciones.Dock = DockStyle.Fill;

                    var pnlGraficoFlow = new Panel { Dock = DockStyle.Fill };
                    var lblGrTitle = new Label { Text = "Gráfico", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = System.Drawing.Color.Black, AutoSize = true, Location = new Point(0, 5) };
                    var lineGr = new Label { BackColor = System.Drawing.Color.Black, Size = new Size(400, 2), Location = new Point(0, 28) };

                    _cielabChart.Location = new Point(0, 35);
                    _cielabChart.Width = 510;
                    _cielabChart.Height = 280;

                    pnlGraficoFlow.Controls.Add(lblGrTitle);
                    pnlGraficoFlow.Controls.Add(lineGr);
                    pnlGraficoFlow.Controls.Add(_cielabChart);

                    pnlBottomSplit.Controls.Add(pnlCondiciones, 0, 0);
                    pnlBottomSplit.Controls.Add(pnlGraficoFlow, 1, 0);

                    pnlReportFlow.Controls.Add(pnlBottomSplit);
                }
            }
        }

        private Control CreateConditionsPanel()
        {
            var pnl = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, AutoScroll = true };

            var lblTitle = new Label { Text = "Condiciones", Font = new Font("Segoe UI", 10, FontStyle.Bold), ForeColor = System.Drawing.Color.Black, AutoSize = true, Margin = new Padding(0, 5, 0, 0) };
            var line = new Label { BackColor = System.Drawing.Color.Black, Size = new Size(380, 2), Margin = new Padding(0, 2, 0, 10) };
            pnl.Controls.Add(lblTitle); pnl.Controls.Add(line);

            // Intento de cargar la imagen que el usuario suele arrastrar luego a la carpeta logicDocs
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logicDocs", "Condiciones.png");
                if (File.Exists(path))
                {
                    var pic = new PictureBox { Image = Image.FromFile(path), SizeMode = PictureBoxSizeMode.Zoom, Width = 380, Height = 250 };
                    pnl.Controls.Add(pic);
                    return pnl;
                }
            }
            catch { }

            //  matriz visual nativa
            Font f8 = new Font("Segoe UI", 8); Font f8b = new Font("Segoe UI", 8, FontStyle.Bold);

            var t1 = new TableLayoutPanel { AutoSize = true, ColumnCount = 5, CellBorderStyle = TableLayoutPanelCellBorderStyle.None, Margin = new Padding(0, 0, 0, 10) };
            t1.Controls.Add(new Label { Text = "Claro (Thin)", BackColor = System.Drawing.Color.LightGray, AutoSize = true, Font = f8 }, 0, 0);
            t1.Controls.Add(new Label { Text = "Aumentar []", AutoSize = true, Font = f8 }, 1, 0);
            t1.Controls.Add(new Label { Text = "Duller", AutoSize = true, Font = f8 }, 2, 0);
            t1.Controls.Add(new Label { Text = "Brighter", AutoSize = true, Font = f8 }, 3, 0);
            t1.Controls.Add(new Label { Text = "100", BackColor = System.Drawing.Color.LightGray, AutoSize = true, Font = f8 }, 4, 0);

            t1.Controls.Add(new Label { Text = "Oscuro (Full)", BackColor = System.Drawing.Color.Black, ForeColor = System.Drawing.Color.White, AutoSize = true, Font = f8 }, 0, 1);
            t1.Controls.Add(new Label { Text = "Disminuir []", AutoSize = true, Font = f8 }, 1, 1);
            t1.Controls.Add(new Label { Text = "Brighter", AutoSize = true, Font = f8 }, 2, 1);
            t1.Controls.Add(new Label { Text = "Duller", AutoSize = true, Font = f8 }, 3, 1);
            t1.Controls.Add(new Label { Text = "0", BackColor = System.Drawing.Color.Black, ForeColor = System.Drawing.Color.White, AutoSize = true, Font = f8 }, 4, 1);
            pnl.Controls.Add(t1);

            var t2 = new TableLayoutPanel { AutoSize = true, ColumnCount = 3, Margin = new Padding(0, 0, 0, 10) };
            t2.Controls.Add(new Label { Text = "Amarillo", ForeColor = System.Drawing.Color.Goldenrod, Font = f8b, AutoSize = true }, 0, 0);
            t2.Controls.Add(new Label { Text = "Aumentar Amarillo", Font = f8, AutoSize = true }, 1, 0);
            t2.Controls.Add(new Label { Text = "Disminuir Amarillo", Font = f8, AutoSize = true }, 2, 0);

            t2.Controls.Add(new Label { Text = "Azul", ForeColor = System.Drawing.Color.Blue, Font = f8b, AutoSize = true }, 0, 1);
            t2.Controls.Add(new Label { Text = "Aumentar Azul", Font = f8, AutoSize = true }, 1, 1);
            t2.Controls.Add(new Label { Text = "Disminuir Azul", Font = f8, AutoSize = true }, 2, 1);

            t2.Controls.Add(new Label { Text = "Verde", ForeColor = System.Drawing.Color.Green, Font = f8b, AutoSize = true }, 0, 2);
            t2.Controls.Add(new Label { Text = "Aumentar Verde", Font = f8, AutoSize = true }, 1, 2);
            t2.Controls.Add(new Label { Text = "Disminuir Verde", Font = f8, AutoSize = true }, 2, 2);

            t2.Controls.Add(new Label { Text = "Rojo", ForeColor = System.Drawing.Color.Red, Font = f8b, AutoSize = true }, 0, 3);
            t2.Controls.Add(new Label { Text = "Aumentar Rojo", Font = f8, AutoSize = true }, 1, 3);
            t2.Controls.Add(new Label { Text = "Disminuir Rojo", Font = f8, AutoSize = true }, 2, 3);
            pnl.Controls.Add(t2);

            var t3 = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single };
            t3.Controls.Add(new Label { Text = "Yellower (Greener)", Font = f8, AutoSize = true }, 0, 0); t3.Controls.Add(new Label { Text = "Bluer (Redder)", Font = f8, AutoSize = true }, 1, 0);
            t3.Controls.Add(new Label { Text = "Yellower (Redder)", Font = f8, AutoSize = true }, 0, 1); t3.Controls.Add(new Label { Text = "Bluer (Greener)", Font = f8, AutoSize = true }, 1, 1);
            t3.Controls.Add(new Label { Text = "Greener (Bluer)", Font = f8, AutoSize = true }, 0, 2); t3.Controls.Add(new Label { Text = "Redder (Yellower)", Font = f8, AutoSize = true }, 1, 2);
            t3.Controls.Add(new Label { Text = "Greener (Yellower)", Font = f8, AutoSize = true }, 0, 3); t3.Controls.Add(new Label { Text = "Redder (Bluer)", Font = f8, AutoSize = true }, 1, 3);
            pnl.Controls.Add(t3);

            var t4 = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, CellBorderStyle = TableLayoutPanelCellBorderStyle.Single, Margin = new Padding(0, 5, 0, 0) };
            t4.Controls.Add(new Label { Text = "Bluer (Redder)", Font = f8, AutoSize = true }, 0, 0); t4.Controls.Add(new Label { Text = "Yellower (Greener)", Font = f8, AutoSize = true }, 1, 0);
            t4.Controls.Add(new Label { Text = "Bluer (Greener)", Font = f8, AutoSize = true }, 0, 1); t4.Controls.Add(new Label { Text = "Yellower (Redder)", Font = f8, AutoSize = true }, 1, 1);
            t4.Controls.Add(new Label { Text = "Redder (Yellower)", Font = f8, AutoSize = true }, 0, 2); t4.Controls.Add(new Label { Text = "Greener (Bluer)", Font = f8, AutoSize = true }, 1, 2);
            t4.Controls.Add(new Label { Text = "Redder (Bluer)", Font = f8, AutoSize = true }, 0, 3); t4.Controls.Add(new Label { Text = "Greener (Yellower)", Font = f8, AutoSize = true }, 1, 3);
            pnl.Controls.Add(t4);

            return pnl;
        }

        private void PopulateFromReport(OcrReport report)
        {
            if (report == null) return;
            var shadeData = new ShadeExtractionResult { ShadeName = report.Batch?.ShadeName, Recipe = report.Recipe };
            var results = EngineCalc.CalculateAllIlluminants(report);
            PopulateFromObjects(shadeData, results);
        }

        private void LoadCorrectiveRecipeGrid(EngineRes res)
        {
            if (dgvCorrectiveRecipe == null || res == null || _shadeData == null || _shadeData.Recipe == null) return;

            // 1. CONFIGURACIÓN MORFOLÓGICA HORIZONTAL RIGIDA
            dgvCorrectiveRecipe.Rows.Clear();
            dgvCorrectiveRecipe.Columns.Clear();

            dgvCorrectiveRecipe.Columns.Add("colColorante", " Colorante");
            dgvCorrectiveRecipe.Columns.Add("colR1_Con", "R1 Con. %"); dgvCorrectiveRecipe.Columns.Add("colR1_Part", "R1 Part. %"); dgvCorrectiveRecipe.Columns.Add("colR1_Var", "R1 %");
            dgvCorrectiveRecipe.Columns.Add("colR2_Con", "R2 Con. %"); dgvCorrectiveRecipe.Columns.Add("colR2_Part", "R2 Part. %"); dgvCorrectiveRecipe.Columns.Add("colR2_Var", "R2 %");
            dgvCorrectiveRecipe.Columns.Add("colR3_Con", "R3 Con. %"); dgvCorrectiveRecipe.Columns.Add("colR3_Part", "R3 Part. %"); dgvCorrectiveRecipe.Columns.Add("colR3_Var", "R3 %");

            string[] colsVar = { "colR1_Var", "colR2_Var", "colR3_Var" };
            foreach (var c in new[] { "colR1_Con", "colR2_Con", "colR3_Con" }) dgvCorrectiveRecipe.Columns[c].DefaultCellStyle.Format = "N5";
            foreach (var p in new[] { "colR1_Part", "colR2_Part", "colR3_Part" }) dgvCorrectiveRecipe.Columns[p].DefaultCellStyle.Format = "P1";

            foreach (var v in colsVar)
            {
                dgvCorrectiveRecipe.Columns[v].DefaultCellStyle.Format = "0%;-0%;0%";
            }
            // Extraer lista de concentraciones iniciales y ejecutar el nuevo motor limpio
            List<double> conOriginales = _shadeData.Recipe.Select(x => ParsePercentageValue(x.Percentage)).ToList();
            Color.ColorimetricCalculator.CalcularNuevasRecetasMaestras(res, conOriginales);

            double totalOriginalBase = conOriginales.Sum();
            double totalR1 = res.RecetaR1_Luminosidad.Sum();
            double totalR2 = res.RecetaR2_Croma.Sum();
            double totalR3 = res.RecetaR3_Tono.Sum();

            double sumaVariacionR1 = 0.0;
            double sumaVariacionR2 = 0.0;
            double sumaVariacionR3 = 0.0;

            // 2. INYECCIÓN FILA POR FILA (COLORANTES)
            for (int i = 0; i < _shadeData.Recipe.Count; i++)
            {
                int idx = dgvCorrectiveRecipe.Rows.Add();
                double orig = conOriginales[i];
                double r1 = res.RecetaR1_Luminosidad[i];
                double r2 = res.RecetaR2_Croma[i];
                double r3 = res.RecetaR3_Tono[i];

                dgvCorrectiveRecipe.Rows[idx].Cells["colColorante"].Value = _shadeData.Recipe[i].Name;

                double varR1 = Math.Abs((orig > 0) ? ((r1 / orig) - 1.0) : 0.0);
                double varR2 = Math.Abs((orig > 0) ? ((r2 / orig) - 1.0) : 0.0);
                double varR3 = Math.Abs((orig > 0) ? ((r3 / orig) - 1.0) : 0.0);

                sumaVariacionR1 += varR1;
                sumaVariacionR2 += varR2;
                sumaVariacionR3 += varR3;

                // R1
                dgvCorrectiveRecipe.Rows[idx].Cells["colR1_Con"].Value = r1;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR1_Part"].Value = totalR1 > 0 ? r1 / totalR1 : 0;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR1_Var"].Value = varR1;

                // R2
                dgvCorrectiveRecipe.Rows[idx].Cells["colR2_Con"].Value = r2;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR2_Part"].Value = totalR2 > 0 ? r2 / totalR2 : 0;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR2_Var"].Value = varR2;

                // R3
                dgvCorrectiveRecipe.Rows[idx].Cells["colR3_Con"].Value = r3;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR3_Part"].Value = totalR3 > 0 ? r3 / totalR3 : 0;
                dgvCorrectiveRecipe.Rows[idx].Cells["colR3_Var"].Value = varR3;

                foreach (var col in colsVar)
                {
                    dgvCorrectiveRecipe.Rows[idx].Cells[col].Style.ForeColor = System.Drawing.Color.Black;
                    dgvCorrectiveRecipe.Rows[idx].Cells[col].Style.Font = new System.Drawing.Font(dgvCorrectiveRecipe.Font, System.Drawing.FontStyle.Bold);
                }
            }

            // 3. FILA TOTAL CON VARIACIÓN NETAS GLOBAL ALGEBRAICA (=+D57/D11-1)
            int totalIdx = dgvCorrectiveRecipe.Rows.Add();
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colColorante"].Value = "TOTAL";

            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR1_Con"].Value = totalR1;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR1_Part"].Value = 1.00;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR1_Var"].Value = Math.Abs(totalOriginalBase > 0 ? (totalR1 / totalOriginalBase) - 1.0 : 0);

            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR2_Con"].Value = totalR2;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR2_Part"].Value = 1.00;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR2_Var"].Value = Math.Abs(totalOriginalBase > 0 ? (totalR2 / totalOriginalBase) - 1.0 : 0);

            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR3_Con"].Value = totalR3;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR3_Part"].Value = 1.00;
            dgvCorrectiveRecipe.Rows[totalIdx].Cells["colR3_Var"].Value = Math.Abs(totalOriginalBase > 0 ? (totalR3 / totalOriginalBase) - 1.0 : 0);

            dgvCorrectiveRecipe.Rows[totalIdx].DefaultCellStyle.Font = new System.Drawing.Font(dgvCorrectiveRecipe.Font, System.Drawing.FontStyle.Bold);
            dgvCorrectiveRecipe.Rows[totalIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            foreach (var col in colsVar) dgvCorrectiveRecipe.Rows[totalIdx].Cells[col].Style.ForeColor = System.Drawing.Color.Black;
        }

        private void UpdateChart(EngineRes res)
        {
            if (res == null || _cielabChart == null) return;
            _cielabChart.DeltaL = res.DeltaL; _cielabChart.DeltaA = res.DeltaA; _cielabChart.DeltaB = res.DeltaB;
            _cielabChart.DeltaE = res.DeltaE; _cielabChart.AbsoluteL = res.StdL; _cielabChart.AbsoluteA = res.StdA; _cielabChart.AbsoluteB = res.StdB;
            _cielabChart.LotL = res.LotL; _cielabChart.LotA = res.LotA; _cielabChart.LotB = res.LotB;
            _cielabChart.Invalidate();
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                // Seleccionar la fuente de resultados disponible
                var todosLosResultados = _resultsLegacy ?? new List<EngineRes>();

                if (todosLosResultados.Count == 0 && _lastMainResult != null)
                    todosLosResultados = new List<EngineRes> { _lastMainResult };

                if (todosLosResultados.Count == 0)
                {
                    MessageBox.Show("No hay resultados de análisis para guardar.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string shadeName = _shadeData?.ShadeName ?? "N/A";
                DateTime fechaActual = DateTime.Now;

                int registrosGuardados = 0;

                foreach (var res in todosLosResultados)
                {
                    // Calcular recetas si aún no se han calculado
                    List<double> conOriginales = _shadeData?.Recipe?
                        .Select(x => ParsePercentageValue(x.Percentage))
                        .ToList() ?? new List<double>();

                    if (conOriginales.Count > 0 && (res.RecetaR1_Luminosidad == null || res.RecetaR1_Luminosidad.Count == 0))
                        EngineCalc.CalcularNuevasRecetasMaestras(res, conOriginales);

                    // Obtener nombres de recetas R1, R2, R3 (primeros 3 colorantes o lo que haya)
                    string r1Name = conOriginales.Count > 0 && _shadeData?.Recipe?.Count > 0
                        ? _shadeData.Recipe[0].Name : "---";
                    string r2Name = conOriginales.Count > 1 && _shadeData?.Recipe?.Count > 1
                        ? _shadeData.Recipe[1].Name : "---";
                    string r3Name = conOriginales.Count > 2 && _shadeData?.Recipe?.Count > 2
                        ? _shadeData.Recipe[2].Name : "---";

                    // Diagnósticos del motor
                    string diagL = res.DiagnosticoL ?? "";
                    string diagC = res.DiagnosisC ?? "";
                    string diagH = res.DiagnosisH ?? "";

                    // Impacto = porcentaje de factor (como texto)
                    string impL = ((double)res.FactorL * 100.0).ToString("F2", System.Globalization.CultureInfo.InvariantCulture) + "%";
                    string impC = ((double)res.FactorC * 100.0).ToString("F2", System.Globalization.CultureInfo.InvariantCulture) + "%";
                    string impH = ((double)res.FactorH * 100.0).ToString("F2", System.Globalization.CultureInfo.InvariantCulture) + "%";

                    // Recomendaciones (usar el motor de diagnóstico)
                    string recL = EngineCalc.GetLuminosityDiagnosis(res.DeltaL);
                    string recC = res.RecommendationC ?? "";
                    string recH = EngineCalc.GetEngineeringDiagnosis("H", res.DeltaHue, "");

                    // Si hay ingredientes, guardar uno por colorante
                    if (_shadeData?.Recipe != null && _shadeData.Recipe.Count > 0)
                    {
                        foreach (var ing in _shadeData.Recipe)
                        {
                            decimal concOrig = (decimal)ParsePercentageValue(ing.Percentage);

                            Color.Services.HistorialService.GuardarRegistroMaestro(
                                shadeName:      shadeName,
                                fecha:          fechaActual,
                                iluminante:     res.Illuminant,
                                dyeName:        ing.Name ?? ing.Code ?? "N/A",
                                concOriginal:   concOrig,
                                r1:             r1Name,
                                r2:             r2Name,
                                r3:             r3Name,
                                impL:           impL,  diagL: diagL, recL: recL,
                                impC:           impC,  diagC: diagC, recC: recC,
                                impH:           impH,  diagH: diagH, recH: recH,
                                factorA:        ((double)res.FactorA).ToString("F5", System.Globalization.CultureInfo.InvariantCulture),
                                factorB:        ((double)res.FactorB).ToString("F5", System.Globalization.CultureInfo.InvariantCulture),
                                deltaE:         res.DeltaE.ToString("F4", System.Globalization.CultureInfo.InvariantCulture)
                            );
                            registrosGuardados++;
                        }
                    }
                    else
                    {
                        // Sin receta: guardar un registro resumen por iluminante
                        Color.Services.HistorialService.GuardarRegistroMaestro(
                            shadeName:      shadeName,
                            fecha:          fechaActual,
                            iluminante:     res.Illuminant,
                            dyeName:        "N/A",
                            concOriginal:   0m,
                            r1:             "---", r2: "---", r3: "---",
                            impL:           impL,  diagL: diagL, recL: recL,
                            impC:           impC,  diagC: diagC, recC: recC,
                            impH:           impH,  diagH: diagH, recH: recH,
                            factorA:        ((double)res.FactorA).ToString("F5", System.Globalization.CultureInfo.InvariantCulture),
                            factorB:        ((double)res.FactorB).ToString("F5", System.Globalization.CultureInfo.InvariantCulture),
                            deltaE:         res.DeltaE.ToString("F4", System.Globalization.CultureInfo.InvariantCulture)
                        );
                        registrosGuardados++;
                    }
                }

                MessageBox.Show(
                    $"Análisis guardado exitosamente.\n{registrosGuardados} registro(s) escritos en la base de datos.",
                    "Guardado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                btnGuardar.Enabled = false;
                btnGuardar.Text = "✔ Guardado";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error al guardar en la base de datos:\n" + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private double ParsePercentageValue(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return 0;
            string clean = val.Replace("%", "").Trim().Replace(",", ".");
            double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double res);
            return res;
        }
    }
}