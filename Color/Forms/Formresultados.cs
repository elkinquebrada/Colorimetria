using System.IO;
using Color.Services;
using OCR;
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
using EngineRow = Color.ColorimetricRow;

namespace Color
{
    public class FormResultados : Form
    {
        // ======= Datos de entrada =======
        private readonly OcrReport _report;
        private readonly string _resumenLegacy;
        private readonly List<EngineRes> _resultsLegacy;
        private List<CorrectiveRecipeResult> _recipeResults;
        private ShadeExtractionResult _shadeData;

        // ======= Controles de la vista (Tablas) =======
        private DataGridView dgvShadeHistory;
        private DataGridView dgvAnalysisLeft;
        private DataGridView dgvAnalysisLeftTL84;
        private DataGridView dgvAnalysisLeftA;
        private DataGridView dgvComparisonSummary;
        private DataGridView dgvCielabSummary;
        private DataGridView dgvAnalysisRight;
        private DataGridView dgvAnalysisRightTL84;
        private DataGridView dgvAnalysisRightA;
        private DataGridView dgvCorrectiveRecipe;
        private Label lblAlertCorrective;
        private Label lblRightShadeValue;


        private SplitContainer splitMedicionesCmc;
        private Button btnGuardar;
        private Button btnCerrar;
        private Button btnRegresar;

        private Button btnVerGrafico;
        private CielabChartControl _cielabChart;
        private EngineRes _lastMainResult;
        public object FormOcrOrigen { get; set; }

        // ======= Tolerancias (L*, Hue y ΔE) =======
        private double DL_MAX => Properties.Settings.Default.ToleranciaDL;
        private double DC_MAX => Properties.Settings.Default.ToleranciaDC;
        private double DH_MAX => Properties.Settings.Default.ToleranciaDH;
        private double DE_MAX => Properties.Settings.Default.ToleranciaDE;

        // ======= Constructores =======
        public FormResultados(OcrReport report)
        {
            _report = report ?? new OcrReport();
            _resultsLegacy = new List<EngineRes>();
            InitializeComponents();

            // Lógica silenciosa: Poblar desde el objeto Report directamente
            PopulateFromReport(_report);
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
                    Width = 55,
                    Height = 55,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    BackColor = System.Drawing.Color.Transparent
                };

                // Encontrar el label de título para ponerlo encima
                Control titleControl = this.Controls.Cast<Control>().FirstOrDefault(c => c is Label && c.Height == 50);
                if (titleControl != null)
                {
                    logo.Location = new Point(titleControl.Width - logo.Width - 15, 0);
                    titleControl.Controls.Add(logo);
                }
                else
                {
                    logo.Location = new Point(this.Width - logo.Width - 20, 5);
                    this.Controls.Add(logo);
                }
                logo.BringToFront();
            }
            catch { }
        }

        public FormResultados(string resumen, List<EngineRes> results, List<CorrectiveRecipeResult> recipeResults = null, ShadeExtractionResult shadeData = null)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = results ?? new List<EngineRes>();
            _recipeResults = recipeResults;
            _shadeData = shadeData;
            InitializeComponents();

            // Lógica silenciosa: Poblar desde los objetos ya calculados
            PopulateFromObjects(_shadeData, _resultsLegacy);
            AddBrandingLogo();
        }

        private void InitializeComponents()
        {
            this.Text = "TINT COATS CADENA";
            this.Size = new Size(1100, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = System.Drawing.Color.White;

            var lblTitulo = new Label
            {
                Text = "ANALISIS DE COLORIMETRIA",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.FromArgb(0, 102, 204),
                Dock = DockStyle.Top,
                Height = 50,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(15, 0, 0, 0)
            };

            // ---- Botones ----
            btnGuardar = CreateStyledButton(" Guardar", System.Drawing.Color.FromArgb(45, 126, 247));
            btnGuardar.Click += BtnGuardar_Click;

            btnRegresar = CreateStyledButton("← Regresar", System.Drawing.Color.FromArgb(180, 100, 30));
            btnRegresar.Click += (s, e) => {
                this.DialogResult = DialogResult.Retry;
                this.Close();
            };

            btnCerrar = CreateStyledButton("Finalizar", System.Drawing.Color.FromArgb(90, 90, 90));
            btnCerrar.Click += (s, e) => this.Close();

            // ---- Título ----
            lblTitulo.Visible = true;
            lblTitulo.Height = 50;

            _cielabChart = new CielabChartControl
            {
                Dock = DockStyle.Fill,
                Mode = CielabChartControl.ViewMode.Relative,
                Title = "", // El título lo pondremos en un label externo para el estilo solicitado
                BackColor = System.Drawing.Color.White
            };

            btnVerGrafico = new Button
            {
                Text = "🔍 Ver Gráfico",
                Size = new Size(130, 34),
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnVerGrafico.Click += (s, e) => {
                if (_lastMainResult == null) return;
                var frm = new FormDetalleCielab(_lastMainResult.DeltaL, _lastMainResult.DeltaA, _lastMainResult.DeltaB, _lastMainResult.DeltaE, _lastMainResult.CmcValue, 1.20, "",
                    _lastMainResult.StdL, _lastMainResult.StdA, _lastMainResult.StdB);
                frm.Show();
            };

            // ---- Grids ----
            dgvShadeHistory = CreateStyledGrid();
            dgvShadeHistory.ColumnHeadersVisible = false;
            dgvShadeHistory.ColumnCount = 5;
            dgvShadeHistory.Columns[0].Name = "Dye Code";
            dgvShadeHistory.Columns[1].Name = "Dye Names";
            dgvShadeHistory.Columns[2].Name = "Concentration";
            dgvShadeHistory.Columns[3].Name = "% Participación";
            dgvShadeHistory.Columns[4].Name = "% Ajuste dL";
            dgvShadeHistory.Columns[4].Visible = false;
            dgvShadeHistory.Columns[0].FillWeight = 15;
            dgvShadeHistory.Columns[1].FillWeight = 50;
            dgvShadeHistory.Columns[2].FillWeight = 15;
            dgvShadeHistory.Columns[3].FillWeight = 20;
            dgvShadeHistory.Columns[4].FillWeight = 15;

            dgvAnalysisLeftTL84 = CreateAnalysisGrid();
            dgvAnalysisLeftA = CreateAnalysisGrid();

            // TABLA 1: Ajuste de la Mezcla (Ejes dl, da, db)
            dgvCielabSummary = CreateStyledGrid();
            dgvCielabSummary.ColumnHeadersVisible = true;
            dgvCielabSummary.ColumnCount = 5;
            dgvCielabSummary.Columns[0].Name = "Eje";
            dgvCielabSummary.Columns[1].Name = "Variacion";
            dgvCielabSummary.Columns[2].Name = "Impacto";
            dgvCielabSummary.Columns[3].Name = "Diagnostico";
            dgvCielabSummary.Columns[4].Name = "Ajuste";

            dgvCielabSummary.Columns[0].HeaderText = "EJE";
            dgvCielabSummary.Columns[1].HeaderText = "Variacion (Δ)";
            dgvCielabSummary.Columns[2].HeaderText = "Impacto";
            dgvCielabSummary.Columns[3].HeaderText = "Diagnostico";
            dgvCielabSummary.Columns[4].HeaderText = "AJUSTE (%)";

            foreach (DataGridViewColumn col in dgvCielabSummary.Columns) col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvCielabSummary.Columns[0].FillWeight = 40f;
            dgvCielabSummary.Columns[1].FillWeight = 60f;
            dgvCielabSummary.Columns[2].FillWeight = 110f;
            dgvCielabSummary.Columns[3].FillWeight = 110f;
            dgvCielabSummary.Columns[4].FillWeight = 60f;

            dgvCielabSummary.EnableHeadersVisualStyles = false;
            dgvCielabSummary.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvCielabSummary.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvCielabSummary.ColumnHeadersDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvCielabSummary.ColumnHeadersDefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
            dgvCielabSummary.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            foreach (DataGridViewColumn col in dgvCielabSummary.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.HeaderCell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                col.HeaderCell.Style.ForeColor = System.Drawing.Color.White;
                col.HeaderCell.Style.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                col.HeaderCell.Style.SelectionForeColor = System.Drawing.Color.White;
                col.HeaderCell.Style.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            // TABLA 2: Revisión de Apariencia (Ejes dL, dC, dH)
            dgvAnalysisLeft = CreateStyledGrid();
            dgvAnalysisLeft.ColumnHeadersVisible = true;
            dgvAnalysisLeft.ColumnCount = 5;
            dgvAnalysisLeft.Columns[0].Name = "Eje";
            dgvAnalysisLeft.Columns[1].Name = "Variacion";
            dgvAnalysisLeft.Columns[2].Name = "Impacto";
            dgvAnalysisLeft.Columns[3].Name = "Dianostico";
            dgvAnalysisLeft.Columns[4].Name = "Ajuste";

            dgvAnalysisLeft.Columns[0].HeaderText = "EJE";
            dgvAnalysisLeft.Columns[1].HeaderText = "Variacion (Δ)";
            dgvAnalysisLeft.Columns[2].HeaderText = "Impacto";
            dgvAnalysisLeft.Columns[3].HeaderText = "Diagnostico";
            dgvAnalysisLeft.Columns[4].HeaderText = "AJUSTE (%)";

            foreach (DataGridViewColumn col in dgvAnalysisLeft.Columns) col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvAnalysisLeft.Columns[0].FillWeight = 40f;
            dgvAnalysisLeft.Columns[1].FillWeight = 60f;
            dgvAnalysisLeft.Columns[2].FillWeight = 110f;
            dgvAnalysisLeft.Columns[3].FillWeight = 110f;
            dgvAnalysisLeft.Columns[4].FillWeight = 60f;

            dgvAnalysisLeft.EnableHeadersVisualStyles = false;
            dgvAnalysisLeft.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvAnalysisLeft.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvAnalysisLeft.ColumnHeadersDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvAnalysisLeft.ColumnHeadersDefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;
            dgvAnalysisLeft.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            foreach (DataGridViewColumn col in dgvAnalysisLeft.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                col.HeaderCell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                col.HeaderCell.Style.ForeColor = System.Drawing.Color.White;
                col.HeaderCell.Style.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                col.HeaderCell.Style.SelectionForeColor = System.Drawing.Color.White;
                col.HeaderCell.Style.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            }

            dgvComparisonSummary = CreateStyledGrid();
            dgvComparisonSummary.ColumnHeadersVisible = true;
            dgvComparisonSummary.ColumnCount = 4;
            dgvComparisonSummary.Columns[0].Name = "Facet";
            dgvComparisonSummary.Columns[1].Name = "Tolerance";
            dgvComparisonSummary.Columns[2].Name = "Illuminant";
            dgvComparisonSummary.Columns[3].Name = "Result";
            foreach (DataGridViewColumn col in dgvComparisonSummary.Columns) col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            dgvAnalysisRight = CreateAnalysisGrid();
            dgvAnalysisRightTL84 = CreateAnalysisGrid();
            dgvAnalysisRightA = CreateAnalysisGrid();

            // Estilo tenue para iluminantes secundarios (no compiten con D65)
            ApplyTenueGridStyle(dgvAnalysisLeftTL84);
            ApplyTenueGridStyle(dgvAnalysisLeftA);
            ApplyTenueGridStyle(dgvAnalysisRightTL84);
            ApplyTenueGridStyle(dgvAnalysisRightA);

            // ---- Layout ----
            splitMedicionesCmc = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = Math.Max(25, this.Width / 2),
                IsSplitterFixed = false
            };
            this.Resize += (s, e) => {
                try { if (this.Width > 100) splitMedicionesCmc.SplitterDistance = this.Width / 2; }
                catch { }
            };

            var pnlCorrective = new Panel { Dock = DockStyle.Bottom, Height = 150 };
            dgvCorrectiveRecipe = CreateCorrectiveGrid();
            lblAlertCorrective = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.Gray,
                Cursor = Cursors.Help
            };

            // ToolTip aclaratorio de Protocolo de Seguridad
            ToolTip ttProtocolo = new ToolTip { IsBalloon = true, ToolTipTitle = "Protocolo de Seguridad Industrial", AutoPopDelay = 15000 };
            string msgProtocolo = "Nota sobre Alertas de Medición y Ajuste:\n\n" +
                "1. Validación Multiluminante: El lote se evalúa bajo tres fuentes de luz (D65, TL84, A). El 'NO CUMPLE' se activa si existe riesgo de metamerismo en cualquier iluminante.\n\n" +
                "2. Límite de Estabilidad (Umbral 15%): Ajustes superiores al 15% indican correcciones drásticas que pueden comprometer la estabilidad química y repetibilidad del tono.";
            ttProtocolo.SetToolTip(lblAlertCorrective, msgProtocolo);

            var lblHeaderCorrective = CreateHeaderLabel(" FORMULACIÓN CORRECTIVA DE RECETA" +
                "");
            lblHeaderCorrective.Dock = DockStyle.Top;
            lblHeaderCorrective.Height = 28;

            pnlCorrective.Controls.Add(dgvCorrectiveRecipe);
            pnlCorrective.Controls.Add(lblAlertCorrective);
            pnlCorrective.Controls.Add(lblHeaderCorrective);

            // --- Panel Izquierdo (RECETA / HISTORY) ---
            var pnlLeftRecipe = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8 };
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Percent, 20));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Percent, 35));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlLeftRecipe.RowStyles.Add(new RowStyle(SizeType.Percent, 45));

            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("ANALISIS DE SHADE HISTORY REPORT (RECETA)"), 0, 0);
            pnlLeftRecipe.Controls.Add(dgvShadeHistory, 0, 1);
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("TABLA 1: Ajuste de la Mezcla (Ejes dl, da, db)"), 0, 2);
            pnlLeftRecipe.Controls.Add(dgvCielabSummary, 0, 3);
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("TABLA 2: Revisión de Apariencia (Ejes dL, dC, dH)"), 0, 4);
            pnlLeftRecipe.Controls.Add(dgvAnalysisLeft, 0, 5);

            var pnlCorrectiveContainer = new Panel { Dock = DockStyle.Fill };
            pnlCorrectiveContainer.Controls.Add(dgvCorrectiveRecipe);
            pnlCorrectiveContainer.Controls.Add(lblAlertCorrective);
            dgvCorrectiveRecipe.Dock = DockStyle.Fill;
            lblAlertCorrective.Dock = DockStyle.Bottom;
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("RESUMEN DE FORMULACIÓN CORRECTIVA (D65)"), 0, 6);
            pnlLeftRecipe.Controls.Add(pnlCorrectiveContainer, 0, 7);

            // --- Panel Derecho (LOTE / COMPARISON) ---
            var pnlRightLot = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8 };
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Percent, 22f));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Percent, 26f));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Percent, 26f));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            pnlRightLot.RowStyles.Add(new RowStyle(SizeType.Percent, 26f));

            pnlRightLot.Controls.Add(CreateHeaderLabel("ANALISIS DE SAMPLE COMPARISON (LOTE)"), 0, 0);

            var pnlComparison = new Panel { Dock = DockStyle.Fill };
            var pnlShade = new TableLayoutPanel { Dock = DockStyle.Top, Height = 22, ColumnCount = 2, RowCount = 1 };
            pnlShade.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            pnlShade.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var lblShadeTitle = CreateHeaderLabel("Shade Name");
            lblRightShadeValue = CreateHeaderLabel("N/A");
            pnlShade.Controls.Add(lblShadeTitle, 0, 0);
            pnlShade.Controls.Add(lblRightShadeValue, 1, 0);

            pnlComparison.Controls.Add(dgvComparisonSummary);
            pnlComparison.Controls.Add(pnlShade);
            dgvComparisonSummary.Dock = DockStyle.Fill;
            pnlRightLot.Controls.Add(pnlComparison, 0, 1);

            pnlRightLot.Controls.Add(CreateHeaderLabel("ANALISIS ILUMINANTE D65"), 0, 2);
            pnlRightLot.Controls.Add(dgvAnalysisRight, 0, 3);

            pnlRightLot.Controls.Add(CreateHeaderLabel("ANALISIS ILUMINANTE TL84", true), 0, 4);
            pnlRightLot.Controls.Add(dgvAnalysisRightTL84, 0, 5);

            pnlRightLot.Controls.Add(CreateHeaderLabel("ANALISIS ILUMINANTE A / CWF", true), 0, 6);
            pnlRightLot.Controls.Add(dgvAnalysisRightA, 0, 7);

            splitMedicionesCmc.Panel1.Controls.Add(pnlLeftRecipe);
            splitMedicionesCmc.Panel2.Controls.Add(pnlRightLot);

            var pnlBottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = System.Drawing.Color.FromArgb(245, 245, 245),
                Padding = new Padding(10)
            };

            // Botones IZQUIERDA: Regresar y Finalizar
            btnRegresar.Location = new Point(15, 12);
            btnCerrar.Text = "Finalizar";
            btnCerrar.Location = new Point(btnRegresar.Right + 10, 12);

            // Helper para reposicionar botones DERECHA dinamicamente
            Action reposicionarDerecha = () => {
                btnVerGrafico.Left = pnlBottom.Width - btnVerGrafico.Width - 15;
                btnGuardar.Left = btnVerGrafico.Left - btnGuardar.Width - 10;
                btnVerGrafico.Top = 12;
                btnGuardar.Top = 12;
            };

            // Reposicionar cuando el panel cambia de tamaño (Maximize incluido)
            pnlBottom.Resize += (s, e) => reposicionarDerecha();

            pnlBottom.Controls.Add(btnRegresar);
            pnlBottom.Controls.Add(btnCerrar);
            pnlBottom.Controls.Add(btnGuardar);
            pnlBottom.Controls.Add(btnVerGrafico);

            this.Controls.Add(splitMedicionesCmc);
            this.Controls.Add(lblTitulo);
            this.Controls.Add(pnlBottom);
        }
        private Panel CreatePanelWithManyGrids(string h1, DataGridView g1, string h2, DataGridView g2, string h3, DataGridView g3, string h4, DataGridView g4)
        {
            var pnl = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8 };
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 135));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33));

            pnl.Controls.Add(CreateHeaderLabel(h1), 0, 0);
            pnl.Controls.Add(g1, 0, 1);

            g1.Columns[0].FillWeight = 20;
            g1.Columns[1].FillWeight = 50;
            g1.Columns[2].FillWeight = 15;
            g1.Columns[3].FillWeight = 15;

            pnl.Controls.Add(CreateHeaderLabel(h2), 0, 2);
            pnl.Controls.Add(g2, 0, 3);
            pnl.Controls.Add(CreateHeaderLabel(h3, true), 0, 4);
            pnl.Controls.Add(g3, 0, 5);
            pnl.Controls.Add(CreateHeaderLabel(h4, true), 0, 6);
            pnl.Controls.Add(g4, 0, 7);

            return pnl;
        }

        private Panel CreatePanelWithGrids(string head1, DataGridView g1, string head2, DataGridView g2)
        {
            var pnl = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 4 };
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 55));

            pnl.Controls.Add(CreateHeaderLabel(head1), 0, 0);
            pnl.Controls.Add(g1, 0, 1);
            pnl.Controls.Add(CreateHeaderLabel(head2), 0, 2);
            pnl.Controls.Add(g2, 0, 3);
            return pnl;
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
                Font = new Font("Segoe UI", 8.2f),
                ScrollBars = ScrollBars.Vertical,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgv.ColumnHeadersDefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;

            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.White;
            dgv.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black;

            return dgv;
        }

        private DataGridView CreateCorrectiveGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "Colorante"; dgv.Columns[0].FillWeight = 25;
            dgv.Columns[1].Name = "Receta Original"; dgv.Columns[1].FillWeight = 15;
            dgv.Columns[1].Visible = false; // Ocultar 
            dgv.Columns[2].Name = "Receta # 1"; dgv.Columns[2].FillWeight = 15;
            dgv.Columns[3].Name = "Receta # 2"; dgv.Columns[3].FillWeight = 15;
            dgv.Columns[4].Name = "Receta # 3"; dgv.Columns[4].FillWeight = 15;
            dgv.Columns[5].Name = "Participación"; dgv.Columns[5].FillWeight = 15;
            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "EJE"; dgv.Columns[0].FillWeight = 10;
            dgv.Columns[1].Name = "Δ%"; dgv.Columns[1].FillWeight = 10;
            dgv.Columns[1].HeaderCell.ToolTipText = "(Std - Lot)";
            dgv.Columns[2].Name = "VARIACION"; dgv.Columns[2].FillWeight = 12;
            dgv.Columns[3].Name = "IMPACTO"; dgv.Columns[3].FillWeight = 18;
            dgv.Columns[4].Name = "DIAGNOSTICO"; dgv.Columns[4].FillWeight = 25;
            dgv.Columns[5].Name = "RECOMENDACION"; dgv.Columns[5].FillWeight = 25;
            return dgv;
        }

        private void ApplyTranslucentStyle(DataGridView dgv)
        {
            var faintColor = System.Drawing.Color.FromArgb(160, 170, 180);
            dgv.DefaultCellStyle.ForeColor = faintColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = faintColor;
            dgv.GridColor = System.Drawing.Color.FromArgb(245, 245, 245);
            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.White;
            dgv.DefaultCellStyle.SelectionForeColor = faintColor;

            dgv.CellMouseEnter += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                    dgv.Rows[e.RowIndex].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                    dgv.Rows[e.RowIndex].Cells[0].Style.ForeColor = System.Drawing.Color.White;
                }
            };
            dgv.CellMouseLeave += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                    dgv.Rows[e.RowIndex].DefaultCellStyle.ForeColor = faintColor;
                    dgv.Rows[e.RowIndex].Cells[0].Style.ForeColor = faintColor;
                }
            };
        }

        private Button CreateStyledButton(string text, System.Drawing.Color color)
        {
            return new Button
            {
                Text = text,
                Size = new Size(130, 35),
                BackColor = color,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
        }

        private Label CreateHeaderLabel(string text, bool tenue = false)
        {
            // Colores corporativos para D65, colores muy suaves para secundarios
            var backColor = tenue ? System.Drawing.Color.FromArgb(210, 210, 215) : System.Drawing.Color.FromArgb(0, 102, 204);
            var foreColor = tenue ? System.Drawing.Color.FromArgb(120, 120, 120) : System.Drawing.Color.White;
            return new Label
            {
                Text = " " + text,
                BackColor = backColor,
                ForeColor = foreColor,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        /// Aplica estilo tenue (gris claro) a las grillas de iluminantes secundarios
        /// para que no compitan visualmente con el iluminante principal D65
        private void ApplyTenueGridStyle(DataGridView dgv)
        {
            var lightGray = System.Drawing.Color.FromArgb(200, 200, 200); // Texto muy claro
            var bgGray = System.Drawing.Color.FromArgb(248, 248, 248);

            // 1. Estilo General
            dgv.DefaultCellStyle.BackColor = bgGray;
            dgv.DefaultCellStyle.ForeColor = lightGray;
            dgv.DefaultCellStyle.SelectionBackColor = bgGray;
            dgv.DefaultCellStyle.SelectionForeColor = lightGray;
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);

            // 2. Estilo de Filas (Forzado)
            dgv.RowsDefaultCellStyle.BackColor = bgGray;
            dgv.RowsDefaultCellStyle.ForeColor = lightGray;
            dgv.RowsDefaultCellStyle.SelectionBackColor = bgGray;
            dgv.RowsDefaultCellStyle.SelectionForeColor = lightGray;

            // 3. Filas Alternas
            dgv.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(242, 242, 242);
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = lightGray;
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(242, 242, 242);

            // 4. Cabeceras
            dgv.ColumnHeadersDefaultCellStyle.BackColor = bgGray;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = lightGray;
            dgv.ColumnHeadersDefaultCellStyle.SelectionBackColor = bgGray;
            dgv.EnableHeadersVisualStyles = false;

            // 5. Bordes
            dgv.GridColor = System.Drawing.Color.FromArgb(235, 235, 235);

            // 6. Desactivar resaltado visual de selección estándar
            dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;

            // 7. INTERACCIÓN PROFESIONAL: Efecto Hover (Revelar datos al pasar el mouse)
            dgv.CellMouseEnter += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var row = dgv.Rows[e.RowIndex];
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204); // Azul Coats
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                        cell.Style.ForeColor = System.Drawing.Color.White;
                    }
                }
            };

            dgv.CellMouseLeave += (s, e) => {
                if (e.RowIndex >= 0)
                {
                    var row = dgv.Rows[e.RowIndex];
                    var originalBg = (e.RowIndex % 2 == 0) ? bgGray : System.Drawing.Color.FromArgb(242, 242, 242);
                    row.DefaultCellStyle.BackColor = originalBg;
                    row.DefaultCellStyle.ForeColor = lightGray;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = originalBg;
                        cell.Style.ForeColor = lightGray;
                    }
                }
            };
        }

        private void HighlightChecks(DataGridView dgv, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgv.Rows.Count) return;
            var row = dgv.Rows[rowIndex];
            foreach (DataGridViewCell cell in row.Cells)
            {
                string val = cell.Value?.ToString();
                if (val != null && (val.Contains("✔") || val.Contains("✔") || val.Contains("✔")))
                {
                    cell.Style.ForeColor = System.Drawing.Color.ForestGreen;
                    cell.Style.Font = new Font(dgv.Font, FontStyle.Bold);
                    cell.Style.SelectionForeColor = System.Drawing.Color.ForestGreen;
                }
            }
        }

        /// Aplica estilo tenue a todas las celdas de una fila específica. ✓
        /// Garantiza que ninguna celda herede colores vivos de otro estilo.
        private void ApplyTenueRowStyle(DataGridView dgv, int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dgv.Rows.Count) return;
            // Detectar por GridColor o SelectionMode si es un grid tenue
            if (dgv.SelectionMode != DataGridViewSelectionMode.CellSelect) return;

            var row = dgv.Rows[rowIndex];
            var lightGray = System.Drawing.Color.FromArgb(200, 200, 200);
            var bgGray = (rowIndex % 2 == 0) ? System.Drawing.Color.FromArgb(248, 248, 248) : System.Drawing.Color.FromArgb(242, 242, 242);

            foreach (DataGridViewCell cell in row.Cells)
            {
                cell.Style.BackColor = bgGray;
                cell.Style.ForeColor = lightGray;
                cell.Style.SelectionBackColor = bgGray;
                cell.Style.SelectionForeColor = lightGray;
                cell.Style.Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);
            }
        }

        private double ParsePercentageValue(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return 0;
            string clean = val.Replace("%", "").Trim().Replace(",", ".");
            if (double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double res))
                return res;
            return 0;
        }

        private void PopulateFromObjects(ShadeExtractionResult shadeData, List<EngineRes> results)
        {
            if (shadeData != null)
            {
                lblRightShadeValue.Text = shadeData.ShadeName ?? "N/A";
                dgvShadeHistory.Rows.Clear();
                int idxShade = dgvShadeHistory.Rows.Add("Shade Name", shadeData.ShadeName ?? "N/A", "", "", "");
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

                int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration", "% Participación", "% Ajuste dL");
                dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
                dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                if (shadeData.Recipe != null)
                {
                    double totalConc = 0;
                    foreach (var ing in shadeData.Recipe)
                    {
                        totalConc += ParsePercentageValue(ing.Percentage);
                    }

                    var d65 = results?.FirstOrDefault(r => r.Illuminant.Contains("D65")) ?? results?.FirstOrDefault();
                    double percentL = d65 != null ? d65.PercentL : 0;
                    string percentLStr = percentL > 0 ? "+" + percentL.ToString("F2") + "%" : percentL.ToString("F2") + "%";

                    foreach (var ing in shadeData.Recipe)
                    {
                        double pctVal = ParsePercentageValue(ing.Percentage);
                        double part = totalConc > 0 ? (pctVal / totalConc) * 100 : 0;
                        dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage, part.ToString("F1") + "%", percentLStr);
                    }

                    // Add Total row
                    int idxTotal = dgvShadeHistory.Rows.Add("Total", "", totalConc.ToString("F5") + "%", "100.0%", "");
                    dgvShadeHistory.Rows[idxTotal].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                    dgvShadeHistory.Rows[idxTotal].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
                }
            }

            if (results != null && results.Count > 0)
            {
                // Buscamos D65 como prioritario, si no el primero que haya
                var d65 = results.FirstOrDefault(r => r.Illuminant.Contains("D65")) ?? results[0];
                _lastMainResult = d65;

                // Identificar los demás iluminantes dinámicamente para llenar los 3 espacios
                var others = results.Where(r => r != d65).ToList();
                var ill2 = others.Count > 0 ? others[0] : null;
                var ill3 = others.Count > 1 ? others[1] : null;

                FillComparisonSummary(results, DE_MAX);

                Func<string, double> toDbl = s => {
                    if (string.IsNullOrEmpty(s)) return 0;
                    string clean = Regex.Replace(s, @"[^\d\.\-eE,]+", "").Replace(',', '.');
                    if (double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double v)) return v;
                    return 0;
                };

                dgvCielabSummary.Rows.Clear();

                // ERROR (Δ) values:
                double errL = Math.Abs((double)d65.FactorL);
                double errA = Math.Abs((double)d65.FactorA);
                double errB = Math.Abs((double)d65.FactorB);

                // Eje dl
                string lotL = d65.DeltaL < 0 ? " Claro" : " Oscuro";
                string actL = d65.DeltaL < 0 ? "Aumentar colorante" : "Disminuir colorante";
                string adjL = (d65.DeltaL < 0 ? "+ " : "- ") + Math.Abs(d65.PercentL).ToString("F2") + "%";
                if (Math.Abs(d65.PercentL) <= 0.1)
                {
                    lotL = "✔";
                    actL = "✔";
                    adjL = "✔";
                }
                int idxL = dgvCielabSummary.Rows.Add("dl", errL.ToString("F5"), lotL, actL, adjL);
                ApplyEjeStyle(dgvCielabSummary, idxL, "dl"); ApplyTenueRowStyle(dgvCielabSummary, idxL);

                // Eje da — DeltaA = Std - Lot
                string lotA = d65.DeltaA < 0 ? " Mas Rojo" : " Mas Verde";
                string actA = d65.DeltaA < 0 ? "Disminuir el Rojo" : "Aumentar el Rojo";

                double stdA = shadeData != null ? toDbl(shadeData.StdA) : 0;
                double stdB = shadeData != null ? toDbl(shadeData.StdB) : 0;
                double dA = d65.DeltaA;
                double dB = d65.DeltaB;
                double pctA = (Math.Abs(stdA) > 0.1) ? (dA / Math.Abs(stdA)) : 0;
                double pctB = (Math.Abs(stdB) > 0.1) ? (dB / Math.Abs(stdB)) : 0;

                // Signo directo desde el delta: negativo→Disminuir→"-"  positivo→Aumentar→"+"
                string adjA = (d65.DeltaA < 0 ? "- " : "+ ") + Math.Abs(pctA * 100).ToString("F2") + "%";
                if (Math.Abs(pctA * 100) <= 0.1)
                {
                    lotA = "✔";
                    actA = "✔";
                    adjA = "✔";
                }
                int idxA = dgvCielabSummary.Rows.Add("da", errA.ToString("F5"), lotA, actA, adjA);
                ApplyEjeStyle(dgvCielabSummary, idxA, "da"); ApplyTenueRowStyle(dgvCielabSummary, idxA);

                // Eje db — DeltaB = Std - Lot
                string lotB = d65.DeltaB < 0 ? " Mas Amarillo" : " Mas Azul";
                string actB = d65.DeltaB < 0 ? "Disminuir el Amarillo" : "Aumentar el Azul";
                string adjB = "- " + Math.Abs(pctB * 100).ToString("F2") + "%";
                if (Math.Abs(pctB * 100) <= 0.1)
                {
                    lotB = "✔";
                    actB = "✔";
                    adjB = "✔";
                }
                int idxB = dgvCielabSummary.Rows.Add("db", errB.ToString("F5"), lotB, actB, adjB);
                ApplyEjeStyle(dgvCielabSummary, idxB, "db"); ApplyTenueRowStyle(dgvCielabSummary, idxB);

                HighlightChecks(dgvCielabSummary, idxL);
                HighlightChecks(dgvCielabSummary, idxA);
                HighlightChecks(dgvCielabSummary, idxB);

                // --- TABLA IZQUIERDA Y DERECHA: Sincronización Total con el Motor ---
                // Forzamos que ambos paneles usen la misma fuente de datos (EngineRes)
                FillAnalysisGrid(dgvAnalysisLeft, d65, true);
                FillAnalysisGrid(dgvAnalysisRight, d65, false);

                if (ill2 != null) FillAnalysisGrid(dgvAnalysisRightTL84, ill2, false);
                if (ill3 != null) FillAnalysisGrid(dgvAnalysisRightA, ill3, false);

                // --- CALCULO DE RECETA CORRECTIVA (D65) ---
                if (shadeData != null)
                {
                    var ingredients = RecipeCorrector.IngredientsFromShade(shadeData);
                    var correctiveResult = RecipeCorrector.CalculateCorrectiveRecipe(ingredients, d65);
                    bool allComply = results != null && results.All(r => r.DeltaE <= DE_MAX);
                    FillCorrectiveRecipeGrid(correctiveResult, allComply);
                }

                // Actualizar gráfico con D65
                if (d65 != null) UpdateChart(d65);

                // Limpiar selección para evitar filas azules resaltadas al inicio
                dgvAnalysisRightTL84.ClearSelection();
                dgvAnalysisRightA.ClearSelection();
            }
        }

        private void FillAnalysisGridFromOcr(DataGridView dgv, ShadeExtractionResult shade, double? varL = null)
        {
            dgv.Rows.Clear();
            if (shade == null || shade.Batch == null) return;
            var batch = shade.Batch;

            Func<string, double> toDbl = s => {
                if (string.IsNullOrEmpty(s)) return 0;
                string clean = Regex.Replace(s, @"[^\d\.\-eE,]+", "").Replace(',', '.');
                if (double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double v)) return v;
                return 0;
            };

            double dL = toDbl(batch.DL);
            double dC = toDbl(batch.DC);
            double dH = toDbl(batch.DH);
            double dE = toDbl(batch.DE);

            // Valores Lab para calcular ejes A/B
            double stdL = toDbl(shade.StdL);
            double stdA = toDbl(shade.StdA);
            double stdB = toDbl(shade.StdB);
            double lotL = toDbl(batch.L);
            double lotA = toDbl(batch.A);
            double lotB = toDbl(batch.B);

            System.Diagnostics.Debug.WriteLine($"--- DEBUG TINT COATS (LEFT PANEL) ---");
            System.Diagnostics.Debug.WriteLine($"Std L: {stdL}, Lot L: {lotL}");
            System.Diagnostics.Debug.WriteLine($"Formula: Abs(({stdL} - {lotL}) / {stdL})");
            System.Diagnostics.Debug.WriteLine($"Resultado: {(stdL > 0 ? Math.Abs((stdL - lotL) / stdL) : 0)}");
            System.Diagnostics.Debug.WriteLine($"--------------------------------------");

            double dA = lotA - stdA;
            double dB = lotB - stdB;
            double pctA = (Math.Abs(stdA) > 0.1) ? (dA / Math.Abs(stdA)) : 0;
            double pctB = (Math.Abs(stdB) > 0.1) ? (dB / Math.Abs(stdB)) : 0;

            // Variaciones Relativas (Fórmula Industrial Coats: (Std - Lot) / Std)
            double relL = (stdL > 0) ? (stdL - lotL) / stdL : 0;

            // Para el panel de diagnóstico industrial (Polar), calculamos C y H desde Lab
            double stdC = Math.Sqrt(stdA * stdA + stdB * stdB);
            double lotC = Math.Sqrt(lotA * lotA + lotB * lotB);
            double relC = (stdC > 0) ? (stdC - lotC) / stdC : 0;

            double stdH = Math.Atan2(stdB, stdA) * 180.0 / Math.PI;
            if (stdH < 0) stdH += 360;
            double lotH = Math.Atan2(lotB, lotA) * 180.0 / Math.PI;
            if (lotH < 0) lotH += 360;

            double relH = lotH - stdH; // Factor H es absoluto en grados
            if (relH > 180) relH -= 360;
            if (relH < -180) relH += 360;

            // PRUEBA DE ESCRITORIO EN CONSOLA (DEBUG)
            System.Diagnostics.Debug.WriteLine("--- INICIO PRUEBA DE ESCRITORIO (D65 OCR) ---");
            System.Diagnostics.Debug.WriteLine($"Std L: {stdL}, Std a: {stdA}, Std b: {stdB}");
            System.Diagnostics.Debug.WriteLine($"Lot L: {lotL}, Lot a: {lotA}, Lot b: {lotB}");
            System.Diagnostics.Debug.WriteLine($"Cálculo dL Relativo: ({stdL} - {lotL}) / {stdL} = {relL}");
            System.Diagnostics.Debug.WriteLine($"Cálculo dC Relativo: ({stdC} - {lotC}) / {stdC} = {relC}");
            System.Diagnostics.Debug.WriteLine($"Cálculo dH Absoluto: {lotH} - {stdH} = {relH}");
            System.Diagnostics.Debug.WriteLine("--- FIN PRUEBA DE ESCRITORIO ---");

            if (dE > 0 && dE <= DE_MAX)
            {
                int i1 = dgvComparisonSummary.Rows.Add("dl", DL_MAX.ToString("F3"), "D65", "CUMPLE");
                int i2 = dgvComparisonSummary.Rows.Add("da", DC_MAX.ToString("F3"), "D65", "CUMPLE");
                int i3 = dgvComparisonSummary.Rows.Add("db", DH_MAX.ToString("F3"), "D65", "CUMPLE");
            }
            else
            {
                // Variaciones de Receta (Panel Izquierdo - Shade History)
                var res = new ColorCorrectionResult
                {
                    DeltaL = lotL - stdL,
                    DeltaChroma = lotC - stdC,
                    DeltaHue = relH,
                    FactorL = (decimal)relL,
                    DeltaA = dA,
                    DeltaB = dB,
                    FactorC = (decimal)relC,
                    FactorH = (decimal)relH
                };

                var resDL = Math.Abs(res.DeltaL) <= DL_MAX ? "CUMPLE" : "NO CUMPLE";
                int idxDL = dgvComparisonSummary.Rows.Add("dl", DL_MAX.ToString("F3"), "D65", resDL);
                if (resDL == "NO CUMPLE") dgvComparisonSummary.Rows[idxDL].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDA = Math.Abs(res.DeltaA) <= DC_MAX ? "CUMPLE" : "NO CUMPLE";
                int idxDA = dgvComparisonSummary.Rows.Add("da", DC_MAX.ToString("F3"), "D65", resDA);
                if (resDA == "NO CUMPLE") dgvComparisonSummary.Rows[idxDA].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDB = Math.Abs(res.DeltaB) <= DH_MAX ? "CUMPLE" : "NO CUMPLE";
                int idxDB = dgvComparisonSummary.Rows.Add("db", DH_MAX.ToString("F3"), "D65", resDB);
                if (resDB == "NO CUMPLE") dgvComparisonSummary.Rows[idxDB].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }
        }

        private void FillComparisonSummary(List<ColorCorrectionResult> results, double tolDE)
        {
            dgvComparisonSummary.Rows.Clear();
            dgvComparisonSummary.Rows.Add("Tolerancia CMC", $"DE {tolDE:F2}", "", "");

            foreach (var res in results)
            {
                if (res == null) continue;
                string status = res.DeltaE <= tolDE ? "CUMPLE" : "NO CUMPLE";
                string illum = res.Illuminant;
                if (string.IsNullOrEmpty(illum)) illum = "N/A";

                int idx = dgvComparisonSummary.Rows.Add("dE", res.DeltaE.ToString("F3"), illum, status);
                if (status == "NO CUMPLE")
                {
                    dgvComparisonSummary.Rows[idx].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                    dgvComparisonSummary.Rows[idx].Cells[3].Style.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                }
                else
                {
                    dgvComparisonSummary.Rows[idx].Cells[3].Style.ForeColor = System.Drawing.Color.ForestGreen;
                    dgvComparisonSummary.Rows[idx].Cells[3].Style.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                }
            }
        }

        private void ApplyEjeStyle(DataGridView dgv, int rowIndex, string eje)
        {
            if (rowIndex < 0 || rowIndex >= dgv.Rows.Count) return;
            var cell = dgv.Rows[rowIndex].Cells[0];
            cell.Value = eje;
            cell.Style.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Detectar si es un grid de iluminante secundario (tenue)
            bool esTenue = dgv.SelectionMode == DataGridViewSelectionMode.CellSelect;

            if (esTenue)
            {
                // Todos los ejes en gris muy suave — no compiten con D65
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(200, 200, 200);
                cell.Style.Font = new Font("Segoe UI", 8.5f, FontStyle.Regular);
                return;
            }

            if (dgv.DefaultCellStyle.SelectionBackColor == System.Drawing.Color.White)
            {
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(160, 170, 180);
                return;
            }

            // Grid principal (D65) — colores vivos con jerarquía
            if (eje.StartsWith("dl"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(45, 45, 45);
            else if (eje.StartsWith("da") || eje.StartsWith("dC"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);
            else if (eje.StartsWith("db") || eje.StartsWith("dH"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(180, 0, 0);
        }

        private void FillCorrectiveRecipeGrid(CorrectiveRecipeResult result, bool allComply = false)
        {
            dgvCorrectiveRecipe.Rows.Clear();
            if (result == null) return;

            Func<string, double, double> extractVal = (opt, orig) =>
            {
                if (string.IsNullOrWhiteSpace(opt) || opt == "---")
                    return orig;
                string[] parts = opt.Split(' ');
                if (parts.Length > 0)
                {
                    string clean = parts[0].Replace("%", "").Trim().Replace(",", ".");
                    if (double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                        return parsed;
                }
                return orig;
            };

            double totalRecipe1 = 0;
            double totalRecipe2 = 0;
            double totalRecipe3 = 0;

            if (result.Ingredients != null)
            {
                int tempIdx = 0;
                foreach (var ing in result.Ingredients)
                {
                    string opt1 = (tempIdx == 0) ? ing.Optiondl : "---";
                    totalRecipe1 += extractVal(opt1, ing.Original);

                    string opt2 = ing.Optionda;
                    totalRecipe2 += extractVal(opt2, ing.Original);

                    string opt3 = ing.Optiondb;
                    totalRecipe3 += extractVal(opt3, ing.Original);

                    tempIdx++;
                }
            }

            int rowCount = 0;
            foreach (var ing in result.Ingredients)
            {
                string optionDlDisplay = ing.Optiondl;

                // Ocultar datos de la fila 2 (índice 1) y fila 3 (índice 2) solo en Receta # 1
                if (rowCount == 1 || rowCount == 2)
                {
                    optionDlDisplay = "---";
                }

                double part = 0;
                if (rowCount == 0)
                {
                    double val = extractVal(optionDlDisplay, ing.Original);
                    part = totalRecipe1 > 0 ? (val / totalRecipe1) * 100 : 0;
                }
                else if (rowCount == 1)
                {
                    double val = extractVal(ing.Optionda, ing.Original);
                    part = totalRecipe2 > 0 ? (val / totalRecipe2) * 100 : 0;
                }
                else if (rowCount == 2)
                {
                    double val = extractVal(ing.Optiondb, ing.Original);
                    part = totalRecipe3 > 0 ? (val / totalRecipe3) * 100 : 0;
                }
                else
                {
                    part = result.TotalOriginal > 0 ? (ing.Original / result.TotalOriginal) * 100 : 0;
                }

                string partStr = part.ToString("F1") + "%";

                int idx = dgvCorrectiveRecipe.Rows.Add(
                    ing.Name,
                    ing.Original.ToString("F5"),
                    optionDlDisplay,
                    ing.Optionda,
                    ing.Optiondb,
                    partStr
                );

                if (ing.IsCritical && !allComply)
                {
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.Font = new Font(dgvCorrectiveRecipe.Font, FontStyle.Bold);
                }
                rowCount++;
            }

            // Agregar fila de Total
            double totalOriginal = result.Ingredients != null ? result.Ingredients.Sum(i => i.Original) : 0;
            int idxTotal = dgvCorrectiveRecipe.Rows.Add(
                "Total",
                totalOriginal.ToString("F5"),
                totalRecipe1.ToString("F5"),
                totalRecipe2.ToString("F5"),
                totalRecipe3.ToString("F5"),
                "100.0%"
            );
            dgvCorrectiveRecipe.Rows[idxTotal].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvCorrectiveRecipe.Rows[idxTotal].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            if (allComply)
            {
                lblAlertCorrective.Text = "Lote Cumple Tolerancia - Sin Ajustes";
                lblAlertCorrective.BackColor = System.Drawing.Color.ForestGreen;
            }
            else
            {
                lblAlertCorrective.Text = result.AlertMessage;
                lblAlertCorrective.BackColor = result.AlertSeverity == "None" ? System.Drawing.Color.ForestGreen : System.Drawing.Color.Firebrick;
            }
        }

        private void FillAnalysisGrid(DataGridView dgv, ColorCorrectionResult res, bool isRecipe)
        {
            dgv.Rows.Clear();
            if (res == null) return;

            if (dgv.ColumnCount == 5)
            {
                // Poblado para TABLA 2: Revisión de Apariencia (Ejes dL, dC, dH)
                // dL — DeltaL = Std - Lot:
                string impL = res.DeltaL > 0 ? " Oscuro" : " Claro";
                string actL = res.DeltaL > 0 ? "(Disminuir carga )" : "(Aumentar carga )";
                string adjL = (res.DeltaL > 0 ? "- " : "+ ") + Math.Abs(res.PercentL).ToString("F2") + "%";
                if (Math.Abs(res.PercentL) <= 0.1)
                {
                    impL = "✔";
                    actL = "✔";
                    adjL = "✔";
                }
                int r1 = dgv.Rows.Add("dL", Math.Abs((double)res.FactorL).ToString("F5"), impL, actL, adjL);
                ApplyEjeStyle(dgv, r1, "dL"); ApplyTenueRowStyle(dgv, r1);

                // dC — DeltaChroma = Std - Lot:
                string impC = res.DeltaChroma > 0 ? " Opaco / Apagado" : " Vivo / brillante";
                string actC = res.DeltaChroma > 0 ? "Avivar Tono" : "Opacar";
                string adjC = (res.DeltaChroma > 0 ? "+ " : "- ") + Math.Abs(res.PercentChroma).ToString("F2") + "%";
                if (Math.Abs(res.PercentChroma) <= 0.1)
                {
                    impC = "✔";
                    actC = "✔";
                    adjC = "✔";
                }
                int r2 = dgv.Rows.Add("dC", Math.Abs((double)res.FactorC).ToString("F5"), impC, actC, adjC);
                ApplyEjeStyle(dgv, r2, "dC"); ApplyTenueRowStyle(dgv, r2);

                // dH — dirección del viraje por eje dominante (da vs db)
                string impH;
                string actH = res.DeltaHue > 0 ? "Aumentar Matizador" : "Disminuir Matizador";
                string adjH = (res.DeltaHue > 0 ? "+ " : "- ") + Math.Abs(res.DeltaHue).ToString("F2") + "%";
                if (Math.Abs(res.DeltaHue) <= 0.1)
                {
                    impH = "✔";
                    actH = "✔";
                    adjH = "✔";
                }
                else
                {
                    if (Math.Abs(res.DeltaA) >= Math.Abs(res.DeltaB))
                        impH = res.DeltaA < 0 ? "Virado a Rojo" : "Virado a Verde";
                    else
                        impH = res.DeltaB < 0 ? "Virado a Amarillo" : "Virado a Azul";
                }
                int r3 = dgv.Rows.Add("dH", Math.Abs(res.DeltaHue).ToString("F5"), impH, actH, adjH);
                ApplyEjeStyle(dgv, r3, "dH"); ApplyTenueRowStyle(dgv, r3);

                HighlightChecks(dgv, r1);
                HighlightChecks(dgv, r2);
                HighlightChecks(dgv, r3);
            }
            else
            {
                string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
                string imp = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
                string rec = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

                string labelL = isRecipe ? "dl (Claro/Oscuro)" : "dl (Intensidad Carga)";
                string label2 = isRecipe ? "da (Rojo/Verde)" : "dC (Saturación/Limp)";
                string label3 = isRecipe ? "db (Amar/Azul)" : "dH (Tono/Matiz)";

                double val2 = isRecipe ? res.DeltaA : res.DeltaChroma;
                double val3 = isRecipe ? res.DeltaB : res.DeltaHue;

                int r1 = dgv.Rows.Add(labelL, res.DeltaL.ToString("F2"), Math.Abs((double)res.FactorL).ToString("F5"), imp, diag, rec);
                int r2 = dgv.Rows.Add(label2, val2.ToString("F2"), Math.Abs((double)res.FactorA).ToString("F5"), res.DescripcionC, res.DiagnosisC, res.RecommendationC);
                int r3 = dgv.Rows.Add(label3, val3.ToString("F2"), Math.Abs((double)res.FactorB).ToString("F5"), res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);

                ApplyEjeStyle(dgv, r1, labelL); ApplyTenueRowStyle(dgv, r1);
                ApplyEjeStyle(dgv, r2, label2); ApplyTenueRowStyle(dgv, r2);
                ApplyEjeStyle(dgv, r3, label3); ApplyTenueRowStyle(dgv, r3);

                HighlightChecks(dgv, r1);
                HighlightChecks(dgv, r2);
                HighlightChecks(dgv, r3);
            }

            if (res.DeltaE <= DE_MAX)
            {
                foreach (DataGridViewRow row in dgv.Rows) ApplyTenueRowStyle(dgv, row.Index);
            }
        }

        private void UpdateChart(EngineRes res)
        {
            if (res == null || _cielabChart == null) return;

            _cielabChart.DeltaL = res.DeltaL;
            _cielabChart.DeltaA = res.DeltaA;
            _cielabChart.DeltaB = res.DeltaB;
            _cielabChart.DeltaE = res.DeltaE;
            _cielabChart.ToleranceDE = DE_MAX;

            _cielabChart.AbsoluteL = res.StdL;
            _cielabChart.AbsoluteA = res.StdA;
            _cielabChart.AbsoluteB = res.StdB;

            _cielabChart.LotL = res.LotL;
            _cielabChart.LotA = res.LotA;
            _cielabChart.LotB = res.LotB;

            _cielabChart.Invalidate();
        }

        private void PopulateFromReport(OcrReport report)
        {
            if (report == null) return;

            lblRightShadeValue.Text = report.Batch?.ShadeName ?? "N/A";

            var expertAnalysis = EngineCalc.CalculateIndustrialCorrection(report);
            var allResults = EngineCalc.CalculateAllIlluminants(report);

            dgvShadeHistory.Rows.Clear();
            int idxShade = dgvShadeHistory.Rows.Add("Shade Name", report.Batch?.ShadeName ?? "N/A", "", "", "");
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

            int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration", "% Participación", "% Ajuste dL");
            dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            if (report.Recipe != null)
            {
                double totalConc = 0;
                foreach (var ing in report.Recipe)
                {
                    totalConc += ParsePercentageValue(ing.Percentage);
                }

                double percentL = expertAnalysis != null ? expertAnalysis.PercentL : 0;
                string percentLStr = percentL > 0 ? "+" + percentL.ToString("F2") + "%" : percentL.ToString("F2") + "%";

                foreach (var ing in report.Recipe)
                {
                    double pctVal = ParsePercentageValue(ing.Percentage);
                    double part = totalConc > 0 ? (pctVal / totalConc) * 100 : 0;
                    dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage, part.ToString("F1") + "%", percentLStr);
                }

                // Add Total row
                int idxTotal = dgvShadeHistory.Rows.Add("Total", "", totalConc.ToString("F5") + "%", "100.0%", "");
                dgvShadeHistory.Rows[idxTotal].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvShadeHistory.Rows[idxTotal].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            }

            var d65 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("D65"));
            var tl84 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("TL84"));
            var illA = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("A") || c.Illuminant.Contains("CWF"));

            FillComparisonSummary(allResults, report.TolDE);
            if (d65 != null)
            {
                var std = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double dA = 0, dB = 0;
                if (std != null && lot != null)
                {
                    dA = lot.A - std.A;
                    dB = lot.B - std.B;
                }

                string pL = expertAnalysis.PercentL > 0 ? "+" + expertAnalysis.PercentL.ToString("F2") : expertAnalysis.PercentL.ToString("F2");
                string pC = expertAnalysis.PercentChroma > 0 ? "+" + expertAnalysis.PercentChroma.ToString("F2") : expertAnalysis.PercentChroma.ToString("F2");
                string pH = expertAnalysis.DeltaHue > 0 ? "+" + expertAnalysis.PercentHue.ToString("F2") : "-" + expertAnalysis.PercentHue.ToString("F2");

                dgvCielabSummary.Rows.Clear();

                // ERROR (Δ) values:
                double errL = Math.Abs((double)expertAnalysis.FactorL);
                double errA = Math.Abs((double)expertAnalysis.FactorA);
                double errB = Math.Abs((double)expertAnalysis.FactorB);

                // Eje dl — DeltaLightness (OCR) = Lot - Std
                string lotLStr = d65.DeltaLightness < 0 ? " Oscuro" : " Claro";
                string actLStr = d65.DeltaLightness < 0 ? "Aumentar colorante" : "Disminuir colorante";
                string adjLStr = (d65.DeltaLightness < 0 ? "+ " : "- ") + Math.Abs(expertAnalysis.PercentL).ToString("F2") + "%";
                if (Math.Abs(expertAnalysis.PercentL) <= 0.1)
                {
                    lotLStr = "✔";
                    actLStr = "✔";
                    adjLStr = "✔";
                }
                int idxL = dgvCielabSummary.Rows.Add("dl", errL.ToString("F5"), lotLStr, actLStr, adjLStr);
                ApplyEjeStyle(dgvCielabSummary, idxL, "dl"); ApplyTenueRowStyle(dgvCielabSummary, idxL);

                // Eje da — REGLA: Bajar→"-"  Subir→"+"
                // dA aquí = d65.DeltaA = Std - Lot (convención motor)
                string lotAStr = dA < 0 ? " Mas Rojo" : " Mas Verde";
                string actAStr = dA < 0 ? "Disminuir el Rojo" : "Aumentar el Rojo";
                double stdA = std != null ? std.A : 0;
                double pctA = (Math.Abs(stdA) > 0.1) ? (dA / Math.Abs(stdA)) : 0;
                string adjAStr = (dA < 0 ? "- " : "+ ") + Math.Abs(pctA * 100).ToString("F2") + "%";
                if (Math.Abs(pctA * 100) <= 0.1)
                {
                    lotAStr = "✔";
                    actAStr = "✔";
                    adjAStr = "✔";
                }
                int idxA = dgvCielabSummary.Rows.Add("da", errA.ToString("F5"), lotAStr, actAStr, adjAStr);
                ApplyEjeStyle(dgvCielabSummary, idxA, "da"); ApplyTenueRowStyle(dgvCielabSummary, idxA);

                // Eje db — REGLA: Bajar→"-"  Subir→"+"
                // dB aquí = d65.DeltaB = Std - Lot (convención motor)
                string lotBStr = dB < 0 ? " Mas Amarillo" : " Mas Azul";
                string actBStr = dB < 0 ? "Disminuir el Amarillo" : "Disminuir el Azul";
                double stdB = std != null ? std.B : 0;
                double pctB = (Math.Abs(stdB) > 0.1) ? (dB / Math.Abs(stdB)) : 0;
                string adjBStr = "- " + Math.Abs(pctB * 100).ToString("F2") + "%";
                if (Math.Abs(pctB * 100) <= 0.1)
                {
                    lotBStr = "✔";
                    actBStr = "✔";
                    adjBStr = "✔";
                }
                int idxB = dgvCielabSummary.Rows.Add("db", errB.ToString("F5"), lotBStr, actBStr, adjBStr);
                ApplyEjeStyle(dgvCielabSummary, idxB, "db"); ApplyTenueRowStyle(dgvCielabSummary, idxB);

                HighlightChecks(dgvCielabSummary, idxL);
                HighlightChecks(dgvCielabSummary, idxA);
                HighlightChecks(dgvCielabSummary, idxB);
            }

            if (d65 != null)
            {
                var std = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (std != null && lot != null)
                {
                    pA = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    pB = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                }

                FillAnalysisGridFromCmc(dgvAnalysisLeft, d65, report.TolDE, true, pA, pB);
                FillAnalysisGridFromCmc(dgvAnalysisRight, d65, report.TolDE, false, pA, pB);
            }

            if (tl84 != null)
            {
                var std = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("TL84") && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("TL84") && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (std != null && lot != null)
                {
                    pA = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    pB = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                }
                FillAnalysisGridFromCmc(dgvAnalysisRightTL84, tl84, report.TolDE, false, pA, pB);
            }

            if (illA != null)
            {
                var std = report.Measures.FirstOrDefault(m => (m.Illuminant.Contains("A") || m.Illuminant.Contains("CWF")) && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => (m.Illuminant.Contains("A") || m.Illuminant.Contains("CWF")) && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (std != null && lot != null)
                {
                    pA = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    pB = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                }
                FillAnalysisGridFromCmc(dgvAnalysisRightA, illA, report.TolDE, false, pA, pB);
            }

            if (expertAnalysis.Success)
            {
                var ingredients = RecipeCorrector.IngredientsFromShade(new ShadeExtractionResult
                {
                    Recipe = report.Recipe
                });

                if (ingredients.Count > 0)
                {
                    var correctiveResult = RecipeCorrector.CalculateCorrectiveRecipe(ingredients, expertAnalysis);
                    bool allComply = allResults != null && allResults.Count > 0 && allResults.All(r => r.DeltaE <= report.TolDE);
                    FillCorrectiveRecipeGrid(correctiveResult, allComply);
                }

                _lastMainResult = expertAnalysis;
                UpdateChart(expertAnalysis);
            }

            dgvAnalysisRightTL84.ClearSelection();
            dgvAnalysisRightA.ClearSelection();
        }

        private void FillAnalysisGridFromCmc(DataGridView dgv, CmcDifferenceRow cmc, double tolDE, bool isRecipe, double pctA = 0, double pctB = 0)
        {
            dgv.Rows.Clear();
            if (cmc == null) return;

            if (dgv.ColumnCount == 5)
            {
                // Poblado para TABLA 2: Revisión de Apariencia (Ejes dL, dC, dH)
                string impL = cmc.DeltaLightness < 0 ? " Oscuro" : " Claro";
                string actL = cmc.DeltaLightness < 0 ? "(Aumentar carga )" : "(Disminuir carga )";
                string adjL = (cmc.DeltaLightness < 0 ? "+ " : "- ") + Math.Abs(cmc.DeltaLightness).ToString("F2") + "%";
                if (Math.Abs(cmc.DeltaLightness) <= 0.1)
                {
                    impL = "✔";
                    actL = "✔";
                    adjL = "✔";
                }
                double fL = cmc.DeltaLightness / 100.0;
                int r1 = dgv.Rows.Add("dL", Math.Abs(fL).ToString("F5"), impL, actL, adjL);
                ApplyEjeStyle(dgv, r1, "dL"); ApplyTenueRowStyle(dgv, r1);

                // dC
                string impC = cmc.DeltaChroma >= 0 ? " Vivo / Brillante" : " Opaco / Apagado";
                string actC = cmc.DeltaChroma >= 0 ? "Opacar " : "Avivar Tono";
                string adjC = (cmc.DeltaChroma >= 0 ? "- " : "+ ") + Math.Abs(cmc.DeltaChroma).ToString("F2") + "%";
                if (Math.Abs(cmc.DeltaChroma) <= 0.1)
                {
                    impC = "✔";
                    actC = "✔";
                    adjC = "✔";
                }
                double fC = cmc.DeltaChroma / 100.0;
                int r2 = dgv.Rows.Add("dC", Math.Abs(fC).ToString("F5"), impC, actC, adjC);
                ApplyEjeStyle(dgv, r2, "dC"); ApplyTenueRowStyle(dgv, r2);

                // dH — dirección del viraje por eje dominante (pctA vs pctB pasados como parámetros)
                string impH;
                string actH = cmc.DeltaHue >= 0 ? "Aumentar Matizador" : "Disminuir Matizador";
                string adjH = (cmc.DeltaHue >= 0 ? "+ " : "- ") + Math.Abs(cmc.DeltaHue).ToString("F2") + "%";
                if (Math.Abs(cmc.DeltaHue) <= 0.1)
                {
                    impH = "✔";
                    actH = "✔";
                    adjH = "✔";
                }
                else
                {
                    if (Math.Abs(pctA) >= Math.Abs(pctB))
                        impH = pctA < 0 ? "Virado a Rojo" : "Virado a Verde";
                    else
                        impH = pctB < 0 ? "Virado a Amarillo" : "Virado a Azul";
                }
                int r3 = dgv.Rows.Add("dH", Math.Abs(cmc.DeltaHue).ToString("F5"), impH, actH, adjH);
                ApplyEjeStyle(dgv, r3, "dH"); ApplyTenueRowStyle(dgv, r3);

                HighlightChecks(dgv, r1);
                HighlightChecks(dgv, r2);
                HighlightChecks(dgv, r3);

                if (cmc.DeltaCMC <= tolDE)
                {
                    ApplyTenueRowStyle(dgv, r1);
                    ApplyTenueRowStyle(dgv, r2);
                    ApplyTenueRowStyle(dgv, r3);
                }
            }
            else
            {
                decimal fL = (decimal)cmc.DeltaLightness / 100m;
                decimal fC = (decimal)cmc.DeltaChroma / 100m;

                var res = new ColorCorrectionResult
                {
                    DeltaL = cmc.DeltaLightness,
                    DeltaChroma = cmc.DeltaChroma,
                    DeltaHue = cmc.DeltaHue,
                    DeltaA = pctA * 50,
                    DeltaB = pctB * 50
                };

                string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
                string imp = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
                string rec = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

                string labelL = isRecipe ? "dl (Claro/Oscuro)" : "dl (Intensidad Carga)";
                string label2 = isRecipe ? "da (Rojo/Verde)" : "dC (Saturación/Limp)";
                string label3 = isRecipe ? "db (Amar/Azul)" : "dH (Tono/Matiz)";

                double val2 = isRecipe ? res.DeltaA : res.DeltaChroma;
                double val3 = isRecipe ? res.DeltaB : res.DeltaHue;

                int r1 = dgv.Rows.Add(labelL, res.DeltaL.ToString("F2"), Math.Abs((double)fL).ToString("F5"), imp, diag, rec);
                int r2 = dgv.Rows.Add(label2, val2.ToString("F2"), Math.Abs((double)fC).ToString("F5"), res.DescripcionC, res.DiagnosisC, res.RecommendationC);
                int r3 = dgv.Rows.Add(label3, val3.ToString("F2"), "0.00000", res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);

                ApplyEjeStyle(dgv, r1, labelL);
                ApplyEjeStyle(dgv, r2, label2);
                ApplyEjeStyle(dgv, r3, label3);

                HighlightChecks(dgv, r1);
                HighlightChecks(dgv, r2);
                HighlightChecks(dgv, r3);

                if (cmc.DeltaCMC <= tolDE)
                {
                    ApplyTenueRowStyle(dgv, r1);
                    ApplyTenueRowStyle(dgv, r2);
                    ApplyTenueRowStyle(dgv, r3);
                }
            }
        }

        private void BtnRegresar_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Retry;
            this.Close();
        }

        private void BtnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvCorrectiveRecipe.Rows.Count == 0)
                {
                    MessageBox.Show("No hay datos para guardar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 1. Obtener Metadatos Maestro (Cabecera)
                string shadeName = _shadeData?.ShadeName ?? "Unknown";
                DateTime fechaActual = DateTime.Now;

                string iluminante = _lastMainResult?.Illuminant ?? "D65";

                // Datos Cualitativos (Panel Izquierdo - Motor Experto)
                string impL = _lastMainResult?.ImpactoRecetaL ?? "";
                string diagL = _lastMainResult?.DiagnosticoL ?? "";
                string recL = _lastMainResult?.RecomendacionRecetaL ?? "";

                string impC = _lastMainResult?.DescripcionC ?? "";
                string diagC = _lastMainResult?.DiagnosisC ?? "";
                string recC = _lastMainResult?.RecommendationC ?? "";

                string impH = _lastMainResult?.ImpactoMatiz ?? "";
                string diagH = _lastMainResult?.DiagnosisH ?? "";
                string recH = _lastMainResult?.RecomendacionMatiz ?? "";

                // 2. Guardar en Base de Datos Unificada (Fila por componente)
                foreach (DataGridViewRow row in dgvCorrectiveRecipe.Rows)
                {
                    if (row.IsNewRow) continue;

                    string name = row.Cells["Colorante"].Value?.ToString() ?? "";
                    if (name.Equals("Total", StringComparison.OrdinalIgnoreCase)) continue;

                    string strOriginal = row.Cells["Receta Original"].Value?.ToString() ?? "0";
                    string strAdjDL = row.Cells["Receta # 1"].Value?.ToString() ?? "0";
                    string strAdjDC = row.Cells["Receta # 2"].Value?.ToString() ?? "0";
                    string strAdjDH = row.Cells["Receta # 3"].Value?.ToString() ?? "0";

                    // Determinar la nueva receta (el valor que no sea "---" entre las opciones dl, da, db)
                    string strNueva = strOriginal;
                    if (strAdjDL != "---") strNueva = strAdjDL;
                    else if (strAdjDC != "---") strNueva = strAdjDC;
                    else if (strAdjDH != "---") strNueva = strAdjDH;

                    // Conversión numérica estricta para persistencia
                    decimal.TryParse(strOriginal.Replace("%", ""), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal concOriginal);

                    // Buscar el código en dgvShadeHistory
                    string code = "";
                    foreach (DataGridViewRow rShade in dgvShadeHistory.Rows)
                    {
                        if (rShade.Cells[1].Value?.ToString() == name)
                        {
                            code = rShade.Cells[0].Value?.ToString() ?? "";
                            break;
                        }
                    }

                    // Filtrar recomendaciones para que solo aparezcan en el colorante relevante
                    string pDye = _lastMainResult?.PrimaryDyeName ?? "";
                    string sDye = _lastMainResult?.SecondaryDyeName ?? "";
                    string tDye = _lastMainResult?.TonerDyeName ?? "";

                    string recC_Filtrada = (name == sDye) ? recC : "✔";
                    string recH_Filtrada = (name == tDye) ? recH : "✔";

                    // Para el Eje L (Luminosidad), la recomendación es global, se mantiene para todos
                    string recL_Filtrada = recL;

                    Color.Services.HistorialService.GuardarRegistroMaestro(
                        shadeName, fechaActual, iluminante,
                        name,
                        concOriginal, strAdjDL, strAdjDC, strAdjDH,
                        impL, diagL, recL_Filtrada,
                        impC, diagC, recC_Filtrada,
                        impH, diagH, recH_Filtrada
                    );
                }

                // 3. Notificación y Opción de Reporte
                var result = MessageBox.Show($"Datos del Shade {shadeName} guardados exitosamente.\n\n¿Desea generar el reporte técnico detallado (.txt)?",
                                            "Finalización Exitosa", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                // Bloquear botón para evitar duplicación por doble clic
                btnGuardar.Enabled = false;
                btnGuardar.Text = "✔ Guardado";
                btnGuardar.BackColor = System.Drawing.Color.FromArgb(50, 160, 80);

                if (result == DialogResult.Yes)
                {
                    GenerarReporteTexto();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en la integridad de datos: " + ex.Message, "Error de Sistema", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void GenerarReporteTexto()
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Texto (*.txt)|*.txt";
                    sfd.FileName = "Reporte_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".txt";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        var sb = new System.Text.StringBuilder();
                        sb.AppendLine("=== REPORTE DE COLORIMETRIA ===");
                        sb.AppendLine("Fecha: " + DateTime.Now.ToString());
                        sb.AppendLine();

                        Action<string, DataGridView> exportGrid = (title, dgv) => {
                            sb.AppendLine("--- " + title + " ---");
                            foreach (DataGridViewRow row in dgv.Rows)
                            {
                                if (!row.IsNewRow)
                                {
                                    var cells = new List<string>();
                                    for (int i = 0; i < row.Cells.Count; i++)
                                    {
                                        if (row.Cells[i].Value != null) cells.Add(row.Cells[i].Value.ToString());
                                    }
                                    sb.AppendLine(string.Join(" | ", cells));
                                }
                            }
                            sb.AppendLine();
                        };

                        exportGrid("ANALISIS DE SHADE HISTORY REPORT", dgvShadeHistory);
                        exportGrid("ANALISIS ILUMINANTE D65 (IZQ)", dgvAnalysisLeft);
                        exportGrid("ANALISIS DE SAMPLE COMPARISON", dgvComparisonSummary);
                        exportGrid("ANALISIS ILUMINANTE D65 (DER)", dgvAnalysisRight);
                        exportGrid("ANALISIS ILUMINANTE TL84 (DER)", dgvAnalysisRightTL84);
                        exportGrid("ANALISIS ILUMINANTE A/CWF (DER)", dgvAnalysisRightA);

                        sb.AppendLine("--- PROTOCOLO DE SEGURIDAD EN INGENIERÍA DE COLOR ---");
                        sb.AppendLine("Nota sobre Alertas de Medición y Ajuste:");
                        sb.AppendLine("El sistema implementa una capa de Inteligencia Preventiva que va más allá de la comparación visual simple.");
                        sb.AppendLine("1. Validación Multiluminante: El lote se evalúa simultáneamente bajo tres fuentes de luz (D65, TL84, A). El estado 'NO CUMPLE' se activa si existe riesgo de metamerismo (desviación crítica) en al menos un iluminante.");
                        sb.AppendLine("2. Límite de Estabilidad de Receta (Umbral 15%): La alerta roja de 'Ajustes > 15%' indica que la corrección química necesaria es drástica. Superar este umbral representa un riesgo para la estabilidad de la mezcla.");
                        sb.AppendLine();

                        System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                        MessageBox.Show("Reporte de texto guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Error al exportar: " + ex.Message); }
        }
    }
}