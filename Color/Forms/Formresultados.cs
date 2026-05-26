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
            dgvShadeHistory.Columns[3].Name = " Participación";
            dgvShadeHistory.Columns[4].Name = " Ajuste dL";
            dgvShadeHistory.Columns[4].Visible = false;
            dgvShadeHistory.Columns[0].FillWeight = 15;
            dgvShadeHistory.Columns[1].FillWeight = 50;
            dgvShadeHistory.Columns[2].FillWeight = 15;
            dgvShadeHistory.Columns[3].FillWeight = 20;
            dgvShadeHistory.Columns[4].FillWeight = 15;

            dgvAnalysisLeftTL84 = CreateAnalysisGrid();
            dgvAnalysisLeftA = CreateAnalysisGrid();

            dgvCielabSummary = CreateStyledGrid();
            dgvCielabSummary.ColumnHeadersVisible = true;
            dgvCielabSummary.ColumnCount = 5;
            dgvCielabSummary.Columns[0].Name = "EJE";
            dgvCielabSummary.Columns[1].Name = "Variacion";
            dgvCielabSummary.Columns[2].Name = "AJUSTE";
            dgvCielabSummary.Columns[3].Name = "Impacto";
            dgvCielabSummary.Columns[4].Name = "Accion";

            dgvCielabSummary.Columns[0].HeaderText = "EJE";
            dgvCielabSummary.Columns[1].HeaderText = "Variacion (Δ)";
            dgvCielabSummary.Columns[2].HeaderText = "Impacto";
            dgvCielabSummary.Columns[3].HeaderText = "Accion";
            dgvCielabSummary.Columns[4].HeaderText = "AJUSTE ";

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

            dgvAnalysisLeft = CreateStyledGrid();
            dgvAnalysisLeft.ColumnHeadersVisible = true;
            dgvAnalysisLeft.ColumnCount = 5;
            dgvAnalysisLeft.Columns[0].Name = "EJE";
            dgvAnalysisLeft.Columns[1].Name = "Variacion";
            dgvAnalysisLeft.Columns[2].Name = "AJUSTE";
            dgvAnalysisLeft.Columns[3].Name = "Impacto";
            dgvAnalysisLeft.Columns[4].Name = "Accion";

            dgvAnalysisLeft.Columns[0].HeaderText = "EJE";
            dgvAnalysisLeft.Columns[1].HeaderText = "Variacion (Δ)";
            dgvAnalysisLeft.Columns[2].HeaderText = "Impacto";
            dgvAnalysisLeft.Columns[3].HeaderText = "Accion";
            dgvAnalysisLeft.Columns[4].HeaderText = "AJUSTE ";

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
            dgvComparisonSummary.Columns[0].Name = "axis";
            dgvComparisonSummary.Columns[1].Name = "Tolerance";
            dgvComparisonSummary.Columns[2].Name = "Illuminant";
            dgvComparisonSummary.Columns[3].Name = "Result";
            foreach (DataGridViewColumn col in dgvComparisonSummary.Columns) col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvComparisonSummary.Columns[0].FillWeight = 30;
            dgvComparisonSummary.Columns[1].FillWeight = 25;
            dgvComparisonSummary.Columns[2].FillWeight = 20;
            dgvComparisonSummary.Columns[3].FillWeight = 25;

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
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("Iluminante D65: (Ejes dl, da, db)"), 0, 2);
            pnlLeftRecipe.Controls.Add(dgvCielabSummary, 0, 3);
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel("Iluminante D65: (Ejes dL, dC, dH)"), 0, 4);
            pnlLeftRecipe.Controls.Add(dgvAnalysisLeft, 0, 5);

            var pnlCorrectiveContainer = new Panel { Dock = DockStyle.Fill };
            pnlCorrectiveContainer.Controls.Add(dgvCorrectiveRecipe);
            pnlCorrectiveContainer.Controls.Add(lblAlertCorrective);
            dgvCorrectiveRecipe.Dock = DockStyle.Fill;
            lblAlertCorrective.Dock = DockStyle.Bottom;
            pnlLeftRecipe.Controls.Add(CreateHeaderLabel(" FORMULACIÓN CORRECTIVA (D65)"), 0, 6);
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

            pnlRightLot.Controls.Add(CreateHeaderLabel("ANALISIS DE PASS / FAIL (LOTE)"), 0, 0);

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
            dgv.Columns[1].Visible = false; 
            dgv.Columns[2].Name = "Receta # 1"; dgv.Columns[2].FillWeight = 15;
            dgv.Columns[3].Name = "Receta # 2"; dgv.Columns[3].FillWeight = 15;
            dgv.Columns[4].Name = "Receta # 3"; dgv.Columns[4].FillWeight = 15;
            dgv.Columns[5].Name = "Participación"; dgv.Columns[5].FillWeight = 15;

            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 5;
            dgv.Columns[0].Name = "EJE";          dgv.Columns[0].FillWeight = 8;
            dgv.Columns[1].Name = "Variacion(Δ)"; dgv.Columns[1].FillWeight = 15;
            dgv.Columns[1].HeaderCell.ToolTipText = "(Std - Lot)";
            dgv.Columns[2].Name = "Impacto";      dgv.Columns[2].FillWeight = 22;
            dgv.Columns[3].Name = "Accion";  dgv.Columns[3].FillWeight = 30;
            dgv.Columns[4].Name = "Ajuste";       dgv.Columns[4].FillWeight = 25;
            return dgv;
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
        private void ApplyTenueGridStyle(DataGridView dgv)
        {
            var lightGray = System.Drawing.Color.FromArgb(200, 200, 200);
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
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204); 
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

                int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration", "Participación", " Ajuste dL");
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

                // Eje dl
                string varL = (d65.FactorL >= 0 ? "+" : "") + d65.FactorL.ToString("F2") + "%";
                string adjL = (d65.DeltaL > 0 ? "- " : "+ ") + Math.Abs(d65.FactorL).ToString("F2") + "%";
                string lotL = d65.DeltaL < 0 ? " Claro" : " Oscuro";
                string actL = d65.DeltaL < 0 ? "Aumentar colorante" : "Reducir colorante";
                
                if (Math.Abs(d65.DeltaL) <= DL_MAX)
                {
                    varL = "✔"; lotL = "✔"; actL = "✔"; adjL = "✔";
                }
                int idxL = dgvCielabSummary.Rows.Add("dl", varL, adjL, lotL, actL);
                ApplyEjeStyle(dgvCielabSummary, idxL, "dl"); ApplyTenueRowStyle(dgvCielabSummary, idxL);

                // Eje da
                double pctA = (double)d65.FactorA;
                string varA = (pctA >= 0 ? "+" : "") + pctA.ToString("F2") + "%";
                string adjA = (d65.DeltaA < 0 ? "- " : "+ ") + Math.Abs(pctA).ToString("F2") + "%";
                string lotA = d65.DeltaA < 0 ? " Mas Rojo" : " Mas Verde";
                string actA = d65.DeltaA < 0 ? "Reducir el Rojo" : "Aumentar el Rojo";
                if (Math.Abs(d65.DeltaA) <= DC_MAX)
                {
                    varA = "✔"; lotA = "✔"; actA = "✔"; adjA = "✔";
                }
                int idxA = dgvCielabSummary.Rows.Add("da", varA, adjA, lotA, actA);
                ApplyEjeStyle(dgvCielabSummary, idxA, "da"); ApplyTenueRowStyle(dgvCielabSummary, idxA);

                // Eje db
                double pctB = (double)d65.FactorB;
                string varB = (pctB >= 0 ? "+" : "") + pctB.ToString("F2") + "%";
                string adjB = (d65.DeltaB < 0 ? "- " : "+ ") + Math.Abs(pctB).ToString("F2") + "%";
                string lotB = d65.DeltaB < 0 ? " Mas Amarillo" : " Mas Azul";
                string actB = d65.DeltaB < 0 ? "Reducir el Amarillo" : "Aumentar el Azul";
                if (Math.Abs(d65.DeltaB) <= DC_MAX)
                {
                    varB = "✔"; lotB = "✔"; actB = "✔"; adjB = "✔";
                }
                int idxB = dgvCielabSummary.Rows.Add("db", varB, adjB, lotB, actB);
                ApplyEjeStyle(dgvCielabSummary, idxB, "db"); ApplyTenueRowStyle(dgvCielabSummary, idxB);

                HighlightChecks(dgvCielabSummary, idxL);
                HighlightChecks(dgvCielabSummary, idxA);
                HighlightChecks(dgvCielabSummary, idxB);

                // --- TABLA IZQUIERDA Y DERECHA: Sincronización Total con el Motor ---
                FillAnalysisGrid(dgvAnalysisLeft, d65, true);
                FillRightPanelGrid_LCH(dgvAnalysisRight, d65);

                if (ill2 != null) FillRightPanelGrid_LCH(dgvAnalysisRightTL84, ill2);
                if (ill3 != null) FillRightPanelGrid_LCH(dgvAnalysisRightA, ill3);

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

        private void FillComparisonSummary(List<ColorCorrectionResult> results, double tolDE)
        {
            dgvComparisonSummary.Rows.Clear();
            
            // Fila de Encabezado de Tolerancia (Fija según estándar de negocio)
            int h1 = dgvComparisonSummary.Rows.Add(" ", $"DE {tolDE:F2}", "", "");
            dgvComparisonSummary.Rows[h1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvComparisonSummary.Rows[h1].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvComparisonSummary.Rows[h1].DefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);

            // Agregamos dL, dC, dH para el iluminante principal (D65) como constantes de negocio
            var d65 = results.FirstOrDefault(r => r.Illuminant.Contains("D65"));
            if (d65 != null)
            {
                AddComparisonRow("dL", d65.DeltaL, DL_MAX, "D65");
                AddComparisonRow("dC", d65.DeltaChroma, DC_MAX, "D65");
                AddComparisonRow("dH", d65.DeltaHue, DH_MAX, "D65");
            }

            foreach (var res in results)
            {
                if (res == null) continue;
                AddComparisonRow("dE", res.DeltaE, tolDE, res.Illuminant);
            }
        }

        private void AddComparisonRow(string facet, double value, double limit, string illuminant)
        {
            string status = Math.Abs(value) <= limit ? "CUMPLE" : "NO CUMPLE";
            string illum = string.IsNullOrEmpty(illuminant) ? "N/A" : illuminant;

            // Columna Tolerance muestra el LÍMITE ESTÁTICO de negocio, no el valor medido.
            int idx = dgvComparisonSummary.Rows.Add(facet, limit.ToString("F3"), illum, status);
            
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
                    string activeAdj = "---";
                    if (tempIdx == 0) activeAdj = ing.Optiondl;
                    else if (ing.Optionda != "---") activeAdj = ing.Optionda;
                    else if (ing.Optiondb != "---") activeAdj = ing.Optiondb;

                    string opt1 = (tempIdx == 0) ? activeAdj : "---";
                    totalRecipe1 += extractVal(opt1, ing.Original);

                    string opt2 = (tempIdx == 1) ? activeAdj : "---";
                    totalRecipe2 += extractVal(opt2, ing.Original);

                    string opt3 = (tempIdx == 2) ? activeAdj : "---";
                    totalRecipe3 += extractVal(opt3, ing.Original);

                    tempIdx++;
                }
            }

            int rowCount = 0;
            foreach (var ing in result.Ingredients)
            {
                string activeAdj = "---";
                if (rowCount == 0) activeAdj = ing.Optiondl;
                else if (ing.Optionda != "---") activeAdj = ing.Optionda;
                else if (ing.Optiondb != "---") activeAdj = ing.Optiondb;

                string col1 = (rowCount == 0) ? activeAdj : "---";
                string col2 = (rowCount == 1) ? activeAdj : "---";
                string col3 = (rowCount == 2) ? activeAdj : "---";

                double part = 0;
                if (rowCount == 0)
                {
                    double val = extractVal(col1, ing.Original);
                    part = totalRecipe1 > 0 ? (val / totalRecipe1) * 100 : 0;
                }
                else if (rowCount == 1)
                {
                    double val = extractVal(col2, ing.Original);
                    part = totalRecipe2 > 0 ? (val / totalRecipe2) * 100 : 0;
                }
                else if (rowCount == 2)
                {
                    double val = extractVal(col3, ing.Original);
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
                    col1,
                    col2,
                    col3,
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

            if (result.AlertSeverity == "Critical")
            {
                lblAlertCorrective.Text = result.AlertMessage;
                lblAlertCorrective.BackColor = System.Drawing.Color.Firebrick;
            }
            else if (result.AlertSeverity == "Warning")
            {
                lblAlertCorrective.Text = result.AlertMessage;
                lblAlertCorrective.BackColor = System.Drawing.Color.DarkGoldenrod;
            }
            else if (!allComply)
            {
                lblAlertCorrective.Text = "Lote Fuera de Tolerancia - Revisar Medición";
                lblAlertCorrective.BackColor = System.Drawing.Color.Firebrick;
            }
            else
            {
                lblAlertCorrective.Text = result.AlertMessage;
                lblAlertCorrective.BackColor = System.Drawing.Color.ForestGreen;
            }
        }

        private void FillAnalysisGrid(DataGridView dgv, ColorCorrectionResult res, bool isRecipe, double pctA = 0, double pctB = 0)
        {
            dgv.Rows.Clear();
            if (res == null) return;

            // ---- ROW 1: dL ----
            string varL = res.DeltaL.ToString("F3");
            string adjL = (res.DeltaL > 0 ? "- " : "+ ") + Math.Abs(res.PercentL).ToString("F2") + "%";
            string impL = res.DeltaL > 0 ? " Oscuro" : " Claro";
            string actL = res.DeltaL > 0 ? "Reducir carga" : "Aumentar carga";
            
            if (Math.Abs(res.DeltaL) <= DL_MAX)
            {
                varL = "✔"; adjL = "✔"; impL = "✔"; actL = "✔";
            }
            int r1 = dgv.Rows.Add("dL", varL, adjL, impL, actL);
            ApplyEjeStyle(dgv, r1, "dL"); ApplyTenueRowStyle(dgv, r1);

            // ---- ROW 2: dC ----
            string varC = res.DeltaChroma.ToString("F3");
            string adjC = (res.DeltaChroma > 0 ? "+ " : "- ") + Math.Abs(res.PercentChroma).ToString("F2") + "%";
            string impC = res.DeltaChroma > 0 ? " Opaco / Apagado" : " Vivo / Brillante";
            string actC = res.DeltaChroma > 0 ? "Avivar Tono" : "Opacar";
            
            if (Math.Abs(res.DeltaChroma) <= DC_MAX)
            {
                varC = "✔"; adjC = "✔"; impC = "✔"; actC = "✔";
            }
            int r2 = dgv.Rows.Add("dC", varC, adjC, impC, actC);
            ApplyEjeStyle(dgv, r2, "dC"); ApplyTenueRowStyle(dgv, r2);

            // ---- ROW 3: dH ----
            string varH = res.DeltaHue.ToString("F3");
            string adjH = (res.DeltaHue > 0 ? "+ " : "- ") + Math.Abs(res.PercentHue).ToString("F2") + "%";
            string impH = "";
            if (Math.Abs(res.DeltaA) >= Math.Abs(res.DeltaB))
                impH = res.DeltaA < 0 ? "Virado a Rojo" : "Virado a Verde";
            else
                impH = res.DeltaB < 0 ? "Virado a Amarillo" : "Virado a Azul";

            string actH = res.DeltaHue > 0 ? "Aumentar Matizador" : "Reducir Matizador";
            
            if (Math.Abs(res.DeltaHue) <= DH_MAX)
            {
                varH = "✔"; adjH = "✔"; impH = "✔"; actH = "✔";
            }
            int r3 = dgv.Rows.Add("dH", varH, adjH, impH, actH);
            ApplyEjeStyle(dgv, r3, "dH"); ApplyTenueRowStyle(dgv, r3);

            HighlightChecks(dgv, r1);
            HighlightChecks(dgv, r2);
            HighlightChecks(dgv, r3);
        }

        private void FillRightPanelGrid_LCH(DataGridView dgv, ColorCorrectionResult res)
        {
            FillAnalysisGrid(dgv, res, false);
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

            int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration", " Participación", " Ajuste dL");
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
                string actLStr = d65.DeltaLightness < 0 ? "Aumentar colorante" : "Reducir colorante";
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
                string lotAStr = dA < 0 ? " Mas Rojo" : " Mas Verde";
                string actAStr = dA < 0 ? "Reducir el Rojo" : "Aumentar el Rojo";
                double stdA = std != null ? std.A : 0;
                double pctA_raw = (Math.Abs(stdA) > 0.1) ? (dA / Math.Abs(stdA)) : 0;
                double pctA = Math.Max(-0.15, Math.Min(0.15, pctA_raw));
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
                string lotBStr = dB < 0 ? " Mas Amarillo" : " Mas Azul";
                string actBStr = dB < 0 ? "Reducir el Amarillo" : "Reducir el Azul";
                double stdB = std != null ? std.B : 0;
                double pctB_raw = (Math.Abs(stdB) > 0.1) ? (dB / Math.Abs(stdB)) : 0;
                double pctB = Math.Max(-0.15, Math.Min(0.15, pctB_raw));
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
                    double pA_raw = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    double pB_raw = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                    pA = Math.Max(-0.15, Math.Min(0.15, pA_raw));
                    pB = Math.Max(-0.15, Math.Min(0.15, pB_raw));
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
                    double pA_raw = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    double pB_raw = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                    pA = Math.Max(-0.15, Math.Min(0.15, pA_raw));
                    pB = Math.Max(-0.15, Math.Min(0.15, pB_raw));
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
                    double pA_raw = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    double pB_raw = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                    pA = Math.Max(-0.15, Math.Min(0.15, pA_raw));
                    pB = Math.Max(-0.15, Math.Min(0.15, pB_raw));
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

            // ---- ROW 1: dL ----
            string varL = cmc.DeltaLightness.ToString("F3");
            string adjL = (cmc.DeltaLightness < 0 ? "+ " : "- ") + Math.Abs(cmc.DeltaLightness).ToString("F2") + "%";
            string impL = cmc.DeltaLightness < 0 ? " Oscuro" : " Claro";
            string actL = cmc.DeltaLightness < 0 ? "Aumentar carga" : "Reducir carga";
            
            if (Math.Abs(cmc.DeltaLightness) <= DL_MAX)
            {
                varL = "✔"; adjL = "✔"; impL = "✔"; actL = "✔";
            }
            int r1 = dgv.Rows.Add("dL", varL, adjL, impL, actL);
            ApplyEjeStyle(dgv, r1, "dL"); ApplyTenueRowStyle(dgv, r1);

            // ---- ROW 2: dC ----
            string varC = cmc.DeltaChroma.ToString("F3");
            string adjC = (cmc.DeltaChroma >= 0 ? "- " : "+ ") + Math.Abs(cmc.DeltaChroma).ToString("F2") + "%";
            string impC = cmc.DeltaChroma >= 0 ? " Vivo / Brillante" : " Opaco / Apagado";
            string actC = cmc.DeltaChroma >= 0 ? "Opacar" : "Avivar Tono";
            
            if (Math.Abs(cmc.DeltaChroma) <= DC_MAX)
            {
                varC = "✔"; adjC = "✔"; impC = "✔"; actC = "✔";
            }
            int r2 = dgv.Rows.Add("dC", varC, adjC, impC, actC);
            ApplyEjeStyle(dgv, r2, "dC"); ApplyTenueRowStyle(dgv, r2);

            // ---- ROW 3: dH ----
            string varH = cmc.DeltaHue.ToString("F3");
            string adjH = (cmc.DeltaHue >= 0 ? "+ " : "- ") + Math.Abs(cmc.DeltaHue).ToString("F2") + "%";
            string impH = "";
            if (Math.Abs(pctA) >= Math.Abs(pctB))
                impH = pctA < 0 ? "Virado a Rojo" : "Virado a Verde";
            else
                impH = pctB < 0 ? "Virado a Amarillo" : "Virado a Azul";

            string actH = cmc.DeltaHue >= 0 ? "Aumentar Matizador" : "Reducir Matizador";
            
            if (Math.Abs(cmc.DeltaHue) <= DH_MAX)
            {
                varH = "✔"; adjH = "✔"; impH = "✔"; actH = "✔";
            }
            int r3 = dgv.Rows.Add("dH", varH, adjH, impH, actH);
            ApplyEjeStyle(dgv, r3, "dH"); ApplyTenueRowStyle(dgv, r3);

            HighlightChecks(dgv, r1);
            HighlightChecks(dgv, r2);
            HighlightChecks(dgv, r3);

            if (cmc.DeltaCMC <= tolDE)
            {
                foreach (DataGridViewRow row in dgv.Rows) ApplyTenueRowStyle(dgv, row.Index);
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
                        exportGrid("ANALISIS DE PASS / FAIL", dgvComparisonSummary);
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