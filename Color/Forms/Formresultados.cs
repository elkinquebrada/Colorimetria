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
        private DataGridView dgvAnalysisRight;
        private DataGridView dgvAnalysisRightTL84;
        private DataGridView dgvAnalysisRightA;
        private DataGridView dgvCorrectiveRecipe;
        private Label lblAlertCorrective;

        private RichTextBox txtReport;
        private RichTextBox txtRecomendacion;
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
            btnGuardar = CreateStyledButton("💾 Guardar", System.Drawing.Color.FromArgb(45, 126, 247));
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
            dgvShadeHistory.ColumnCount = 3;
            dgvShadeHistory.Columns[0].Name = "Dye Code";
            dgvShadeHistory.Columns[1].Name = "Dye Names";
            dgvShadeHistory.Columns[2].Name = "Concentration";

            dgvAnalysisLeft = CreateAnalysisGrid();
            dgvAnalysisLeftTL84 = CreateAnalysisGrid();
            dgvAnalysisLeftA = CreateAnalysisGrid();

            dgvComparisonSummary = CreateStyledGrid();
            dgvComparisonSummary.ColumnHeadersVisible = false;
            dgvComparisonSummary.ColumnCount = 4;
            dgvComparisonSummary.Columns[0].Name = "Fact";
            dgvComparisonSummary.Columns[1].Name = "Tolerance";
            dgvComparisonSummary.Columns[2].Name = "Illuminant";
            dgvComparisonSummary.Columns[3].Name = "Result";

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
            lblAlertCorrective = new Label { 
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

            var pnlLeft = CreatePanelWithGrids("ANALISIS DE SHADE HISTORY REPORT", dgvShadeHistory, 
                                               "ANALISIS ILUMINANTE D65", dgvAnalysisLeft);

            var pnlRight = CreatePanelWithManyGrids("ANALISIS DE SAMPLE COMPARISON", dgvComparisonSummary, 
                                                   "ANALISIS ILUMINANTE D65", dgvAnalysisRight,
                                                   "ANALISIS ILUMINANTE TL84", dgvAnalysisRightTL84,
                                                   "ANALISIS ILUMINANTE A / CWF", dgvAnalysisRightA);

            // --- Panel Izquierdo Unificado (Grillas + Receta) ---
            var pnlLeftUnified = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 6 };
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));  
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Percent, 33));   
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));  
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Percent, 33));   
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));  
            pnlLeftUnified.RowStyles.Add(new RowStyle(SizeType.Percent, 34));   

            // Forzar eliminación de barras de scroll en grillas técnicas para asegurar visión total
            dgvShadeHistory.ScrollBars = ScrollBars.None;
            dgvAnalysisLeft.ScrollBars = ScrollBars.None;
            dgvCorrectiveRecipe.ScrollBars = ScrollBars.None;
            dgvShadeHistory.BorderStyle = BorderStyle.None;
            dgvAnalysisLeft.BorderStyle = BorderStyle.None;
            dgvCorrectiveRecipe.BorderStyle = BorderStyle.None;

            // 1. Shade History
            pnlLeftUnified.Controls.Add(CreateHeaderLabel("ANALISIS DE SHADE HISTORY REPORT"), 0, 0);
            pnlLeftUnified.Controls.Add(dgvShadeHistory, 0, 1);
            
            // 2. D65 Analysis
            pnlLeftUnified.Controls.Add(CreateHeaderLabel("ANALISIS ILUMINANTE D65"), 0, 2);
            pnlLeftUnified.Controls.Add(dgvAnalysisLeft, 0, 3);

            // 3. Receta Correctiva
            var pnlCorrectiveContainer = new Panel { Dock = DockStyle.Fill };
            pnlCorrectiveContainer.Controls.Add(dgvCorrectiveRecipe);
            pnlCorrectiveContainer.Controls.Add(lblAlertCorrective);
            dgvCorrectiveRecipe.Dock = DockStyle.Fill;
            lblAlertCorrective.Dock = DockStyle.Bottom;
            pnlLeftUnified.Controls.Add(CreateHeaderLabel("RESUMEN DE FORMULACIÓN CORRECTIVA (D65)"), 0, 4);
            pnlLeftUnified.Controls.Add(pnlCorrectiveContainer, 0, 5);

            splitMedicionesCmc.Panel1.Controls.Add(pnlLeftUnified);
            splitMedicionesCmc.Panel2.Controls.Add(pnlRight);

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
                btnGuardar.Left  = btnVerGrafico.Left - btnGuardar.Width - 10;
                btnVerGrafico.Top = 12;
                btnGuardar.Top    = 12;
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
            dgv.ColumnCount = 5;
            dgv.Columns[0].Name = "Colorante";            dgv.Columns[0].FillWeight = 25;
            dgv.Columns[1].Name = "Receta Original";      dgv.Columns[1].FillWeight = 15;
            dgv.Columns[2].Name = "Receta # 1";    dgv.Columns[2].FillWeight = 15;
            dgv.Columns[3].Name = "Receta # 2";    dgv.Columns[3].FillWeight = 15;
            dgv.Columns[4].Name = "Receta # 3";     dgv.Columns[4].FillWeight = 15;
            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "EJE";          dgv.Columns[0].FillWeight = 10;
            dgv.Columns[1].Name = "VARIACION";    dgv.Columns[1].FillWeight = 12;
            dgv.Columns[2].Name = "Δ %";        dgv.Columns[2].FillWeight = 10;
            dgv.Columns[3].Name = "IMPACTO";      dgv.Columns[3].FillWeight = 18;
            dgv.Columns[4].Name = "DIAGNOSTICO";   dgv.Columns[4].FillWeight = 25;
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
                if (e.RowIndex >= 0) {
                    dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                    dgv.Rows[e.RowIndex].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                    dgv.Rows[e.RowIndex].Cells[0].Style.ForeColor = System.Drawing.Color.White;
                }
            };
            dgv.CellMouseLeave += (s, e) => {
                if (e.RowIndex >= 0) {
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
                if (e.RowIndex >= 0) {
                    var row = dgv.Rows[e.RowIndex];
                    row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204); // Azul Coats
                    row.DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                    foreach (DataGridViewCell cell in row.Cells) {
                        cell.Style.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                        cell.Style.ForeColor = System.Drawing.Color.White;
                    }
                }
            };

            dgv.CellMouseLeave += (s, e) => {
                if (e.RowIndex >= 0) {
                    var row = dgv.Rows[e.RowIndex];
                    var originalBg = (e.RowIndex % 2 == 0) ? bgGray : System.Drawing.Color.FromArgb(242, 242, 242);
                    row.DefaultCellStyle.BackColor = originalBg;
                    row.DefaultCellStyle.ForeColor = lightGray;
                    foreach (DataGridViewCell cell in row.Cells) {
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
                if (cell.Value != null && cell.Value.ToString() == "✔")
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

        private void PopulateFromObjects(ShadeExtractionResult shadeData, List<EngineRes> results)
        {
            if (shadeData != null)
            {
                dgvShadeHistory.Rows.Clear();
                int idxShade = dgvShadeHistory.Rows.Add("Shade Name", shadeData.ShadeName ?? "N/A", "");
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

                int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration");
                dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
                dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                if (shadeData.Recipe != null)
                {
                    foreach (var ing in shadeData.Recipe)
                        dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage);
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
                
                dgvComparisonSummary.Rows.Clear();
                string shadeName = !string.IsNullOrEmpty(d65.ShadeName) ? d65.ShadeName : (shadeData?.ShadeName ?? "N/A");
                
                int idxShade2 = dgvComparisonSummary.Rows.Add("Shade Name", shadeName, "", "");
                dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
                dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
                dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

                int idxHdr2 = dgvComparisonSummary.Rows.Add("Facet", "Tolerance", "Illuminant", "Result");
                dgvComparisonSummary.Rows[idxHdr2].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
                dgvComparisonSummary.Rows[idxHdr2].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                
                // --- Cuadro de Tolerancia CMC Estándar (Formato Profesional) ---
                string tolSummary = $"DE {DE_MAX:F2}";
                int tolIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolSummary, "", "");

                // --- Filas Detalladas (dl, da, db) ---
                var resDL = Math.Abs(d65.DeltaL) <= DL_MAX ? "CUMPLE" : "NO CUMPLE";
                int idxDL = dgvComparisonSummary.Rows.Add("dl", DL_MAX.ToString("F3"), "D65", resDL);
                if (resDL == "NO CUMPLE") dgvComparisonSummary.Rows[idxDL].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDC = (ill2 != null && Math.Abs(ill2.DeltaChroma) <= DC_MAX) ? "CUMPLE" : "NO CUMPLE";
                int idxDC = dgvComparisonSummary.Rows.Add("da", DC_MAX.ToString("F3"), (ill2?.Illuminant ?? "TL84"), resDC);
                if (resDC == "NO CUMPLE") dgvComparisonSummary.Rows[idxDC].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDH = (ill3 != null && Math.Abs(ill3.DeltaHue) <= DH_MAX) ? "CUMPLE" : "NO CUMPLE";
                int idxDH = dgvComparisonSummary.Rows.Add("db", DH_MAX.ToString("F3"), (ill3?.Illuminant ?? "A"), resDH);
                if (resDH == "NO CUMPLE") dgvComparisonSummary.Rows[idxDH].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                
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
                    FillCorrectiveRecipeGrid(correctiveResult);
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
                int i1 = dgv.Rows.Add("", relL.ToString("F5"), dL.ToString("F2"), "DENTRO DE TOLERANCIA", "Normal", "OK");
                int i2 = dgv.Rows.Add("", relC.ToString("F5"), dC.ToString("F2"), "DENTRO DE TOLERANCIA", "Normal", "OK");
                int i3 = dgv.Rows.Add("", relH.ToString("F5"), dH.ToString("F2"), "DENTRO DE TOLERANCIA", "Normal", "OK");
                ApplyEjeStyle(dgv, i1, "dl Luminosidad");
                ApplyEjeStyle(dgv, i2, "da (Rojo/Verde)");
                ApplyEjeStyle(dgv, i3, "db (Amar/Azul)");
            }
            else
            {
                // Variaciones de Receta (Panel Izquierdo - Shade History)
                var res = new ColorCorrectionResult {
                    DeltaL = lotL - stdL,
                    DeltaChroma = lotC - stdC,
                    DeltaHue = relH,
                    FactorL = (decimal)relL, 
                    DeltaA = dA,
                    DeltaB = dB,
                    FactorC = (decimal)relC,
                    FactorH = (decimal)relH
                };

                string diag = res.DiagnosticoL;
                string imp = res.ImpactoRecetaL;
                string rec = res.RecomendacionRecetaL;

                // Mostrar valores base para transparencia total (Std vs Lot)
                string valBaseL = $"[S:{stdL:F2} L:{lotL:F2}]";
                int r1 = dgv.Rows.Add("dl Luminosidad", ((double)res.FactorL).ToString("F5"), res.DeltaL.ToString("F2"), imp, diag, rec);
                dgv.Rows[r1].Cells[0].ToolTipText = valBaseL; 
                
                string valBaseC = $"[S:{stdC:F2} L:{lotC:F2}]";
                int r2 = dgv.Rows.Add("da (Rojo/Verde)", ((double)res.FactorC).ToString("F5"), res.DeltaChroma.ToString("F2"), res.DescripcionC, res.DiagnosisC, res.RecommendationC);
                dgv.Rows[r2].Cells[0].ToolTipText = valBaseC;
 
                string valBaseH = $"[S:{stdH:F2} L:{lotH:F2}]";
                int r3 = dgv.Rows.Add("db (Amar/Azul)", ((double)res.FactorH).ToString("F5"), res.DeltaHue.ToString("F2"), res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);
                dgv.Rows[r3].Cells[0].ToolTipText = valBaseH;
                
                ApplyEjeStyle(dgv, r1, "dl Luminosidad");
                ApplyEjeStyle(dgv, r2, "da (Rojo/Verde)");
                ApplyEjeStyle(dgv, r3, "db (Amar/Azul)");
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
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(45, 45, 45);       // Casi negro
            else if (eje.StartsWith("da"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);    // Gris medio
            else if (eje.StartsWith("db"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(180, 0, 0);        // Rojo oscuro
        }

        private void FillCorrectiveRecipeGrid(CorrectiveRecipeResult result)
        {
            dgvCorrectiveRecipe.Rows.Clear();
            if (result == null) return;

            int rowCount = 0;
            foreach (var ing in result.Ingredients)
            {
                string optionDlDisplay = ing.Optiondl;
                
                // Ocultar datos de la fila 2 (índice 1) y fila 3 (índice 2) solo en Receta # 1
                if (rowCount == 1 || rowCount == 2)
                {
                    optionDlDisplay = "---";
                }

                int idx = dgvCorrectiveRecipe.Rows.Add(
                    ing.Name,
                    ing.Original.ToString("F5"),
                    optionDlDisplay,
                    ing.Optionda,
                    ing.Optiondb
                );

                if (ing.IsCritical)
                {
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.Font = new Font(dgvCorrectiveRecipe.Font, FontStyle.Bold);
                }
                rowCount++;
            }

            lblAlertCorrective.Text = result.AlertMessage;
            lblAlertCorrective.BackColor = result.AlertSeverity == "None" ? System.Drawing.Color.ForestGreen : System.Drawing.Color.Firebrick;
        }

        private void FillAnalysisGrid(DataGridView dgv, ColorCorrectionResult res, bool isRecipe)
        {
            dgv.Rows.Clear();
            if (res == null) return;

            string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
            string imp  = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
            string rec  = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

            int r1 = dgv.Rows.Add("dl Luminosidad", Math.Abs((double)res.FactorL).ToString("F5"), res.DeltaL.ToString("F2"), imp, diag, rec);
            int r2 = dgv.Rows.Add("da (Rojo/Verde)", Math.Abs((double)res.FactorA).ToString("F5"), res.DeltaA.ToString("F2"), res.DescripcionC, res.DiagnosisC, res.RecommendationC);
            int r3 = dgv.Rows.Add("db (Amar/Azul)", Math.Abs((double)res.FactorB).ToString("F5"), res.DeltaB.ToString("F2"), res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);
            
            ApplyEjeStyle(dgv, r1, "dl"); ApplyTenueRowStyle(dgv, r1);
            ApplyEjeStyle(dgv, r2, "da"); ApplyTenueRowStyle(dgv, r2);
            ApplyEjeStyle(dgv, r3, "db"); ApplyTenueRowStyle(dgv, r3);

            // Resaltado de Checks Verdes
            HighlightChecks(dgv, r1);
            HighlightChecks(dgv, r2);
            HighlightChecks(dgv, r3);

            // Resaltar si cumple o no de forma visual pero sin ocultar datos
            if (res.DeltaE <= DE_MAX)
            {
                foreach (DataGridViewRow row in dgv.Rows) ApplyTenueRowStyle(dgv, row.Index);
            }
        }

        // --- HELPERS DE MATIZ ---

        private void UpdateChart(EngineRes res)
        {
            if (res == null || _cielabChart == null) return;
            
            _cielabChart.DeltaL = res.DeltaL;
            _cielabChart.DeltaA = res.DeltaA;
            _cielabChart.DeltaB = res.DeltaB;
            _cielabChart.DeltaE = res.DeltaE;
            _cielabChart.ToleranceDE = DE_MAX;
            
            // Usar valores absolutos reales para la visualización del color
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

            dgvShadeHistory.Rows.Clear();
            int idxShade = dgvShadeHistory.Rows.Add("Shade Name", report.Batch?.ShadeName ?? "N/A", "");
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvShadeHistory.Rows[idxShade].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

            int idxHdr1 = dgvShadeHistory.Rows.Add("Dye Code", "Dye Names", "Concentration");
            dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgvShadeHistory.Rows[idxHdr1].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            if (report.Recipe != null)
            {
                foreach (var ing in report.Recipe)
                    dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage);
            }

            // Buscamos los iluminantes en las medidas
            var d65 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("D65"));
            var tl84 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("TL84"));
            var illA = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("A") || c.Illuminant.Contains("CWF"));

            dgvComparisonSummary.Rows.Clear();
            int idxShade2 = dgvComparisonSummary.Rows.Add("Shade Name", report.Batch?.ShadeName ?? "N/A", "", "");
            dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            dgvComparisonSummary.Rows[idxShade2].DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

            int idxHdr2 = dgvComparisonSummary.Rows.Add("Dato", "Tolerancia", "Iluminante", "Resultado");
            dgvComparisonSummary.Rows[idxHdr2].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgvComparisonSummary.Rows[idxHdr2].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            
            // --- Cuadro de Tolerancia CMC Estándar (Formato Profesional) ---
            string tolSummary = $"DE {report.TolDE:F2}";
            int tIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolSummary, "", "");

            // --- Filas Detalladas (DL, DC, DH) ---
            if (d65 != null)
            {
                var resDL = Math.Abs(d65.DeltaLightness) <= report.TolDL ? "CUMPLE" : "NO CUMPLE";
                int idxDL = dgvComparisonSummary.Rows.Add("dl", report.TolDL.ToString("F3"), "D65", resDL);
                if (resDL == "NO CUMPLE") dgvComparisonSummary.Rows[idxDL].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }
            if (tl84 != null)
            {
                var resDC = Math.Abs(tl84.DeltaChroma) <= report.TolDC ? "CUMPLE" : "NO CUMPLE";
                int idxDC = dgvComparisonSummary.Rows.Add("da", report.TolDC.ToString("F3"), "TL84", resDC);
                if (resDC == "NO CUMPLE") dgvComparisonSummary.Rows[idxDC].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }
            if (illA != null)
            {
                var resDH = Math.Abs(illA.DeltaHue) <= report.TolDH ? "CUMPLE" : "NO CUMPLE";
                int idxDH = dgvComparisonSummary.Rows.Add("db", report.TolDH.ToString("F3"), illA.Illuminant, resDH);
                if (resDH == "NO CUMPLE") dgvComparisonSummary.Rows[idxDH].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }

            if (d65 != null)
            {
                // Buscar medidas Lab para calcular ejes A/B
                var std = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (std != null && lot != null) {
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
                if (std != null && lot != null) {
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
                if (std != null && lot != null) {
                    pA = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    pB = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                }
                FillAnalysisGridFromCmc(dgvAnalysisRightA, illA, report.TolDE, false, pA, pB);
            }

            // --- CALCULO DE RECETA CORRECTIVA FASE 2 (MOTOR EXPERTO) ---
            var expertAnalysis = EngineCalc.CalculateIndustrialCorrection(report);
            
            if (expertAnalysis.Success)
            {
                var ingredients = RecipeCorrector.IngredientsFromShade(new ShadeExtractionResult { 
                    Recipe = report.Recipe
                });
                
                if (ingredients.Count > 0)
                {
                    var correctiveResult = RecipeCorrector.CalculateCorrectiveRecipe(ingredients, expertAnalysis);
                    FillCorrectiveRecipeGrid(correctiveResult);
                }

                _lastMainResult = expertAnalysis;
                UpdateChart(expertAnalysis);
            }

            // Limpiar selección para evitar filas azules resaltadas al inicio
            dgvAnalysisRightTL84.ClearSelection();
            dgvAnalysisRightA.ClearSelection();
        }

        private void FillAnalysisGridFromCmc(DataGridView dgv, CmcDifferenceRow cmc, double tolDE, bool isRecipe, double pctA = 0, double pctB = 0)
        {
            dgv.Rows.Clear();
            if (cmc == null) return;

            decimal fL = (decimal)cmc.DeltaLightness / 100m;
            decimal fC = (decimal)cmc.DeltaChroma / 100m;

            // Crear objeto de resultado temporal para usar lógica dinámica del motor
            var res = new ColorCorrectionResult {
                DeltaL = cmc.DeltaLightness,
                DeltaChroma = cmc.DeltaChroma,
                DeltaHue = cmc.DeltaHue,
                DeltaA = pctA * 50, 
                DeltaB = pctB * 50
            };

            string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
            string imp = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
            string rec = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

            int r1 = dgv.Rows.Add("", Math.Abs((double)fL).ToString("F5"), res.DeltaL.ToString("F2"), diag, imp, rec);
            int r2 = dgv.Rows.Add("", Math.Abs((double)fC).ToString("F5"), res.DeltaChroma.ToString("F2"), res.DiagnosisC, res.DescripcionC, res.RecommendationC);
            int r3 = dgv.Rows.Add("", "0.00000", res.DeltaHue.ToString("F2"), res.DiagnosisH, res.ImpactoMatiz, res.RecomendacionMatiz);
            
            ApplyEjeStyle(dgv, r1, "dl (Luminosidad)"); 
            ApplyEjeStyle(dgv, r2, "da (Brillo)"); 
            ApplyEjeStyle(dgv, r3, "db (Matiz)");

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

                    string name = row.Cells[0].Value?.ToString() ?? "";
                    string strOriginal = row.Cells[1].Value?.ToString() ?? "0";
                    string strAdjDL = row.Cells[2].Value?.ToString() ?? "0";
                    string strAdjDC = row.Cells[3].Value?.ToString() ?? "0"; 
                    string strAdjDH = row.Cells[4].Value?.ToString() ?? "0"; 

                    // Determinar la nueva receta (el valor que no sea "---" entre las opciones dl, da, db)
                    string strNueva = strOriginal;
                    if (strAdjDL != "---") strNueva = strAdjDL;
                    else if (strAdjDC != "---") strNueva = strAdjDC;
                    else if (strAdjDH != "---") strNueva = strAdjDH;

                    // Conversión numérica estricta para persistencia
                    decimal.TryParse(strOriginal.Replace("%",""), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal concOriginal);

                    // Buscar el código en dgvShadeHistory
                    string code = "";
                    foreach (DataGridViewRow rShade in dgvShadeHistory.Rows) {
                        if (rShade.Cells[1].Value?.ToString() == name) {
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
                            foreach (DataGridViewRow row in dgv.Rows) {
                                if (!row.IsNewRow) {
                                    var cells = new List<string>();
                                    for (int i = 0; i < row.Cells.Count; i++) {
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