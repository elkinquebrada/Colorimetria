using Color.Forms;
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
        private List<Color.IlluminantCorrectionResult> _recipeResults;
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
        }

        public FormResultados(string resumen, List<EngineRes> results, List<Color.IlluminantCorrectionResult> recipeResults = null, ShadeExtractionResult shadeData = null)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = results ?? new List<EngineRes>();
            _recipeResults = recipeResults;
            _shadeData = shadeData;
            InitializeComponents();

            // Lógica silenciosa: Poblar desde los objetos ya calculados
            PopulateFromObjects(_shadeData, _resultsLegacy);
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
            btnCerrar = CreateStyledButton("Finalizar", System.Drawing.Color.FromArgb(200, 30, 30));
            btnCerrar.Click += (s, e) => this.Close();
            btnRegresar = CreateStyledButton("← Regresar", System.Drawing.Color.FromArgb(180, 100, 30));
            btnRegresar.Click += BtnRegresar_Click;

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
                var frm = new FormDetalleCielab(_lastMainResult.DeltaL, _lastMainResult.DeltaA, _lastMainResult.DeltaB, _lastMainResult.DeltaE, _lastMainResult.CmcValue, 1.20, "");
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
                BackColor = System.Drawing.Color.Gray
            };
            
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
            dgv.Columns[0].Name = "Colorante";         dgv.Columns[0].FillWeight = 25;
            dgv.Columns[1].Name = "% Receta Original";     dgv.Columns[1].FillWeight = 12;
            dgv.Columns[2].Name = "% Ajuste DL";      dgv.Columns[2].FillWeight = 12;
            dgv.Columns[3].Name = "% Ajuste DH";      dgv.Columns[3].FillWeight = 12;
            dgv.Columns[4].Name = "% Nueva Receta"; dgv.Columns[4].FillWeight = 18;
            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "EJE";          dgv.Columns[0].FillWeight = 10;
            dgv.Columns[1].Name = "VARIACION";    dgv.Columns[1].FillWeight = 12;
            dgv.Columns[2].Name = "Δ";        dgv.Columns[2].FillWeight = 10;
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

        /// <summary>
        /// Aplica estilo tenue (gris claro) a las grillas de iluminantes secundarios
        /// para que no compitan visualmente con el iluminante principal D65.
        /// </summary>
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

        /// Aplica estilo tenue a todas las celdas de una fila específica.
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

                // --- Filas Detalladas (DL, DC, DH) ---
                var resDL = Math.Abs(d65.DeltaL) <= DL_MAX ? "CUMPLE" : "NO CUMPLE";
                int idxDL = dgvComparisonSummary.Rows.Add("DL", DL_MAX.ToString("F3"), "D65", resDL);
                if (resDL == "NO CUMPLE") dgvComparisonSummary.Rows[idxDL].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDC = (ill2 != null && Math.Abs(ill2.DeltaChroma) <= DC_MAX) ? "CUMPLE" : "NO CUMPLE";
                int idxDC = dgvComparisonSummary.Rows.Add("DC", DC_MAX.ToString("F3"), (ill2?.Illuminant ?? "TL84"), resDC);
                if (resDC == "NO CUMPLE") dgvComparisonSummary.Rows[idxDC].Cells[3].Style.ForeColor = System.Drawing.Color.Red;

                var resDH = (ill3 != null && Math.Abs(ill3.DeltaHue) <= DH_MAX) ? "CUMPLE" : "NO CUMPLE";
                int idxDH = dgvComparisonSummary.Rows.Add("DH", DH_MAX.ToString("F3"), (ill3?.Illuminant ?? "A"), resDH);
                if (resDH == "NO CUMPLE") dgvComparisonSummary.Rows[idxDH].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                
                
                // --- TABLA IZQUIERDA: Datos del Shade History Report (OCR) ---
                if (shadeData != null && shadeData.Batch != null)
                {
                    var recipeD65 = _recipeResults?.FirstOrDefault(r => r.Illuminant.Contains("D65"));
                    FillAnalysisGridFromOcr(dgvAnalysisLeft, shadeData, recipeD65?.VariacionL);
                }
                else
                {
                    FillAnalysisGrid(dgvAnalysisLeft, d65, true);
                }

                // --- TABLA DERECHA: Datos del Sample Comparison (Cálculo actual) ---
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

            double dL = toDbl(batch.DL) * 10;
            double dC = toDbl(batch.DC) * 10;
            double dH = toDbl(batch.DH) * 10;
            double dE = toDbl(batch.DE);

            // Valores Lab para calcular ejes A/B
            double stdA = toDbl(shade.StdA);
            double stdB = toDbl(shade.StdB);
            double lotA = toDbl(batch.A);
            double lotB = toDbl(batch.B);
            
            double dA = lotA - stdA;
            double dB = lotB - stdB;
            double pctA = (Math.Abs(stdA) > 0.1) ? (dA / Math.Abs(stdA)) : 0;
            double pctB = (Math.Abs(stdB) > 0.1) ? (dB / Math.Abs(stdB)) : 0;

            if (dE > 0 && dE <= DE_MAX)
            {
                int i1 = dgv.Rows.Add("", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "LOTE APROBADO", "Normal");
                int i2 = dgv.Rows.Add("", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "Normal");
                int i3 = dgv.Rows.Add("", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "Normal");
                ApplyEjeStyle(dgv, i1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, i2, "DC (Brillo)");
                ApplyEjeStyle(dgv, i3, "DH (Matiz)");
            }
            else
            {
                // Variaciones de Receta (Panel Izquierdo - Shade History)
                var res = new ColorCorrectionResult {
                    DeltaL = toDbl(batch.DL),
                    DeltaChroma = toDbl(batch.DC),
                    DeltaHue = toDbl(batch.DH),
                    PercentL = (varL ?? (toDbl(batch.DL) * 10)) / 100.0, 
                    DeltaA = lotA - stdA,
                    DeltaB = lotB - stdB,
                    PercentA = pctA,
                    PercentB = pctB,
                    PercentChroma = toDbl(batch.DC) / 100.0
                };

                int r1 = dgv.Rows.Add("", res.DeltaL.ToString("F2"), $"{res.PorcentajeRecetaL:F1}%", res.ImpactoRecetaL, res.DiagnosticoL, res.RecomendacionRecetaL);
                int r2 = dgv.Rows.Add("", res.DeltaChroma.ToString("F2"), $"{Math.Abs(res.PercentChroma * 100):F1}%", res.DescripcionC, res.DiagnosisC, res.RecommendationC);
                int r3 = dgv.Rows.Add("", res.DeltaHue.ToString("F2"), $"{Math.Abs(res.PercentHue * 100):F1}%", res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);
                ApplyEjeStyle(dgv, r1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, r2, "DC (Brillo)");
                ApplyEjeStyle(dgv, r3, "DH (Matiz)");
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
            if (eje.StartsWith("DL"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(45, 45, 45);       // Casi negro
            else if (eje.StartsWith("DC"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);    // Gris medio
            else if (eje.StartsWith("DH"))
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(180, 0, 0);        // Rojo oscuro
        }

        private void FillCorrectiveRecipeGrid(CorrectiveRecipeResult result)
        {
            dgvCorrectiveRecipe.Rows.Clear();
            if (result == null) return;

            foreach (var ing in result.Ingredients)
            {
                int idx = dgvCorrectiveRecipe.Rows.Add(
                    ing.Name,
                    ing.Original.ToString("F5"),
                    (ing.FactorDL >= 1 ? "+" : "") + ((ing.FactorDL - 1) * 100).ToString("F5"),
                    (ing.FactorDH >= 1 ? "+" : "") + ((ing.FactorDH - 1) * 100).ToString("F5"),
                    ing.NewConcentration.ToString("F5")
                );

                if (ing.Status == "SATURACIÓN")
                {
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.BackColor = System.Drawing.Color.MistyRose;
                    dgvCorrectiveRecipe.Rows[idx].DefaultCellStyle.ForeColor = System.Drawing.Color.Red;
                }
            }

            lblAlertCorrective.Text = result.AlertMessage;
            switch (result.AlertSeverity)
            {
                case "Critical":
                case "Error":
                    lblAlertCorrective.BackColor = System.Drawing.Color.Firebrick;
                    break;
                case "Warning":
                    lblAlertCorrective.BackColor = System.Drawing.Color.Goldenrod;
                    break;
                case "None":
                    lblAlertCorrective.BackColor = System.Drawing.Color.ForestGreen;
                    break;
                default:
                    lblAlertCorrective.BackColor = System.Drawing.Color.Gray;
                    break;
            }
        }

        private void FillAnalysisGrid(DataGridView dgv, EngineRes res, bool isRecipe)
        {
            dgv.Rows.Clear();
            if (res == null) return;

            if (res.CmcValue <= DE_MAX || res.DeltaE <= DE_MAX)
            {
                int i1 = dgv.Rows.Add("", (res.DeltaL * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "LOTE APROBADO", "Normal");
                int i2 = dgv.Rows.Add("", (res.DeltaChroma * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "Normal");
                int i3 = dgv.Rows.Add("", (res.DeltaHue * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "Normal");
                ApplyEjeStyle(dgv, i1, "DL (Fuerza)"); ApplyTenueRowStyle(dgv, i1);
                ApplyEjeStyle(dgv, i2, "DC (Brillo)"); ApplyTenueRowStyle(dgv, i2);
                ApplyEjeStyle(dgv, i3, "DH (Matiz)"); ApplyTenueRowStyle(dgv, i3);
            }
            else
            {
                string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
                string imp  = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
                string rec  = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

                int r1 = dgv.Rows.Add("", res.DeltaL.ToString("F2"), $"{res.PorcentajeRecetaL:F1}%", imp, diag, rec);
                int r2 = dgv.Rows.Add("", res.DeltaChroma.ToString("F2"), $"{Math.Abs(res.PercentChroma * 100):F1}%", res.DescripcionC, res.DiagnosisC, res.RecommendationC);
                int r3 = dgv.Rows.Add("", res.DeltaHue.ToString("F2"), $"{Math.Abs(res.PercentHue * 100):F1}%", res.ImpactoMatiz, res.DiagnosisH, res.RecomendacionMatiz);
                ApplyEjeStyle(dgv, r1, "DL (Fuerza)"); ApplyTenueRowStyle(dgv, r1);
                ApplyEjeStyle(dgv, r2, "DC (Brillo)"); ApplyTenueRowStyle(dgv, r2);
                ApplyEjeStyle(dgv, r3, "DH (Matiz)"); ApplyTenueRowStyle(dgv, r3);
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
            
            // Si tenemos valores absolutos (AbsDelta no es el Abs del delta sino el valor absoluto del lote)
            // Nota: En este motor, AbsDeltaL parece ser el valor absoluto.
            _cielabChart.AbsoluteL = res.AbsDeltaL - res.DeltaL;
            _cielabChart.AbsoluteA = res.AbsDeltaA - res.DeltaA;
            _cielabChart.AbsoluteB = res.AbsDeltaB - res.DeltaB;
            
            _cielabChart.LotL = res.AbsDeltaL;
            _cielabChart.LotA = res.AbsDeltaA;
            _cielabChart.LotB = res.AbsDeltaB;

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
                int idxDL = dgvComparisonSummary.Rows.Add("DL", report.TolDL.ToString("F3"), "D65", resDL);
                if (resDL == "NO CUMPLE") dgvComparisonSummary.Rows[idxDL].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }
            if (tl84 != null)
            {
                var resDC = Math.Abs(tl84.DeltaChroma) <= report.TolDC ? "CUMPLE" : "NO CUMPLE";
                int idxDC = dgvComparisonSummary.Rows.Add("DC", report.TolDC.ToString("F3"), "TL84", resDC);
                if (resDC == "NO CUMPLE") dgvComparisonSummary.Rows[idxDC].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
            }
            if (illA != null)
            {
                var resDH = Math.Abs(illA.DeltaHue) <= report.TolDH ? "CUMPLE" : "NO CUMPLE";
                int idxDH = dgvComparisonSummary.Rows.Add("DH", report.TolDH.ToString("F3"), illA.Illuminant, resDH);
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

            // --- CALCULO DE RECETA CORRECTIVA (D65) ---
            if (d65 != null)
            {
                var stdD65 = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && m.Type.ToUpper().Contains("STD"));
                var lotD65 = report.Measures.FirstOrDefault(m => m.Illuminant.Contains("D65") && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (stdD65 != null && lotD65 != null) {
                    pA = (Math.Abs(stdD65.A) > 0.1) ? (lotD65.A - stdD65.A) / Math.Abs(stdD65.A) : 0;
                    pB = (Math.Abs(stdD65.B) > 0.1) ? (lotD65.B - stdD65.B) / Math.Abs(stdD65.B) : 0;
                }

                var resD65 = new ColorCorrectionResult {
                    Illuminant = "D65",
                    DeltaL = d65.DeltaLightness,
                    DeltaHue = d65.DeltaHue,
                    PercentL = d65.DeltaLightness,
                    PercentA = pA,
                    PercentB = pB
                };

                var ingredients = RecipeCorrector.IngredientsFromShade(new ShadeExtractionResult { 
                    Recipe = report.Recipe
                });
                
                if (ingredients.Count > 0)
                {
                    var correctiveResult = RecipeCorrector.CalculateCorrectiveRecipe(ingredients, resD65);
                    FillCorrectiveRecipeGrid(correctiveResult);
                }

                // Actualizar gráfico con D65
                UpdateChart(resD65);

                // Limpiar selección para evitar filas azules resaltadas al inicio
                dgvAnalysisRightTL84.ClearSelection();
                dgvAnalysisRightA.ClearSelection();
            }
        }

        private void FillAnalysisGridFromCmc(DataGridView dgv, CmcDifferenceRow cmc, double tolDE, bool isRecipe, double pctA = 0, double pctB = 0)
        {
            dgv.Rows.Clear();
            if (cmc == null) return;

            double dL = cmc.DeltaLightness * 10;
            double dC = cmc.DeltaChroma * 10;
            double dH = cmc.DeltaHue * 10;
            double dE = cmc.DeltaCMC ?? 0;

            if (dE > 0 && dE <= tolDE)
            {
                int i1 = dgv.Rows.Add("", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "LOTE APROBADO", "-");
                int i2 = dgv.Rows.Add("", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "-");
                int i3 = dgv.Rows.Add("", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "No requiere corrección", "-");
                ApplyEjeStyle(dgv, i1, "DL (Fuerza)"); ApplyTenueRowStyle(dgv, i1);
                ApplyEjeStyle(dgv, i2, "DC (Brillo)"); ApplyTenueRowStyle(dgv, i2);
                ApplyEjeStyle(dgv, i3, "DH (Matiz)"); ApplyTenueRowStyle(dgv, i3);
            }
            else
            {
                // Crear objeto de resultado temporal para usar lógica dinámica del motor
                var res = new ColorCorrectionResult {
                    DeltaL = cmc.DeltaLightness,
                    DeltaChroma = cmc.DeltaChroma,
                    DeltaHue = cmc.DeltaHue,
                    PercentL = cmc.DeltaLightness, 
                    DeltaA = pctA * 50, 
                    DeltaB = pctB * 50,
                    PercentA = pctA,
                    PercentB = pctB,
                    PercentChroma = cmc.DeltaChroma
                };

                string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
                string imp = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
                string rec = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

                int r1 = dgv.Rows.Add("", res.DeltaL.ToString("F2"), $"{res.PorcentajeRecetaL:F1}%", diag, imp, rec);
                int r2 = dgv.Rows.Add("", res.DeltaChroma.ToString("F2"), $"{Math.Abs(res.PercentChroma * 100):F1}%", res.DiagnosisC, res.DescripcionC, res.RecommendationC);
                int r3 = dgv.Rows.Add("", res.DeltaHue.ToString("F2"), $"{Math.Abs(res.PercentHue * 100):F1}%", res.DiagnosisH, res.ImpactoMatiz, res.RecomendacionMatiz);
                ApplyEjeStyle(dgv, r1, "DL (Fuerza)"); ApplyTenueRowStyle(dgv, r1);
                ApplyEjeStyle(dgv, r2, "DC (Brillo)"); ApplyTenueRowStyle(dgv, r2);
                ApplyEjeStyle(dgv, r3, "DH (Matiz)"); ApplyTenueRowStyle(dgv, r3);
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
                
                // Buscamos el análisis primario (D65) para los metadatos de eje
                var primaryAnalysis = _resultsLegacy?.FirstOrDefault(x => x.Illuminant == "D65") ?? 
                                     _resultsLegacy?.FirstOrDefault();
                
                double dlEje = primaryAnalysis?.DeltaL ?? 0;
                double dcEje = primaryAnalysis?.DeltaChroma ?? 0;
                double dhEje = primaryAnalysis?.DeltaHue ?? 0;
                string iluminante = primaryAnalysis?.Illuminant ?? "D65";

                // 2. Guardar en Base de Datos Unificada (Fila por componente)
                foreach (DataGridViewRow row in dgvCorrectiveRecipe.Rows)
                {
                    if (row.IsNewRow) continue;

                    string name = row.Cells[0].Value?.ToString() ?? "";
                    string strOriginal = row.Cells[1].Value?.ToString() ?? "0";
                    string strAdjDL = row.Cells[2].Value?.ToString() ?? "0";
                    string strAdjDH = row.Cells[3].Value?.ToString() ?? "0"; // Mapeado a Ajuste LD/DH
                    string strNueva = row.Cells[4].Value?.ToString() ?? "0";

                    // Conversión numérica estricta (Decimal 5 decimales)
                    decimal.TryParse(strOriginal.Replace("%",""), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal concOriginal);
                    decimal.TryParse(strAdjDL, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal adjDL);
                    decimal.TryParse(strAdjDH, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal adjDH);
                    decimal.TryParse(strNueva.Replace("%",""), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal nuevaReceta);

                    // Buscar el código en dgvShadeHistory
                    string code = "";
                    foreach (DataGridViewRow rShade in dgvShadeHistory.Rows) {
                        if (rShade.Cells[1].Value?.ToString() == name) {
                            code = rShade.Cells[0].Value?.ToString() ?? "";
                            break;
                        }
                    }

                    Color.Services.HistorialService.GuardarRegistroMaestro(
                        shadeName, fechaActual, iluminante,
                        dlEje, dcEje, dhEje,
                        code, name, 
                        concOriginal, adjDL, adjDH, nuevaReceta
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

                        System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                        MessageBox.Show("Reporte de texto guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Error al exportar: " + ex.Message); }
        }
    }
}