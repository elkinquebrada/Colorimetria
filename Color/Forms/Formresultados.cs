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

        // ======= Controles de la vista (Tablas) =======
        private DataGridView dgvShadeHistory;
        private DataGridView dgvAnalysisLeft;
        private DataGridView dgvAnalysisLeftTL84;
        private DataGridView dgvAnalysisLeftA;
        private DataGridView dgvComparisonSummary;
        private DataGridView dgvAnalysisRight;
        private DataGridView dgvAnalysisRightTL84;
        private DataGridView dgvAnalysisRightA;

        private RichTextBox txtReport;
        private RichTextBox txtRecomendacion;
        private SplitContainer splitMedicionesCmc;
        private Button btnExportar;
        private Button btnHistorial;
        private Button btnCerrar;
        private Button btnRegresar;

        // ======= Gráfico CIELAB =======
        private Button btnVerGrafico;
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
            InitializeComponents();
            
            // Lógica silenciosa: Poblar desde el objeto Report directamente
            PopulateFromReport(_report);
        }

        public FormResultados(string resumen, List<EngineRes> results, List<Color.IlluminantCorrectionResult> recipeResults = null, ShadeExtractionResult shadeData = null)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = results ?? new List<EngineRes>();
            _recipeResults = recipeResults;
            InitializeComponents();

            // Lógica silenciosa: Poblar desde los objetos ya calculados
            PopulateFromObjects(shadeData, _resultsLegacy);
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
            btnRegresar = CreateStyledButton("← Regresar", System.Drawing.Color.FromArgb(180, 100, 0));
            btnRegresar.Click += (s, e) => {
                this.DialogResult = DialogResult.Retry;
                this.Close();
            };

            btnHistorial = CreateStyledButton("📜 Historial", System.Drawing.Color.FromArgb(34, 139, 34));
            btnHistorial.Click += BtnHistorial_Click;
            btnExportar = CreateStyledButton("💾 Exportar .txt", System.Drawing.Color.FromArgb(70, 130, 180));
            btnExportar.Click += BtnExportar_Click;
            btnCerrar = CreateStyledButton("Cerrar", System.Drawing.Color.FromArgb(200, 30, 30));
            btnCerrar.Click += (s, e) => this.Close();

            btnVerGrafico = new Button
            {
                Text = "🔍 Ver Gráfico Detallado",
                Size = new Size(180, 34),
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
            dgvShadeHistory.ColumnCount = 3;
            dgvShadeHistory.Columns[0].Name = "Codigo";
            dgvShadeHistory.Columns[1].Name = "nombre";
            dgvShadeHistory.Columns[2].Name = "porcentaje";

            dgvAnalysisLeft = CreateAnalysisGrid();
            dgvAnalysisLeftTL84 = CreateAnalysisGrid();
            dgvAnalysisLeftA = CreateAnalysisGrid();

            dgvComparisonSummary = CreateStyledGrid();
            dgvComparisonSummary.ColumnCount = 4;
            dgvComparisonSummary.Columns[0].Name = "Dato";
            dgvComparisonSummary.Columns[1].Name = "Tolerancia";
            dgvComparisonSummary.Columns[2].Name = "Iluminante";
            dgvComparisonSummary.Columns[3].Name = "Resultado";

            dgvAnalysisRight = CreateAnalysisGrid();
            dgvAnalysisRightTL84 = CreateAnalysisGrid();
            dgvAnalysisRightA = CreateAnalysisGrid();

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

            var pnlLeft = CreatePanelWithGrids("ANALISIS DE SHADE HISTORY REPORT", dgvShadeHistory, 
                                               "ANALISIS ILUMINANTE D65", dgvAnalysisLeft);

            var pnlRight = CreatePanelWithManyGrids("ANALISIS DE SAMPLE COMPARISON", dgvComparisonSummary, 
                                                   "ANALISIS ILUMINANTE D65", dgvAnalysisRight,
                                                   "ANALISIS ILUMINANTE TL84", dgvAnalysisRightTL84,
                                                   "ANALISIS ILUMINANTE A / CWF", dgvAnalysisRightA);

            ApplyTranslucentStyle(dgvAnalysisRightTL84);
            ApplyTranslucentStyle(dgvAnalysisRightA);
            
            var pnlBtnGrafico = new Panel { Dock = DockStyle.Bottom, Height = 45 };
            pnlBtnGrafico.Controls.Add(btnVerGrafico);
            btnVerGrafico.Location = new Point(10, 5);
            pnlRight.Controls.Add(pnlBtnGrafico);

            splitMedicionesCmc.Panel1.Controls.Add(pnlLeft);
            splitMedicionesCmc.Panel2.Controls.Add(pnlRight);

            var pnlBottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(10)
            };
            pnlBottom.Controls.Add(btnCerrar);
            pnlBottom.Controls.Add(btnExportar);
            pnlBottom.Controls.Add(btnHistorial);
            pnlBottom.Controls.Add(btnRegresar);

            this.Controls.Add(splitMedicionesCmc);
            this.Controls.Add(lblTitulo);
            this.Controls.Add(pnlBottom);
        }

        private Panel CreatePanelWithManyGrids(string h1, DataGridView g1, string h2, DataGridView g2, string h3, DataGridView g3, string h4, DataGridView g4)
        {
            var pnl = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 8 };
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); // Header 1
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 25)); // Grid 1
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 2 (D65)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 35)); // Grid 2 (D65)
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 3 (TL84)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 20)); // Grid 3 (TL84)
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 4 (A)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 20)); // Grid 4 (A)

            pnl.Controls.Add(CreateHeaderLabel(h1), 0, 0);
            pnl.Controls.Add(g1, 0, 1);
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
                ScrollBars = ScrollBars.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
                DefaultCellStyle = new DataGridViewCellStyle { WrapMode = DataGridViewTriState.True }
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 5;
            dgv.Columns[0].Name = "VARIACION";
            dgv.Columns[1].Name = "%";
            dgv.Columns[2].Name = "DIAGNOSTICO";
            dgv.Columns[3].Name = "IMPACTO";
            dgv.Columns[4].Name = "RECOMENDACION";
            return dgv;
        }

        private void ApplyTranslucentStyle(DataGridView dgv)
        {
            var faintColor = System.Drawing.Color.FromArgb(180, 180, 180);
            dgv.DefaultCellStyle.ForeColor = faintColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = faintColor;
            dgv.GridColor = System.Drawing.Color.FromArgb(245, 245, 245);
            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(250, 250, 250);
            dgv.DefaultCellStyle.SelectionForeColor = faintColor;
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
            var backColor = tenue ? System.Drawing.Color.FromArgb(235, 240, 245) : System.Drawing.Color.FromArgb(0, 102, 204);
            var foreColor = tenue ? System.Drawing.Color.FromArgb(160, 170, 180) : System.Drawing.Color.White;
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

        private void PopulateFromObjects(ShadeExtractionResult shadeData, List<EngineRes> results)
        {
            if (shadeData != null)
            {
                dgvShadeHistory.Rows.Clear();
                dgvShadeHistory.Rows.Add("Shade Name", shadeData.ShadeName ?? "N/A", "");
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
                dgvComparisonSummary.Rows.Add("Shade Name", shadeName, "", "");
                
                // --- Cuadro de Tolerancia CMC Estándar (Formato Profesional) ---
                string tolStr = $"DE {DE_MAX:F2}  |  DL {DL_MAX:F3}  |  DC {DC_MAX:F3}  |  DH {DH_MAX:F3}";
                int tolIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolStr, "", "");
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(230, 240, 250);
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(0, 51, 102);
                
                string resDL = Math.Abs(d65.DeltaL) <= DL_MAX ? "CUMPLE" : "NO CUMPLE";
                string resDC = (ill2 != null) ? (Math.Abs(ill2.DeltaChroma) <= DC_MAX ? "CUMPLE" : "NO CUMPLE") : "N/A";
                string resDH = (ill3 != null) ? (Math.Abs(ill3.DeltaHue) <= DH_MAX ? "CUMPLE" : "NO CUMPLE") : "N/A";

                dgvComparisonSummary.Rows.Add("DL", DL_MAX.ToString("F3"), d65.Illuminant, resDL);
                dgvComparisonSummary.Rows.Add("DC", DC_MAX.ToString("F3"), ill2?.Illuminant ?? "TL84", resDC);
                dgvComparisonSummary.Rows.Add("DH", DH_MAX.ToString("F3"), ill3?.Illuminant ?? "A", resDH);
                
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
                dgv.Rows.Add("DL", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "LOTE APROBADO");
                dgv.Rows.Add("DC", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
                dgv.Rows.Add("DH", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
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

                dgv.Rows.Add("DL", res.DeltaL.ToString("F2"), "Carga Colorante", res.DescripcionL, res.RecomendacionL);
                dgv.Rows.Add("DC", res.DeltaChroma.ToString("F2"), "Saturación", res.DiagnosisC, res.RecommendationC);
                dgv.Rows.Add("DH", res.DeltaHue.ToString("F2"), "Matiz/Hue", res.ImpactoMatiz, res.RecomendacionMatiz);
            }
        }

        private void FillAnalysisGrid(DataGridView dgv, EngineRes res, bool isRecipe)
        {
            dgv.Rows.Clear();
            if (res == null) return;

            if (res.CmcValue <= DE_MAX || res.DeltaE <= DE_MAX)
            {
                dgv.Rows.Add("DL", (res.DeltaL * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "LOTE APROBADO");
                dgv.Rows.Add("DC", (res.DeltaChroma * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
                dgv.Rows.Add("DH", (res.DeltaHue * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
            }
            else
            {
                if (isRecipe)
                {
                    dgv.Rows.Add("DL", res.DeltaL.ToString("F2"), "Carga Colorante", res.DescripcionL, res.RecomendacionL);
                    dgv.Rows.Add("DC", res.DeltaChroma.ToString("F2"), "Saturación", res.DiagnosisC, res.RecommendationC);
                    dgv.Rows.Add("DH", res.DeltaHue.ToString("F2"), "Matiz/Hue", res.ImpactoMatiz, res.RecomendacionMatiz);
                }
                else
                {
                    dgv.Rows.Add("DL", res.DeltaL.ToString("F2"), "Tiempo/Temperatura", res.DescripcionL, res.RecomendacionL);
                    dgv.Rows.Add("DC", res.DeltaChroma.ToString("F2"), "Agotamiento químico", res.DiagnosisC, res.RecommendationC);
                    dgv.Rows.Add("DH", res.DeltaHue.ToString("F2"), "Desvío de matiz", res.ImpactoMatiz, res.RecomendacionMatiz);
                }
            }
        }

        // --- PANEL IZQUIERDO: RECETA / CONCENTRACIÓN ---
        private string GetDiagLRecipe(double dL) => dL < 0 ? "Exceso de concentración en formulación" : dL > 0 ? "Falta de carga colorante en receta" : "Concentración correcta";
        private string GetImpLRecipe(double dL) => dL < 0 ? "Oscurecimiento por sobre-concentración" : dL > 0 ? "Aclaramiento por falta de tinte" : "-";
        private string GetInstLRecipe_ConCifra(double dL, double varL) => dL < 0 ? $"REDUCIR % TOTAL RECETA EN {Math.Abs(varL):F1}%" : $"AUMENTAR % TOTAL RECETA EN {Math.Abs(varL):F1}%";

        // --- PANEL DERECHO: LOTE / PROCESO ---
        private string GetDiagLLot(double dL) => dL < 0 ? "Lote Oscuro (Sobre-teñido)" : dL > 0 ? "Lote Claro (Bajo agotamiento)" : "Luminosidad OK";
        private string GetImpLLot(double dL) => dL < 0 ? "Revisar curva de temperatura y agotamiento" : dL > 0 ? "Verificar fijación, pH y relación de baño" : "-";
        private string GetInstLLot_ConCifra(double dL, double pctL) => dL < 0 ? $"REVISAR CURVA: Lote oscuro, corregir intensidad en {Math.Abs(pctL):F1}%" : $"VERIFICAR PH: Lote claro, aumentar carga en {Math.Abs(pctL):F1}%";

        private string GetInstC(double dC) => dC < -1 ? "AUMENTAR FUERZA / CONC." : dC > 1 ? "REDUCIR CARGA / DILUIR" : "OK";
        private string GetInstH(double dH) => dH < -1 ? "AUMENTAR ROJO / DISM. VERDE" : dH > 1 ? "AUMENTAR AZUL / DISM. AMAR." : "OK";

        private string GetDiagC(double dC) => dC < 0 ? "Más Opaco / Sucio" : dC > 0 ? "Más Vivo / Saturado" : "Saturación OK";
        private string GetDiagH(double dH) => dH < 0 ? "Desviación hacia tonos AZULES" : dH > 0 ? "Desviación hacia tonos AMARILLOS" : "Matiz OK";
        private string GetImpH(double dH) => Math.Abs(dH) > 1 ? "Cambio tonal perceptible" : "Normal";

        private string GetCombinedAxesText(double pctARatio, double pctBRatio)
        {
            double pA = pctARatio * 100.0;
            double pB = pctBRatio * 100.0;
            string partA = Math.Abs(pA) > 0.05 ? (pA > 0 ? $"Rojo {Math.Abs(pA):F1}%" : $"Verde {Math.Abs(pA):F1}%") : "";
            string partB = Math.Abs(pB) > 0.05 ? (pB > 0 ? $"Amarillo {Math.Abs(pB):F1}%" : $"Azul {Math.Abs(pB):F1}%") : "";
            if (!string.IsNullOrEmpty(partA) && !string.IsNullOrEmpty(partB)) return $"{partA} / {partB}";
            return string.IsNullOrEmpty(partA) ? (string.IsNullOrEmpty(partB) ? "Matiz OK" : partB) : partA;
        }

        private string GetCombinedImpactText(double pctARatio, double pctBRatio)
        {
            double pA = pctARatio * 100.0;
            double pB = pctBRatio * 100.0;
            string partA = Math.Abs(pA) > 0.05 ? (pA > 0 ? "Más Rojizo" : "Más Verdoso") : "";
            string partB = Math.Abs(pB) > 0.05 ? (pB > 0 ? "Más Amarillento" : "Más Azulado") : "";
            if (!string.IsNullOrEmpty(partA) && !string.IsNullOrEmpty(partB)) return $"{partA} / {partB}";
            return string.IsNullOrEmpty(partA) ? (string.IsNullOrEmpty(partB) ? "Normal" : partB) : partA;
        }

        private string GetCombinedInstText(double pctARatio, double pctBRatio, bool isRecipe)
        {
            double pA = pctARatio * 100.0;
            double pB = pctBRatio * 100.0;
            string suffix = isRecipe ? "" : " en baño";
            
            string instA = Math.Abs(pA) > 0.05 ? (pA > 0 ? $"Bajar Rojo {Math.Abs(pA):F1}%{suffix}" : $"Subir Rojo {Math.Abs(pA):F1}%{suffix}") : "";
            string instB = Math.Abs(pB) > 0.05 ? (pB > 0 ? $"Bajar Amarillo {Math.Abs(pB):F1}%{suffix}" : $"Subir Amarillo {Math.Abs(pB):F1}%{suffix}") : "";
            
            if (!string.IsNullOrEmpty(instA) && !string.IsNullOrEmpty(instB)) return $"{instA} / {instB}";
            return string.IsNullOrEmpty(instA) ? (string.IsNullOrEmpty(instB) ? "OK" : instB) : instA;
        }

        private void PopulateFromReport(OcrReport report)
        {
            if (report == null) return;

            // Buscamos los iluminantes en las medidas
            var d65 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("D65"));
            var tl84 = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("TL84"));
            var illA = report.CmcDifferences.FirstOrDefault(c => c.Illuminant.Contains("A") || c.Illuminant.Contains("CWF"));

            dgvComparisonSummary.Rows.Clear();
            dgvComparisonSummary.Rows.Add("Shade Name", report.Batch?.ShadeName ?? "N/A", "", "");
            
            // --- Cuadro de Tolerancia CMC Estándar (Formato Profesional) ---
            string tolStr = $"DE {report.TolDE:F2}  |  DL {report.TolDL:F3}  |  DC {report.TolDC:F3}  |  DH {report.TolDH:F3}";
            int tIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolStr, "", "");
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(230, 240, 250);
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(0, 51, 102);

            if (d65 != null)
            {
                string resDL = Math.Abs(d65.DeltaLightness) <= report.TolDL ? "CUMPLE" : "NO CUMPLE";
                dgvComparisonSummary.Rows.Add("DL", report.TolDL.ToString("F3"), d65.Illuminant, resDL);
                
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
                string resDC = Math.Abs(tl84.DeltaChroma) <= report.TolDC ? "CUMPLE" : "NO CUMPLE";
                dgvComparisonSummary.Rows.Add("DC", report.TolDC.ToString("F3"), tl84.Illuminant, resDC);

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
                string resDH = Math.Abs(illA.DeltaHue) <= report.TolDH ? "CUMPLE" : "NO CUMPLE";
                dgvComparisonSummary.Rows.Add("DH", report.TolDH.ToString("F3"), illA.Illuminant, resDH);

                var std = report.Measures.FirstOrDefault(m => (m.Illuminant.Contains("A") || m.Illuminant.Contains("CWF")) && m.Type.ToUpper().Contains("STD"));
                var lot = report.Measures.FirstOrDefault(m => (m.Illuminant.Contains("A") || m.Illuminant.Contains("CWF")) && (m.Type.ToUpper().Contains("SPL") || m.Type.ToUpper().Contains("LOT")));
                double pA = 0, pB = 0;
                if (std != null && lot != null) {
                    pA = (Math.Abs(std.A) > 0.1) ? (lot.A - std.A) / Math.Abs(std.A) : 0;
                    pB = (Math.Abs(std.B) > 0.1) ? (lot.B - std.B) / Math.Abs(std.B) : 0;
                }
                FillAnalysisGridFromCmc(dgvAnalysisRightA, illA, report.TolDE, false, pA, pB);
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
                dgv.Rows.Add("DL", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "LOTE APROBADO");
                dgv.Rows.Add("DC", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "No requiere corrección");
                dgv.Rows.Add("DH", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "No requiere corrección");
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

                if (isRecipe)
                {
                    dgv.Rows.Add("DL", res.DeltaL.ToString("F2"), "Carga Colorante", res.DescripcionL, res.RecomendacionL);
                    dgv.Rows.Add("DC", res.DeltaChroma.ToString("F2"), "Saturación", res.DiagnosisC, res.RecommendationC);
                    dgv.Rows.Add("DH", res.DeltaHue.ToString("F2"), "Matiz/Hue", res.ImpactoMatiz, res.RecomendacionMatiz);
                }
                else
                {
                    dgv.Rows.Add("DL", res.DeltaL.ToString("F2"), "Tiempo/Temperatura", res.DescripcionL, res.RecomendacionL);
                    dgv.Rows.Add("DC", res.DeltaChroma.ToString("F2"), "Agotamiento químico", res.DiagnosisC, res.RecommendationC);
                    dgv.Rows.Add("DH", res.DeltaHue.ToString("F2"), "Desvío de matiz", res.ImpactoMatiz, res.RecomendacionMatiz);
                }
            }
        }

        private void BtnExportar_Click(object sender, EventArgs e)
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
                        MessageBox.Show("Archivo de texto guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnHistorial_Click(object sender, EventArgs e)
        {
            try
            {
                var frm = new Color.Forms.FormHistorial();
                var dt = Color.Services.HistorialService.ObtenerHistorial();
                frm.CargarHistorial(dt);
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al abrir el historial: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}