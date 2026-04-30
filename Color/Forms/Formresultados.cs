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
        private DataGridView dgvCorrectiveRecipe;
        private Label lblAlertCorrective;

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

            var pnlCorrective = new Panel { Dock = DockStyle.Bottom, Height = 220 };
            dgvCorrectiveRecipe = CreateCorrectiveGrid();
            lblAlertCorrective = new Label { 
                Dock = DockStyle.Bottom, 
                Height = 35, 
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.Gray
            };
            pnlCorrective.Controls.Add(dgvCorrectiveRecipe);
            pnlCorrective.Controls.Add(lblAlertCorrective);
            pnlCorrective.Controls.Add(CreateHeaderLabel("RESUMEN DE FORMULACIÓN CORRECTIVA (D65)"));

            var pnlLeft = CreatePanelWithGrids("ANALISIS DE SHADE HISTORY REPORT", dgvShadeHistory, 
                                               "ANALISIS ILUMINANTE D65", dgvAnalysisLeft);
            pnlLeft.Controls.Add(pnlCorrective);
            pnlCorrective.BringToFront();
            pnlCorrective.Dock = DockStyle.Bottom;

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
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 135)); // Grid 1 (Summary) - Ampliado para 5 filas
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 2 (D65)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33)); // Grid 2 (D65 - Principal)
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 3 (TL84)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33)); // Grid 3 (TL84)
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25)); // Header 4 (A)
            pnl.RowStyles.Add(new RowStyle(SizeType.Percent, 33)); // Grid 4 (A)

            pnl.Controls.Add(CreateHeaderLabel(h1), 0, 0);
            pnl.Controls.Add(g1, 0, 1);
            
            // Ajustar pesos de columnas para Summary
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
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            return dgv;
        }

        private DataGridView CreateCorrectiveGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "Nombre";         dgv.Columns[0].FillWeight = 25;
            dgv.Columns[1].Name = "% Original";     dgv.Columns[1].FillWeight = 12;
            dgv.Columns[2].Name = "Ajuste DL";      dgv.Columns[2].FillWeight = 12;
            dgv.Columns[3].Name = "Ajuste DH";      dgv.Columns[3].FillWeight = 12;
            dgv.Columns[4].Name = "% Nueva Receta"; dgv.Columns[4].FillWeight = 18;
            dgv.Columns[5].Name = "Status";         dgv.Columns[5].FillWeight = 21;
            return dgv;
        }

        private DataGridView CreateAnalysisGrid()
        {
            var dgv = CreateStyledGrid();
            dgv.ColumnCount = 6;
            dgv.Columns[0].Name = "EJE";          dgv.Columns[0].FillWeight = 10;
            dgv.Columns[1].Name = "VARIACION";    dgv.Columns[1].FillWeight = 12;
            dgv.Columns[2].Name = "%";            dgv.Columns[2].FillWeight = 10;
            dgv.Columns[3].Name = "DIAGNOSTICO";   dgv.Columns[3].FillWeight = 25;
            dgv.Columns[4].Name = "IMPACTO";      dgv.Columns[4].FillWeight = 18;
            dgv.Columns[5].Name = "RECOMENDACION"; dgv.Columns[5].FillWeight = 25;
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
                string tolSummary = $"DE {DE_MAX:F2}";
                int tolIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolSummary, "", "");
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(230, 240, 250);
                dgvComparisonSummary.Rows[tolIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(0, 51, 102);

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
                int i1 = dgv.Rows.Add("", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "LOTE APROBADO");
                int i2 = dgv.Rows.Add("", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
                int i3 = dgv.Rows.Add("", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
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

                int r1 = dgv.Rows.Add("", res.DeltaL.ToString("F2"), $"{res.PorcentajeRecetaL:F1}%", res.DiagnosticoL, res.ImpactoRecetaL, res.RecomendacionRecetaL);
                int r2 = dgv.Rows.Add("", res.DeltaChroma.ToString("F2"), $"{Math.Abs(res.PercentChroma * 100):F1}%", res.DiagnosisC, res.DescripcionC, res.RecommendationC);
                int r3 = dgv.Rows.Add("", res.DeltaHue.ToString("F2"), $"{Math.Abs(res.PercentHue * 100):F1}%", res.DiagnosisH, res.ImpactoMatiz, res.RecomendacionMatiz);
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

            if (eje.StartsWith("DL"))
            {
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(45, 45, 45); // Dark Gray
            }
            else if (eje.StartsWith("DC"))
            {
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100); // Medium Gray
            }
            else if (eje.StartsWith("DH"))
            {
                cell.Style.ForeColor = System.Drawing.Color.FromArgb(180, 0, 0); // Dark Red
            }
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
                    ing.NewConcentration.ToString("F5"),
                    ing.Status
                );

                if (ing.Status == "SATURACIÓN")
                {
                    dgvCorrectiveRecipe.Rows[idx].Cells[5].Style.BackColor = System.Drawing.Color.MistyRose;
                    dgvCorrectiveRecipe.Rows[idx].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
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
                int i1 = dgv.Rows.Add("", (res.DeltaL * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "LOTE APROBADO");
                int i2 = dgv.Rows.Add("", (res.DeltaChroma * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
                int i3 = dgv.Rows.Add("", (res.DeltaHue * 10).ToString("F1") + "%", "DENTRO DE TOLERANCIA", "Normal", "No requiere corrección");
                ApplyEjeStyle(dgv, i1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, i2, "DC (Brillo)");
                ApplyEjeStyle(dgv, i3, "DH (Matiz)");
            }
            else
            {
                string diag = isRecipe ? res.DiagnosticoL : res.DiagnosticoLoteL;
                string imp = isRecipe ? res.ImpactoRecetaL : res.ImpactoLoteL;
                string rec = isRecipe ? res.RecomendacionRecetaL : res.RecomendacionLoteL;

                int r1 = dgv.Rows.Add("", res.DeltaL.ToString("F2"), $"{res.PorcentajeRecetaL:F1}%", diag, imp, rec);
                int r2 = dgv.Rows.Add("", res.DeltaChroma.ToString("F2"), $"{Math.Abs(res.PercentChroma * 100):F1}%", res.DiagnosisC, res.DescripcionC, res.RecommendationC);
                int r3 = dgv.Rows.Add("", res.DeltaHue.ToString("F2"), $"{Math.Abs(res.PercentHue * 100):F1}%", res.DiagnosisH, res.ImpactoMatiz, res.RecomendacionMatiz);
                ApplyEjeStyle(dgv, r1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, r2, "DC (Brillo)");
                ApplyEjeStyle(dgv, r3, "DH (Matiz)");
            }
        }

        // --- HELPERS DE MATIZ ---

        private void PopulateFromReport(OcrReport report)
        {
            if (report == null) return;

            dgvShadeHistory.Rows.Clear();
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
            dgvComparisonSummary.Rows.Add("Shade Name", report.Batch?.ShadeName ?? "N/A", "", "");
            
            // --- Cuadro de Tolerancia CMC Estándar (Formato Profesional) ---
            string tolSummary = $"DE {report.TolDE:F2}";
            int tIdx = dgvComparisonSummary.Rows.Add("Tolerancia CMC", tolSummary, "", "");
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(230, 240, 250);
            dgvComparisonSummary.Rows[tIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(0, 51, 102);

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
                int i1 = dgv.Rows.Add("", dL.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "LOTE APROBADO");
                int i2 = dgv.Rows.Add("", dC.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "No requiere corrección");
                int i3 = dgv.Rows.Add("", dH.ToString("F1") + "%", "DENTRO DE TOLERANCIA", "-", "No requiere corrección");
                ApplyEjeStyle(dgv, i1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, i2, "DC (Brillo)");
                ApplyEjeStyle(dgv, i3, "DH (Matiz)");
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
                ApplyEjeStyle(dgv, r1, "DL (Fuerza)");
                ApplyEjeStyle(dgv, r2, "DC (Brillo)");
                ApplyEjeStyle(dgv, r3, "DH (Matiz)");
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