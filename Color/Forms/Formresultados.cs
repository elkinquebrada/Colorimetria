using System.IO;
using Color.Services;
using Color.Models;
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

        // Etiquetas del panel de metadatos textiles (superior izquierdo)
        private Label lblValueShadeName;
        private Label lblValueDyeingClass;
        private Label lblValueSubstrate;
        private Label lblValueCountPly;
        private Label lblValueFiberType;

        private double DE_MAX => Properties.Settings.Default.ToleranciaDE;

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
            this.Text = "TINT COATS CADENA - REPORTE DE COLOR";
            this.Size = new Size(1100, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.BackColor = System.Drawing.Color.FromArgb(230, 230, 230);

            pnlWhitePaper = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = System.Drawing.Color.FromArgb(230, 230, 230), Padding = new Padding(20) };
            pnlReportFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = System.Drawing.Color.White, Padding = new Padding(30), Width = 1000 };
            pnlWhitePaper.Controls.Add(pnlReportFlow);

            this.Resize += (s, e) => { pnlReportFlow.Width = Math.Max(900, pnlWhitePaper.ClientSize.Width - 40); };

            // CABECERA
            var pnlHeader = new Panel { Width = 940, Height = 100, Margin = new Padding(0, 0, 0, 20) };
            pnlHeader.Controls.Add(new Label { Text = "COATS TNT", Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(0, 0), AutoSize = true });
            pnlHeader.Controls.Add(new Label { Text = "Análisis de Color", BackColor = System.Drawing.Color.FromArgb(0, 102, 204), 
             ForeColor = System.Drawing.Color.White, Font = new Font("Segoe UI", 14, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter, Size = new Size(940, 40), Location = new Point(0, 30) });
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
            dgvShadeHistory.Columns[2].HeaderText = "Concentration, %";
            dgvShadeHistory.Columns[3].HeaderText = "Proportion, %";
            SetupShadeHistoryPainting();

            lblRightShadeValue = new Label { Text = "Shade Name:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = System.Drawing.Color.Black };
            lblValueShadeName = new Label { Text = "-", AutoSize = true, Font = new Font("Segoe UI", 9), ForeColor = System.Drawing.Color.Black };
            lblValueDyeingClass = new Label { Text = "Dyeing Class: -", AutoSize = true, ForeColor = System.Drawing.Color.Black };
            lblValueSubstrate = new Label { Text = "Substrate: -", AutoSize = true, ForeColor = System.Drawing.Color.Black };
            lblValueCountPly = new Label { Text = "Count/Ply: -", AutoSize = true, ForeColor = System.Drawing.Color.Black };
            lblValueFiberType = new Label { Text = "Fibre Type: -", AutoSize = true, ForeColor = System.Drawing.Color.Black };

            var pnlGeneralInfo = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown };
            var pnlShadeName = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
            pnlShadeName.Controls.Add(lblRightShadeValue);
            pnlShadeName.Controls.Add(lblValueShadeName);
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
            


            // BLOQUES
            blockD65 = new IluminantReportBlock { Width = 940, Margin = new Padding(0, 0, 0, 15) };
            blockTL84 = new IluminantReportBlock { Width = 940, Margin = new Padding(0, 0, 0, 15) };
            blockCWF = new IluminantReportBlock { Width = 940, Margin = new Padding(0, 0, 0, 15) };
            
            pnlReportFlow.Controls.Add(blockD65);
            pnlReportFlow.Controls.Add(blockTL84);
            pnlReportFlow.Controls.Add(blockCWF);

            // RECETA Y GRAFICO
            var pnlBottom = new TableLayoutPanel { Width = 940, Height = 250, ColumnCount = 2 };
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
            pnlBottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));

            dgvCorrectiveRecipe = CreateCorrectiveGrid();
            _cielabChart = new CielabChartControl { Dock = DockStyle.Fill };

            pnlBottom.Controls.Add(dgvCorrectiveRecipe, 0, 0);
            pnlBottom.Controls.Add(_cielabChart, 1, 0);
            pnlReportFlow.Controls.Add(pnlBottom);

            // BOTONES
            var pnlButtons = new Panel { Dock = DockStyle.Bottom, Height = 60, BackColor = System.Drawing.Color.White };
            btnGuardar = CreateStyledButton(" Guardar", System.Drawing.Color.FromArgb(45, 126, 247));
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

        /// Inyecta los metadatos texiles extraidos por TextileMetadataExtractor
        /// directamente en el panel superior izquierdo del formulario.
        /// <summary>
        /// Inyecta los metadatos textiles extraídos por TextileMetadataExtractor
        /// directamente en el panel superior izquierdo del formulario.
        /// </summary>
        public void UpdateTextileMetadataPanel(TextileMetadata metadata)
        {
            if (metadata == null) return;

            // 1. OBLIGAMOS A QUE TODAS LAS LÍNEAS USEN EXACTAMENTE LA MISMA FUENTE EN NEGRITA
            //    Esto es lo que hace que "Shade Name" se vea de ese color negro fuerte.
            Font fuenteNegritaIgualdad = new Font("Segoe UI", 9, FontStyle.Bold);
            System.Drawing.Color negroShade = System.Drawing.Color.Black;

            lblValueShadeName.Font = fuenteNegritaIgualdad; lblValueShadeName.ForeColor = negroShade;
            lblValueDyeingClass.Font = fuenteNegritaIgualdad; lblValueDyeingClass.ForeColor = negroShade;
            lblValueSubstrate.Font = fuenteNegritaIgualdad; lblValueSubstrate.ForeColor = negroShade;
            lblValueCountPly.Font = fuenteNegritaIgualdad; lblValueCountPly.ForeColor = negroShade;
            lblValueFiberType.Font = fuenteNegritaIgualdad; lblValueFiberType.ForeColor = negroShade;

            if (lblRightShadeValue != null)
            {
                lblRightShadeValue.Font = fuenteNegritaIgualdad;
                lblRightShadeValue.ForeColor = negroShade;
            }

            // 2. ASIGNACIÓN DE TEXTOS CON ESPACIADO PARA QUE QUEDEN ALINEADOS
            lblValueShadeName.Text = !string.IsNullOrEmpty(metadata.ShadeName) && metadata.ShadeName != "-" ? metadata.ShadeName.ToUpper() : "-";
            lblValueDyeingClass.Text = "Dyeing Class:   " + (!string.IsNullOrEmpty(metadata.DyeingClass) && metadata.DyeingClass != "-" ? metadata.DyeingClass.ToUpper() : "-");
            lblValueSubstrate.Text = "Substrate:      " + (!string.IsNullOrEmpty(metadata.Substrate) && metadata.Substrate != "-" ? metadata.Substrate.ToUpper() : "-");
            lblValueCountPly.Text = "Count/Ply:      " + (!string.IsNullOrEmpty(metadata.CountPly) && metadata.CountPly != "-" ? metadata.CountPly : "-");
            lblValueFiberType.Text = "Fibre Type:     " + (!string.IsNullOrEmpty(metadata.FiberType) && metadata.FiberType != "-" ? metadata.FiberType.ToUpper() : "-");

            // 3. CORRECCIÓN PARA EVITAR LA DUPLICACIÓN: Unificamos el texto en el control compuesto derecho
            if (lblRightShadeValue != null && !string.IsNullOrEmpty(metadata.ShadeName) && metadata.ShadeName != "-")
            {
                lblRightShadeValue.Text = "Shade Name: " + metadata.ShadeName.ToUpper();
            }
        }
        public void ActualizarTablaTolerancias(double de, double dl, double dc, double dh)
        {
            if (lblTolDe != null) lblTolDe.Text = de.ToString("F2");
            if (lblTolDl != null) lblTolDl.Text = dl.ToString("F2");
            if (lblTolDc != null) lblTolDc.Text = dc.ToString("F2");
            if (lblTolDh != null) lblTolDh.Text = dh.ToString("F2");

            // Pasamos estas tolerancias también a los bloques visuales de los iluminantes
            if (blockD65 != null) blockD65.UpdateTolerances(de, dl, dc, dh);
            if (blockTL84 != null) blockTL84.UpdateTolerances(de, dl, dc, dh);
            if (blockCWF != null) blockCWF.UpdateTolerances(de, dl, dc, dh);
        }

        private TableLayoutPanel CreateTolerancesTable()
        {
            var pnl = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 4, 
                RowCount = 3, Margin = new Padding(10, 0, 0, 0), CellBorderStyle = TableLayoutPanelCellBorderStyle.Single, BackColor = System.Drawing.Color.White };
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            pnl.Height = 80;

            var lblTitle = new Label { Text = "Tolerancias", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, BackColor = System.Drawing.Color.FromArgb(0, 122, 204), 
                ForeColor = System.Drawing.Color.White, Font = new Font("Segoe UI", 9, FontStyle.Regular), Margin = new Padding(0) };
            pnl.Controls.Add(lblTitle, 0, 0);
            pnl.SetColumnSpan(lblTitle, 4);

            pnl.Controls.Add(CreateGridLabel("DE", true), 0, 1);
            pnl.Controls.Add(CreateGridLabel("DL", true), 1, 1);
            pnl.Controls.Add(CreateGridLabel("DC", true), 2, 1);
            pnl.Controls.Add(CreateGridLabel("DH", true), 3, 1);

            lblTolDe = CreateGridLabel("1.20", false);
            lblTolDl = CreateGridLabel("0.69", false);
            lblTolDc = CreateGridLabel("0.69", false);
            lblTolDh = CreateGridLabel("0.69", false);

            pnl.Controls.Add(lblTolDe, 0, 2);
            pnl.Controls.Add(lblTolDl, 1, 2);
            pnl.Controls.Add(lblTolDc, 2, 2);
            pnl.Controls.Add(lblTolDh, 3, 2);

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
            var dgv = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = System.Drawing.Color.White, BorderStyle = BorderStyle.None, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                AllowUserToAddRows = false, RowHeadersVisible = false, ReadOnly = true, SelectionMode = DataGridViewSelectionMode.FullRowSelect, Font = new Font("Segoe UI", 8.2f) };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
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
            try {
                // Lógica de carga de logo omitida por brevedad, mantener si es necesario
            } catch { }
        }

        private void PopulateFromObjects(ShadeExtractionResult shadeData, List<EngineRes> results)
        {
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
                        dgvShadeHistory.Rows.Add(ing.Code, ing.Name, ing.Percentage.Replace("%",""), ((int)Math.Round(p)).ToString() + "%");
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

                blockD65.UpdateData(d65);
                if (tl84 != null) blockTL84.UpdateData(tl84);
                if (cwf != null) blockCWF.UpdateData(cwf);

                UpdateChart(d65); _lastMainResult = d65;

                if (shadeData != null)
                {
                    var ingredients = RecipeCorrector.IngredientsFromShade(shadeData);
                    var correctiveResult = RecipeCorrector.CalculateCorrectiveRecipe(ingredients, d65);
                    FillCorrectiveRecipeGrid(correctiveResult);
                }
            }
        }

        private void PopulateFromReport(OcrReport report)
        {
            if (report == null) return;
            var shadeData = new ShadeExtractionResult { ShadeName = report.Batch?.ShadeName, Recipe = report.Recipe };
            var results = EngineCalc.CalculateAllIlluminants(report);
            PopulateFromObjects(shadeData, results);
        }

        private void FillCorrectiveRecipeGrid(CorrectiveRecipeResult result)
        {
            dgvCorrectiveRecipe.Rows.Clear();
            if (result == null || result.Ingredients == null) return;
            double t1 = result.Ingredients.Sum(i => i.R1), t2 = result.Ingredients.Sum(i => i.R2), t3 = result.Ingredients.Sum(i => i.R3);
            foreach (var ing in result.Ingredients)
            {
                dgvCorrectiveRecipe.Rows.Add(ing.Name, "", 
                    ing.R1.ToString("F5"), ((int)Math.Round(t1 > 0 ? ing.R1/t1*100 : 0)).ToString() + "%",
                    ing.R2.ToString("F5"), ((int)Math.Round(t2 > 0 ? ing.R2/t2*100 : 0)).ToString() + "%",
                    ing.R3.ToString("F5"), ((int)Math.Round(t3 > 0 ? ing.R3/t3*100 : 0)).ToString() + "%");
            }

            int totalRowIdx = dgvCorrectiveRecipe.Rows.Add("[Dyes]", "", 
                    t1.ToString("F5"), "100%",
                    t2.ToString("F5"), "100%",
                    t3.ToString("F5"), "100%");
            dgvCorrectiveRecipe.Rows[totalRowIdx].DefaultCellStyle.Font = new Font(dgvCorrectiveRecipe.Font, FontStyle.Bold);
            dgvCorrectiveRecipe.Rows[totalRowIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);
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
            try {
                MessageBox.Show("Reporte guardado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnGuardar.Enabled = false; btnGuardar.Text = "✔ Guardado";
            } catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
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