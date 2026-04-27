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

        // ======= Datos OCR retenidos para escritura automática al Excel =======
        private List<EngineRow> _ocrRows;      // L/A/B por iluminante del OCR
        private List<CmcResult> _cmcResults;   // CMC(2:1) calculado

        // ======= Controles de la vista =======
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

        // ======= Tolerancias (L*, Hue y ΔE) =======
        private double DL_MAX => Properties.Settings.Default.ToleranciaDL;
        private double DC_MAX => Properties.Settings.Default.ToleranciaDC;
        private double DH_MAX => Properties.Settings.Default.ToleranciaDH;
        private double DE_MAX => Properties.Settings.Default.ToleranciaDE;

        // ======= Proporción del divisor (55% izquierda) =======
        private double _splitLeftRatio = 0.55;

        // ======= Constructores =======
        public FormResultados(OcrReport report)
        {
            _report = report ?? new OcrReport();
            InitializeComponents();

            // Poblar vistas
            var resumen = BuildResumenFromReport(_report);
            var recom = BuildRecomendacionFromReport(_report);
            SetFormattedText(txtReport, resumen);
            SetFormattedText(txtRecomendacion, recom);

            // Retener filas OCR y calcular CMC para escritura automática al Excel
            _ocrRows = _report.Measures?.Select(m => new EngineRow
            {
                Illuminant = m.Illuminant, Type = m.Type,
                L = m.L, A = m.A, B = m.B,
                Chroma = m.Chroma > 0 ? m.Chroma : Math.Sqrt(m.A * m.A + m.B * m.B),
                Hue = m.Hue > 0 ? m.Hue : (Math.Atan2(m.B, m.A) * 180.0 / Math.PI + (Math.Atan2(m.B, m.A) < 0 ? 360.0 : 0.0))
            }).ToList();

            if (_ocrRows != null && _ocrRows.Count > 0)
            {
                var corr = EngineCalc.Calculate(_ocrRows);
                _cmcResults = EngineCalc.CalculateCmc(corr, _ocrRows);
            }

            // Actualizar gráfico con D65 del reporte
            UpdateChartFromReport(_report);

            // ── ESCRITURA AUTOMÁTICA AL EXCEL DE REFERENCIA ──
            WriteExcelSilently();
        }

        public FormResultados(string resumen, List<EngineRes> results)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = results ?? new List<EngineRes>();
            InitializeComponents();

            // Extraer metadata global (puente)
            string prefix = GetGlobalMetadataPrefix();

            // Poblar vistas
            SetFormattedText(txtReport, string.IsNullOrEmpty(prefix) ? _resumenLegacy : prefix + _resumenLegacy);
            SetFormattedText(txtRecomendacion, BuildRecomendacionFromResults(_resultsLegacy, DL_MAX, DC_MAX, DH_MAX, DE_MAX));

            // Retener filas OCR desde el reporte global y calcular CMC
            _ocrRows = OcrReport.LastReport?.Measures?.Select(m => new EngineRow
            {
                Illuminant = m.Illuminant, Type = m.Type,
                L = m.L, A = m.A, B = m.B,
                Chroma = m.Chroma > 0 ? m.Chroma : Math.Sqrt(m.A * m.A + m.B * m.B),
                Hue = m.Hue > 0 ? m.Hue : (Math.Atan2(m.B, m.A) * 180.0 / Math.PI + (Math.Atan2(m.B, m.A) < 0 ? 360.0 : 0.0))
            }).ToList();

            _cmcResults = _resultsLegacy?.Count > 0 && _ocrRows?.Count > 0
                ? EngineCalc.CalculateCmc(_resultsLegacy, _ocrRows)
                : null;

            // Actualizar gráfico
            UpdateChartFromResults(_resultsLegacy);

            // ── ESCRITURA AUTOMÁTICA AL EXCEL DE REFERENCIA ──
            WriteExcelSilently();
        }

        // Constructor de 3 argumentos — resumen + correcciones colorimetricas + correcciones de receta.
        public FormResultados(string resumen, List<EngineRes> corrections, List<Color.IlluminantCorrectionResult> recipeCorrections)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = corrections ?? new List<EngineRes>();
            _recipeResults = recipeCorrections;
            InitializeComponents();

            // Extraer metadata global (puente)
            string prefix = GetGlobalMetadataPrefix();

            // Panel izquierdo: receta + corrección de receta por iluminante
            var sbLeft = new System.Text.StringBuilder();
            if (!string.IsNullOrEmpty(prefix)) sbLeft.Append(prefix);
            sbLeft.Append(_resumenLegacy);
            if (_resultsLegacy != null && _resultsLegacy.Count > 0)
            {
                var d65Res = _resultsLegacy.FirstOrDefault(r => string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase)) ?? _resultsLegacy[0];
                Color.IlluminantCorrectionResult recipeD65 = null;
                if (_recipeResults != null)
                {
                    recipeD65 = _recipeResults.FirstOrDefault(r => string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase)) ?? _recipeResults.FirstOrDefault();
                }
                
                sbLeft.AppendLine();
                sbLeft.Append(BuildExecutiveSummary(d65Res, recipeD65));
            }
            SetFormattedText(txtReport, sbLeft.ToString());

            SetFormattedText(txtRecomendacion, BuildRecomendacionFromResults(_resultsLegacy, DL_MAX, DC_MAX, DH_MAX, DE_MAX));

            // Retener filas OCR desde el reporte global + calcular CMC
            _ocrRows = OcrReport.LastReport?.Measures?.Select(m => new EngineRow
            {
                Illuminant = m.Illuminant, Type = m.Type,
                L = m.L, A = m.A, B = m.B,
                Chroma = m.Chroma > 0 ? m.Chroma : Math.Sqrt(m.A * m.A + m.B * m.B),
                Hue = m.Hue > 0 ? m.Hue : (Math.Atan2(m.B, m.A) * 180.0 / Math.PI + (Math.Atan2(m.B, m.A) < 0 ? 360.0 : 0.0))
            }).ToList();

            _cmcResults = _resultsLegacy?.Count > 0 && _ocrRows?.Count > 0
                ? EngineCalc.CalculateCmc(_resultsLegacy, _ocrRows)
                : null;

            // Actualizar gráfico
            UpdateChartFromResults(_resultsLegacy);

            // ── ESCRITURA AUTOMÁTICA AL EXCEL DE REFERENCIA ──
            WriteExcelSilently();
        }

        private string BuildExecutiveSummary(EngineRes d65Res, Color.IlluminantCorrectionResult recipeRes)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("============================================================");
            sb.AppendLine($"       ANALIASIS DE LA RECDETA - ILUMINANTE {d65Res.Illuminant}");
            sb.AppendLine("============================================================");
            sb.AppendLine();

            // 1. LIGHTNESS
            double dLVar = d65Res.DeltaL * 10.0;
            string diagL = dLVar < 0 ? "El lote está más OSCURO que el estándar." 
                         : dLVar > 0 ? "El lote está más CLARO que el estándar." 
                         : "La luminosidad coincide con el estándar.";
            string impL = dLVar < 0 ? "Exceso de concentración de colorantes en la mezcla." 
                        : dLVar > 0 ? "Falta de concentración de colorantes en la mezcla." 
                        : "Concentración balanceada.";

            sb.AppendLine("[1. ANÁLISIS DE LUMINOSIDAD (LIGHTNESS)]");
            sb.AppendLine($"Variación (dL): {dLVar:F1}%");
            sb.AppendLine($"Diagnóstico: {diagL}");
            sb.AppendLine($"Impacto: {impL}");
            sb.AppendLine();

            // 2. CHROMA
            double dCVar = d65Res.DeltaChroma * 10.0;
            string diagC = dCVar < 0 ? "El color está notablemente más OPACO/GRIS." 
                         : dCVar > 0 ? "El color está notablemente más VIVO/SATURADO." 
                         : "La saturación coincide con el estándar.";
            string impC = dCVar < 0 ? "Pérdida de viveza; la mezcla está \"sucia\" o mal balanceada." 
                        : dCVar > 0 ? "Exceso de saturación." 
                        : "Saturación balanceada.";

            sb.AppendLine("[2. ANÁLISIS DE SATURACIÓN (CHROMA)]");
            sb.AppendLine($"Variación (dC): {dCVar:F1}%");
            sb.AppendLine($"Diagnóstico: {diagC}");
            sb.AppendLine($"Impacto: {impC}");
            sb.AppendLine();

            // 3. HUE
            double stdHue = d65Res.StdHue;
            double lotHue = d65Res.LotHue;
            double dHueVar = d65Res.PercentHue; 
            string hueShiftDir = "";
            if (d65Res.DeltaB < -0.05) hueShiftDir = "AZUL";
            else if (d65Res.DeltaB > 0.05) hueShiftDir = "AMARILLO";
            else if (d65Res.DeltaA > 0.05) hueShiftDir = "ROJO";
            else if (d65Res.DeltaA < -0.05) hueShiftDir = "VERDE";

            if (string.IsNullOrEmpty(hueShiftDir)) hueShiftDir = "NEUTRO";

            string sign = d65Res.DeltaHue >= 0 ? "+" : "";
            sb.AppendLine("[3. ANÁLISIS DE MATIZ (HUE)]");
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture, "Ángulos: Std: {0:F2}°  |  Lote: {1:F2}°", stdHue, lotHue));
            sb.AppendLine($"Variación: {dHueVar:F1}%");
            sb.AppendLine($"Diagnóstico: Desviación ({sign}{d65Res.DeltaHue:F2}°) hacia el {hueShiftDir}.");
            sb.AppendLine();

            sb.AppendLine("============================================");
            sb.AppendLine("[RECOMENDACIÓN]");
            sb.AppendLine("============================================");
            sb.AppendLine();

            int recNum = 1;

            if (recipeRes != null && recipeRes.Ingredients != null && recipeRes.Ingredients.Count > 0 && hueShiftDir != "NEUTRO")
            {
                string culpritDye = "";
                if (hueShiftDir == "AZUL") culpritDye = recipeRes.Ingredients.FirstOrDefault(i => i.Name.ToUpper().Contains("BLUE") || i.Name.ToUpper().Contains("NAVY"))?.Name;
                else if (hueShiftDir == "ROJO") culpritDye = recipeRes.Ingredients.FirstOrDefault(i => i.Name.ToUpper().Contains("RED") || i.Name.ToUpper().Contains("RUBINE") || i.Name.ToUpper().Contains("SCARLET"))?.Name;
                else if (hueShiftDir == "AMARILLO") culpritDye = recipeRes.Ingredients.FirstOrDefault(i => i.Name.ToUpper().Contains("YELLOW") || i.Name.ToUpper().Contains("GOLD"))?.Name;
                else if (hueShiftDir == "VERDE") culpritDye = recipeRes.Ingredients.FirstOrDefault(i => i.Name.ToUpper().Contains("GREEN"))?.Name;

                if (!string.IsNullOrEmpty(culpritDye))
                {
                    sb.AppendLine($"{recNum++}. Reducir carga de {culpritDye} (Causante del viraje hacia el {hueShiftDir.ToLower()}).");
                }
                else
                {
                    string cleanCorrA = d65Res.CorrectionA.Replace("«", "").Replace("»", "");
                    string cleanCorrB = d65Res.CorrectionB.Replace("«", "").Replace("»", "");
                    if (!string.IsNullOrEmpty(cleanCorrA) && cleanCorrA != Color.ToleranceResult.MSG_OK) sb.AppendLine($"{recNum++}. Ajuste de Tinte: {cleanCorrA}");
                    if (!string.IsNullOrEmpty(cleanCorrB) && cleanCorrB != Color.ToleranceResult.MSG_OK) sb.AppendLine($"{recNum++}. Ajuste de Tinte: {cleanCorrB}");
                }
            }
            else
            {
                string cleanCorrA = d65Res.CorrectionA.Replace("«", "").Replace("»", "");
                string cleanCorrB = d65Res.CorrectionB.Replace("«", "").Replace("»", "");
                if (!string.IsNullOrEmpty(cleanCorrA) && cleanCorrA != Color.ToleranceResult.MSG_OK) sb.AppendLine($"{recNum++}. Ajuste de Tinte: {cleanCorrA}");
                if (!string.IsNullOrEmpty(cleanCorrB) && cleanCorrB != Color.ToleranceResult.MSG_OK) sb.AppendLine($"{recNum++}. Ajuste de Tinte: {cleanCorrB}");
            }

            if (Math.Abs(dLVar) >= 0.1)
            {
                sb.AppendLine($"{recNum++}. Ajustar concentración general para corregir el {dLVar:F1}% de Luz.");
            }

            string cost = dHueVar > 20.0 ? "alto" : dHueVar > 5.0 ? "medio" : "bajo";
            sb.AppendLine($"{recNum++}. El costo de reproceso es {cost} debido a la desviación del {dHueVar:F1}%.");
            sb.AppendLine("============================================================");

            return sb.ToString();
        }

        private static string GetGlobalMetadataPrefix()
        {
            var globalRes = Color.ShadeReportExtractor.LastResult;
            if (globalRes == null || string.IsNullOrWhiteSpace(globalRes.ShadeName)) return string.Empty;

            var sb = new System.Text.StringBuilder();
            sb.Append(" Shade Name : ").Append(globalRes.ShadeName);
            sb.AppendLine();
            sb.AppendLine("───────────────────────────────────────────────────────────────");
            return sb.ToString();
        }

        /// Referencia al formulario OCR de origen — para el boton Regresar.
        public Form FormOcrOrigen { get; set; }

        // ======= Escritura automática al Excel de referencia =======
        /// <summary>
        /// Vuelca silenciosamente los valores OCR a las celdas azules del Excel de referencia.
        /// El usuario no interviene en nada: ocurre automáticamente en segundo plano.
        /// </summary>
        private void WriteExcelSilently()
        {
            // Capturar snapshot de los datos para el hilo de fondo
            var rows     = _ocrRows;
            var cmc      = _cmcResults;
            var recipe   = _recipeResults;

            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    string excelPath = ExcelWriter.FindReferenceExcelPath();
                    if (string.IsNullOrEmpty(excelPath)) return;

                    ExcelWriter.WriteAll(excelPath, rows, cmc, recipe);
                }
                catch
                {
                    // Silencioso — nunca interrumpir el flujo del usuario
                }
            });
        }

        // ======= Inicialización de la UI (layout elástico) =======
        private void InitializeComponents()
        {
            // ---- Ventana y escalado ----
            this.Text = "TINT COATS CADENA";
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ControlBox = true;
            this.ShowIcon = true;

            this.BackColor = System.Drawing.Color.White;
            this.ForeColor = System.Drawing.Color.Black;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            // Tamaño inicial cómodo (90% del área de trabajo)
            var wa = Screen.PrimaryScreen.WorkingArea;
            int targetWidth = (int)(wa.Width * 0.90);
            int targetHeight = (int)(wa.Height * 0.90);
            this.MinimumSize = new Size(980, 640);
            this.Size = new Size(targetWidth, targetHeight);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized; 
            this.ResizeRedraw = true;

            // ---- Título ----
            var lblTitulo = new Label
            {
                Text = "ANALISIS DE COLORIMETRIA",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.FromArgb(0, 102, 204),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(15, 0, 0, 0)
            };

            // ---- Controles inferiores ----
            btnExportar = new Button
            {
                Text = "💾 Exportar .txt",
                Size = new Size(150, 38),
                BackColor = System.Drawing.Color.FromArgb(70, 130, 180), 
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            btnExportar.FlatAppearance.BorderSize = 0;
            btnExportar.Click += BtnExportar_Click;

            btnHistorial = new Button
            {
                Text = "📜 Historial",
                Size = new Size(150, 38),
                BackColor = System.Drawing.Color.FromArgb(34, 139, 34), 
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnHistorial.FlatAppearance.BorderSize = 0;
            btnHistorial.Click += BtnHistorial_Click;

            btnCerrar = new Button
            {
                Text = "Cerrar",
                Size = new Size(120, 38),
                BackColor = System.Drawing.Color.FromArgb(200, 30, 30),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            btnCerrar.FlatAppearance.BorderSize = 0;
            btnCerrar.Click += (s, e) => this.Close();

            btnRegresar = new Button
            {
                Text = "← Regresar",
                Size = new Size(150, 38),
                BackColor = System.Drawing.Color.FromArgb(180, 100, 0),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            btnRegresar.FlatAppearance.BorderSize = 0;
            btnRegresar.Click += BtnRegresar_Click;

            // ---- Split central (Dock=Fill) ----
            splitMedicionesCmc = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = System.Drawing.Color.White
            };

            // Panel IZQUIERDO: Reporte/OCR (encabezado + textbox)
            var panelLeftHeader = new Label
            {
                Text = " ANALISIS DE SHADE HISTORY REPORT ",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 32,
                Padding = new Padding(4, 6, 4, 6),
                BackColor = System.Drawing.Color.FromArgb(0, 102, 204)
            };

            txtReport = BuildRichControl(null); 
            txtReport.Dock = DockStyle.Fill;

            var panelLeft = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.White };
            panelLeft.Controls.Add(txtReport);
            panelLeft.Controls.Add(panelLeftHeader);

            splitMedicionesCmc.Panel1.Controls.Add(panelLeft);

            // Panel DERECHO: CMC (2:1) / Recomendación (encabezado + textbox)
            var panelRightHeader = new Label
            {
                Text = "ANALISIS DE SAMPLE COMPARISON  ",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 32,
                Padding = new Padding(4, 6, 4, 6),
                BackColor = System.Drawing.Color.FromArgb(0, 102, 204)
            };

            txtRecomendacion = BuildRichControl(null);
            txtRecomendacion.Dock = DockStyle.Fill;


            btnVerGrafico = new Button
            {
                Text = "🔍 Ver Gráfico Detallado",
                Size = new Size(180, 34),
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Visible = false 
            };
            btnVerGrafico.FlatAppearance.BorderColor = System.Drawing.Color.DarkGray;
            btnVerGrafico.Click += BtnVerGrafico_Click;

            // Panel para el botón (encima del diagnóstico)
            var pnlButtonArea = new Panel { Dock = DockStyle.Bottom, Height = 45, BackColor = System.Drawing.Color.White };
            pnlButtonArea.Controls.Add(btnVerGrafico);
            btnVerGrafico.Location = new Point(10, 5);

            var panelRight = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.White };
            panelRight.Controls.Add(txtRecomendacion);
            panelRight.Controls.Add(pnlButtonArea);
            panelRight.Controls.Add(panelRightHeader);

            splitMedicionesCmc.Panel2.Controls.Add(panelRight);

            // ---- Root layout (TableLayoutPanel) ----
            var panelRoot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = System.Drawing.Color.White
            };
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));

            // Fila 0: Título
            var panelHeader = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.FromArgb(0, 102, 204) };
            panelHeader.Controls.Add(lblTitulo);
            panelRoot.Controls.Add(panelHeader, 0, 0);

            // Fila 1: Split
            panelRoot.Controls.Add(splitMedicionesCmc, 0, 1);

            // Fila 2: Regresar (izquierda) | Exportar + Cerrar (derecha)
            var panelButtons = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = System.Drawing.Color.White,
                Padding = new Padding(10, 8, 10, 8)
            };
            panelButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            panelButtons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // Izquierda: Regresar
            var panelBtnLeft = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                BackColor = System.Drawing.Color.White,
                Padding = new Padding(0, 0, 0, 0)
            };
            panelBtnLeft.Controls.Add(btnRegresar);

            // Derecha: Exportar + Cerrar
            var panelBtnRight = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                BackColor = System.Drawing.Color.White,
                Padding = new Padding(0, 0, 0, 0)
            };
            panelBtnRight.Controls.Add(btnCerrar);
            panelBtnRight.Controls.Add(btnExportar);
            panelBtnRight.Controls.Add(btnHistorial);

            panelButtons.Controls.Add(panelBtnLeft, 0, 0);
            panelButtons.Controls.Add(panelBtnRight, 1, 0);
            panelRoot.Controls.Add(panelButtons, 0, 2);

            // Agregar root al formulario
            this.Controls.Add(panelRoot);

            // Atajos para foco rápido (opcional)
            this.KeyPreview = true;
            this.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.D1) { txtReport.Focus(); e.SuppressKeyPress = true; }
                if (e.Control && e.KeyCode == Keys.D2) { txtRecomendacion.Focus(); e.SuppressKeyPress = true; }
            };

            // Mantener proporción del divisor y limpiar selección de texto en Load/Resize
            this.Load += (s, e) => {
                ApplySplitRatio();
                txtReport.SelectionLength = 0;
                txtRecomendacion.SelectionLength = 0;
                btnCerrar.Focus(); 
            };
            this.Resize += (s, e) => ApplySplitRatio();
        }

        // ======= Actualización del Gráfico =======
        private void UpdateChartFromReport(OcrReport rep)
        {
            if (rep == null || rep.Measures == null) return;
            var rows = rep.Measures.Select(m => new EngineRow
            {
                Illuminant = m.Illuminant,
                Type = m.Type,
                L = m.L,
                A = m.A,
                B = m.B,
                Hue = m.Hue
            }).ToList();
            var results = EngineCalc.Calculate(rows);
            UpdateChartFromResults(results);
        }

        private void UpdateChartFromResults(List<EngineRes> results)
        {
            if (results == null || results.Count == 0) return;

            // Buscamos D65 como estándar industrial, sino el primero
            var res = results.FirstOrDefault(r => string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase)) 
                      ?? results[0];

            _lastMainResult = res;
            if (btnVerGrafico != null) btnVerGrafico.Visible = true;
        }

        private void BtnVerGrafico_Click(object sender, EventArgs e)
        {
            if (_lastMainResult == null) return;
            
            // --- Inmersión de Datos: Buscar valores absolutos del Estándar (Std) ---
            double absL = 50, absA = 0, absB = 0;
            if (_report != null && _report.Measures != null)
            {
                // Buscamos el registro 'Std' que coincida con el iluminante del gráfico actual
                var stdRow = _report.Measures.FirstOrDefault(m => 
                    m.Type == "Std" && 
                    string.Equals(m.Illuminant, _lastMainResult.Illuminant, StringComparison.OrdinalIgnoreCase));
                
                if (stdRow != null)
                {
                    absL = stdRow.L;
                    absA = stdRow.A;
                    absB = stdRow.B;
                }
            }

            // Combinar instrucciones profesionales para el gráfico
            string combinedInstructions = $"{_lastMainResult.LightnessInstruction} | {_lastMainResult.CorrectionA} | {_lastMainResult.CorrectionB}";

            using (var frm = new FormDetalleCielab(
                _lastMainResult.DeltaL, 
                _lastMainResult.DeltaA, 
                _lastMainResult.DeltaB, 
                _lastMainResult.DeltaE, 
                _lastMainResult.CmcValue,
                DE_MAX,
                combinedInstructions,
                absL, absA, absB))
            {
                frm.ShowDialog();
            }
        }

        // ======= Aplica proporción del Split (55% izquierda / 45% derecha) =======
        private void ApplySplitRatio()
        {
            try
            {
                if (splitMedicionesCmc != null && splitMedicionesCmc.Orientation == Orientation.Vertical)
                {
                    int w = splitMedicionesCmc.ClientSize.Width;
                    int distance = (int)(w * _splitLeftRatio);

                    // evitar colapsos: reservar 200 px por lado
                    distance = Math.Max(200, Math.Min(w - 200, distance));
                    splitMedicionesCmc.SplitterDistance = distance;
                }
            }
            catch
            {
                // ignorar si aún no está listo para medir
            }
        }

        // ======= Factory de RichTextBox monoespaciado =======
        private static RichTextBox BuildRichControl(System.Drawing.Color? foreColor)
        {
            return new RichTextBox
            {
                Multiline = true,
                ScrollBars = RichTextBoxScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.0f),
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                ReadOnly = true,
                WordWrap = true,
                TabStop = false, 
                BorderStyle = BorderStyle.None
            };
        }

        private void SetFormattedText(RichTextBox rtb, string text)
        {
            if (rtb == null) return;
            rtb.Clear();
            if (string.IsNullOrEmpty(text)) return;

            rtb.SuspendLayout();
            
            bool inD65Block = true; // D65 es el predeterminado al inicio
            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            
            System.Drawing.Color colorDim = System.Drawing.Color.Gray;
            System.Drawing.Color colorMain = System.Drawing.Color.Black;

            foreach (var line in lines)
            {
                int startLine = rtb.TextLength;
                string trimmed = line.Trim();
                
                // --- 1. Detección de bloque ---
                // Si la línea indica un iluminante secundario, cambiamos a modo tenue
                if (trimmed.Contains("Iluminante") && !trimmed.Contains("D65"))
                {
                    inD65Block = false;
                }
                else if (trimmed.Contains("D65") || trimmed.Contains("DIAGNÓSTICO TÉCNICO") || trimmed.Contains("RECOMENDACIÓN FINAL"))
                {
                    inD65Block = true;
                }
                else if (trimmed.StartsWith("[ ") && trimmed.EndsWith(" ]")) // Tags de recomendación [ TL84 ]
                {
                    inD65Block = trimmed.Contains("D65");
                }

                // --- 2. Procesamiento de Etiquetas Dinámicas « » ---
                string processedLine = line;
                List<(int start, int length)> highlightRanges = new List<(int, int)>();
                while (processedLine.Contains("«"))
                {
                    int idxStart = processedLine.IndexOf("«");
                    int idxEnd = processedLine.IndexOf("»", idxStart);
                    if (idxEnd == -1) break;

                    string content = processedLine.Substring(idxStart + 1, idxEnd - idxStart - 1);
                    processedLine = processedLine.Remove(idxStart, (idxEnd - idxStart) + 1).Insert(idxStart, content);
                    highlightRanges.Add((idxStart, content.Length));
                }

                rtb.AppendText(processedLine + Environment.NewLine);

                // --- 3. Aplicación de Estilos ---
                rtb.Select(startLine, processedLine.Length);
                
                if (inD65Block)
                {
                    rtb.SelectionColor = colorMain;
                    // Opcional: resaltar headers o D65 específico
                    if (trimmed.Contains("D65") || trimmed.Contains("DIAGNÓSTICO") || trimmed.Contains("RECOMENDACIÓN"))
                    {
                        rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                    }
                }
                else
                {
                    rtb.SelectionColor = colorDim;
                    rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
                }

                // Resaltado de celdas importantes (Etiquetadas con « »)
                foreach (var range in highlightRanges)
                {
                    rtb.Select(startLine + range.start, range.length);
                    rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                    if (inD65Block) rtb.SelectionColor = colorMain;
                }

                // Restaurar posición
                rtb.SelectionStart = rtb.TextLength;
            }
            
            rtb.SelectionStart = 0;
            rtb.SelectionLength = 0;
            rtb.ScrollToCaret();
            rtb.ResumeLayout();
        }

        // =========================================================
        // RECOMENDACIÓN — desde lista de resultados 
        // =========================================================
        private static string BuildRecomendacionFromResults(List<EngineRes> results, double DL_MAX = 0.69, double DC_MAX = 0.69, double DH_MAX = 0.69, double DE_MAX = 1.20)
        {
            if (results == null || results.Count == 0)
                return "No hay resultados para generar recomendación.";

            var sb = new StringBuilder();
            
            // --- Encabezado de Shade (si existe) ---
            var shade = Color.ShadeReportExtractor.LastResult;
            if (shade != null && !string.IsNullOrEmpty(shade.ShadeName))
            {
                sb.AppendLine($"Shade Name : {shade.ShadeName}");
                sb.AppendLine("───────────────────────────────────────────────────────────────────");
            }
            sb.AppendLine($"Tolerancia: ({DE_MAX:F2} DL:{DL_MAX:F2} DC:{DC_MAX:F2} DH:{DH_MAX:F2}) ");

            var band = new ToleranceResult
            {
                DE = DE_MAX,
                DL = DL_MAX,
                DC = DC_MAX,
                DH = DH_MAX
            };
            var evaluation = EngineCalc.EvaluateTolerance(results, band);
            // sb.AppendLine(evaluation.FormatReport()); // Ocultar para ahorrar espacio

            // % formateado con un decimal para mayor precisión técnica
            Func<double, string> FmtPctSigned = v =>
                double.IsNaN(v) ? "N/D" : (Math.Round(v, 1, MidpointRounding.AwayFromZero).ToString("0.0", CultureInfo.InvariantCulture) + "%");

            // --- 1) DIAGNÓSTICO TÉCNICO (D65 o Principal) ---
            var mainRes = results.FirstOrDefault(r => string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase)) ?? results[0];
            var mainCheck = evaluation.Checks.FirstOrDefault(c => string.Equals(c.Illuminant, mainRes.Illuminant, StringComparison.OrdinalIgnoreCase));

            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            sb.AppendLine(" DIAGNÓSTICO TÉCNICO POR ILUMINANTE");
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            
            // --- Resumen de Cumplimiento ---
            foreach (var r in results)
            {
                var check = evaluation.Checks.FirstOrDefault(c => string.Equals(c.Illuminant, r.Illuminant, StringComparison.OrdinalIgnoreCase));
                string statusTxt = (check != null && check.Passes) ? "CUMPLE" : "NO CUMPLE";
                sb.AppendLine($"Iluminante {r.Illuminant} {statusTxt}");
            }
            sb.AppendLine();

            if (evaluation.AllPass)
            {
                sb.AppendLine(" ESTADO DEL LOTE: PASS");
                sb.AppendLine();
                sb.AppendLine(" Todos los iluminantes cumplen las tolerancias definidas.");
                sb.AppendLine(" Recomendación final: NO SE REQUIERE CORRECCIÓN.");
                sb.AppendLine("═══════════════════════════════════════════════════════════════════");
                return sb.ToString();
            }

            sb.AppendLine("Basado en los deltas calculados, el color presenta las siguientes desviaciones:");
            sb.AppendLine();
            sb.AppendLine("(ILUMINANTE PRINCIPAL D65)");
            
            sb.AppendLine(string.Format(" * Matiz (ΔH): {0:F2} ({1:F1}%). {2}.", 
                mainRes.DeltaHue, 
                mainRes.PercentHue,
                GetHueDescription(mainRes)));
            
            sb.AppendLine(string.Format(" * Luminosidad (ΔL): {0:F2} ({1:F1}%). {2}.", 
                mainRes.DeltaL, 
                mainRes.PercentL,
                GetLightnessDescription(mainRes.DeltaL)));

            sb.AppendLine(string.Format(" * Chroma (ΔC): {0:F2} ({1:F1}%). {2}.", 
                mainRes.DeltaChroma, 
                mainRes.PercentChroma,
                GetChromaDescription(mainRes.DeltaChroma)));

            // --- Otros iluminantes en el diagnóstico ---
            if (results.Count > 1)
            {
                foreach (var r in results)
                {
                    if (r == mainRes) continue;
                    sb.AppendLine();
                    sb.AppendLine($"--- Iluminante {r.Illuminant} ---");
                    sb.AppendLine(string.Format(" * Matiz ΔH: {0:F2} ({1:F1}%). {2}", r.DeltaHue, r.PercentHue, GetHueDescription(r)));
                    sb.AppendLine(string.Format(" * Luminosidad ΔL: {0:F2} ({1:F1}%). {2}", r.DeltaL, r.PercentL, GetLightnessDescription(r.DeltaL)));
                    sb.AppendLine(string.Format(" * Chroma ΔC: {0:F2} ({1:F1}%). {2}", r.DeltaChroma, r.PercentChroma, GetChromaDescription(r.DeltaChroma)));
                }
            }
            
            // --- ESTADO DEL LOTE ---
            string status = evaluation.AllPass ? "PASS" : "FALL";
            string reason = "";
            if (!evaluation.AllPass && mainCheck != null)
            {
                 if (mainCheck.MeasuredDE > DE_MAX) reason = $"(DE {mainCheck.MeasuredDE:F2} > {DE_MAX:F2})";
                 else if (Math.Abs(mainRes.DeltaHue) > DH_MAX) reason = $"(ΔH {Math.Abs(mainRes.DeltaHue):F2} > {DH_MAX:F2})";
            }
            sb.AppendLine();
            sb.AppendLine($" ESTADO DEL LOTE: {status} ");
            sb.AppendLine();


            // --- 2) RECOMENDACIÓN FINAL ---
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            sb.AppendLine(" RECOMENDACIÓN FINAL POR ILUMINANTE");
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            
            foreach (var r in results)
            {
                string cleanCorrA = r.CorrectionA.Replace("«", "").Replace("»", "");
                string cleanCorrB = r.CorrectionB.Replace("«", "").Replace("»", "");
                string cleanLight = r.LightnessInstruction.Replace("«", "").Replace("»", "");

                if (string.Equals(r.Illuminant, "D65", StringComparison.OrdinalIgnoreCase))
                {
                    sb.AppendLine("(ILUMINANTE PRINCIPAL D65)");
                }
                else
                {
                    sb.AppendLine($"[ {r.Illuminant} ]");
                }
                
                sb.AppendLine(" • Ajuste de Tinte: (" + cleanCorrA + ")");
                sb.AppendLine(" • Ajuste de Tinte: (" + cleanCorrB + ")");
                sb.AppendLine(" • Ajuste de Concentración: " + cleanLight);
                sb.AppendLine();
            }



            sb.AppendLine("───────────────────────────────────────────────────────────────────");
            sb.AppendLine("Tras el ajuste, se recomienda re-medición bajo todos los iluminantes.");
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        private static string GetHueDescription(EngineRes r)
        {
            List<string> shifts = new List<string>();
            if (r.DeltaA > 0.05) shifts.Add("más rojizos");
            else if (r.DeltaA < -0.05) shifts.Add("más verdosos");

            if (r.DeltaB > 0.05) shifts.Add("más amarillentos");
            else if (r.DeltaB < -0.05) shifts.Add("más azulados");

            if (shifts.Count == 0) return "Sin rotación de tono significativa";
            
            string intensity = Math.Abs(r.DeltaHue) > 1.0 ? "significativamente " : "";
            return $"El color ha rotado {intensity}hacia tonos {string.Join(" y ", shifts)}";
        }

        private static string GetLightnessDescription(double dL)
        {
            if (dL < -0.20) return "La muestra es notablemente más oscura que el estándar";
            if (dL < -0.01) return "La muestra es ligeramente más oscura";
            if (dL > 0.20) return "La muestra es notablemente más clara que el estándar";
            if (dL > 0.01) return "La muestra es ligeramente más clara";
            return "Luminosidad dentro de rango";
        }

        private static string GetChromaDescription(double dC)
        {
            if (dC < -0.15) return "La muestra tiene menos saturación (está más 'lavada') que el estándar";
            if (dC > 0.15) return "La muestra tiene más saturación (color más intenso)";
            return "Saturación dentro de rango";
        }

        // ======= Helper: imprime L*, a*, b* con % y acción (formato específico) =======
        private static string BuildPerAxisPercentAdvice(EngineRes r)
        {
            var sb = new StringBuilder();

            // Lote L* (Ligthness)
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                " * Ligthness : {0}", r.LightnessInstruction));

            // Croma
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                " * Chroma    : {0}", r.ChromaInstruction));

            // Eje a* (Rojo/Verde)
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                " * Eje a*    : {0}", r.CorrectionA));

            // Eje b* (Amarillo/Azul)
            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                " * Eje b*    : {0}", r.CorrectionB));

            return sb.ToString();
        }

        // =========================================================
        // RECOMENDACIÓN — desde OcrReport 
        // =========================================================
        private static string BuildRecomendacionFromReport(OcrReport rep)
        {
            //  Validaciones básicas
            if (rep == null || rep.Measures == null || rep.Measures.Count == 0)
                return "No hay datos en el reporte para generar recomendación.";

            //  Convertir a filas para el motor (Std/Lot por iluminante)
            List<EngineRow> rowsForEngine = rep.Measures.Select(m => {

                // Asegurar que C* y h° estén calculados (requeridos para el cálculo de elipses CMC)
                double chroma = m.Chroma;
                if (chroma <= 0) chroma = Math.Sqrt(m.A * m.A + m.B * m.B);

                double hue = m.Hue;
                if (hue <= 0) {
                    hue = Math.Atan2(m.B, m.A) * 180.0 / Math.PI;
                    if (hue < 0) hue += 360.0;
                }

                return new EngineRow
                {
                    Illuminant = m.Illuminant,
                    Type = m.Type,
                    L = m.L,
                    A = m.A,
                    B = m.B,
                    Chroma = chroma,
                    Hue = hue
                };
            }).ToList();

            //  Calcular con el motor (List<ColorimetricRow> -> List<CorrectionResult>)
            List<EngineRes> calcResults = EngineCalc.Calculate(rowsForEngine);
            
            // --- CÁLCULO CMC (ELIPSE) ---
            var cmcResults = EngineCalc.CalculateCmc(calcResults, rowsForEngine);
            if (cmcResults != null)
            {
                foreach (var r in calcResults)
                {
                    var cmcRes = cmcResults.FirstOrDefault(x => string.Equals(x.Illuminant, r.Illuminant, StringComparison.OrdinalIgnoreCase));
                    if (cmcRes != null) r.CmcValue = cmcRes.CmcValue;
                }
            }

            //  Prefijo con Shade Name y DT Main de manera segura
            string shadeName = string.Empty;
            string dtMain = string.Empty;

            var globalRes = Color.ShadeReportExtractor.LastResult;
            if (globalRes != null)
            {
                shadeName = globalRes.ShadeName;
                dtMain = globalRes.DtMain;
            }

            if (string.IsNullOrWhiteSpace(shadeName) || string.IsNullOrWhiteSpace(dtMain))
            {
                try
                {
                    var pShade = rep.GetType().GetProperty("ShadeName");
                    if (pShade != null && string.IsNullOrWhiteSpace(shadeName)) shadeName = pShade.GetValue(rep, null) as string;

                    var pDt = rep.GetType().GetProperty("DtMain");
                    if (pDt != null && string.IsNullOrWhiteSpace(dtMain)) dtMain = pDt.GetValue(rep, null) as string;
                }
                catch { }
            }

            string prefix = string.Empty;
            if (!string.IsNullOrWhiteSpace(shadeName) || !string.IsNullOrWhiteSpace(dtMain))
            {
                var sb2 = new StringBuilder();
                sb2.AppendLine();
                if (!string.IsNullOrWhiteSpace(shadeName))
                    sb2.AppendLine(" Shade Name : " + shadeName);
                if (!string.IsNullOrWhiteSpace(dtMain))
                    sb2.AppendLine(" DT Main    : " + dtMain);
                sb2.AppendLine("───────────────────────────────────────────────────────────────");
                prefix = sb2.ToString();
            }

            //  Reutilizar el generador de texto unificado
            string body = BuildRecomendacionFromResults(calcResults,
                Properties.Settings.Default.ToleranciaDL,
                Properties.Settings.Default.ToleranciaDC,
                Properties.Settings.Default.ToleranciaDH,
                Properties.Settings.Default.ToleranciaDE);

            //  POBLAR CAMPOS DETALLADOS PARA EL HISTORIAL (Know-How)
            try
            {
                if (rep.Batch == null) rep.Batch = new BatchInfo();
                rep.Batch.ShadeName = shadeName;
                rep.Batch.BatchId = dtMain;
                rep.Batch.LotNo = string.Empty; 

                if (calcResults != null && calcResults.Count > 0)
                {
                    var r = calcResults[0]; 
                    rep.Batch.dL = r.DeltaL;
                    rep.Batch.dC = r.DeltaChroma;
                    rep.Batch.dH = r.DeltaHue;
                    rep.Batch.dE = r.DeltaE;

                    var bandObj = new ToleranceResult
                    {
                        DE = Properties.Settings.Default.ToleranciaDE,
                        DL = Properties.Settings.Default.ToleranciaDL,
                        DC = Properties.Settings.Default.ToleranciaDC,
                        DH = Properties.Settings.Default.ToleranciaDH
                    };
                    var evalSingle = EngineCalc.EvaluateTolerance(new List<EngineRes> { r }, bandObj);
                    bool pass = evalSingle.AllPass;
                    rep.Batch.PF = pass ? "PASS" : "FAIL";

                    // Diagnóstico corto (L*) para el reporte
                    if (r.DeltaL < -0.01) rep.DiagnosticoL = "Lote más OSCURO";
                    else if (r.DeltaL > 0.01) rep.DiagnosticoL = "Lote más CLARO";
                    else rep.DiagnosticoL = "Sin desviación L*";

                    // Recomendación corta para el reporte
                    double pctAbs = Math.Round(Math.Abs(r.PercentL), 0, MidpointRounding.AwayFromZero);
                    string accion = (r.DeltaL < -0.01) ? "ACLARAR" : (r.DeltaL > 0.01 ? "OSCURECER" : "N/A");
                    rep.Recomendacion = string.Format("{0} {1}%", accion, pctAbs);
                }
            }
            catch { }

            return string.IsNullOrEmpty(prefix) ? body : prefix + body;
        }

        // =========================================================
        // RESUMEN TEXTO (OCR)
        // =========================================================
        private static string BuildResumenFromReport(OcrReport rep)
        {
            if (rep == null) return string.Empty;

            var sb = new StringBuilder();
            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            sb.AppendLine(" RESULTADOS DE CORRECCIÓN COLORIMÉTRICA — OCR");
            sb.AppendLine(" Fecha: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture));
            sb.AppendLine("═══════════════════════════════════════════════════════════════");

            // ---- Shade Name y DT Main ----
            string shadeName = string.Empty;
            string dtMain = string.Empty;

            var globalRes = Color.ShadeReportExtractor.LastResult;
            if (globalRes != null)
            {
                shadeName = globalRes.ShadeName;
                dtMain = globalRes.DtMain;
            }

            if (string.IsNullOrWhiteSpace(shadeName) || string.IsNullOrWhiteSpace(dtMain))
            {
                try
                {
                    var pShade = rep.GetType().GetProperty("ShadeName");
                    if (pShade != null && string.IsNullOrWhiteSpace(shadeName)) shadeName = pShade.GetValue(rep, null) as string;

                    var pDt = rep.GetType().GetProperty("DtMain");
                    if (pDt != null && string.IsNullOrWhiteSpace(dtMain)) dtMain = pDt.GetValue(rep, null) as string;
                }
                catch { }
            }

            bool hasMeta = !string.IsNullOrWhiteSpace(shadeName) || !string.IsNullOrWhiteSpace(dtMain);
            if (hasMeta)
            {
                sb.AppendLine();
                if (!string.IsNullOrWhiteSpace(shadeName))
                    sb.AppendLine(" Shade Name : " + shadeName);
                if (!string.IsNullOrWhiteSpace(dtMain))
                    sb.AppendLine(" DT Main    : " + dtMain);
                sb.AppendLine("───────────────────────────────────────────────────────────────");
            }

            if (rep.Measures != null && rep.Measures.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Mediciones (Iluminante / Tipo / L* / a* / b* / Chroma / Hue°)");
                sb.AppendLine("───────────────────────────────────────────────────────────────");

                string last = "";
                foreach (var m in rep.Measures)
                {
                    if (last != "" && last != m.Illuminant)
                        sb.AppendLine("───────────────────────────────────────────────────────────────");

                    sb.AppendLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "{0,-6} {1,-4} {2,6:0.00} {3,6:0.00} {4,6:0.00} {5,6:0.00} {6,3:0}",
                        m.Illuminant, m.Type, m.L, m.A, m.B, m.Chroma, m.Hue));

                    last = m.Illuminant;
                }
            }

            sb.AppendLine("═══════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        // =========================================================
        // EXPORTAR TEXTO
        // =========================================================
        private void BtnRegresar_Click(object sender, EventArgs e)
        {
            if (FormOcrOrigen == null || FormOcrOrigen.IsDisposed) return;

            try
            {
                this.Hide();
                var result = FormOcrOrigen.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    // Usuario confirmó — Form1 detectará RowsConfirmed y reabrirá Resultados
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
                else
                {
                    // Usuario canceló — cerrar sin señal de recálculo
                    this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
                    this.Close();
                }
            }
            catch
            {
                this.Show();
            }
        }

        private bool _historialGuardado = false;
        
        private void BtnHistorial_Click(object sender, EventArgs e)
        {
            try
            {
                //  Obtener Metadatos
                string shade = string.Empty;
                string batch = string.Empty;

                var globalRes = Color.ShadeReportExtractor.LastResult;
                if (globalRes != null)
                {
                    shade = globalRes.ShadeName;
                    batch = globalRes.DtMain;
                }

                if (string.IsNullOrWhiteSpace(shade) && _report != null)
                {
                    try { shade = _report.GetType().GetProperty("ShadeName")?.GetValue(_report, null) as string; } catch { }
                    try { batch = _report.GetType().GetProperty("DtMain")?.GetValue(_report, null) as string; } catch { }
                }

                string shadeName = string.IsNullOrWhiteSpace(shade) ? "N/A" : shade;
                string batchId = string.IsNullOrWhiteSpace(batch) ? "N/A" : batch;
                string lotNo = _report?.Batch?.LotNo ?? "N/A";
                double deltaEcalc = _report?.Batch?.dE ?? 0.0;
                double deltaLcalc = _report?.Batch?.dL ?? 0.0;

                if (!_historialGuardado && _resultsLegacy != null && _resultsLegacy.Count > 0)
                {
                    // Preguntar si está seguro de guardar
                    var resultConfirm = MessageBox.Show("¿Está seguro de que desea guardar los datos en el historial?", "Confirmar Guardado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (resultConfirm == DialogResult.No) return;

                    foreach (var r in _resultsLegacy)
                    {
                        string iluminante = r.Illuminant;

                        double pctL = Math.Round(Math.Abs(r.PercentL), 1, MidpointRounding.AwayFromZero);
                        double pctA = Math.Round(Math.Abs(r.PercentA), 1, MidpointRounding.AwayFromZero);
                        double pctB = Math.Round(Math.Abs(r.PercentB), 1, MidpointRounding.AwayFromZero);

                        string lPctStr = double.IsNaN(pctL) ? "0%" : pctL.ToString("0.0", CultureInfo.InvariantCulture) + "%";
                        string aPctStr = double.IsNaN(pctA) ? "0%" : pctA.ToString("0.0", CultureInfo.InvariantCulture) + "%";
                        string bPctStr = double.IsNaN(pctB) ? "0%" : pctB.ToString("0.0", CultureInfo.InvariantCulture) + "%";

                        string chromaPctStr = "0%";
                        try
                        {
                            double chromaBase = Math.Sqrt(Math.Pow(r.PercentA, 2) + Math.Pow(r.PercentB, 2));
                            if (!double.IsNaN(chromaBase)) chromaPctStr = Math.Round(chromaBase, 1).ToString("0.0", CultureInfo.InvariantCulture) + "%";
                        }
                        catch { }

                        string diagL = "Sin desviación"; string corrL = "N/A";
                        if (r.DeltaL < -0.01) { diagL = "Lote más OSCURO"; corrL = $"ACLARAR {lPctStr}"; }
                        else if (r.DeltaL > 0.01) { diagL = "Lote más CLARO"; corrL = $"OSCURECER {lPctStr}"; }

                        string diagA = "Sin sesgo"; string corrA = "N/A";
                        if (r.DeltaA > 0.01) { diagA = "MÁS ROJO"; corrA = $"{aPctStr} DISMINUIR ROJO"; }
                        else if (r.DeltaA < -0.01) { diagA = "MÁS VERDE"; corrA = $"{aPctStr} DISMINUIR VERDE"; }

                        string diagB = "Sin sesgo"; string corrB = "N/A";
                        if (r.DeltaB > 0.01) { diagB = "MÁS AMARILLO"; corrB = $"{bPctStr} DISMINUIR AMARILLO"; }
                        else if (r.DeltaB < -0.01) { diagB = "MÁS AZUL"; corrB = $"{bPctStr} DISMINUIR AZUL"; }

                        HistorialService.GuardarMedicionDetallada(
                            DateTime.Now,
                            shadeName,
                            iluminante,
                            lPctStr,
                            chromaPctStr,
                            diagL, corrL,
                            diagA, corrA,
                            diagB, corrB
                        );
                    }
                    _historialGuardado = true;
                    MessageBox.Show("Datos guardados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                //  Abrir historial
                FormHistorial frm = new FormHistorial();
                frm.CargarHistorial(HistorialService.ObtenerHistorial());
                frm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al abrir el historial: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportar_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtReport.Text) && string.IsNullOrWhiteSpace(txtRecomendacion.Text))
                {
                    MessageBox.Show("No hay información en el reporte para exportar.", "Reporte Vacío", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Preguntar al usuario qué formato quiere
                var resultado = MessageBox.Show(
                    "¿Desea exportar a Excel?\n\n" +
                    "  • [Sí]  → Exportar a Excel (.xls) con tabla estructurada\n" +
                    "  • [No]  → Exportar como texto plano (.txt)",
                    "Formato de exportación",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);

                if (resultado == DialogResult.Cancel) return;

                if (resultado == DialogResult.Yes)
                    ExportarExcel();
                else
                    ExportarTxt();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportarTxt()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Texto (*.txt)|*.txt";
                sfd.FileName = "Reporte_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".txt";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string contenido = txtReport.Text + Environment.NewLine +
                                       new string('=', 70) + Environment.NewLine +
                                       txtRecomendacion.Text;
                    System.IO.File.WriteAllText(sfd.FileName, contenido, System.Text.Encoding.UTF8);
                    MessageBox.Show("Archivo de texto guardado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void ExportarExcel()
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel (*.xls)|*.xls";
                sfd.FileName = "Reporte_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xls";
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var sb = new StringBuilder();

                // Cabecera XML de Excel
                sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                sb.AppendLine("<?mso-application progid=\"Excel.Sheet\"?>");
                sb.AppendLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
                sb.AppendLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
                sb.AppendLine("<Styles>");

                // Estilo título
                sb.AppendLine("<Style ss:ID=\"sTitle\">");
                sb.AppendLine("<Font ss:Bold=\"1\" ss:Size=\"13\" ss:Color=\"#FFFFFF\"/>");
                sb.AppendLine("<Interior ss:Color=\"#1F3864\" ss:Pattern=\"Solid\"/>");
                sb.AppendLine("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
                sb.AppendLine("</Style>");

                // Estilo encabezado
                sb.AppendLine("<Style ss:ID=\"sHeader\">");
                sb.AppendLine("<Font ss:Bold=\"1\" ss:Size=\"10\" ss:Color=\"#FFFFFF\"/>");
                sb.AppendLine("<Interior ss:Color=\"#2E75B6\" ss:Pattern=\"Solid\"/>");
                sb.AppendLine("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/>");
                sb.AppendLine("</Style>");

                // Estilo fila normal
                sb.AppendLine("<Style ss:ID=\"sRow\">");
                sb.AppendLine("<Font ss:Size=\"10\"/>");
                sb.AppendLine("<Interior ss:Color=\"#FFFFFF\" ss:Pattern=\"Solid\"/>");
                sb.AppendLine("<Alignment ss:Vertical=\"Center\"/>");
                sb.AppendLine("</Style>");

                // Estilo fila alterna
                sb.AppendLine("<Style ss:ID=\"sAlt\">");
                sb.AppendLine("<Font ss:Size=\"10\"/>");
                sb.AppendLine("<Interior ss:Color=\"#DEEAF1\" ss:Pattern=\"Solid\"/>");
                sb.AppendLine("<Alignment ss:Vertical=\"Center\"/>");
                sb.AppendLine("</Style>");

                // Estilo aprobado/rechazado
                sb.AppendLine("<Style ss:ID=\"sOk\">");
                sb.AppendLine("<Font ss:Bold=\"1\" ss:Color=\"#1E8449\"/>");
                sb.AppendLine("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
                sb.AppendLine("</Style>");
                sb.AppendLine("<Style ss:ID=\"sFail\">");
                sb.AppendLine("<Font ss:Bold=\"1\" ss:Color=\"#C0392B\"/>");
                sb.AppendLine("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>");
                sb.AppendLine("</Style>");
                sb.AppendLine("</Styles>");

                // ═══════════════════════════════════════
                // HOJA 1 – DIAGNÓSTICO POR ILUMINANTE
                // ═══════════════════════════════════════
                sb.AppendLine("<Worksheet ss:Name=\"Diagnóstico\">");
                sb.AppendLine("<Table>");

                // Título
                string shadeName = "N/A";
                var globalRes = Color.ShadeReportExtractor.LastResult;
                if (globalRes != null && !string.IsNullOrWhiteSpace(globalRes.ShadeName))
                    shadeName = globalRes.ShadeName;

                sb.AppendLine("<Row ss:Height=\"35\">");
                sb.AppendLine($"<Cell ss:MergeAcross=\"8\" ss:StyleID=\"sTitle\"><Data ss:Type=\"String\">REPORTE COLORIMÉTRICO  •  Shade: {Esc(shadeName)}  •  {DateTime.Now:dd/MM/yyyy HH:mm}</Data></Cell>");
                sb.AppendLine("</Row>");
                sb.AppendLine("<Row ss:Height=\"5\"/>");

                // Encabezados (9 columnas: sin deltas numéricos)
                sb.AppendLine("<Row ss:Height=\"28\">");
                foreach (var h in new[] {
                    "Iluminante", "Estado",
                    "Diagnóstico L*", "Corrección L*",
                    "Diagnóstico a*", "Corrección a*",
                    "Diagnóstico b*", "Corrección b*",
                    "Dominancia Cromática" })
                    sb.AppendLine($"<Cell ss:StyleID=\"sHeader\"><Data ss:Type=\"String\">{Esc(h)}</Data></Cell>");
                sb.AppendLine("</Row>");

                if (_resultsLegacy != null)
                {
                    // Usar el MISMO método de evaluación que el panel CMC 2:1 de la app
                    var evaluation = EngineCalc.EvaluateTolerance(_resultsLegacy, DE_MAX);

                    int ri = 0;
                    foreach (var r in _resultsLegacy)
                    {
                        // Buscar el check del iluminante actual (idéntico al panel CMC 2:1)
                        var check = evaluation.Checks.FirstOrDefault(c =>
                            string.Equals(c.Illuminant, r.Illuminant, StringComparison.OrdinalIgnoreCase));
                        bool ok = check != null ? check.Passes : true;

                        string est = ok ? "CUMPLE" : "NO CUMPLE";
                        string stRow = (ri % 2 == 0) ? "sRow" : "sAlt";
                        string stEst = ok ? "sOk" : "sFail";

                        // Porcentajes (redondeados a 1 decimal, valor absoluto)
                        double pctL = Math.Round(Math.Abs(r.PercentL), 1);
                        double pctA = Math.Round(Math.Abs(r.PercentA), 1);
                        double pctB = Math.Round(Math.Abs(r.PercentB), 1);

                        // ─── L* ───
                        string diagL = r.DeltaL < -0.01 ? "Lote más OSCURO" : r.DeltaL > 0.01 ? "Lote más CLARO" : "Sin desviación";
                        string corrL = r.DeltaL < -0.01 ? (pctL.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + "% ACLARAR") : r.DeltaL > 0.01 ? (pctL.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + "% OSCURECER") : "N/A";

                        // ─── a* ───
                        string diagA = "";
                        string corrA = "";
                        if (r.DeltaA > 0.01)       { diagA = "MÁS ROJO";    corrA = $"{pctA:0.0}% DISMINUIR ROJO o AUMENTAR VERDE"; }
                        else if (r.DeltaA < -0.01) { diagA = "MÁS VERDE";   corrA = $"{pctA:0.0}% DISMINUIR VERDE o AUMENTAR ROJO"; }
                        else                       { diagA = "Sin sesgo";    corrA = "N/A"; }

                        // ─── b* ───
                        string diagB = "";
                        string corrB = "";
                        if (r.DeltaB > 0.01)       { diagB = "MÁS AMARILLO"; corrB = $"{pctB:0.0}% DISMINUIR AMARILLO o AUMENTAR AZUL"; }
                        else if (r.DeltaB < -0.01) { diagB = "MÁS AZUL";     corrB = $"{pctB:0.0}% DISMINUIR AZUL o AUMENTAR AMARILLO"; }
                        else                       { diagB = "Sin sesgo";     corrB = "N/A"; }

                        // ─── Dominancia cromática ───
                        double modA = Math.Abs(r.DeltaA), modB = Math.Abs(r.DeltaB);
                        string dominancia;
                        if (modA > modB + 0.01)      dominancia = "Eje a* (Rojo↔Verde)";
                        else if (modB > modA + 0.01) dominancia = "Eje b* (Amarillo↔Azul)";
                        else if (modA + modB > 0.02) dominancia = "Mixta (a* y b* similares)";
                        else                         dominancia = "Sin dominancia";

                        sb.AppendLine("<Row ss:Height=\"20\">");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(r.Illuminant)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stEst}\"><Data ss:Type=\"String\">{Esc(est)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(diagL)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(corrL)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(diagA)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(corrA)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(diagB)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(corrB)}</Data></Cell>");
                        sb.AppendLine($"<Cell ss:StyleID=\"{stRow}\"><Data ss:Type=\"String\">{Esc(dominancia)}</Data></Cell>");
                        sb.AppendLine("</Row>");
                        ri++;
                    }
                }

                sb.AppendLine("</Table>");
                sb.AppendLine("</Worksheet>");

                // ═══════════════════════════════════════
                // HOJA 2 – REPORTE COMPLETO (texto)
                // ═══════════════════════════════════════
                sb.AppendLine("<Worksheet ss:Name=\"Reporte Completo\">");
                sb.AppendLine("<Table>");
                sb.AppendLine("<Row ss:Height=\"30\">");
                sb.AppendLine("<Cell ss:StyleID=\"sTitle\"><Data ss:Type=\"String\">REPORTE COMPLETO</Data></Cell>");
                sb.AppendLine("</Row>");

                foreach (var linea in (txtReport.Text + Environment.NewLine + txtRecomendacion.Text).Split('\n'))
                {
                    sb.AppendLine("<Row ss:Height=\"16\">");
                    sb.AppendLine($"<Cell ss:StyleID=\"sRow\"><Data ss:Type=\"String\">{Esc(linea.TrimEnd('\r'))}</Data></Cell>");
                    sb.AppendLine("</Row>");
                }

                sb.AppendLine("</Table>");
                sb.AppendLine("</Worksheet>");
                sb.AppendLine("</Workbook>");

                System.IO.File.WriteAllText(sfd.FileName, sb.ToString(), new System.Text.UTF8Encoding(true));
                MessageBox.Show("Reporte exportado a Excel correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Escapa caracteres especiales XML
        private static string Esc(string s) =>
            System.Security.SecurityElement.Escape(s ?? "");
    }
}