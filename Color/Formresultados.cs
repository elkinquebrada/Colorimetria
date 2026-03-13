using System;
using System.Collections.Generic;
using System.Drawing;
using DrawingColor = System.Drawing.Color;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;


// Aliases para evitar ambigüedad de tipos si hay clases con el mismo nombre en otros namespaces
using EngineCalc = Color.ColorimetricCalculator;
using EngineRow = Color.ColorimetricRow;
using EngineRes = Color.ColorCorrectionResult;

namespace Color
{
    public class FormResultados : Form
    {
        // ======= Datos de entrada =======
        private readonly OcrReport _report;
        private readonly string _resumenLegacy;
        private readonly List<EngineRes> _resultsLegacy;

        // ======= Controles de la vista =======
        private TextBox txtReport;
        private TextBox txtRecomendacion;
        private TextBox txtReceta;
        private SplitContainer splitMedicionesCmc; // IZQ: Receta/Reporte | DER: CMC/Recomendación
        private SplitContainer splitLeft;           // IZQ arriba: Receta | IZQ abajo: Reporte OCR
        private Button btnExportar;
        private Button btnCerrar;

        // ======= Datos receta (IlluminantCorrectionResult) =======
        private List<IlluminantCorrectionResult> _recipeResults;

        // ======= Tolerancias (L*, Hue y ΔE) =======
        private const double DL_MAX = 0.69;
        private const double DH_MAX = 0.69;
        private const double DE_MAX = 1.20;

        // ======= Proporción del divisor (55% izquierda) =======
        private double _splitLeftRatio = 0.55;

        // ======= Constructores =======
        public FormResultados(OcrReport report)
        {
            _report = report ?? new OcrReport();
            InitializeComponents();

            // Poblar vistas
            txtReceta.Text = BuildResumenFromReport(_report);
            txtReport.Text = BuildResumenFromReport(_report);
            txtRecomendacion.Text = BuildRecomendacionFromReport(_report);
        }

        public FormResultados(string resumen, List<EngineRes> results,
            List<IlluminantCorrectionResult> recipeResults = null)
        {
            _resumenLegacy = resumen ?? "";
            _resultsLegacy = results ?? new List<EngineRes>();
            _recipeResults = recipeResults;
            InitializeComponents();

            // Panel izquierdo: cálculos de receta (si hay) o resumen OCR
            txtReceta.Text = (_recipeResults != null && _recipeResults.Count > 0)
                ? RecipeCorrector.BuildSummaryText(_recipeResults)
                : _resumenLegacy;
            txtReport.Text = _resumenLegacy;
            txtRecomendacion.Text = BuildRecomendacionFromResults(_resultsLegacy);
        }

        // ======= Inicialización de la UI (layout elástico) =======
        private void InitializeComponents()
        {
            // ---- Ventana y escalado ----
            this.Text = "Resultados — Corrección Colorimétrica";
            this.FormBorderStyle = FormBorderStyle.Sizable; // permite min/max y resize
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ControlBox = true;
            this.ShowIcon = true;

            this.BackColor = DrawingColor.FromArgb(30, 30, 30);
            this.AutoScaleMode = AutoScaleMode.Dpi; // respeta 125%, 150%, etc.

            // Tamaño inicial cómodo (90% del área de trabajo)
            var wa = Screen.PrimaryScreen.WorkingArea;
            int targetWidth = (int)(wa.Width * 0.90);
            int targetHeight = (int)(wa.Height * 0.90);
            this.MinimumSize = new Size(980, 640);
            this.Size = new Size(targetWidth, targetHeight);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResizeRedraw = true;

            // ---- Título ----
            var lblTitulo = new Label
            {
                Text = "RESULTADOS DE CORRECCIÓN COLORIMÉTRICA",
                Font = new Font("Segoe UI", 13, FontStyle.Bold),
                ForeColor = DrawingColor.White,
                AutoSize = true,
                Location = new Point(20, 15)
            };

            // ---- Controles inferiores ----
            btnExportar = new Button
            {
                Text = "💾 Exportar .txt",
                Size = new Size(150, 38),
                BackColor = DrawingColor.FromArgb(0, 100, 180),
                ForeColor = DrawingColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            btnExportar.FlatAppearance.BorderSize = 0;
            btnExportar.Click += BtnExportar_Click;

            btnCerrar = new Button
            {
                Text = "Cerrar",
                Size = new Size(120, 38),
                BackColor = DrawingColor.FromArgb(60, 60, 60),
                ForeColor = DrawingColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10)
            };
            btnCerrar.FlatAppearance.BorderSize = 0;
            btnCerrar.Click += (s, e) => this.Close();

            // ---- Split central (Dock=Fill) ----
            splitMedicionesCmc = new SplitContainer
            {
                Dock = DockStyle.Fill,                // clave para crecer con la ventana
                Orientation = Orientation.Vertical,   // izquierda/derecha
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };

            // Panel IZQUIERDO: solo Cálculos de Receta
            var lblReceta = new Label
            {
                Text = "🧪 Cálculos de Corrección de Receta",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = DrawingColor.White,
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(4, 6, 4, 6),
                BackColor = DrawingColor.FromArgb(35, 35, 35)
            };
            txtReceta = BuildTextBox(DrawingColor.FromArgb(200, 230, 255));
            txtReceta.Dock = DockStyle.Fill;
            txtReport = BuildTextBox(null); // mantener campo para exportación
            var panelLeft = new Panel { Dock = DockStyle.Fill, BackColor = DrawingColor.FromArgb(45, 45, 45) };
            panelLeft.Controls.Add(txtReceta);
            panelLeft.Controls.Add(lblReceta);

            splitMedicionesCmc.Panel1.Controls.Add(panelLeft);

            // Panel DERECHO: CMC (2:1) / Recomendación (encabezado + textbox)
            var panelRightHeader = new Label
            {
                Text = "✅ CMC (2:1) / Recomendación",
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = DrawingColor.White,
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(4, 6, 4, 6),
                BackColor = DrawingColor.FromArgb(35, 35, 35)
            };

            txtRecomendacion = BuildTextBox(DrawingColor.FromArgb(255, 220, 100)); // ámbar
            txtRecomendacion.Dock = DockStyle.Fill;

            var panelRight = new Panel { Dock = DockStyle.Fill, BackColor = DrawingColor.FromArgb(45, 45, 45) };
            panelRight.Controls.Add(txtRecomendacion);
            panelRight.Controls.Add(panelRightHeader);

            splitMedicionesCmc.Panel2.Controls.Add(panelRight);

            // ---- Root layout (TableLayoutPanel) ----
            var panelRoot = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = this.BackColor
            };
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));  // Título
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // Centro (Split)
            panelRoot.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));  // Botones

            // Fila 0: Título
            var panelHeader = new Panel { Dock = DockStyle.Fill, BackColor = this.BackColor };
            panelHeader.Controls.Add(lblTitulo);
            panelRoot.Controls.Add(panelHeader, 0, 0);

            // Fila 1: Split
            panelRoot.Controls.Add(splitMedicionesCmc, 0, 1);

            // Fila 2: Botones a la derecha (Exportar | Cerrar)
            var panelButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 10, 20, 10),
                BackColor = this.BackColor
            };
            panelButtons.Controls.Add(btnCerrar);
            panelButtons.Controls.Add(btnExportar);
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

            // Mantener proporción del divisor en Load/Resize
            this.Load += (s, e) => ApplySplitRatio();
            this.Resize += (s, e) => ApplySplitRatio();
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
                    distance = Math.Max(200, Math.Min(w - 200, distance));
                    splitMedicionesCmc.SplitterDistance = distance;
                }

            }
            catch { }
        }

        // ======= Factory de TextBox monoespaciado (oscuro) =======
        private static TextBox BuildTextBox(DrawingColor? foreColor)
        {
            return new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.0f),
                BackColor = DrawingColor.FromArgb(20, 20, 20),
                ForeColor = foreColor ?? DrawingColor.LightGreen,
                ReadOnly = true,
                WordWrap = false,
                Text = string.Empty
            };
        }

        // =========================================================
        // RECOMENDACIÓN — desde lista de resultados (Croma eliminado del texto impreso)
        // =========================================================
        private static string BuildRecomendacionFromResults(List<EngineRes> results)
        {
            if (results == null || results.Count == 0)
                return "No hay resultados para generar recomendación.";

            var sb = new StringBuilder();

            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            sb.AppendLine(" RECOMENDACIÓN FINAL INTEGRADA");
            sb.AppendLine(" Fecha: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture));
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            sb.AppendLine();

            // --- Encabezado de tolerancias (sin imprimir Croma) ---
            sb.AppendLine("ESTADO L/ΔE (tolerancias):");
            sb.AppendLine(string.Format(
                CultureInfo.InvariantCulture,
                " DL≤{0:0.00} DH≤{1:0.00} DE≤{2:0.00}",
                DL_MAX, DH_MAX, DE_MAX));
            sb.AppendLine();

            bool cumpleTodo = true;

            // % formateado con signo, sin '+'
            Func<double, string> FmtPctSigned = v =>
                double.IsNaN(v) ? "N/D" : (v.ToString("0;-0;0", CultureInfo.InvariantCulture) + "%");

            // ---- 1) ESTADO por iluminante (sin DC impreso) ----
            foreach (var r in results)
            {
                // Lógica de pass: DL, DH y DE (sin DC)
                bool pass =
                    Math.Abs(r.DeltaL) <= DL_MAX &&
                    Math.Abs(r.DeltaHue) <= DH_MAX &&
                    r.DeltaE <= DE_MAX;

                if (!pass) cumpleTodo = false;

                // Imprimimos SOLO DE y DL (sin DC)
                sb.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    "{0,-6} -> {1} (DE={2:0.00} DL={3:+0.00;-0.00})",
                    r.Illuminant,
                    pass ? "CUMPLE " : "NO CUMPLE",
                    r.DeltaE,
                    r.DeltaL
                ));
            }

            // ---- 2) TABLA COMPACTA de % a corregir (L/A/B) ----
            sb.AppendLine();
            sb.AppendLine("% a corregir (por iluminante):");
            sb.AppendLine(string.Format(
                CultureInfo.InvariantCulture,
                " {0,-8} {1,10} {2,12} {3,12}",
                "Illum", "%L", "%A", "%B"));

            foreach (var r in results)
            {
                sb.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    " {0,-8} {1,10} {2,12} {3,12}",
                    r.Illuminant,
                    FmtPctSigned(r.PercentL),
                    FmtPctSigned(r.PercentA),
                    FmtPctSigned(r.PercentB)
                ));
            }

            sb.AppendLine();
            sb.AppendLine("───────────────────────────────────────────────────────────────────");

            if (cumpleTodo)
            {
                sb.AppendLine();
                sb.AppendLine(" Todos los iluminantes cumplen las tolerancias definidas.");
                sb.AppendLine(" Recomendación final: NO SE REQUIERE CORRECCIÓN.");
                return sb.ToString();
            }

            // ---- 3) DIAGNÓSTICO por iluminante (formato solicitado) ----
            sb.AppendLine();
            sb.AppendLine("DIAGNOSTICO POR ILUMINANTE (Lot vs Std):");
            sb.AppendLine();

            foreach (var r in results)
            {
                sb.AppendLine(" [" + r.Illuminant + "]");

                // 👉 L*, a*, b* con % y acción en el formato que pediste
                sb.Append(BuildPerAxisPercentAdvice(r));

                // (Opcional) Si quieres mostrar también la línea "Ejes L*, a*, b* ... % a corregir",
                // descomenta la siguiente línea:
                // sb.Append(BuildAxesBlock(r));

                // Plano polar (a*, b*) + dominancia
                sb.Append(BuildPlanoPolarAdvice(r));

                sb.AppendLine();
            }

            sb.AppendLine("───────────────────────────────────────────────────────────────────");
            sb.AppendLine("Tras el ajuste, se recomienda re-medición bajo todos los iluminantes.");
            sb.AppendLine("═══════════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        // ======= Helper: imprime L*, a*, b* con % y acción (formato específico) =======
        private static string BuildPerAxisPercentAdvice(EngineRes r)
        {
            var sb = new StringBuilder();

            // L* → "5 %"
            Func<double, string> FmtPctL = v =>
            {
                if (double.IsNaN(v)) return "0 %";
                double iv = Math.Round(Math.Abs(v), 0, MidpointRounding.AwayFromZero);
                return iv.ToString("0", CultureInfo.InvariantCulture) + " %";
            };

            // a*/b* → "6%" / "14%"
            Func<double, string> FmtPctNoSpace = v =>
            {
                if (double.IsNaN(v)) return "0%";
                double iv = Math.Round(Math.Abs(v), 0, MidpointRounding.AwayFromZero);
                return iv.ToString("0", CultureInfo.InvariantCulture) + "%";
            };

            // --- L* (claro/oscuro) con % y acción
            string pctL = FmtPctL(r.PercentL);
            if (r.DeltaL < -0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * L*: Lote más OSCURO  -> {0} Corrección: ACLARAR", pctL));
            }
            else if (r.DeltaL > 0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * L*: Lote más CLARO  -> {0} Corrección: OSCURECER", pctL));
            }
            else
            {
                sb.AppendLine(" * L*: Sin desviación (DL ≈ 0)");
            }

            // --- a* (rojo/verde) con % y acción (usa "a*=")
            string pctA = FmtPctNoSpace(r.PercentA);
            if (r.DeltaA > 0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * a*= MÁS ROJO   → Corrección: {0} DISMINUIR ROJO o AUMENTAR VERDE.", pctA));
            }
            else if (r.DeltaA < -0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * a*= MÁS VERDE  → Corrección: {0} DISMINUIR VERDE o AUMENTAR ROJO.", pctA));
            }
            else
            {
                sb.AppendLine(" * a*: Sin sesgo (Δa ≈ 0).");
            }

            // --- b* (amarillo/azul) con % y acción (usa "b*:")
            string pctB = FmtPctNoSpace(r.PercentB);
            if (r.DeltaB > 0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * b*: MÁS AMARILLO  → Corrección: {0} DISMINUIR AMARILLO o AUMENTAR AZUL.", pctB));
            }
            else if (r.DeltaB < -0.01)
            {
                sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                    " * b*: MÁS AZUL  → Corrección: {0} DISMINUIR AZUL o AUMENTAR AMARILLO.", pctB));
            }
            else
            {
                sb.AppendLine(" * b*: Sin sesgo (Δb ≈ 0).");
            }

            return sb.ToString();
        }

        // ======= (Opcional) Bloque de ejes L*, a*, b* con Δ y % a corregir =======
        private static string BuildAxesBlock(EngineRes r)
        {
            var sb = new StringBuilder();

            Func<double, string> FmtPct = v =>
                double.IsNaN(v) ? "N/D" : (v.ToString("0;-0;0", CultureInfo.InvariantCulture) + "%");

            // Línea de deltas (signadas)
            sb.AppendLine(string.Format(
                CultureInfo.InvariantCulture,
                " * Ejes L*, a*, b*: ΔL={0:+0.00;-0.00}, Δa={1:+0.00;-0.00}, Δb={2:+0.00;-0.00}",
                r.DeltaL, r.DeltaA, r.DeltaB));

            // Línea de porcentajes a corregir por iluminante
            sb.AppendLine(string.Format(
                CultureInfo.InvariantCulture,
                "   % a corregir (este iluminante): L={0}  A={1}  B={2}",
                FmtPct(r.PercentL), FmtPct(r.PercentA), FmtPct(r.PercentB)));

            return sb.ToString();
        }

        // ======= Diagnóstico en plano polar (a*, b*) =======
        private static string BuildPlanoPolarAdvice(EngineRes r)
        {
            var sb = new StringBuilder();

            double da = r.DeltaA;     // +a* = más ROJO ;  -a* = más VERDE
            double db = r.DeltaB;     // +b* = más AMARILLO ; -b* = más AZUL
            double eps = 0.01;        // umbral de “casi cero”

            // Ángulo polar del vector (Δa, Δb) y módulo en el plano a*-b*
            double angleRad = Math.Atan2(db, da); // y = Δb , x = Δa
            double angleDeg = angleRad * 180.0 / Math.PI;
            if (angleDeg < 0) angleDeg += 360.0;
            double modulo = Math.Sqrt(da * da + db * db);

            // Cuadrantes para contexto
            string cuadrante;
            if (da >= 0 && db >= 0) cuadrante = "rojo‑amarillo (+a,+b)";
            else if (da < 0 && db >= 0) cuadrante = "verde‑amarillo (−a,+b)";
            else if (da < 0 && db < 0) cuadrante = "verde‑azul (−a,−b)";
            else cuadrante = "rojo‑azul (+a,−b)";

            sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                " Plano polar (a*, b*): módulo={0:0.00}, ángulo={1:0.0}° → cuadrante {2}.",
                modulo, angleDeg, cuadrante));

            // Dominancia cromática (qué eje pesa más)
            if (Math.Abs(da) > Math.Abs(db) + eps)
                sb.AppendLine("   Dominancia cromática: eje a* (Rojo↔Verde). Prioriza la corrección sobre ROJO/VERDE.");
            else if (Math.Abs(db) > Math.Abs(da) + eps)
                sb.AppendLine("   Dominancia cromática: eje b* (Amarillo↔Azul). Prioriza la corrección sobre AMARILLO/AZUL.");
            else if (modulo > eps)
                sb.AppendLine("   Dominancia cromática: mixta (a* y b* similares). Corrige en ambos ejes.");

            return sb.ToString();
        }

        // =========================================================
        // RECOMENDACIÓN — desde OcrReport (usa el motor y reusa arriba)
        // =========================================================
        private static string BuildRecomendacionFromReport(OcrReport rep)
        {
            // 1) Validaciones básicas
            if (rep == null || rep.Measures == null || rep.Measures.Count == 0)
                return "No hay datos en el reporte para generar recomendación.";

            // 2) Convertir a filas para el motor (Std/Lot por iluminante)
            List<EngineRow> rowsForEngine = rep.Measures.Select(m => new EngineRow
            {
                Illuminant = m.Illuminant, // D65 / TL84 / A...
                Type = m.Type,       // "Std" o "Lot"
                L = m.L,
                A = m.A,
                B = m.B,
                Hue = m.Hue
            }).ToList();

            // 3) Calcular con el motor (List<ColorimetricRow> -> List<CorrectionResult>)
            List<EngineRes> calcResults = EngineCalc.Calculate(rowsForEngine);

            // 4) Reutilizar el generador de texto unificado
            return BuildRecomendacionFromResults(calcResults);
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
        private void BtnExportar_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Texto (*.txt)|*.txt|Todos los archivos|*.*";
                sfd.FileName = "Colorimetria_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string contenido =
                        "=== CÁLCULOS DE RECETA ===" + Environment.NewLine
                        + (txtReceta != null ? txtReceta.Text : string.Empty)
                        + Environment.NewLine + Environment.NewLine
                        + "=== REPORTE OCR ===" + Environment.NewLine
                        + (txtReport != null ? txtReport.Text : string.Empty)
                        + Environment.NewLine + Environment.NewLine
                        + "=== CMC / RECOMENDACIÓN ===" + Environment.NewLine
                        + (txtRecomendacion != null ? txtRecomendacion.Text : string.Empty);

                    System.IO.File.WriteAllText(sfd.FileName, contenido, Encoding.UTF8);

                    MessageBox.Show("Archivo guardado correctamente.", "Exportado",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}