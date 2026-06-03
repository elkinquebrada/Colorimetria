using System;
using System.Drawing;
using System.Windows.Forms;
using Color.Services;

namespace Color
{
    /// <summary>
    /// Bloque de reporte por iluminante.
    /// Layout horizontal (izquierda → derecha):
    ///   [Tabla L,a,b | Tabla Chroma/Hue | Tabla Lightness/Chroma/Hue/CMC + Iluminante]
    /// Cada sección incluye su fila de "Acción" en la parte inferior.
    /// </summary>
    public class IluminantReportBlock : UserControl
    {
        // --- Grillas Lab ---
        private DataGridView dgvLab;
        private DataGridView dgvActions;

        // --- Grilla Chroma/Hue ---
        private DataGridView dgvChromaHue;
        private DataGridView dgvChromaActions;

        // --- Panel CMC (Labels individuales) ---
        private Label lblCmcValue;
        private Label lblLightness;
        private Label lblChromaVal;
        private Label lblHueVal;
        private Label lblIlluminantName;

        // Tolerancias activas
        private double _deAprobado = 1.20;
        private double _dlAprobado = 1.0;
        private double _dcAprobado = 1.0;
        private double _dhAprobado = 1.0;
        private ColorCorrectionResult _lastResult = null;

        // Colores corporativos
        private static readonly System.Drawing.Color ColBlue    = System.Drawing.Color.FromArgb(0, 122, 204);
        private static readonly System.Drawing.Color ColWhite   = System.Drawing.Color.White;
        private static readonly System.Drawing.Color ColBlack   = System.Drawing.Color.Black;
        private static readonly System.Drawing.Color ColPassBg  = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
        private static readonly System.Drawing.Color ColPassFg  = System.Drawing.Color.FromArgb(0, 102, 0);
        private static readonly System.Drawing.Color ColFailBg  = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
        private static readonly System.Drawing.Color ColFailFg  = System.Drawing.Color.FromArgb(153, 0, 0);

        public IluminantReportBlock()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // El bloque tiene un alto fijo de 115px para 4 filas de datos + cabecera + acción
            this.Size        = new Size(1000, 115);
            this.BackColor   = ColWhite;
            this.Margin      = new Padding(0, 0, 0, 10);
            this.BorderStyle = BorderStyle.FixedSingle;

            // ===================================================================
            // LAYOUT PRINCIPAL: 3 columnas → Lab | ChromaHue | CMC+Iluminante
            // ===================================================================
            var root = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 3,
                RowCount    = 1,
                Padding     = new Padding(0),
                Margin      = new Padding(0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };
            // Lab ocupa ~53%, ChromaHue ~22%, CMC+Illuminant ~25%
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 53));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // ===================================================================
            // COLUMNA 1 – Tabla L,a,b  (arriba) + Tabla Acción (abajo)
            // ===================================================================
            var colLeft = new TableLayoutPanel
            {
                Dock     = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                Margin   = new Padding(0),
                Padding  = new Padding(0)
            };
            colLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 68)); // Lab
            colLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 32)); // Acción

            dgvLab     = BuildGrid(4);
            dgvActions = BuildGrid(3);

            // ─ Cabecera Lab ─
            var rHeader = dgvLab.Rows[dgvLab.Rows.Add("D65", "L", "a", "b")];
            StyleHeaderRow(rHeader, ColBlue, ColWhite, true);

            // ─ Filas de datos ─
            dgvLab.Rows.Add("Std", "0.00", "0.00", "0.00");
            dgvLab.Rows.Add("Lot", "0.00", "0.00", "0.00");
            var rDelta = dgvLab.Rows[dgvLab.Rows.Add("Δ", "0.00", "0.00", "0.00")];
            rDelta.DefaultCellStyle.Font      = new Font("Segoe UI", 8, FontStyle.Bold);
            rDelta.DefaultCellStyle.ForeColor = ColBlue;

            // ─ Ajuste ancho de la primera columna (D65/Std/Lot/Delta) ─
            dgvLab.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvLab.Columns[0].Width        = 40;

            // ─ Tabla Acción ─
            var rAcHead = dgvActions.Rows[dgvActions.Rows.Add("Aumentar []", "Aumentar Verde", "Aumentar Azul")];
            StyleHeaderRow(rAcHead, ColBlue, ColWhite, false);
            dgvActions.Rows.Add("-1%", "0%", "0%");
            dgvActions.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            dgvActions.Columns[0].Width        = 40;

            colLeft.Controls.Add(dgvLab,     0, 0);
            colLeft.Controls.Add(dgvActions, 0, 1);

            // ===================================================================
            // COLUMNA 2 – Tabla Chroma/Hue (arriba) + Tabla Acción (abajo)
            // ===================================================================
            var colMid = new TableLayoutPanel
            {
                Dock     = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                Margin   = new Padding(0),
                Padding  = new Padding(0)
            };
            colMid.RowStyles.Add(new RowStyle(SizeType.Percent, 68));
            colMid.RowStyles.Add(new RowStyle(SizeType.Percent, 32));

            dgvChromaHue     = BuildGrid(2);
            dgvChromaActions = BuildGrid(2);

            // ─ Cabecera Chroma/Hue ─
            var rCHead = dgvChromaHue.Rows[dgvChromaHue.Rows.Add("Chroma", "Hue")];
            StyleHeaderRow(rCHead, ColBlue, ColWhite, true);
            dgvChromaHue.Rows.Add("0.00", "0.00");
            dgvChromaHue.Rows.Add("0.00", "0.00");
            var rDC = dgvChromaHue.Rows[dgvChromaHue.Rows.Add("0.00", "0.00")];
            rDC.DefaultCellStyle.Font      = new Font("Segoe UI", 8, FontStyle.Bold);
            rDC.DefaultCellStyle.ForeColor = ColBlue;

            // ─ Tabla Acción Chroma ─
            var rCaHead = dgvChromaActions.Rows[dgvChromaActions.Rows.Add("Brighter", "Yellower (Redder)")];
            StyleHeaderRow(rCaHead, System.Drawing.Color.FromArgb(80, 80, 80), ColWhite, false);
            dgvChromaActions.Rows.Add("Duller", "Bluer (Greener)");
            var rCaData = dgvChromaActions.Rows[dgvChromaActions.Rows.Count - 1];
            rCaData.Cells[0].Style.ForeColor = System.Drawing.Color.FromArgb(0, 0, 180);
            rCaData.Cells[1].Style.ForeColor = System.Drawing.Color.FromArgb(0, 0, 180);

            colMid.Controls.Add(dgvChromaHue,     0, 0);
            colMid.Controls.Add(dgvChromaActions, 0, 1);

            // ===================================================================
            // COLUMNA 3 – Panel CMC  (5 columnas: lightnees | Chroma | Hue | cmc(2:1) + Illuminant)
            // ===================================================================
            var colRight = new TableLayoutPanel
            {
                Dock        = DockStyle.Fill,
                ColumnCount = 5,
                RowCount    = 4,
                Margin      = new Padding(0),
                Padding     = new Padding(0),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };
            colRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            colRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            colRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            colRight.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            colRight.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 50));

            colRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // fila 0: headers
            colRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // fila 1: vacía (Std)
            colRight.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // fila 2: vacía (Lot)
            colRight.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // fila 3: valores CMC

            // Fila 0 – Cabeceras
            colRight.Controls.Add(MakeHeaderLabel("lightnees"), 0, 0);
            colRight.Controls.Add(MakeHeaderLabel("Chroma"),    1, 0);
            colRight.Controls.Add(MakeHeaderLabel("Hue"),       2, 0);
            colRight.Controls.Add(MakeHeaderLabel("cmc (2:1)"), 3, 0);

            // Etiqueta del iluminante (span vertical filas 0-3 en la columna 4)
            lblIlluminantName = new Label
            {
                Text      = "D65",
                BackColor = ColBlue,
                ForeColor = ColWhite,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = new Font("Segoe UI", 12, FontStyle.Bold)
            };
            colRight.Controls.Add(lblIlluminantName, 4, 0);
            colRight.SetRowSpan(lblIlluminantName, 4);

            // Filas 1 y 2 – Celdas vacías (Std, Lot – sin datos en el CMC block)
            for (int fila = 1; fila <= 2; fila++)
                for (int col = 0; col < 4; col++)
                    colRight.Controls.Add(MakeEmptyCell(), col, fila);

            // Fila 3 – Valores semafóricos CMC
            lblLightness  = MakeCmcValueLabel("0.00");
            lblChromaVal  = MakeCmcValueLabel("0.00");
            lblHueVal     = MakeCmcValueLabel("0.00");
            lblCmcValue   = MakeCmcValueLabel("0.00");

            colRight.Controls.Add(lblLightness,  0, 3);
            colRight.Controls.Add(lblChromaVal,  1, 3);
            colRight.Controls.Add(lblHueVal,     2, 3);
            colRight.Controls.Add(lblCmcValue,   3, 3);

            // ===================================================================
            // ENSAMBLE FINAL
            // ===================================================================
            root.Controls.Add(colLeft,  0, 0);
            root.Controls.Add(colMid,   1, 0);
            root.Controls.Add(colRight, 2, 0);

            this.Controls.Add(root);
        }

        // ─────────────────────────────────────────────────────────────────────
        // FACTORIES
        // ─────────────────────────────────────────────────────────────────────

        private DataGridView BuildGrid(int cols)
        {
            var dgv = new DataGridView
            {
                Dock                    = DockStyle.Fill,
                ColumnCount             = cols,
                BackgroundColor         = ColWhite,
                BorderStyle             = BorderStyle.None,
                CellBorderStyle         = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor               = System.Drawing.Color.FromArgb(210, 210, 210),
                ColumnHeadersVisible    = false,
                RowHeadersVisible       = false,
                AllowUserToAddRows      = false,
                ReadOnly                = true,
                AutoSizeColumnsMode     = DataGridViewAutoSizeColumnsMode.Fill,
                Font                    = new Font("Segoe UI", 8),
                ScrollBars              = ScrollBars.None,
                DefaultCellStyle        = { SelectionBackColor = ColWhite, SelectionForeColor = ColBlack }
            };
            return dgv;
        }

        private static void StyleHeaderRow(DataGridViewRow row,
                                           System.Drawing.Color bg,
                                           System.Drawing.Color fg,
                                           bool bold)
        {
            row.DefaultCellStyle.BackColor = bg;
            row.DefaultCellStyle.ForeColor = fg;
            if (bold)
                row.DefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
        }

        private Label MakeHeaderLabel(string text)
        {
            return new Label
            {
                Text      = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = new Font("Segoe UI", 8, FontStyle.Bold),
                BackColor = ColBlue,
                ForeColor = ColWhite,
                Margin    = new Padding(0)
            };
        }

        private Label MakeEmptyCell()
        {
            return new Label
            {
                Dock      = DockStyle.Fill,
                BackColor = ColWhite,
                Margin    = new Padding(0)
            };
        }

        private Label MakeCmcValueLabel(string text)
        {
            return new Label
            {
                Text      = text,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock      = DockStyle.Fill,
                Font      = new Font("Segoe UI", 10, FontStyle.Bold),
                BackColor = ColPassBg,
                ForeColor = ColPassFg,
                Margin    = new Padding(0)
            };
        }

        // ─────────────────────────────────────────────────────────────────────
        // ACTUALIZACIÓN DE TOLERANCIAS
        // ─────────────────────────────────────────────────────────────────────
        public void UpdateTolerances(double de, double dl, double dc, double dh)
        {
            _deAprobado = de;
            _dlAprobado = dl;
            _dcAprobado = dc;
            _dhAprobado = dh;
            if (_lastResult != null) UpdateData(_lastResult);
        }

        // ─────────────────────────────────────────────────────────────────────
        // POBLACIÓN DE DATOS
        // ─────────────────────────────────────────────────────────────────────
        public void UpdateData(ColorCorrectionResult res)
        {
            if (res == null) return;
            _lastResult = res;

            // ── Nombre del iluminante ──
            string illum = !string.IsNullOrEmpty(res.Illuminant)
                           ? res.Illuminant.ToUpper() : "D65";
            lblIlluminantName.Text = illum;
            if (dgvLab.Rows.Count > 0)
                dgvLab.Rows[0].Cells[0].Value = illum;

            // ── Tabla L, a, b ──
            // Fila 0 = cabecera (ya tiene el nombre del iluminante)
            dgvLab.Rows[1].Cells[1].Value = res.StdL.ToString("F2");
            dgvLab.Rows[1].Cells[2].Value = res.StdA.ToString("F2");
            dgvLab.Rows[1].Cells[3].Value = res.StdB.ToString("F2");

            dgvLab.Rows[2].Cells[1].Value = res.LotL.ToString("F2");
            dgvLab.Rows[2].Cells[2].Value = res.LotA.ToString("F2");
            dgvLab.Rows[2].Cells[3].Value = res.LotB.ToString("F2");

            dgvLab.Rows[3].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaL);
            dgvLab.Rows[3].Cells[2].Value = ColorimetricCalculator.FormatDelta(res.DeltaA);
            dgvLab.Rows[3].Cells[3].Value = ColorimetricCalculator.FormatDelta(res.DeltaB);

            // ── Tabla Chroma / Hue ──
            dgvChromaHue.Rows[1].Cells[0].Value = res.StdC.ToString("F2");
            dgvChromaHue.Rows[1].Cells[1].Value = res.StdH.ToString("F2");
            dgvChromaHue.Rows[2].Cells[0].Value = res.LotC.ToString("F2");
            dgvChromaHue.Rows[2].Cells[1].Value = res.LotH.ToString("F2");
            dgvChromaHue.Rows[3].Cells[0].Value = ColorimetricCalculator.FormatDelta(res.DeltaChroma);
            dgvChromaHue.Rows[3].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaHue);

            // ── Tabla Acción inferior (Lab) ──
            dgvActions.Rows[1].Cells[0].Value = FormatPct(res.PercentL);
            dgvActions.Rows[1].Cells[1].Value = FormatPct(res.PercentA);
            dgvActions.Rows[1].Cells[2].Value = FormatPct(res.PercentB);

            // ── Valores CMC semafóricos ──
            SetCmcLabel(lblLightness, res.CmcLightness, _dlAprobado);
            SetCmcLabel(lblChromaVal, res.CmcChroma,    _dcAprobado);
            SetCmcLabel(lblHueVal,    res.CmcHue,       _dhAprobado);
            SetCmcLabel(lblCmcValue,  res.CmcValue,     _deAprobado, absolute: false);

            lblIlluminantName.Refresh();
        }

        private static string FormatPct(double val)
        {
            int rounded = (int)Math.Round(val, MidpointRounding.AwayFromZero);
            return rounded + "%";
        }

        private static void SetCmcLabel(Label lbl, double value, double tolerance, bool absolute = true)
        {
            bool fail = absolute
                        ? Math.Abs(value) > tolerance
                        : value > tolerance;

            lbl.Text      = value.ToString("F2");
            lbl.BackColor = fail ? ColFailBg : ColPassBg;
            lbl.ForeColor = fail ? ColFailFg : ColPassFg;
        }
    }
}