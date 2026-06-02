using System;
using System.Drawing;
using System.Windows.Forms;
using Color.Services;

namespace Color
{
    public class IluminantReportBlock : UserControl
    {
        private Label lblIlluminantName;
        private DataGridView dgvLab;
        private DataGridView dgvActions;
        private DataGridView dgvChromaHue;

        private Label lblCmcValue;
        private Label lblLightness, lblChroma, lblHue;

        private double _deAprobado = 1.20;
        private double _dlAprobado = 1.0;
        private double _dcAprobado = 1.0;
        private double _dhAprobado = 1.0;
        private ColorCorrectionResult _lastResult = null;

        public IluminantReportBlock()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Ajustamos las dimensiones generales del bloque de reporte
            this.Size = new Size(1000, 140);
            this.BackColor = System.Drawing.Color.White;

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35)); // L,a,b y Acción
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20)); // C,H
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35)); // CMC Summary
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65)); // Ancho fijo perfecto para el iluminante lateral

            // =================================================================
            // --- Bloque 1: L,a,b y Acciones ---
            // =================================================================
            var pnlLeft = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, Margin = new Padding(2) };
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 65));
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Percent, 35));

            dgvLab = CreateBaseGrid(4);
            dgvLab.Columns[0].Width = 55;

            // Fila 0: Cabecera Azul de Títulos
            int headerLabIdx = dgvLab.Rows.Add("D65", "L", "a", "b");
            dgvLab.Rows[headerLabIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            dgvLab.Rows[headerLabIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvLab.Rows[headerLabIdx].DefaultCellStyle.Font = new Font(dgvLab.Font, FontStyle.Bold);

            // Filas de datos estructurales
            dgvLab.Rows.Add("Std", "0.00", "0.00", "0.00"); // Fila 1
            dgvLab.Rows.Add("Lot", "0.00", "0.00", "0.00"); // Fila 2

            int deltaLabIdx = dgvLab.Rows.Add("Δ", "0.00", "0.00", "0.00"); // Fila 3
            dgvLab.Rows[deltaLabIdx].DefaultCellStyle.Font = new Font(dgvLab.Font, FontStyle.Bold);

            // Tabla de acciones inferior
            dgvActions = CreateBaseGrid(3);
            dgvActions.Rows.Add("Aumentar []", "Aumentar Verde", "Aumentar Azul");
            dgvActions.Rows.Add("0%", "0%", "0%");
            dgvActions.Rows[0].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            dgvActions.Rows[0].DefaultCellStyle.ForeColor = System.Drawing.Color.White;

            pnlLeft.Controls.Add(dgvLab, 0, 0);
            pnlLeft.Controls.Add(dgvActions, 0, 1);

            // =================================================================
            // --- Bloque 2: Chroma/Hue ---
            // =================================================================
            var pnlMiddle = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 1, Margin = new Padding(2) };
            dgvChromaHue = CreateBaseGrid(2);

            // Fila 0: Cabecera Azul
            int headerChromaIdx = dgvChromaHue.Rows.Add("Chroma", "Hue");
            dgvChromaHue.Rows[headerChromaIdx].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            dgvChromaHue.Rows[headerChromaIdx].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvChromaHue.Rows[headerChromaIdx].DefaultCellStyle.Font = new Font(dgvChromaHue.Font, FontStyle.Bold);

            // Filas de datos
            dgvChromaHue.Rows.Add("0.00", "0.00"); // Fila 1: Std
            dgvChromaHue.Rows.Add("0.00", "0.00"); // Fila 2: Lot

            int deltaChromaIdx = dgvChromaHue.Rows.Add("0.00", "0.00"); // Fila 3: Delta
            dgvChromaHue.Rows[deltaChromaIdx].DefaultCellStyle.Font = new Font(dgvChromaHue.Font, FontStyle.Bold);

            pnlMiddle.Controls.Add(dgvChromaHue, 0, 0);

            // =================================================================
            // --- Bloque 3 y 4 Unificados: CMC Summary e Iluminante ---
            // =================================================================
            var pnlCmcCombo = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 5, Margin = new Padding(2) };
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 65));

            // Ajustamos las alturas para que sean idénticas a los bloques 1 y 2
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // Header
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // Fila Vacía (Std)
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // Fila Vacía (Lot)
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 22)); // Fila Delta (Data)
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Espacio restante

            var lblTitleLightness = CreateHeaderLabel("lightnees");
            var lblTitleChroma = CreateHeaderLabel("Chroma");
            var lblTitleHue = CreateHeaderLabel("Hue");
            var lblTitleCmc = CreateHeaderLabel("cmc (2:1)");
            var pnlTopSpacer = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.White };

            pnlCmcCombo.Controls.Add(lblTitleLightness, 0, 0);
            pnlCmcCombo.Controls.Add(lblTitleChroma, 1, 0);
            pnlCmcCombo.Controls.Add(lblTitleHue, 2, 0);
            pnlCmcCombo.Controls.Add(lblTitleCmc, 3, 0);
            pnlCmcCombo.Controls.Add(pnlTopSpacer, 4, 0);

            // Celdas vacías para conservar la estética de parrilla
            for (int r = 1; r <= 2; r++)
            {
                for (int c = 0; c < 4; c++)
                {
                    pnlCmcCombo.Controls.Add(CreateValueLabel("", false), c, r);
                }
            }

            lblLightness = CreateValueLabel("0.00", true);
            lblChroma = CreateValueLabel("0.00", true);
            lblHue = CreateValueLabel("0.00", true);
            lblCmcValue = CreateValueLabel("0.00", true);

            pnlCmcCombo.Controls.Add(lblLightness, 0, 3);
            pnlCmcCombo.Controls.Add(lblChroma, 1, 3);
            pnlCmcCombo.Controls.Add(lblHue, 2, 3);
            pnlCmcCombo.Controls.Add(lblCmcValue, 3, 3);

            lblIlluminantName = new Label
            {
                Text = "D65",
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 14, FontStyle.Bold)
            };

            pnlCmcCombo.Controls.Add(lblIlluminantName, 4, 1);
            pnlCmcCombo.SetRowSpan(lblIlluminantName, 3); // Ocupará la altura de Std + Lot + Delta

            // --- Ensamblado General ---
            mainLayout.Controls.Add(pnlLeft, 0, 0);
            mainLayout.Controls.Add(pnlMiddle, 1, 0);
            mainLayout.Controls.Add(pnlCmcCombo, 2, 0);
            mainLayout.SetColumnSpan(pnlCmcCombo, 2); // Ocupa la columna del CMC y la del Iluminante

            this.Controls.Add(mainLayout);
        }

        private DataGridView CreateBaseGrid(int cols)
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ColumnCount = cols,
                BackgroundColor = System.Drawing.Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                ColumnHeadersVisible = false,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                Font = new Font("Segoe UI", 8),
                ScrollBars = ScrollBars.None
            };
            return dgv;
        }

        private Label CreateValueLabel(string value, bool isMain = false)
        {
            var lbl = new Label
            {
                Text = value,
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", isMain ? 10 : 9, isMain ? FontStyle.Bold : FontStyle.Regular),
                BackColor = System.Drawing.Color.White
            };
            return lbl;
        }

        private Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White
            };
        }

        // =================================================================
        // ACTUALIZACIÓN DE DATOS DINÁMICOS (Corregida de forma segura)
        // =================================================================
        public void UpdateTolerances(double de, double dl, double dc, double dh)
        {
            _deAprobado = de;
            _dlAprobado = dl;
            _dcAprobado = dc;
            _dhAprobado = dh;

            if (_lastResult != null)
            {
                UpdateData(_lastResult);
            }
        }

        public void UpdateData(ColorCorrectionResult res)
        {
            if (res == null) return;
            _lastResult = res;

            // Actualización segura sin Invoke (Funciona sin importar si el Handle ya se creó o no)
            string textIlluminant = !string.IsNullOrEmpty(res.Illuminant) ? res.Illuminant.ToUpper() : "D65";
            lblIlluminantName.Text = textIlluminant;

            // Sincronizar el nombre del iluminante en la celda superior izquierda de la grilla Lab
            if (dgvLab.Rows.Count > 0)
            {
                dgvLab.Rows[0].Cells[0].Value = textIlluminant;
            }

            // Update L,a,b (Fila 0 es cabecera, los datos reales van en 1, 2, 3)
            dgvLab.Rows[1].Cells[1].Value = res.StdL.ToString("F2");
            dgvLab.Rows[1].Cells[2].Value = res.StdA.ToString("F2");
            dgvLab.Rows[1].Cells[3].Value = res.StdB.ToString("F2");

            dgvLab.Rows[2].Cells[1].Value = res.LotL.ToString("F2");
            dgvLab.Rows[2].Cells[2].Value = res.LotA.ToString("F2");
            dgvLab.Rows[2].Cells[3].Value = res.LotB.ToString("F2");

            dgvLab.Rows[3].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaL);
            dgvLab.Rows[3].Cells[2].Value = ColorimetricCalculator.FormatDelta(res.DeltaA);
            dgvLab.Rows[3].Cells[3].Value = ColorimetricCalculator.FormatDelta(res.DeltaB);

            // Update Chroma/Hue (Fila 0 es cabecera, los datos reales van en 1, 2, 3)
            dgvChromaHue.Rows[1].Cells[0].Value = res.StdC.ToString("F2");
            dgvChromaHue.Rows[1].Cells[1].Value = res.StdH.ToString("F2");
            dgvChromaHue.Rows[2].Cells[0].Value = res.LotC.ToString("F2");
            dgvChromaHue.Rows[2].Cells[1].Value = res.LotH.ToString("F2");
            dgvChromaHue.Rows[3].Cells[0].Value = ColorimetricCalculator.FormatDelta(res.DeltaChroma);
            dgvChromaHue.Rows[3].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaHue);

            // Update CMC Values
            lblLightness.Text = res.CmcLightness.ToString("F2");
            lblChroma.Text = res.CmcChroma.ToString("F2");
            lblHue.Text = res.CmcHue.ToString("F2");
            lblCmcValue.Text = res.CmcValue.ToString("F2");

            // Estado Semafórico del CMC Tolerancia
            if (res.CmcValue > _deAprobado)
            {
                lblCmcValue.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
                lblCmcValue.ForeColor = System.Drawing.Color.FromArgb(153, 0, 0); // Texto rojo oscuro para legibilidad
            }
            else
            {
                lblCmcValue.BackColor = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
                lblCmcValue.ForeColor = System.Drawing.Color.FromArgb(0, 102, 0); // Texto verde oscuro
            }

            // Estado semafórico de Lightness (DL)
            if (Math.Abs(res.CmcLightness) > _dlAprobado)
            {
                lblLightness.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
                lblLightness.ForeColor = System.Drawing.Color.FromArgb(153, 0, 0);
            }
            else
            {
                lblLightness.BackColor = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
                lblLightness.ForeColor = System.Drawing.Color.FromArgb(0, 102, 0);
            }

            // Estado semafórico de Chroma (DC)
            if (Math.Abs(res.CmcChroma) > _dcAprobado)
            {
                lblChroma.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
                lblChroma.ForeColor = System.Drawing.Color.FromArgb(153, 0, 0);
            }
            else
            {
                lblChroma.BackColor = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
                lblChroma.ForeColor = System.Drawing.Color.FromArgb(0, 102, 0);
            }

            // Estado semafórico de Hue (DH)
            if (Math.Abs(res.CmcHue) > _dhAprobado)
            {
                lblHue.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
                lblHue.ForeColor = System.Drawing.Color.FromArgb(153, 0, 0);
            }
            else
            {
                lblHue.BackColor = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
                lblHue.ForeColor = System.Drawing.Color.FromArgb(0, 102, 0);
            }

            // Update Actions (Tabla Inferior)
            dgvActions.Rows[1].Cells[0].Value = Math.Round(res.PercentL, 0) + "%";
            dgvActions.Rows[1].Cells[1].Value = Math.Round(res.PercentA, 0) + "%";
            dgvActions.Rows[1].Cells[2].Value = Math.Round(res.PercentB, 0) + "%";

            // ¡Truco clave! Forzamos un refresco visual inmediato del control para pintar las etiquetas anidadas
            lblIlluminantName.Refresh();
        }
    }
}