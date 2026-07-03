using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Color.Services;

namespace Color
{
    public class IluminantReportBlock : UserControl
    {
        private Label lblIlluminantName;
        private DataGridView dgvLab;
        private DataGridView dgvDiagnostic;
        private DataGridView dgvActions;
        private DataGridView dgvChromaHue;
        private DataGridView dgvDeltaChromaHue;
        private DataGridView dgvDeviation;

        private Label lblCmcValue;
        private Label lblLightness, lblChroma, lblHue;
        private Label lblCmcStatus;   
        private Label lblMiLeft;      
        private Label lblMiRight;     


        private Label lblSL, lblSC, lblH_Angle, lblSH, lblT, lblF;

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
            // El alto (240px) y el ancho 
            this.MinimumSize = new Size(900, 240);
            this.BackColor = System.Drawing.Color.White;

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 44)); 
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 22)); 
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));

            // =================================================================
            // --- Bloque 1: Configuracion del Contenedor Principal L, a, b ---
            // =================================================================
            TableLayoutPanel pnlLeft = new TableLayoutPanel();
            pnlLeft.Dock = DockStyle.Fill;
            pnlLeft.ColumnCount = 1;
            pnlLeft.RowCount = 5; 
            pnlLeft.Margin = new Padding(0, 0, 10, 0);

            pnlLeft.ColumnStyles.Clear();
            pnlLeft.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            pnlLeft.RowStyles.Clear();

            // Distribución de filas idéntica al Bloque 2 para mantener simetría exacta
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));   
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 6F));   
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 66F));  
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 6F));   
            pnlLeft.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));  

            pnlLeft.Height = 208;

            System.Drawing.Color azulColorimetro = System.Drawing.Color.FromArgb(0, 122, 204);
            System.Drawing.Color grisBordeSuave = System.Drawing.Color.FromArgb(195, 195, 195);

            // =================================================================
            // --- 1. Grid Superior: Datos CIELAB (dgvLab) ---
            // =================================================================
            dgvLab = CreateBaseGrid(4);
            dgvLab.Margin = new Padding(0);

            // Configuración de bordes suaves y eliminación de fila extra
            dgvLab.AllowUserToAddRows = false;
            dgvLab.GridColor = grisBordeSuave;
            dgvLab.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvLab.BorderStyle = BorderStyle.None;

            dgvLab.Columns[0].Width = 55;

            // Fila 0: Cabecera Principal (D65, L, a, b)
            int r0 = dgvLab.Rows.Add("D65", "L", "a", "b");
            dgvLab.Rows[r0].DefaultCellStyle.BackColor = azulColorimetro;
            dgvLab.Rows[r0].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvLab.Rows[r0].DefaultCellStyle.Font = new Font(dgvLab.Font, FontStyle.Bold);
            dgvLab.Rows[r0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvLab.Rows[r0].Height = 24;

            // Filas 1 y 2: Valores estandar y de lote
            int r_std = dgvLab.Rows.Add("Std", "0.00", "0.00", "0.00");
            dgvLab.Rows[r_std].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvLab.Rows[r_std].Height = 24;

            int r_lot = dgvLab.Rows.Add("Lot", "0.00", "0.00", "0.00");
            dgvLab.Rows[r_lot].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvLab.Rows[r_lot].Height = 24;

            pnlLeft.Controls.Add(dgvLab, 0, 0);

            // Separador 1 (línea gris de transición)
            var sepLeft1 = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.FromArgb(220, 220, 220), Margin = new Padding(0) };
            pnlLeft.Controls.Add(sepLeft1, 0, 1);

            // =================================================================
            // --- 2. Grid Central: Delta y Diagnostico ---
            // =================================================================
            dgvDiagnostic = CreateBaseGrid(4);
            dgvDiagnostic.Margin = new Padding(0);
            dgvDiagnostic.ColumnHeadersVisible = false;

            // Configuración de bordes suaves y eliminación de fila extra
            dgvDiagnostic.AllowUserToAddRows = false;
            dgvDiagnostic.GridColor = grisBordeSuave;
            dgvDiagnostic.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvDiagnostic.BorderStyle = BorderStyle.None;

            dgvDiagnostic.Columns[0].Width = 55;

            // Fila 0: Etiquetas dL, da, db
            int rd0 = dgvDiagnostic.Rows.Add("Δ", "dL", "da", "db");
            dgvDiagnostic.Rows[rd0].DefaultCellStyle.Font = new Font(dgvDiagnostic.Font, FontStyle.Bold);
            dgvDiagnostic.Rows[rd0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDiagnostic.Rows[rd0].Height = 22;

            // Fila 1: Valores numéricos de Delta
            int rd1 = dgvDiagnostic.Rows.Add("Δ", "0.00", "0.00", "0.00");
            dgvDiagnostic.Rows[rd1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDiagnostic.Rows[rd1].Height = 22;

            // Fila 2: Texto de Diagnóstico 
            int rd2 = dgvDiagnostic.Rows.Add("Δ", "Claro (Thin)", "Rojo", "Amarillo");
            dgvDiagnostic.Rows[rd2].Height = 22;
            dgvDiagnostic.Rows[rd2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvDiagnostic.Rows[rd2].Cells[1].Style.BackColor = azulColorimetro;
            dgvDiagnostic.Rows[rd2].Cells[1].Style.ForeColor = System.Drawing.Color.White;
            dgvDiagnostic.Rows[rd2].Cells[1].Style.Font = new Font(dgvDiagnostic.Font, FontStyle.Bold);

            // Celda Rojo (Texto Rojo)
            dgvDiagnostic.Rows[rd2].Cells[2].Style.Font = new Font(dgvDiagnostic.Font, FontStyle.Bold);
            dgvDiagnostic.Rows[rd2].Cells[2].Style.ForeColor = System.Drawing.Color.Red;

            // Celda Amarillo (Texto Amarillo Oscuro/Gold)
            dgvDiagnostic.Rows[rd2].Cells[3].Style.Font = new Font(dgvDiagnostic.Font, FontStyle.Bold);
            dgvDiagnostic.Rows[rd2].Cells[3].Style.ForeColor = System.Drawing.Color.DarkGoldenrod;

            dgvDiagnostic.CellPainting += new DataGridViewCellPaintingEventHandler(DgvDiagnostic_CellPainting);

            pnlLeft.Controls.Add(dgvDiagnostic, 0, 2);

            // Separador 2 (línea gris de transición)
            var sepLeft2 = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.FromArgb(220, 220, 220), Margin = new Padding(0) };
            pnlLeft.Controls.Add(sepLeft2, 0, 3);

            // =======================================================
            // --- 3. Grid Inferior: Acciones de Correccion ---
            // =======================================================
            dgvActions = CreateBaseGrid(4);
            dgvActions.Margin = new Padding(0);
            dgvActions.ColumnHeadersVisible = false;

            // Configuración de bordes suaves y eliminación de fila extra
            dgvActions.AllowUserToAddRows = false;
            dgvActions.GridColor = grisBordeSuave;
            dgvActions.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvActions.BorderStyle = BorderStyle.None;

            dgvActions.Columns[0].Width = 55;

            // Fila 0: Encabezados de Acción
            int ra0 = dgvActions.Rows.Add("Accion", "Disminuir [ ]", "Aumentar Verde", "Aumentar Azul");
            dgvActions.Rows[ra0].Height = 32; 
            dgvActions.Rows[ra0].DefaultCellStyle.ForeColor = System.Drawing.Color.Black;
            dgvActions.Rows[ra0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Aumento de tamaño solo para la palabra "Accion" en la primera celda
            dgvActions.Rows[ra0].Cells[0].Style.Font = new Font(dgvActions.Font.FontFamily, 12F, FontStyle.Bold);

            // Fila 1: Porcentajes numéricos
            int ral = dgvActions.Rows.Add("Accion", "2%", "6%", "7%");
            dgvActions.Rows[ral].Height = 26;
            dgvActions.Rows[ral].DefaultCellStyle.ForeColor = azulColorimetro;
            dgvActions.Rows[ral].DefaultCellStyle.Font = new Font(dgvActions.Font, FontStyle.Bold);
            dgvActions.Rows[ral].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvActions.CellPainting += new DataGridViewCellPaintingEventHandler(DgvActions_CellPainting);

            pnlLeft.Controls.Add(dgvActions, 0, 4);

            // =================================================================
            // --- Bloque 2: Configuracion del Contenedor Principal (Middle) ---
            // =================================================================
            TableLayoutPanel pnlMiddle = new TableLayoutPanel();
            pnlMiddle.Dock = DockStyle.Fill;
            pnlMiddle.ColumnCount = 1;
            pnlMiddle.RowCount = 5;
            pnlMiddle.Margin = new Padding(10, 0, 10, 0);

            pnlMiddle.ColumnStyles.Clear();
            pnlMiddle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pnlMiddle.RowStyles.Clear();

            // Distribución de filas optimizada sin dejar aire inferior
            pnlMiddle.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));   
            pnlMiddle.RowStyles.Add(new RowStyle(SizeType.Absolute, 6F));    
            pnlMiddle.RowStyles.Add(new RowStyle(SizeType.Absolute, 76F));   
            pnlMiddle.RowStyles.Add(new RowStyle(SizeType.Absolute, 6F));    
            pnlMiddle.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));   

            // El alto total se compacta perfectamente de 222px a 218px
            pnlMiddle.Height = 218;

            // -----------------------------------------------------------------
            // TABLA 1  Header azul Chroma/Hue + 2 filas datos
            // -----------------------------------------------------------------
            dgvChromaHue = CreateBaseGrid(2);
            dgvChromaHue.Margin = new Padding(0);

            dgvChromaHue.AllowUserToAddRows = false;
            dgvChromaHue.GridColor = System.Drawing.Color.FromArgb(195, 195, 195); 
            dgvChromaHue.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvChromaHue.BorderStyle = BorderStyle.None;

            int hChroma = dgvChromaHue.Rows.Add("Chroma", "Hue");
            dgvChromaHue.Rows[hChroma].DefaultCellStyle.BackColor = azulColorimetro;
            dgvChromaHue.Rows[hChroma].DefaultCellStyle.ForeColor = System.Drawing.Color.White;
            dgvChromaHue.Rows[hChroma].DefaultCellStyle.Font = new Font(dgvChromaHue.Font, FontStyle.Bold);
            dgvChromaHue.Rows[hChroma].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvChromaHue.Rows[hChroma].Height = 24;

            int r1 = dgvChromaHue.Rows.Add("0.00", "0.00");
            dgvChromaHue.Rows[r1].Height = 24;
            dgvChromaHue.Rows[r1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            int r2 = dgvChromaHue.Rows.Add("0.00", "0.00");
            dgvChromaHue.Rows[r2].Height = 24;
            dgvChromaHue.Rows[r2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            pnlMiddle.Controls.Add(dgvChromaHue, 0, 0);

            // Separador 1 (gris tenue)
            var sep1 = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.FromArgb(220, 220, 220), Margin = new Padding(0) };
            pnlMiddle.Controls.Add(sep1, 0, 1);

            // -----------------------------------------------------------------
            // TABLA 2  dC/dH + valores + fila Brighter/Yellower
            // -----------------------------------------------------------------
            dgvDeltaChromaHue = CreateBaseGrid(2);
            dgvDeltaChromaHue.Margin = new Padding(0);

            dgvDeltaChromaHue.AllowUserToAddRows = false;
            dgvDeltaChromaHue.GridColor = System.Drawing.Color.FromArgb(195, 195, 195);
            dgvDeltaChromaHue.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvDeltaChromaHue.BorderStyle = BorderStyle.None;

            int dLbl = dgvDeltaChromaHue.Rows.Add("dC", "dH");
            dgvDeltaChromaHue.Rows[dLbl].Height = 22;
            dgvDeltaChromaHue.Rows[dLbl].DefaultCellStyle.Font = new Font(dgvDeltaChromaHue.Font, FontStyle.Bold);
            dgvDeltaChromaHue.Rows[dLbl].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDeltaChromaHue.Rows[dLbl].Cells[0].Style.BackColor = System.Drawing.Color.White;
            dgvDeltaChromaHue.Rows[dLbl].Cells[0].Style.ForeColor = System.Drawing.Color.Black;

            int dVal = dgvDeltaChromaHue.Rows.Add("0.00", "0.00");
            dgvDeltaChromaHue.Rows[dVal].Height = 22;
            dgvDeltaChromaHue.Rows[dVal].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            int dStatus = dgvDeltaChromaHue.Rows.Add("Brighter", "Yellower (Redder)");
            dgvDeltaChromaHue.Rows[dStatus].Height = 32;
            dgvDeltaChromaHue.Rows[dStatus].Cells[0].Style.BackColor = System.Drawing.Color.White;
            dgvDeltaChromaHue.Rows[dStatus].Cells[0].Style.ForeColor = System.Drawing.Color.Black;
            dgvDeltaChromaHue.Rows[dStatus].Cells[0].Style.Font = new Font(dgvDeltaChromaHue.Font, FontStyle.Bold);
            dgvDeltaChromaHue.Rows[dStatus].Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDeltaChromaHue.Rows[dStatus].Cells[1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

            pnlMiddle.Controls.Add(dgvDeltaChromaHue, 0, 2);

            // Separador 2 (gris tenue)
            var sep2 = new Panel { Dock = DockStyle.Fill, BackColor = System.Drawing.Color.FromArgb(220, 220, 220), Margin = new Padding(0) };
            pnlMiddle.Controls.Add(sep2, 0, 3);

            // -----------------------------------------------------------------
            // TABLA 3 Duller/Bluer (texto) + fila de %
            // -----------------------------------------------------------------
            dgvDeviation = CreateBaseGrid(2);
            dgvDeviation.Margin = new Padding(0);

            dgvDeviation.AllowUserToAddRows = false;
            dgvDeviation.GridColor = System.Drawing.Color.FromArgb(195, 195, 195);
            dgvDeviation.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgvDeviation.BorderStyle = BorderStyle.None;

            int dDuller = dgvDeviation.Rows.Add("Duller", "Bluer\n(Greener)");
            dgvDeviation.Rows[dDuller].Height = 32;
            dgvDeviation.Rows[dDuller].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvDeviation.Rows[dDuller].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDeviation.Rows[dDuller].Cells[0].Style.BackColor = System.Drawing.Color.White;
            dgvDeviation.Rows[dDuller].Cells[0].Style.ForeColor = System.Drawing.Color.Black;

            int dPct = dgvDeviation.Rows.Add("0%", "0%");
            dgvDeviation.Rows[dPct].Height = 26;
            dgvDeviation.Rows[dPct].DefaultCellStyle.ForeColor = azulColorimetro;
            dgvDeviation.Rows[dPct].DefaultCellStyle.Font = new Font(dgvDeviation.Font, FontStyle.Bold);
            dgvDeviation.Rows[dPct].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            pnlMiddle.Controls.Add(dgvDeviation, 0, 4);

            // =================================================================
            //  Bloque 3 y 4 Unificados: CMC, Paremetros e Iluminante 
            // =================================================================

            var pnlCmcCombo = new TableLayoutPanel 
            { 
                Dock = DockStyle.Fill, 
                ColumnCount = 6, 
                RowCount = 7, 
                Margin = new Padding(2),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };

            // Distribucion exacta de columnas (Usando floats explicitos validos)
            for (int i = 0; i < 5; i++)
            {
                pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.0f));
            }
            pnlCmcCombo.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 75.0f)); 

            // Definicion de alturas de filas
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 28.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 40.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 24.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 30.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 74.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Absolute, 20.0f)); 
            pnlCmcCombo.RowStyles.Add(new RowStyle(SizeType.Percent, 100.0f));

            // --- FILA 0: Headers Principales ---
            var lblTitleLightness = CreateHeaderLabel("Lightness");
            var lblTitleChroma    = CreateHeaderLabel("Chroma");
            var lblTitleHue       = CreateHeaderLabel("Hue");
            var lblTitleCmc       = CreateHeaderLabel("CMC (2:1)");

            pnlCmcCombo.Controls.Add(lblTitleLightness, 0, 0);
            pnlCmcCombo.Controls.Add(lblTitleChroma,    1, 0);
            pnlCmcCombo.Controls.Add(lblTitleHue,       2, 0);
            pnlCmcCombo.Controls.Add(lblTitleCmc,       3, 0);
            pnlCmcCombo.SetColumnSpan(lblTitleCmc, 2); 

            // --- FILA 1: Valores Principales (Estilo Verde/Azul de la imagen) ---
            System.Drawing.Color bgGreen = System.Drawing.Color.Black; 
            lblLightness = CreateValueLabelCustom("0.36", bgGreen, System.Drawing.Color.Green, 11.0f);
            lblChroma    = CreateValueLabelCustom("0.00", bgGreen, System.Drawing.Color.Green, 11.0f);
            lblHue       = CreateValueLabelCustom("-0.66", bgGreen, System.Drawing.Color.Green, 11.0f);

            // El valor de CMC en la imagen es azul y centrado
            lblCmcValue  = CreateValueLabelCustom("0.75", System.Drawing.Color.White, System.Drawing.Color.Black, 12.0f);

            pnlCmcCombo.Controls.Add(lblLightness, 0, 1);
            pnlCmcCombo.Controls.Add(lblChroma,    1, 1);
            pnlCmcCombo.Controls.Add(lblHue,       2, 1);
            pnlCmcCombo.Controls.Add(lblCmcValue,  3, 1);
            pnlCmcCombo.SetColumnSpan(lblCmcValue, 2);

            // --- FILA 0 y 1 (Bloque Lateral): Iluminante D65 ---
            lblIlluminantName = new Label
            {
                Text = "D65",
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204), 
                ForeColor = System.Drawing.Color.White,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 11.0f, FontStyle.Bold),
                Margin = new Padding(2, 0, 2, 2)
            };
            pnlCmcCombo.Controls.Add(lblIlluminantName, 5, 0);
            pnlCmcCombo.SetRowSpan(lblIlluminantName, 2); 

            // --- FILA 2: Sub-Headers (SL, SC, h, SH, T, F) ---
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("SL"), 0, 2);
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("SC"), 1, 2);
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("h"),  2, 2);
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("SH"), 3, 2);
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("T"),  4, 2);
            pnlCmcCombo.Controls.Add(CreateSubHeaderLabel("F"),  5, 2);

            // --- FILA 3: Sub-Valores Numericos ---
            lblSL      = CreateSubValueLabel("1.055");
            lblSC      = CreateSubValueLabel("1.195");
            lblH_Angle = CreateSubValueLabel("137.43");
            lblSH      = CreateSubValueLabel("0.930");
            lblT       = CreateSubValueLabel("0.757");
            lblF       = CreateSubValueLabel("0.912");

            pnlCmcCombo.Controls.Add(lblSL,      0, 3);
            pnlCmcCombo.Controls.Add(lblSC,      1, 3);
            pnlCmcCombo.Controls.Add(lblH_Angle, 2, 3);
            pnlCmcCombo.Controls.Add(lblSH,      3, 3);
            pnlCmcCombo.Controls.Add(lblT,       4, 3);
            pnlCmcCombo.Controls.Add(lblF,       5, 3);

            // --- FILA 4: Estado Global OK / FAIL ---
            lblCmcStatus = new Label
            {
                Text = "Ok",
                Font = new Font("Segoe UI", 16.0f, FontStyle.Bold),
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(2)
            };
            pnlCmcCombo.Controls.Add(lblCmcStatus, 0, 4);
            pnlCmcCombo.SetColumnSpan(lblCmcStatus, 6); 

            // --- FILA 5: Header indice de Metamerismo ---
            var lblMiHeader = new Label
            {
                Text = "Indice de Metamerismo (MI)",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(1, 0, 1, 0)
            };
            pnlCmcCombo.Controls.Add(lblMiHeader, 0, 5);
            pnlCmcCombo.SetColumnSpan(lblMiHeader, 6);

            // --- FILA 6: Valores Numericos MI ---
            lblMiLeft = CreateSubValueLabel("0.80");
            lblMiLeft.Font = new Font("Segoe UI", 11.0f, FontStyle.Bold);

            lblMiRight = CreateSubValueLabel("1.20");
            lblMiRight.Font = new Font("Segoe UI", 11.0f, FontStyle.Bold);

            pnlCmcCombo.Controls.Add(lblMiLeft,  0, 6);
            pnlCmcCombo.SetColumnSpan(lblMiLeft,  3); 
            pnlCmcCombo.Controls.Add(lblMiRight, 3, 6);
            pnlCmcCombo.SetColumnSpan(lblMiRight, 3); 

            // --- Ensamblado General ---
            mainLayout.Controls.Add(pnlLeft, 0, 0);
            mainLayout.Controls.Add(pnlMiddle, 1, 0);
            mainLayout.Controls.Add(pnlCmcCombo, 2, 0);

            this.Controls.Add(mainLayout);
        }

        private void DgvDiagnostic_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 0 && (e.RowIndex == 0 || e.RowIndex == 1 || e.RowIndex == 2))
            {
                DataGridView dgv = (DataGridView)sender;
                System.Drawing.Color azulColorimetro = System.Drawing.Color.FromArgb(0, 122, 204);

                // 1. Dibujar fondo azul para la celda actual
                using (SolidBrush brush = new SolidBrush(azulColorimetro))
                {
                    e.Graphics.FillRectangle(brush, e.CellBounds);
                }

                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;

                // 2. Dibujar el simbolo Delta (Calculamos el rectangulo fusionado )
                Rectangle rectFila0 = (e.RowIndex == 0) ? e.CellBounds : dgv.GetCellDisplayRectangle(e.ColumnIndex, 0, true);
                Rectangle rectFila1 = (e.RowIndex == 1) ? e.CellBounds : dgv.GetCellDisplayRectangle(e.ColumnIndex, 1, true);
                Rectangle rectFila2 = (e.RowIndex == 2) ? e.CellBounds : dgv.GetCellDisplayRectangle(e.ColumnIndex, 2, true);

                Rectangle rectFusionado = rectFila0;
                if (!rectFila1.IsEmpty) rectFusionado.Height = (rectFila1.Bottom - rectFila0.Top);
                if (!rectFila2.IsEmpty) rectFusionado.Height = (rectFila2.Bottom - rectFila0.Top);

                using (Font fontDelta = new Font("Segoe UI", 16.0f, FontStyle.Bold))
                using (SolidBrush brushWhite = new SolidBrush(System.Drawing.Color.White))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Δ", fontDelta, brushWhite, rectFusionado, sf);
                }

                using (Pen penBorde = new Pen(dgv.GridColor))
                {
                    e.Graphics.DrawLine(penBorde, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Left, e.CellBounds.Bottom);
                    if (e.RowIndex == 2)
                    {
                        e.Graphics.DrawLine(penBorde, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    }
                }

                e.Handled = true;
            }
        }

        private void DgvActions_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 0 && (e.RowIndex == 0 || e.RowIndex == 1))
            {
                DataGridView dgv = (DataGridView)sender;
                System.Drawing.Color azulColorimetro = System.Drawing.Color.FromArgb(0, 122, 204);

                // 1. Dibujar fondo azul para la celda actual
                using (SolidBrush brush = new SolidBrush(azulColorimetro))
                {
                    e.Graphics.FillRectangle(brush, e.CellBounds);
                }

                e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;

                // 2. Dibujar el texto Acción (Calculamos el rectangulo fusionado completo en ambas filas)
                Rectangle rectFila0 = (e.RowIndex == 0) ? e.CellBounds : dgv.GetCellDisplayRectangle(e.ColumnIndex, 0, true);
                Rectangle rectFila1 = (e.RowIndex == 1) ? e.CellBounds : dgv.GetCellDisplayRectangle(e.ColumnIndex, 1, true);

                Rectangle rectFusionado = rectFila0;
                if (!rectFila1.IsEmpty) rectFusionado.Height = (rectFila1.Bottom - rectFila0.Top);

                using (Font fontAccion = new Font("Segoe UI", 9.0f, FontStyle.Bold))
                using (SolidBrush brushWhite = new SolidBrush(System.Drawing.Color.White))
                {
                    StringFormat sf = new StringFormat();
                    sf.Alignment = StringAlignment.Center;
                    sf.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString("Accion", fontAccion, brushWhite, rectFusionado, sf);
                }

                using (Pen penBorde = new Pen(dgv.GridColor))
                {
                    e.Graphics.DrawLine(penBorde, e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Left, e.CellBounds.Bottom);
                    if (e.RowIndex == (dgv.Rows.Count - 1))
                    {
                        e.Graphics.DrawLine(penBorde, e.CellBounds.Left, e.CellBounds.Bottom - 1, e.CellBounds.Right, e.CellBounds.Bottom - 1);
                    }
                }

                e.Handled = true;
            }
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
                ScrollBars = ScrollBars.None,
                Margin = new Padding(0)
            };
            return dgv;
        }

        private Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204), 
                ForeColor = System.Drawing.Color.White,
                Font = new Font("Segoe UI", 10.0f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(1)
            };
        }

        private Label CreateValueLabelCustom(string text, System.Drawing.Color backColor, System.Drawing.Color foreColor, float fontSize)
        {
            return new Label
            {
                Text = text,
                BackColor = backColor,
                ForeColor = foreColor,
                Font = new Font("Segoe UI", fontSize, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Margin = new Padding(1)
            };
        }

        private Label CreateSubHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.DimGray,
                Font = new Font("Segoe UI", 9.0f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle, 
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 1, 1) 
            };
        }

        private Label CreateSubValueLabel(string text)
        {
            return new Label
            {
                Text = text,
                BackColor = System.Drawing.Color.White,
                ForeColor = System.Drawing.Color.Black, 
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.FixedSingle,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 1, 1)
            };
        }

        // =================================================================
        // ACTUALIZACION DE DATOS DINAMICOS (Corregida de forma segura)
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

            string textIlluminant = !string.IsNullOrEmpty(res.Illuminant) ? res.Illuminant.ToUpper() : "D65";
            lblIlluminantName.Text = textIlluminant;

            // Sincronizar el nombre del iluminante en la celda superior izquierda de la grilla Lab
            if (dgvLab.Rows.Count > 0)
            {
                dgvLab.Rows[0].Cells[0].Value = textIlluminant;
            }

            // Aseguramos la integridad de las etiquetas laterales en dgvDiagnostic y dgvActions
            if (dgvDiagnostic.Rows.Count > 0) dgvDiagnostic.Rows[0].Cells[0].Value = "Δ";
            if (dgvActions.Rows.Count > 0) dgvActions.Rows[0].Cells[0].Value = "Accion";

            // Update L,a,b (Fila 0 es cabecera, los datos reales van en 1, 2, 3)
            dgvLab.Rows[1].Cells[1].Value = res.StdL.ToString("F2", CultureInfo.InvariantCulture);
            dgvLab.Rows[1].Cells[2].Value = res.StdA.ToString("F2", CultureInfo.InvariantCulture);
            dgvLab.Rows[1].Cells[3].Value = res.StdB.ToString("F2", CultureInfo.InvariantCulture);

            dgvLab.Rows[2].Cells[1].Value = res.LotL.ToString("F2", CultureInfo.InvariantCulture);
            dgvLab.Rows[2].Cells[2].Value = res.LotA.ToString("F2", CultureInfo.InvariantCulture);
            dgvLab.Rows[2].Cells[3].Value = res.LotB.ToString("F2", CultureInfo.InvariantCulture);

            // Los deltas van en Fila 1, columnas 1, 2, 3
            dgvDiagnostic.Rows[1].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaL);
            dgvDiagnostic.Rows[1].Cells[2].Value = ColorimetricCalculator.FormatDelta(res.DeltaA);
            dgvDiagnostic.Rows[1].Cells[3].Value = ColorimetricCalculator.FormatDelta(res.DeltaB);

            if (dgvDiagnostic != null && dgvDiagnostic.Rows.Count >= 3)
            {
                int rdDiagnostico = 2; 

                // Asignacion de textos independientes por variable
                string diagL = ColorimetricCalculator.GetLuminosityDiagnosis(res.DeltaL);
                string diagA = ColorimetricCalculator.GetEjeADiagnosis(res.DeltaA);
                string diagB = ColorimetricCalculator.GetEjeBDiagnosis(res.DeltaB);

                dgvDiagnostic.Rows[rdDiagnostico].Cells[1].Value = diagL;
                dgvDiagnostic.Rows[rdDiagnostico].Cells[2].Value = diagA;
                dgvDiagnostic.Rows[rdDiagnostico].Cells[3].Value = diagB;

                // Formato dinamico y estricto de colores
                dgvDiagnostic.Rows[rdDiagnostico].Cells[1].Style.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
                dgvDiagnostic.Rows[rdDiagnostico].Cells[1].Style.ForeColor = System.Drawing.Color.White;

                // Eje A
                dgvDiagnostic.Rows[rdDiagnostico].Cells[2].Style.ForeColor = (diagA == "Rojo") 
                    ? System.Drawing.Color.Red 
                    : System.Drawing.Color.Green;

                // Eje B
                dgvDiagnostic.Rows[rdDiagnostico].Cells[3].Style.ForeColor = (diagB == "Amarillo") 
                    ? System.Drawing.Color.DarkGoldenrod 
                    : System.Drawing.Color.Blue;
            }

            // Update Chroma/Hue (Fila 0 es cabecera, datos en 1 y 2)
            dgvChromaHue.Rows[1].Cells[0].Value = res.StdC.ToString("F2", CultureInfo.InvariantCulture);
            dgvChromaHue.Rows[1].Cells[1].Value = res.StdH.ToString("F2", CultureInfo.InvariantCulture);
            dgvChromaHue.Rows[2].Cells[0].Value = res.LotC.ToString("F2", CultureInfo.InvariantCulture);
            dgvChromaHue.Rows[2].Cells[1].Value = res.LotH.ToString("F2", CultureInfo.InvariantCulture);

            // Update dC/dH y su etiqueta de estado asociada (Fila 1 y 2)
            dgvDeltaChromaHue.Rows[1].Cells[0].Value = ColorimetricCalculator.FormatDelta(res.DeltaChroma);
            dgvDeltaChromaHue.Rows[1].Cells[1].Value = ColorimetricCalculator.FormatDelta(res.DeltaHue);

            // Update CMC Values
            lblLightness.Text = res.CmcLightness.ToString("F2", CultureInfo.InvariantCulture);
            lblChroma.Text    = res.CmcChroma.ToString("F2", CultureInfo.InvariantCulture);
            lblHue.Text       = res.CmcHue.ToString("F2", CultureInfo.InvariantCulture);
            lblCmcValue.Text  = res.CmcValue.ToString("F2", CultureInfo.InvariantCulture);

            // CMC Params
            if (lblSL != null) lblSL.Text = res.SL.ToString("0.000", CultureInfo.InvariantCulture);
            if (lblSC != null) lblSC.Text = res.SC.ToString("0.000", CultureInfo.InvariantCulture);
            if (lblH_Angle != null) lblH_Angle.Text = res.h_angle.ToString("0.00", CultureInfo.InvariantCulture);
            if (lblSH != null) lblSH.Text = res.SH.ToString("0.000", CultureInfo.InvariantCulture);
            if (lblT != null) lblT.Text = res.T_factor.ToString("0.000", CultureInfo.InvariantCulture);
            if (lblF != null) lblF.Text = res.F_factor.ToString("0.000", CultureInfo.InvariantCulture);

            // --- Estado Global OK / FAIL (Fila 4): Solo para D65 ---
            if (textIlluminant == "D65")
            {
                // Índice de Metamerismo (MI)
                lblCmcStatus.Text = "-";
                lblCmcStatus.BackColor = System.Drawing.Color.White;
                lblCmcStatus.ForeColor = System.Drawing.Color.Black;
            }
            else if (textIlluminant == "TL84" || textIlluminant == "A" || textIlluminant == "CWF")
            {
                if (lblCmcStatus != null) lblCmcStatus.Visible = false;
            }

            // Estado Semaforico del CMC Tolerancia
            if (res.CmcValue > _deAprobado)
            {
                lblCmcValue.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFD6D6");
                lblCmcValue.ForeColor = System.Drawing.Color.FromArgb(153, 0, 0); 
            }
            else
            {
                lblCmcValue.BackColor = System.Drawing.ColorTranslator.FromHtml("#D6F5D6");
                lblCmcValue.ForeColor = System.Drawing.Color.FromArgb(0, 102, 0); 
            }

            // Estado semaforico de Lightness (DL)
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

            // Estado semaforico de Chroma (DC)
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

            // Estado semaforico de Hue (DH)
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

            // =========================================================================
            //  PARIDAD EXCEL EXCLUSIVA: LÓGICA "AUMENTAR" EN 2 FILAS 
            // =========================================================================
            if (dgvActions != null && dgvActions.Rows.Count > 0)
            {
              
                if (dgvActions.Rows.Count < 2)
                {
                    dgvActions.Rows.Add("Accion", "-", "-", "-");
                }

                int filaTextosAccion = 0; 
                int filaPorcentajes = 1;  

                // 1. Recuperar diagnósticos dinámicos reales del motor de color 
                string diagL = ColorimetricCalculator.GetLuminosityDiagnosis(res.DeltaL); 
                string diagA = ColorimetricCalculator.GetEjeADiagnosis(res.DeltaA);       
                string diagB = ColorimetricCalculator.GetEjeBDiagnosis(res.DeltaB);       

                // 2. Procesar matemáticamente los tres porcentajes absolutos en formato Entero
                string pctLStr = Math.Abs(Math.Round(res.PercentL, 0)).ToString("0", CultureInfo.InvariantCulture) + "%";
                string pctAStr = Math.Abs(Math.Round(res.PercentA, 0)).ToString("0", CultureInfo.InvariantCulture) + "%";
                string pctBStr = Math.Abs(Math.Round(res.PercentB, 0)).ToString("0", CultureInfo.InvariantCulture) + "%";

                // ---------------------------------------------------------------------
                // Mapeo Fórmula 1 (Luminosidad L*) -> Condicional Estricto Excel
                // ---------------------------------------------------------------------
                string accionLStr = "";
                if (res.DeltaL < 0 && diagL.Contains("Oscuro")) accionLStr = "Disminuir [ ]";
                else if (res.DeltaL > 0 && diagL.Contains("Claro")) accionLStr = "Aumentar [ ]";
                else accionLStr = "Mantener [ ]";

                // ---------------------------------------------------------------------
                // Mapeo Fórmula 2 (Eje a* Rojo/Verde) -> Condicional Estricto Excel
                // ---------------------------------------------------------------------
                string accionAStr = "";
                if (diagA == "Rojo")       accionAStr = "Aumentar Verde"; 
                else if (diagA == "Verde") accionAStr = "Aumentar Rojo";
                else accionAStr = (res.DeltaA < 0) ? "Aumentar Rojo" : "Aumentar Verde";

                // ---------------------------------------------------------------------
                // Mapeo Fórmula 3 (Eje b* Amarillo/Azul) -> Condicional Estricto Excel
                // ---------------------------------------------------------------------
                string accionBStr = "";
                if (diagB == "Amarillo")   accionBStr = "Aumentar Azul"; 
                else if (diagB == "Azul")  accionBStr = "Aumentar Amarillo";
                else accionBStr = (res.DeltaB < 0) ? "Aumentar Amarillo" : "Aumentar Azul";

                // 3. INYECCIÓN EN LA INTERFAZ GRÁFICA (Fila 0 - Textos)
                dgvActions.Rows[filaTextosAccion].Cells[1].Value = accionLStr;
                dgvActions.Rows[filaTextosAccion].Cells[2].Value = accionAStr;
                dgvActions.Rows[filaTextosAccion].Cells[3].Value = accionBStr;

                // 4. INYECCIÓN EN LA INTERFAZ GRÁFICA (Fila 1 - Porcentajes)
                dgvActions.Rows[filaPorcentajes].Cells[1].Value = pctLStr;
                dgvActions.Rows[filaPorcentajes].Cells[2].Value = pctAStr;
                dgvActions.Rows[filaPorcentajes].Cells[3].Value = pctBStr;

                // 5. HOMOLOGACIÓN VISUAL DE FUENTES (Textos)
                dgvActions.Rows[filaTextosAccion].Cells[1].Style.ForeColor = System.Drawing.Color.Black;
                dgvActions.Rows[filaTextosAccion].Cells[2].Style.ForeColor = (accionAStr.Contains("Verde")) ? System.Drawing.Color.Black : System.Drawing.Color.Black;
                dgvActions.Rows[filaTextosAccion].Cells[3].Style.ForeColor = (accionBStr.Contains("Azul")) ? System.Drawing.Color.Black : System.Drawing.Color.Black;

                // HOMOLOGACIÓN VISUAL DE FUENTES (Porcentajes)
                dgvActions.Rows[filaPorcentajes].Cells[1].Style.ForeColor = System.Drawing.Color.Black;
                dgvActions.Rows[filaPorcentajes].Cells[2].Style.ForeColor = (accionAStr.Contains("Verde")) ? System.Drawing.Color.Black : System.Drawing.Color.Black;
                dgvActions.Rows[filaPorcentajes].Cells[3].Style.ForeColor = (accionBStr.Contains("Azul")) ? System.Drawing.Color.Black : System.Drawing.Color.Black;
            }

            ProcesarYMostrarDiagnostico(res);

            // =========================================================
            // OCULTAMIENTO DINAMICO BLOQUE 3: Solo iluminante A / CWF
            // =========================================================
            bool esIluminanteA = (textIlluminant == "A" || textIlluminant == "CWF");

            if (lblCmcStatus != null)
            {
                lblCmcStatus.Visible = !esIluminanteA;
                
                if (textIlluminant == "TL84")
                {
                    lblCmcStatus.Visible = false;
                    
                    if (lblCmcStatus.Parent != null && lblCmcStatus.Parent is System.Windows.Forms.TableLayoutPanel)
                    {
                        System.Windows.Forms.TableLayoutPanel tlp = (System.Windows.Forms.TableLayoutPanel)lblCmcStatus.Parent;
                        int rowIndex = tlp.GetRow(lblCmcStatus);
                        if (rowIndex >= 0 && rowIndex < tlp.RowStyles.Count)
                        {
                            tlp.RowStyles[rowIndex].Height = 0.0f;
                        }
                    }
                }
            }

            if (lblMiLeft != null)
                lblMiLeft.Visible = !esIluminanteA;

            if (lblMiRight != null)
                lblMiRight.Visible = !esIluminanteA;

            if (lblCmcStatus != null && lblCmcStatus.Parent != null)
            {
                foreach (System.Windows.Forms.Control ctl in lblCmcStatus.Parent.Controls)
                {
                    var lbl = ctl as System.Windows.Forms.Label;
                    if (lbl != null && lbl.Text != null && lbl.Text.Contains("Metamerismo"))
                        lbl.Visible = !esIluminanteA;
                }
            }

            // ¡Truco clave! Forzamos un refresco visual inmediato del control para pintar las etiquetas anidadas
            lblIlluminantName.Refresh();
        }

        private void ProcesarYMostrarDiagnostico(ColorCorrectionResult res)
        {
            if (res == null) return;

            // -------------------------------------------------------------------------
            // PARTE A: DIAGNOSTICO VISUAL Y MATRIZ DE COLOR OPUESTO
            // -------------------------------------------------------------------------
            string dirA = res.DeltaA >= 0 ? "Redder" : "Greener";
            string dirB = res.DeltaB >= 0 ? "Yellower" : "Bluer";
            string tRight = $"{dirB} ({dirA})";

            string oppA = res.DeltaA >= 0 ? "Greener" : "Redder";
            string oppB = res.DeltaB >= 0 ? "Bluer" : "Yellower";
            string bRight = $"{oppB}\n({oppA})";

            // 2. Logica exacta de Excel para la celda tipo Brighter/Duller 
            bool isBrighter = res.DeltaChroma > 0;
            string tLeft = isBrighter ? "Brighter" : "Duller";
            string bLeft = isBrighter ? "Duller" : "Brighter";

            // 4. Asignaciones a la Grilla Inferior (opuestos)
            if (dgvDeltaChromaHue != null && dgvDeltaChromaHue.Rows.Count >= 3)
            {
                dgvDeltaChromaHue.Rows[2].Cells[0].Value = tLeft;
                dgvDeltaChromaHue.Rows[2].Cells[0].Style.BackColor = System.Drawing.Color.White;
                dgvDeltaChromaHue.Rows[2].Cells[0].Style.ForeColor = System.Drawing.Color.Black;

                dgvDeltaChromaHue.Rows[2].Cells[1].Value = tRight;
                dgvDeltaChromaHue.Rows[2].Cells[1].Style.BackColor = System.Drawing.Color.White;
                dgvDeltaChromaHue.Rows[2].Cells[1].Style.ForeColor = System.Drawing.Color.Black;
            }

            if (dgvDeviation != null && dgvDeviation.Rows.Count > 0)
            {
                dgvDeviation.Rows[0].Cells[0].Value = bLeft;
                dgvDeviation.Rows[0].Cells[1].Value = bRight;
            }
            //  CALCULO INTERNO ABSOLUTO: PARIDAD CON FORMULAS EXCEL 
            if (dgvDeviation != null && dgvDeviation.Rows.Count > 0)
            {
                int targetRowIdx = dgvDeviation.Rows.Count - 1;

                // 1. Formula =+ABS(H20)*0.1 (Basado en dC / DeltaChroma)
                double calculoInternoChroma = Math.Abs(res.DeltaChroma) * 10.0;
                dgvDeviation.Rows[targetRowIdx].Cells[0].Value = Math.Round(calculoInternoChroma, 0, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture) + "%";

                // 2. Formula =+ABS(I20)*0.1 (Basado en dH / DeltaHue)
                double calculoInternoHue = Math.Abs(res.DeltaHue) * 10.0;
                dgvDeviation.Rows[targetRowIdx].Cells[1].Value = Math.Round(calculoInternoHue, 0, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture) + "%";

                dgvDeviation.Rows[targetRowIdx].Cells[0].Style.ForeColor = System.Drawing.Color.Black;
                dgvDeviation.Rows[targetRowIdx].Cells[1].Style.ForeColor = System.Drawing.Color.Black;
                
                dgvDeviation.Rows[targetRowIdx].Cells[0].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgvDeviation.Rows[targetRowIdx].Cells[1].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // -------------------------------------------------------------------------
            // PARTE B: CONCATENACION DINAMICA DE PARAMETROS CMC (Bloque Inferior)
            // -------------------------------------------------------------------------
            if (lblSL != null)      { lblSL.Text      = $"{res.SL.ToString("0.000")}";       lblSL.ForeColor      = System.Drawing.Color.Black; }
            if (lblSC != null)      { lblSC.Text      = $"{res.SC.ToString("0.000")}";       lblSC.ForeColor      = System.Drawing.Color.Black; }
            if (lblH_Angle != null) { lblH_Angle.Text = $"{res.StdH.ToString("0.00")}";      lblH_Angle.ForeColor = System.Drawing.Color.Black; }
            if (lblSH != null)      { lblSH.Text      = $"{res.SH.ToString("0.000")}";      lblSH.ForeColor      = System.Drawing.Color.Black; }
            if (lblT != null)       { lblT.Text       = $"{res.T_factor.ToString("0.000")}";  lblT.ForeColor       = System.Drawing.Color.Black; }
            if (lblF != null)       { lblF.Text       = $"{res.F_factor.ToString("0.000")}";  lblF.ForeColor       = System.Drawing.Color.Black; }
        }
        // =========================================================================
        // CALCULO CRUZADO CMC(D65) vs CMC(TL84) - Solo para el bloque TL84
        // =========================================================================
                public void SetSpecialCrossCmcResult(double cmcD65, double cmcTL84)
        {
            if (lblMiLeft != null && lblMiRight != null)
            {
                // 1. Valor absoluto (Indice de Metamerismo)
                double diferenciaAbsolutaCmc = Math.Abs(cmcD65 - cmcTL84);

                // 2. Formato de 2 decimales sin %
                string resultadoStr = diferenciaAbsolutaCmc.ToString("0.00", CultureInfo.InvariantCulture);

                // 3. Asignar solo a lblMiLeft y expandirlo
                lblMiLeft.Text = resultadoStr;
                lblMiLeft.ForeColor = System.Drawing.Color.FromArgb(238, 108, 38); 
                lblMiLeft.BackColor = System.Drawing.Color.FromArgb(245, 245, 245);
                lblMiLeft.Visible = true;

                // Ocultar la celda derecha para generar impresion de celda unica
                lblMiRight.Visible = false;

                // Expandir la celda izquierda en el TableLayoutPanel (ColSpan = 6)
                if (lblMiLeft.Parent != null && lblMiLeft.Parent is System.Windows.Forms.TableLayoutPanel)
                {
                    System.Windows.Forms.TableLayoutPanel tlp = (System.Windows.Forms.TableLayoutPanel)lblMiLeft.Parent;
                    tlp.SetColumnSpan(lblMiLeft, 6);
                }

                // Ocultar la etiqueta "Indice de Metamerismo (MI)" si no ha sido ocultada
                if (lblMiLeft.Parent != null)
                {
                    foreach (System.Windows.Forms.Control ctl in lblMiLeft.Parent.Controls)
                    {
                        if (ctl is System.Windows.Forms.Label && 
                            ctl.Text != null && ctl.Text.Contains("Metamerismo"))
                        {
                            ctl.Visible = false;
                        }
                    }
                }
            }
        }

        public void SetVeredictoD65PorMetamerismo(double mi)
        {
            if (lblCmcStatus != null)
            {
                lblCmcStatus.Visible = true;
                if (mi > 1.20)
                {
                    lblCmcStatus.Text = "FALLA";
                    lblCmcStatus.ForeColor = System.Drawing.Color.FromArgb(180, 0, 0);
                    lblCmcStatus.BackColor = System.Drawing.Color.FromArgb(255, 220, 220);
                }
                else if (mi > 0.80)
                {
                    lblCmcStatus.Text = "alerta";
                    lblCmcStatus.ForeColor = System.Drawing.Color.FromArgb(133, 100, 4);
                    lblCmcStatus.BackColor = System.Drawing.Color.FromArgb(255, 243, 205);
                }
                else
                {
                    lblCmcStatus.Text = "ok";
                    lblCmcStatus.ForeColor = System.Drawing.Color.FromArgb(0, 153, 0);
                    lblCmcStatus.BackColor = System.Drawing.Color.FromArgb(230, 255, 230);
                }
            }
        }
    }
}