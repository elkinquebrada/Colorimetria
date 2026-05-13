using Color;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using SysColor = System.Drawing.Color;

namespace Colorimetria
{

    public partial class FormConfirmacionOCR : Form
    {
        // ===== Salidas =====
        public List<ColorimetricRow> RowsConfirmed { get; private set; }
        public OcrReport Report { get { return _report; } }

        // ===== Referencia al MainForm (para minimizar en Load) =====
        public Form MainFormOwner { get; set; }

        // ===== Entradas =====
        private List<ColorimetricRow> _rows;
        private OcrReport _report;
        private ShadeExtractionResult _shadeResult;

        // ===== UI =====
        private DataGridView dgvData;
        private DataGridView dgvReceta;
        private DataGridView dgvLab;
        private DataGridView dgvCmc;
        private TextBox txtRaw;
        private TabControl tabControl;
        private Button btnConfirmar;
        private Button btnCancelar;
        private Button btnRegresar;
        private Label lblTitulo;
        private Label lblDatosReceta;
        private Label lblCount;
        private Label lblTol;

        // =========================================================
        // CONSTRUCTORES
        // =========================================================
        public FormConfirmacionOCR(ColorimetricDataExtractor extractor, Bitmap bitmap)
        {
            if (extractor == null) throw new ArgumentNullException("extractor");
            if (bitmap == null) throw new ArgumentNullException("bitmap");

            _report = extractor.ExtractReportFromBitmap(bitmap);
            _rows = (_report != null && _report.Measures != null)
                ? _report.Measures : new List<ColorimetricRow>();

            InitializeComponents();
            LoadData();
            HookSizingEvents();
            AddBrandingLogo();

            // ===== NUEVO: Minimización diferida del MainForm =====
            this.Load += FormConfirmacionOCR_Load;
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
                    Width = 45,
                    Height = 45,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    BackColor = System.Drawing.Color.Transparent
                };
                
                logo.Location = new Point(this.Width - logo.Width - 15, 0);
                this.Controls.Add(logo);
                logo.BringToFront();
            }
            catch { }
        }

        public FormConfirmacionOCR(ColorimetricDataExtractor extractor, string imagePath)
        {
            if (extractor == null) throw new ArgumentNullException("extractor");
            if (string.IsNullOrWhiteSpace(imagePath)) throw new ArgumentNullException("imagePath");
            if (!System.IO.File.Exists(imagePath)) throw new System.IO.FileNotFoundException("No se encontró la imagen", imagePath);

            _report = extractor.ExtractReportFromFile(imagePath);
            _rows = (_report != null && _report.Measures != null)
                ? _report.Measures : new List<ColorimetricRow>();

            InitializeComponents();
            LoadData();
            HookSizingEvents();
            AddBrandingLogo();

            // ===== NUEVO =====
            this.Load += FormConfirmacionOCR_Load;
        }

        public FormConfirmacionOCR(OcrReport report, ShadeExtractionResult shadeResult)
        {
            _report = report ?? new OcrReport();
            _rows = _report.Measures ?? new List<ColorimetricRow>();
            _shadeResult = shadeResult;

            // Guardar para que FormResultados lo pueda ver (puente global)
            Color.ShadeReportExtractor.LastResult = shadeResult;

            InitializeComponents();
            LoadData();
            HookSizingEvents();
            AddBrandingLogo();
            this.Load += FormConfirmacionOCR_Load;
        }

        public FormConfirmacionOCR(OcrReport report)
        {
            _report = report ?? new OcrReport();
            _rows = _report.Measures ?? new List<ColorimetricRow>();

            InitializeComponents();
            LoadData();
            HookSizingEvents();
            AddBrandingLogo();

            // ===== NUEVO =====
            this.Load += FormConfirmacionOCR_Load;
        }

        [Obsolete("Usa los constructores con OcrReport o con extractor+imagen para ver también la CMC(2:1).")]
        public FormConfirmacionOCR(List<ColorimetricRow> rows)
        {
            _rows = rows ?? new List<ColorimetricRow>();
            _report = null;

            InitializeComponents();
            LoadData();
            HookSizingEvents();

            // ===== NUEVO =====
            this.Load += FormConfirmacionOCR_Load;
        }
        // =========================================================
        // NUEVO: Load → mostrar delante y minimizar MainForm
        // =========================================================
        private void FormConfirmacionOCR_Load(object sender, EventArgs e)
        {
            try
            {
                this.Activate();
                this.BringToFront();
                this.Focus();
            }
            catch { }
        }

        // =========================================================
        // UI
        // =========================================================
        private void InitializeComponents()
        {
            // ---- Ventana y escalado ----
            this.Text = "TINT COATS CADENA";

            // Barra estándar con min/max y redimensionamiento
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ControlBox = true;
            this.ShowIcon = true;
            this.BackColor = SysColor.White;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.SuspendLayout();

            // Arrancar directamente maximizado
            this.MinimumSize = new Size(980, 640);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;
            this.ResizeRedraw = true;

            // ---- Títulos ----
            lblTitulo = new Label
            {
                Text = "DATOS EXTRAÍDOS POR OCR",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = SysColor.White,
                AutoSize = false,
                BackColor = SysColor.FromArgb(30, 90, 180),
                Location = new Point(0, 0),
                Size = new Size(this.ClientSize.Width, 38),
                Padding = new Padding(20, 6, 0, 0)
            };
            var lblSubtitulo = new Label
            {
                Font = new Font("Segoe UI", 9),
                ForeColor = SysColor.FromArgb(80, 80, 80),
                AutoSize = true,
                Location = new Point(20, 45)
            };

            // ---- TabControl con UNA sola pestaña ("Combinado") ----
            tabControl = new TabControl
            {
                Location = new Point(20, 75),
                Font = new Font("Segoe UI", 9)
            };
            var tabCombined = new TabPage("📎 REPORTES")
            {
                BackColor = SysColor.White
            };

            // Layout pestaña Combinado: 6 filas (2 títulos + 4 grillas)
            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                BackColor = SysColor.White,
                Padding = new Padding(0)
            };
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));  // Título 1
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 32f));   // Grid 1
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 22f));   // Grid 2
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 30f));  // Título 2
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 22f));   // Grid 3
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 24f));   // Grid 4
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            // Fila 0: Título "DATOS DE MEDICIÓN"
            var lblDatosMedicion = new Label
            {
                Text = "SAMPLE COMPARISION",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = SysColor.FromArgb(0, 120, 215),
                BackColor = SysColor.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                Padding = new Padding(4, 0, 0, 2)
            };
            tlp.Controls.Add(lblDatosMedicion, 0, 0);

            // Fila 1: Mediciones
            dgvData = BuildMeasuresGrid();
            dgvData.Dock = DockStyle.Fill;
            tlp.Controls.Add(dgvData, 0, 1);

            // Fila 2: CMC(2:1)
            dgvCmc = BuildCmcGrid();
            dgvCmc.Dock = DockStyle.Fill;
            tlp.Controls.Add(dgvCmc, 0, 2);

            // Fila 3: Título "DATOS DE RECETA"
            lblDatosReceta = new Label
            {
                Text = "SHADE HISTORY REPORT",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = SysColor.FromArgb(0, 120, 215),
                BackColor = SysColor.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.BottomLeft,
                Padding = new Padding(4, 0, 0, 2)
            };
            tlp.Controls.Add(lblDatosReceta, 0, 3);

            // Fila 4: Receta (Código / Nombre / %)
            dgvReceta = BuildRecetaGrid();
            dgvReceta.Dock = DockStyle.Fill;
            tlp.Controls.Add(dgvReceta, 0, 4);

            // Fila 5: LAB (Resultados Finales)
            dgvLab = BuildLabGrid();
            dgvLab.Dock = DockStyle.Fill;
            tlp.Controls.Add(dgvLab, 0, 5);

            tabCombined.Controls.Add(tlp);
            tabControl.TabPages.Add(tabCombined);

            // txtRaw se mantiene solo para exportación, no se muestra en UI
            txtRaw = new TextBox { Multiline = true, Visible = false };

            // ---- Botones y etiquetas inferiores ----
            btnCancelar = new Button
            {
                Text = " Cancelar",
                Size = new Size(140, 40),
                BackColor = SysColor.FromArgb(200, 30, 30),
                ForeColor = SysColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnCancelar.FlatAppearance.BorderSize = 0;
            btnCancelar.FlatAppearance.MouseOverBackColor = SysColor.FromArgb(170, 10, 10);
            btnCancelar.FlatAppearance.MouseDownBackColor = SysColor.FromArgb(140, 0, 0);
            btnCancelar.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnConfirmar = new Button
            {
                Text = " Continuar analisis ",
                Size = new Size(160, 40),
                BackColor = SysColor.FromArgb(34, 139, 34),
                ForeColor = SysColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnConfirmar.FlatAppearance.BorderSize = 0;
            btnConfirmar.FlatAppearance.MouseOverBackColor = SysColor.FromArgb(0, 120, 0);
            btnConfirmar.FlatAppearance.MouseDownBackColor = SysColor.FromArgb(0, 100, 0);
            btnConfirmar.Click += BtnConfirmar_Click;

            btnRegresar = new Button
            {
                Text = "↩ Regresar",
                Size = new Size(150, 40),
                BackColor = SysColor.FromArgb(200, 110, 0),
                ForeColor = SysColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold)
            };
            btnRegresar.FlatAppearance.BorderSize = 0;
            btnRegresar.FlatAppearance.MouseOverBackColor = SysColor.FromArgb(170, 90, 0);
            btnRegresar.FlatAppearance.MouseDownBackColor = SysColor.FromArgb(140, 70, 0);
            btnRegresar.Click += BtnRegresar_Click;

            lblCount = new Label
            {
                Text = "Filas detectadas (Sample Comparison): 0",
                ForeColor = SysColor.FromArgb(60, 60, 60),
                Font = new Font("Segoe UI", 9),
                AutoSize = true
            };
            lblTol = new Label
            {
                Text = "",
                ForeColor = SysColor.FromArgb(80, 80, 80),
                Font = new Font("Segoe UI", 9, FontStyle.Italic),
                AutoSize = true
            };

            // Posicionar controles inferiores según tamaño actual
            PositionBottomControls();

            // Agregar a la ventana
            this.Controls.Add(lblTitulo);
            this.Controls.Add(lblSubtitulo);
            this.Controls.Add(tabControl);
            this.Controls.Add(lblTol);
            this.Controls.Add(lblCount);
            this.Controls.Add(btnConfirmar);
            this.Controls.Add(btnCancelar);
            this.Controls.Add(btnRegresar);

            this.CancelButton = btnCancelar;
            this.ResumeLayout(false);
        }

        // Reposiciona botones y etiquetas al cambiar tamaño
        private void PositionBottomControls()
        {
            int margin = 20;
            int bottomY = this.ClientSize.Height - 60;
            int right = this.ClientSize.Width - margin;

            // Botones a la derecha: Cancelar | Continuar analisis
            if (btnCancelar != null) btnCancelar.Location = new Point(right - btnCancelar.Width, bottomY);
            if (btnConfirmar != null && btnCancelar != null)
                btnConfirmar.Location = new Point(btnCancelar.Left - 10 - btnConfirmar.Width, bottomY);

            // Botón Regresar a la izquierda
            if (btnRegresar != null) btnRegresar.Location = new Point(margin, bottomY);

            // Labels entre Regresar y botones derecha
            if (lblCount != null) lblCount.Location = new Point(margin + 160, bottomY + 10);
            if (lblTol != null) lblTol.Location = new Point(margin + 160, bottomY - 14);

            // Redimensionar TabControl
            int topOfTabs = 75;
            int tabsHeight = (btnConfirmar != null ? (btnConfirmar.Top - 10) : (this.ClientSize.Height - 80)) - topOfTabs;
            if (tabControl != null)
            {
                tabControl.Location = new Point(20, topOfTabs);
                tabControl.Size = new Size(this.ClientSize.Width - 40, tabsHeight);
            }

            // lblTitulo cubre todo el ancho
            if (lblTitulo != null)
                lblTitulo.Size = new Size(this.ClientSize.Width, 38);
        }

        private void HookSizingEvents()
        {
            // Aplica proporción del divisor al cargar y cada vez que cambie el tamaño
            this.Load += (s, e) => { ApplySplitRatio(); };
            this.Resize += (s, e) =>
            {
                PositionBottomControls();
                ApplySplitRatio();
            };
        }

        private void ApplySplitRatio()
        {
            // Layout es TableLayoutPanel, no requiere ajuste de splitter
        }

        private DataGridView BuildMeasuresGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                RowHeadersVisible = false,
                BackgroundColor = SysColor.White,
                GridColor = SysColor.FromArgb(180, 180, 180),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                Font = new Font("Consolas", 10.5f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = SysColor.FromArgb(30, 90, 180);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = SysColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = SysColor.White;
            dgv.DefaultCellStyle.ForeColor = SysColor.Black;
            dgv.DefaultCellStyle.SelectionBackColor = SysColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.SelectionForeColor = SysColor.White;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = SysColor.FromArgb(230, 240, 255);

            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Illuminant", HeaderText = "Iluminante", DataPropertyName = "Illuminant" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "Tipo", DataPropertyName = "Type" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "L", HeaderText = "L*", DataPropertyName = "L" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "A", HeaderText = "a*", DataPropertyName = "A" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "B", HeaderText = "b*", DataPropertyName = "B" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Chroma", HeaderText = "Chroma", DataPropertyName = "Chroma" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Hue", HeaderText = "Hue°", DataPropertyName = "Hue" });

            DataGridViewColumn col;
            col = dgv.Columns["L"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["A"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["B"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["Chroma"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["Hue"]; col.ValueType = typeof(int); col.DefaultCellStyle.Format = "0"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            return dgv;
        }

        private DataGridView BuildCmcGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = SysColor.White,
                GridColor = SysColor.FromArgb(180, 180, 180),
                BorderStyle = BorderStyle.None,

                // Partimos midiendo por contenido; tras cargar filas, cambiaremos a Fill
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10.5f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = SysColor.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = SysColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = SysColor.White;
            dgv.DefaultCellStyle.ForeColor = SysColor.Black;
            dgv.DefaultCellStyle.SelectionBackColor = SysColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.SelectionForeColor = SysColor.White;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = SysColor.FromArgb(230, 240, 255);

            // Columnas (cabeceras)
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Illuminant", HeaderText = "Iluminante" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Deltadl", HeaderText = "dl" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Deltada", HeaderText = "da" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Deltadb", HeaderText = "db" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaCMC", HeaderText = "CMC(2:1)" });

            // Evento para Alerta Visual (Texto Rojo si supera tolerancia)
            dgv.CellFormatting += DgvCmc_CellFormatting;

            // Formatos (ahora tratamos como string para concatenar etiquetas)
            DataGridViewColumn col;
            col = dgv.Columns["Deltadl"]; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            col = dgv.Columns["Deltada"]; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            col = dgv.Columns["Deltadb"]; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            col = dgv.Columns["DeltaCMC"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // Mantener números en una sola línea
            dgv.Columns["Illuminant"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["Deltadl"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["Deltada"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["Deltadb"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["DeltaCMC"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;

            return dgv;
        }

        // =========================================================
        // CARGA DE DATOS
        // =========================================================
        private void LoadData()
        {
            LoadMeasuresSection();

            if (_report != null)
            {
                LoadCmcSection(_report.CmcDifferences ?? new List<CmcDifferenceRow>());
                SetTolerances(_report);
            }

            if (_shadeResult != null)
            {
                // SINCRONIZACIÓN: Si el ShadeResult tiene valores Std nulos o incorrectos, 
                // intentamos rescatarlos de las mediciones (D65 Std) que suelen ser más precisas
                if (_report != null && _report.Measures != null)
                {
                    var d65Std = _report.Measures.Find(m => m.Illuminant == "D65" && m.Type == "Std");

                    // Priorizar los valores del reporte si existen. Solo sincronizar si están vacíos.
                    if (d65Std != null && string.IsNullOrWhiteSpace(_shadeResult.StdL))
                    {
                        _shadeResult.StdL = d65Std.L.ToString("F2", CultureInfo.InvariantCulture);
                        _shadeResult.StdA = d65Std.A.ToString("F2", CultureInfo.InvariantCulture);
                        _shadeResult.StdB = d65Std.B.ToString("F2", CultureInfo.InvariantCulture);
                    }
                }

                LoadRecetaSection(_shadeResult);
                LoadLabSection(_shadeResult);
            }

            txtRaw.Text = BuildTextView();
        }



        private void LoadMeasuresSection()
        {
            dgvData.Rows.Clear();



            for (int i = 0; i < _rows.Count; i++)
            {
                ColorimetricRow r = _rows[i];
                int idx = dgvData.Rows.Add(r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue);

                SysColor rowColor = SysColor.White;
                if (r.Illuminant == "D65") rowColor = SysColor.FromArgb(210, 225, 255);
                else if (r.Illuminant == "TL84") rowColor = SysColor.FromArgb(210, 240, 220);
                else if (r.Illuminant == "A") rowColor = SysColor.FromArgb(255, 245, 210);
                else if (r.Illuminant == "CWF") rowColor = SysColor.FromArgb(255, 230, 210);

                dgvData.Rows[idx].DefaultCellStyle.BackColor = rowColor;
                dgvData.Rows[idx].DefaultCellStyle.ForeColor = SysColor.Black;

            }

            lblCount.Text = "Filas detectadas (Sample Comparison): " + _rows.Count;

            // Ajuste por contenido visible y luego rellenar sin aplastar
            dgvData.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // Pesos/mínimos
            dgvData.Columns["Illuminant"].FillWeight = 85f; dgvData.Columns["Illuminant"].MinimumWidth = 90;
            dgvData.Columns["Type"].FillWeight = 70f; dgvData.Columns["Type"].MinimumWidth = 70;
            dgvData.Columns["L"].FillWeight = 80f; dgvData.Columns["L"].MinimumWidth = 80;
            dgvData.Columns["A"].FillWeight = 80f; dgvData.Columns["A"].MinimumWidth = 80;
            dgvData.Columns["B"].FillWeight = 80f; dgvData.Columns["B"].MinimumWidth = 80;
            dgvData.Columns["Chroma"].FillWeight = 95f; dgvData.Columns["Chroma"].MinimumWidth = 90;
            dgvData.Columns["Hue"].FillWeight = 70f; dgvData.Columns["Hue"].MinimumWidth = 70;

            // Filas: números en una línea
            dgvData.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        }

        private void LoadCmcSection(List<CmcDifferenceRow> cmc)
        {
            dgvCmc.Rows.Clear();
            if (cmc == null || cmc.Count == 0) return;

            for (int i = 0; i < cmc.Count; i++)
            {
                CmcDifferenceRow r = cmc[i];
                object cmcVal = r.DeltaCMC.HasValue ? (object)r.DeltaCMC.Value : null;

                // ETAPA 1.2: Formato idéntico al reporte (Valor Etiqueta) sin paréntesis
                string tagL = !string.IsNullOrWhiteSpace(r.LightnessFlagOcr) ? r.LightnessFlagOcr : r.LightnessFlag;
                string tagC = !string.IsNullOrWhiteSpace(r.ChromaFlagOcr) ? r.ChromaFlagOcr : r.ChromaFlag;
                string tagH = !string.IsNullOrWhiteSpace(r.HueFlagOcr) ? r.HueFlagOcr : r.HueFlag;

                string dLStr = string.Format("{0:F2} {1}", r.DeltaLightness, tagL);
                string dCStr = string.Format("{0:F2} {1}", r.DeltaChroma, tagC);
                string dHStr = string.Format("{0:F2} {1}", r.DeltaHue, tagH);

                dgvCmc.Rows.Add(
                    r.Illuminant,
                    dLStr,
                    dCStr,
                    dHStr,
                    cmcVal);

                SysColor rowColor = SysColor.White;
                if (r.Illuminant == "D65") rowColor = SysColor.FromArgb(210, 225, 255);
                else if (r.Illuminant == "TL84") rowColor = SysColor.FromArgb(210, 240, 220);
                else if (r.Illuminant == "A") rowColor = SysColor.FromArgb(255, 245, 210);

                dgvCmc.Rows[dgvCmc.Rows.Count - 1].DefaultCellStyle.BackColor = rowColor;
                dgvCmc.Rows[dgvCmc.Rows.Count - 1].DefaultCellStyle.ForeColor = SysColor.Black;
            }

            // 1) Mide por contenido visible
            dgvCmc.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            // 2) Rellenar con Fill equilibrando pesos y mínimos
            dgvCmc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvCmc.Columns["Illuminant"].FillWeight = 80f; dgvCmc.Columns["Illuminant"].MinimumWidth = 90;
            dgvCmc.Columns["Deltadl"].FillWeight = 70f; dgvCmc.Columns["Deltadl"].MinimumWidth = 90;
            dgvCmc.Columns["Deltada"].FillWeight = 70f; dgvCmc.Columns["Deltada"].MinimumWidth = 90;
            dgvCmc.Columns["Deltadb"].FillWeight = 70f; dgvCmc.Columns["Deltadb"].MinimumWidth = 80;
            dgvCmc.Columns["DeltaCMC"].FillWeight = 110f; dgvCmc.Columns["DeltaCMC"].MinimumWidth = 130;

            // 3) Alto automático por celdas mostradas 
            dgvCmc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
        }

        private void DgvCmc_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.Value == null) return;
            string colName = dgvCmc.Columns[e.ColumnIndex].Name;

            // Columnas de interés para alerta visual
            if (colName == "Deltadl" || colName == "Deltada" || colName == "Deltadb" || colName == "DeltaCMC")
            {
                double val = 0;
                string strVal = e.Value.ToString();
                
                // Extraer número para comparación de límite (0.25)
                string numPart = Regex.Match(strVal.Replace(',', '.'), @"\-?\d+\.?\d*").Value;
                if (double.TryParse(numPart, NumberStyles.Any, CultureInfo.InvariantCulture, out val))
                {
                    if (Math.Abs(val) > 0.25)
                    {
                        e.CellStyle.ForeColor = SysColor.Black;
                        e.CellStyle.SelectionForeColor = SysColor.Black;
                        e.CellStyle.Font = new Font(dgvCmc.Font, FontStyle.Bold);
                    }
                }
            }
        }

        private void SetTolerances(OcrReport rep)
        {
            if (rep == null) { lblTol.Text = ""; return; }
            bool anyTol = (rep.TolDL != 0) || (rep.TolDC != 0) || (rep.TolDH != 0) || (rep.TolDE != 0);

            lblTol.Text = " ";
        }

        // =========================================================
        // VISTA TEXTO
        // =========================================================
        private string BuildTextView()
        {
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("══════════════════════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine(" DATOS COLORIMÉTRICOS EXTRAÍDOS (OCR)");
            sb.AppendLine(string.Format(" Fecha: {0:dd/MM/yyyy HH:mm:ss}", DateTime.Now));
            sb.AppendLine("══════════════════════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine();
            sb.AppendLine("Mediciones");
            sb.AppendLine(string.Format("{0,-10} {1,-6} {2,8} {3,8} {4,8} {5,8} {6,8}",
                "Iluminante", "Tipo", "L*", "a*", "b*", "Chroma", "Hue°"));
            sb.AppendLine("──────────────────────────────────────────────────────────────────────────────────────────────────────────");

            string lastIll = "";
            for (int i = 0; i < _rows.Count; i++)
            {
                ColorimetricRow r = _rows[i];
                if (r.Illuminant != lastIll && lastIll != "")
                    sb.AppendLine("──────────────────────────────────────────────────────────────────────────────────────────────────────────");

                sb.AppendLine(string.Format("{0,-10} {1,-6} {2,8:0.00} {3,8:0.00} {4,8:0.00} {5,8:0.00} {6,8:0}",
                    r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue));
                lastIll = r.Illuminant;
            }

            if (_report != null && _report.CmcDifferences != null && _report.CmcDifferences.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("CMC(2:1) — Difference in (Lightness / Chroma / Hue) [Col Diff CMC(2:1)]");
                sb.AppendLine(string.Format("{0,-10} {1,8} {2,8} {3,8} {4,8} {5,-10} {6,-14}",
                    "Iluminante", "ΔL", "ΔC", "ΔH", "CMC", "Claridad", "Croma/Hue"));
                sb.AppendLine("──────────────────────────────────────────────────────────────────────────────────────────────────────────");

                for (int i = 0; i < _report.CmcDifferences.Count; i++)
                {
                    CmcDifferenceRow c = _report.CmcDifferences[i];
                    string cmcStr = c.DeltaCMC.HasValue
                        ? c.DeltaCMC.Value.ToString("0.00", CultureInfo.InvariantCulture) : "";

                    sb.AppendLine(string.Format("{0,-10} {1,8:0.00} {2,8:0.00} {3,8:0.00} {4,8} {5,-10} {6,-14}",
                        c.Illuminant, c.DeltaLightness, c.DeltaChroma, c.DeltaHue,
                        cmcStr, c.LightnessFlag, c.ChromaHueFlag));
                }
            }

            if (_report != null &&
               ((_report.TolDL != 0) || (_report.TolDC != 0) || (_report.TolDH != 0) || (_report.TolDE != 0)))
            {
                sb.AppendLine();
                sb.AppendLine("Tolerancias:");
                sb.AppendLine(string.Format(" DL: {0:0.00} DC: {1:0.00} DH: {2:0.00} DE: {3:0.00}",
                    _report.TolDL, _report.TolDC, _report.TolDH, _report.TolDE));
            }

            if (_report != null && !string.IsNullOrWhiteSpace(_report.PrintDate))
            {
                sb.AppendLine();
                sb.AppendLine(_report.PrintDate);
            }

            sb.AppendLine("══════════════════════════════════════════════════════════════════════════════════════════════════════════");
            return sb.ToString();
        }

        // =========================================================
        // CONFIRMAR
        // =========================================================
        private void BtnConfirmar_Click(object sender, EventArgs e)
        {
            if (dgvData.IsCurrentCellInEditMode) dgvData.EndEdit();
            if (dgvReceta != null && dgvReceta.IsCurrentCellInEditMode) dgvReceta.EndEdit();
            if (dgvCmc != null && dgvCmc.IsCurrentCellInEditMode) dgvCmc.EndEdit();
            if (dgvLab != null && dgvLab.IsCurrentCellInEditMode) dgvLab.EndEdit();

            // 1. Sincronizar Mediciones (dgvData -> _rows)
            _rows = new List<ColorimetricRow>();
            foreach (DataGridViewRow row in dgvData.Rows)
            {
                if (row.IsNewRow) continue;
                try
                {
                    _rows.Add(new ColorimetricRow
                    {
                        Illuminant = row.Cells["Illuminant"].Value?.ToString(),
                        Type = row.Cells["Type"].Value?.ToString(),
                        L = ParseCellDouble(row.Cells["L"].Value),
                        A = ParseCellDouble(row.Cells["A"].Value),
                        B = ParseCellDouble(row.Cells["B"].Value),
                        Chroma = ParseCellDouble(row.Cells["Chroma"].Value),
                        Hue = ParseCellInt(row.Cells["Hue"].Value)
                    });
                } catch { }
            }
            RowsConfirmed = _rows;
            if (_report != null) _report.Measures = _rows;

            // 2. Sincronizar CMC (dgvCmc -> _report.CmcDifferences)
            if (_report != null && dgvCmc != null)
            {
                _report.CmcDifferences = new List<CmcDifferenceRow>();
                foreach (DataGridViewRow row in dgvCmc.Rows)
                {
                    if (row.IsNewRow) continue;
                    try
                    {
                        _report.CmcDifferences.Add(new CmcDifferenceRow
                        {
                            Illuminant = row.Cells["Illuminante"].Value?.ToString(),
                            DeltaLightness = ParseCellDouble(row.Cells["dl"].Value),
                            DeltaChroma = ParseCellDouble(row.Cells["da"].Value),
                            DeltaHue = ParseCellDouble(row.Cells["db"].Value),
                            DeltaCMC = ParseCellDouble(row.Cells["cmc"].Value)
                        });
                    } catch { }
                }
            }

            // 3. Sincronizar Receta (dgvReceta -> _shadeResult.Recipe)
            if (_shadeResult != null && dgvReceta != null)
            {
                _shadeResult.Recipe = new List<RecipeItem>();
                foreach (DataGridViewRow row in dgvReceta.Rows)
                {
                    if (row.IsNewRow) continue;
                    try
                    {
                        _shadeResult.Recipe.Add(new RecipeItem
                        {
                            Code = row.Cells["Codigo"].Value?.ToString(),
                            Name = row.Cells["Nombre"].Value?.ToString(),
                            Percentage = row.Cells["Porcentaje"].Value?.ToString()
                        });
                    } catch { }
                }
            }

            // 4. Sincronizar LAB/Std/Lot (dgvLab -> _shadeResult)
            if (_shadeResult != null && dgvLab != null)
            {
                foreach (DataGridViewRow row in dgvLab.Rows)
                {
                    string tipo = row.Cells["Tipo"].Value?.ToString() ?? "";
                    if (tipo.Equals("Std", StringComparison.OrdinalIgnoreCase))
                    {
                        _shadeResult.StdL = row.Cells["L"].Value?.ToString();
                        _shadeResult.StdA = row.Cells["A"].Value?.ToString();
                        _shadeResult.StdB = row.Cells["B"].Value?.ToString();
                    }
                    else if (tipo.Equals("Lot", StringComparison.OrdinalIgnoreCase))
                    {
                        if (_shadeResult.Batch == null) _shadeResult.Batch = new BatchMeasure();
                        var b = _shadeResult.Batch;
                        b.L = row.Cells["L"].Value?.ToString();
                        b.A = row.Cells["A"].Value?.ToString();
                        b.B = row.Cells["B"].Value?.ToString();
                        b.DL = row.Cells["dL"].Value?.ToString();
                        b.DC = row.Cells["dC"].Value?.ToString();
                        b.DH = row.Cells["dH"].Value?.ToString();
                        b.DE = row.Cells["dE"].Value?.ToString();
                        b.PF = row.Cells["PF"].Value?.ToString();
                    }
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // =========================================================
        // REGRESAR A PANTALLA PRINCIPAL (Form1)
        // =========================================================
        private void BtnRegresar_Click(object sender, EventArgs e)
        {
            try
            {
                if (MainFormOwner != null && !MainFormOwner.IsDisposed)
                {
                    MainFormOwner.WindowState = FormWindowState.Normal;
                    MainFormOwner.BringToFront();
                    MainFormOwner.Activate();
                }
            }
            catch { }

            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        // Asegurar resultados coherentes al cerrar con "X"
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.None)
                this.DialogResult = DialogResult.Cancel;

            base.OnFormClosing(e);
        }

        // =========================================================
        // GRILLA DE RECETA
        // =========================================================
        private DataGridView BuildRecetaGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = SysColor.White,
                GridColor = SysColor.FromArgb(180, 180, 180),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = SysColor.FromArgb(30, 90, 180);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = SysColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = SysColor.White;
            dgv.DefaultCellStyle.ForeColor = SysColor.Black;
            dgv.DefaultCellStyle.SelectionBackColor = SysColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.SelectionForeColor = SysColor.White;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            dgv.AlternatingRowsDefaultCellStyle.BackColor = SysColor.FromArgb(230, 240, 255);

            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Codigo", HeaderText = "Código", FillWeight = 90f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Nombre", HeaderText = "Nombre", FillWeight = 300f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Porcentaje", HeaderText = "%", FillWeight = 80f });

            dgv.Columns["Porcentaje"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            dgv.ReadOnly = false;
            dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;

            return dgv;
        }

        private void LoadRecetaSection(ShadeExtractionResult shade)
        {
            if (dgvReceta == null || shade == null) return;
            dgvReceta.Rows.Clear();

            // Actualizar título con el Número de Lote si se encontró
            if (!string.IsNullOrWhiteSpace(shade.LotNo))
            {
                lblDatosReceta.Text = $"SHADE HISTORY REPORT";
            }
            else
            {
                lblDatosReceta.Text = "SHADE HISTORY REPORT";
            }

            // ✔ CORREGIDO: ya no evaluamos shade.Success
            if (shade == null ||
                shade.Recipe == null ||
                shade.Recipe.Count == 0)
            {
                dgvReceta.Rows.Add("", "Sin datos de Shade History Report", "");
                return;
            }

            foreach (var item in shade.Recipe)
            {
                int idx = dgvReceta.Rows.Add(item.Code, item.Name, item.Percentage);
                dgvReceta.Rows[idx].DefaultCellStyle.ForeColor = SysColor.Black;
            }
        }

        // =========================================================
        // GRILLA LAB (Batch Measure)
        // =========================================================
        private DataGridView BuildLabGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = false,
                RowHeadersVisible = false,
                BackgroundColor = SysColor.White,
                GridColor = SysColor.FromArgb(180, 180, 180),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                Font = new Font("Consolas", 10f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = SysColor.FromArgb(30, 90, 180);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = SysColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.DefaultCellStyle.BackColor = SysColor.White;
            dgv.DefaultCellStyle.ForeColor = SysColor.Black;
            dgv.DefaultCellStyle.SelectionBackColor = SysColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.SelectionForeColor = SysColor.White;
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Tipo", HeaderText = "Tipo", FillWeight = 60f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "L", HeaderText = "L", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "A", HeaderText = "A", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "B", HeaderText = "B", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "dL", HeaderText = "dL", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "dC", HeaderText = "dC", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "dH", HeaderText = "dH", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "dE", HeaderText = "dE", FillWeight = 80f });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "PF", HeaderText = "P/F", FillWeight = 60f });
            return dgv;
        }

        private void LoadLabSection(ShadeExtractionResult shade)
        {
            if (dgvLab == null || shade == null) return;
            dgvLab.Rows.Clear();

            // 1. Agregar Estándar (Std L A B) si existe
            if (!string.IsNullOrWhiteSpace(shade.StdL))
            {
                string label = "Std";
                int sIdx = dgvLab.Rows.Add(label, shade.StdL, shade.StdA, shade.StdB, "", "", "", "", "");
                dgvLab.Rows[sIdx].DefaultCellStyle.BackColor = SysColor.FromArgb(240, 240, 240);
                dgvLab.Rows[sIdx].DefaultCellStyle.ForeColor = SysColor.Black;
                dgvLab.Rows[sIdx].DefaultCellStyle.Font = new Font(dgvLab.Font, FontStyle.Bold);
            }
        }

        private double ParseCellDouble(object val)
        {
            if (val == null) return 0;
            if (val is double) return (double)val;
            if (val is float) return Math.Round((float)val, 2);
            if (val is decimal) return (double)(decimal)val;
            if (val is int) return (int)val;
            if (val is long) return (long)val;

            string s = Convert.ToString(val, CultureInfo.InvariantCulture);
            double dv;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out dv)) return dv;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out dv)) return dv;
            return 0;
        }

        private int ParseCellInt(object val)
        {
            double d = ParseCellDouble(val);
            int h = (int)Math.Round(d, 0, MidpointRounding.AwayFromZero);
            if (h < 0) h = 0;
            if (h >= 360) h = h % 360;
            return h;
        }
    }
}