// Formconfirmacionocr.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Colorimetria
{
    /// <summary>
    /// Verifica datos extraídos por OCR: Mediciones + CMC(2:1) + Tolerances/PrintDate.
    /// Vista combinada: una sola pestaña con Split arriba (Mediciones ? / CMC ?) y Texto abajo.
    /// </summary>
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

        // ===== UI =====
        private DataGridView dgvData;
        private DataGridView dgvCmc;
        private TextBox txtRaw;
        private TabControl tabControl;
        private Button btnConfirmar;
        private Button btnCancelar;
        private Label lblTitulo;
        private Label lblSubtitulo;
        private Label lblCount;
        private Label lblTol;

        // Contenedor superior (arriba/abajo) y su proporción (70% arriba, 30% abajo)
        private SplitContainer splitTop;
        private double splitTopRatio = 0.70;

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

            // ===== NUEVO: Minimización diferida del MainForm =====
            this.Load += FormConfirmacionOCR_Load;
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

            // ===== NUEVO =====
            this.Load += FormConfirmacionOCR_Load;
        }

        public FormConfirmacionOCR(OcrReport report)
        {
            _report = report ?? new OcrReport();
            _rows = _report.Measures ?? new List<ColorimetricRow>();

            InitializeComponents();
            LoadData();
            HookSizingEvents();

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
        // NUEVO: Load ? mostrar delante y minimizar MainForm
        // =========================================================
        private void FormConfirmacionOCR_Load(object sender, EventArgs e)
        {
            try
            {
                // Asegurar que este diálogo esté visible y al frente
                this.Show();
                this.Activate();
                this.BringToFront();

                // Minimizar el formulario principal después de mostrar este
                if (MainFormOwner != null)
                {
                    MainFormOwner.WindowState = FormWindowState.Minimized;
                }
            }
            catch
            {
                // Ignorar: no debe bloquear la verificación por problemas de foco
            }
        }

        // =========================================================
        // UI
        // =========================================================
        private void InitializeComponents()
        {
            // ---- Ventana y escalado ----
            this.Text = "Verificar datos extraídos";
            // Barra estándar con min/max y redimensionamiento
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.ControlBox = true;
            this.ShowIcon = true;
            this.BackColor = System.Drawing.Color.FromArgb(30, 30, 30);
            this.AutoScaleMode = AutoScaleMode.Dpi; // respeta 125%, 150%, etc.

            // Tamaño dinámico: 90% del área de trabajo
            var wa = Screen.PrimaryScreen.WorkingArea;
            int targetWidth = (int)(wa.Width * 0.90);
            int targetHeight = (int)(wa.Height * 0.90);
            this.MinimumSize = new Size(980, 640);
            this.Size = new Size(targetWidth, targetHeight);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResizeRedraw = true;

            // ---- Títulos ----
            lblTitulo = new Label
            {
                Text = "DATOS EXTRAÍDOS POR OCR",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 15)
            };
            lblSubtitulo = new Label
            {
                Text = "Revisa los valores antes de continuar con los cálculos de corrección colorimétrica.",
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.LightGray,
                AutoSize = true,
                Location = new Point(20, 45)
            };

            // ---- TabControl con UNA sola pestaña ("Combinado") ----
            tabControl = new TabControl
            {
                Location = new Point(20, 75),
                Font = new Font("Segoe UI", 9)
            };
            var tabCombined = new TabPage("?? Combinado")
            {
                BackColor = System.Drawing.Color.FromArgb(45, 45, 45)
            };

            // Layout principal: 2 filas (superior: split arriba/abajo, inferior: texto)
            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = System.Drawing.Color.FromArgb(45, 45, 45)
            };
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 75f)); // más espacio para tablas apiladas
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 25f)); // menos para el bloque de texto
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            // Split superior (ARRIBA: Mediciones / ABAJO: CMC)
            splitTop = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal, // arriba/abajo
                BackColor = System.Drawing.Color.FromArgb(45, 45, 45)
            };

            // Grilla de Mediciones (ARRIBA)
            dgvData = BuildMeasuresGrid();
            dgvData.Dock = DockStyle.Fill;
            splitTop.Panel1.Controls.Add(dgvData);

            // Grilla de CMC (ABAJO)
            dgvCmc = BuildCmcGrid();
            dgvCmc.Dock = DockStyle.Fill;
            splitTop.Panel2.Controls.Add(dgvCmc);

            // Vista texto (ABAJO del tlp)
            txtRaw = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.5f),
                BackColor = System.Drawing.Color.FromArgb(20, 20, 20),
                ForeColor = Color.LightGreen,
                ReadOnly = true,
                WordWrap = false
            };

            tlp.Controls.Add(splitTop, 0, 0);
            tlp.Controls.Add(txtRaw, 0, 1);
            tabCombined.Controls.Add(tlp);
            tabControl.TabPages.Add(tabCombined);

            // ---- Botones y etiquetas inferiores ----
            btnCancelar = new Button
            {
                Text = "? Cancelar",
                Size = new Size(160, 40),
                BackColor = System.Drawing.Color.FromArgb(180, 50, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.Cancel
            };
            btnCancelar.FlatAppearance.BorderSize = 0;
            btnCancelar.Click += delegate { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnConfirmar = new Button
            {
                Text = "? Confirmar",
                Size = new Size(160, 40),
                BackColor = System.Drawing.Color.FromArgb(46, 125, 50),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnConfirmar.FlatAppearance.BorderSize = 0;
            btnConfirmar.Click += BtnConfirmar_Click;

            lblCount = new Label
            {
                Text = "Filas detectadas (mediciones): 0",
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9),
                AutoSize = true
            };
            lblTol = new Label
            {
                Text = "",
                ForeColor = Color.DarkGray,
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

            this.AcceptButton = btnConfirmar;
            this.CancelButton = btnCancelar;
        }

        // Reposiciona botones y etiquetas al cambiar tamaño
        private void PositionBottomControls()
        {
            int margin = 20;
            int bottomY = this.ClientSize.Height - 60; // altura aproximada
            int right = this.ClientSize.Width - margin;

            // Botones a la derecha
            if (btnCancelar != null) btnCancelar.Location = new Point(right - btnCancelar.Width, bottomY);
            if (btnConfirmar != null && btnCancelar != null)
                btnConfirmar.Location = new Point(btnCancelar.Left - 10 - btnConfirmar.Width, bottomY);

            // Labels a la izquierda
            if (lblCount != null) lblCount.Location = new Point(20, bottomY);
            if (lblTol != null) lblTol.Location = new Point(20, bottomY - 24);

            // Redimensionar TabControl para ocupar el centro
            int topOfTabs = 75; // debajo del subtítulo
            int tabsHeight = (btnConfirmar != null ? (btnConfirmar.Top - 10) : (this.ClientSize.Height - 80)) - topOfTabs;
            if (tabControl != null)
            {
                tabControl.Location = new Point(20, topOfTabs);
                tabControl.Size = new Size(this.ClientSize.Width - 40, tabsHeight);
            }
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
            try
            {
                if (splitTop != null && splitTop.Orientation == Orientation.Horizontal)
                {
                    int h = splitTop.ClientSize.Height;
                    int distance = (int)(h * splitTopRatio); // 70% para Mediciones (arriba)
                    distance = Math.Max(180, Math.Min(h - 180, distance)); // margen mínimo
                    splitTop.SplitterDistance = distance;
                }
            }
            catch
            {
                // Ignorar si el control aún no está listo para medir
            }
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
                BackgroundColor = System.Drawing.Color.FromArgb(45, 45, 45),
                GridColor = System.Drawing.Color.FromArgb(80, 80, 80),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, // mide por contenido
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10.5f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(45, 45, 45);

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
                BackgroundColor = System.Drawing.Color.FromArgb(45, 45, 45),
                GridColor = System.Drawing.Color.FromArgb(80, 80, 80),
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

            dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = Color.White;
            dgv.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(45, 45, 45);

            // Columnas (cabeceras)
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Illuminant", HeaderText = "Iluminante" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaLightness", HeaderText = "?L (Lightness)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaChroma", HeaderText = "?C (Chroma)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaHue", HeaderText = "?H (Hue)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaCMC", HeaderText = "CMC(2:1)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "LightnessFlag", HeaderText = "Claridad" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "ChromaHueFlag", HeaderText = "Croma/Hue" });

            // Formatos numéricos
            DataGridViewColumn col;
            col = dgv.Columns["DeltaLightness"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["DeltaChroma"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["DeltaHue"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            col = dgv.Columns["DeltaCMC"]; col.ValueType = typeof(double); col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            // Wrap en flags (textos largos)
            dgv.Columns["LightnessFlag"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgv.Columns["ChromaHueFlag"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            // Mantener números en una sola línea
            dgv.Columns["Illuminant"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["DeltaLightness"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["DeltaChroma"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            dgv.Columns["DeltaHue"].DefaultCellStyle.WrapMode = DataGridViewTriState.False;
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

            txtRaw.Text = BuildTextView();
        }

        private void LoadMeasuresSection()
        {
            dgvData.Rows.Clear();

            for (int i = 0; i < _rows.Count; i++)
            {
                ColorimetricRow r = _rows[i];
                int idx = dgvData.Rows.Add(r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue);

                Color rowColor = System.Drawing.Color.FromArgb(55, 55, 55);
                if (r.Illuminant == "D65") rowColor = System.Drawing.Color.FromArgb(40, 60, 100);
                else if (r.Illuminant == "TL84") rowColor = System.Drawing.Color.FromArgb(40, 80, 60);
                else if (r.Illuminant == "A") rowColor = System.Drawing.Color.FromArgb(80, 80, 40);
                else if (r.Illuminant == "CWF") rowColor = System.Drawing.Color.FromArgb(80, 55, 40);

                dgvData.Rows[idx].DefaultCellStyle.BackColor = rowColor;
            }

            lblCount.Text = "Filas detectadas (mediciones): " + _rows.Count;

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

                dgvCmc.Rows.Add(
                    r.Illuminant,
                    r.DeltaLightness,
                    r.DeltaChroma,
                    r.DeltaHue,
                    cmcVal,
                    r.LightnessFlag,
                    r.ChromaHueFlag);

                Color rowColor = System.Drawing.Color.FromArgb(55, 55, 55);
                if (r.Illuminant == "D65") rowColor = System.Drawing.Color.FromArgb(40, 60, 100);
                else if (r.Illuminant == "TL84") rowColor = System.Drawing.Color.FromArgb(40, 80, 60);
                else if (r.Illuminant == "A") rowColor = System.Drawing.Color.FromArgb(80, 80, 40);

                dgvCmc.Rows[dgvCmc.Rows.Count - 1].DefaultCellStyle.BackColor = rowColor;
            }

            // 1) Mide por contenido visible
            dgvCmc.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);

            // 2) Rellenar con Fill equilibrando pesos y mínimos
            dgvCmc.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvCmc.Columns["Illuminant"].FillWeight = 80f; dgvCmc.Columns["Illuminant"].MinimumWidth = 90;
            dgvCmc.Columns["DeltaLightness"].FillWeight = 70f; dgvCmc.Columns["DeltaLightness"].MinimumWidth = 90;
            dgvCmc.Columns["DeltaChroma"].FillWeight = 70f; dgvCmc.Columns["DeltaChroma"].MinimumWidth = 90;
            dgvCmc.Columns["DeltaHue"].FillWeight = 70f; dgvCmc.Columns["DeltaHue"].MinimumWidth = 80;
            dgvCmc.Columns["DeltaCMC"].FillWeight = 110f; dgvCmc.Columns["DeltaCMC"].MinimumWidth = 130;

            // Flags (textos) con wrap y alto automático
            dgvCmc.Columns["LightnessFlag"].FillWeight = 85f; dgvCmc.Columns["LightnessFlag"].MinimumWidth = 110;
            dgvCmc.Columns["ChromaHueFlag"].FillWeight = 100f; dgvCmc.Columns["ChromaHueFlag"].MinimumWidth = 120;

            // 3) Alto automático por celdas mostradas (solo necesario en CMC por el wrap)
            dgvCmc.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells;
        }

        private void SetTolerances(OcrReport rep)
        {
            if (rep == null) { lblTol.Text = ""; return; }
            bool anyTol = (rep.TolDL != 0) || (rep.TolDC != 0) || (rep.TolDH != 0) || (rep.TolDE != 0);

            lblTol.Text = anyTol
                ? string.Format("Tolerancias — DL: {0:0.00} DC: {1:0.00} DH: {2:0.00} DE: {3:0.00}",
                    rep.TolDL, rep.TolDC, rep.TolDH, rep.TolDE)
                : "";
        }

        // =========================================================
        // VISTA TEXTO
        // =========================================================
        private string BuildTextView()
        {
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("----------------------------------------------------------------------------------------------------------");
            sb.AppendLine(" DATOS COLORIMÉTRICOS EXTRAÍDOS (OCR)");
            sb.AppendLine(string.Format(" Fecha: {0:dd/MM/yyyy HH:mm:ss}", DateTime.Now));
            sb.AppendLine("----------------------------------------------------------------------------------------------------------");
            sb.AppendLine();
            sb.AppendLine("Mediciones");
            sb.AppendLine(string.Format("{0,-10} {1,-6} {2,8} {3,8} {4,8} {5,8} {6,8}",
                "Iluminante", "Tipo", "L*", "a*", "b*", "Chroma", "Hue°"));
            sb.AppendLine("----------------------------------------------------------------------------------------------------------");

            string lastIll = "";
            for (int i = 0; i < _rows.Count; i++)
            {
                ColorimetricRow r = _rows[i];
                if (r.Illuminant != lastIll && lastIll != "")
                    sb.AppendLine("----------------------------------------------------------------------------------------------------------");

                sb.AppendLine(string.Format("{0,-10} {1,-6} {2,8:0.00} {3,8:0.00} {4,8:0.00} {5,8:0.00} {6,8:0}",
                    r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue));
                lastIll = r.Illuminant;
            }

            if (_report != null && _report.CmcDifferences != null && _report.CmcDifferences.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("CMC(2:1) — Difference in (Lightness / Chroma / Hue) [Col Diff CMC(2:1)]");
                sb.AppendLine(string.Format("{0,-10} {1,8} {2,8} {3,8} {4,8} {5,-10} {6,-14}",
                    "Iluminante", "?L", "?C", "?H", "CMC", "Claridad", "Croma/Hue"));
                sb.AppendLine("----------------------------------------------------------------------------------------------------------");

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
                sb.AppendLine("Tolerances:");
                sb.AppendLine(string.Format(" DL: {0:0.00} DC: {1:0.00} DH: {2:0.00} DE: {3:0.00}",
                    _report.TolDL, _report.TolDC, _report.TolDH, _report.TolDE));
            }

            if (_report != null && !string.IsNullOrWhiteSpace(_report.PrintDate))
            {
                sb.AppendLine();
                sb.AppendLine(_report.PrintDate);
            }

            sb.AppendLine("----------------------------------------------------------------------------------------------------------");
            return sb.ToString();
        }

        // =========================================================
        // CONFIRMAR
        // =========================================================
        private void BtnConfirmar_Click(object sender, EventArgs e)
        {
            RowsConfirmed = new List<ColorimetricRow>();

            foreach (DataGridViewRow row in dgvData.Rows)
            {
                try
                {
                    var r = new ColorimetricRow
                    {
                        Illuminant = row.Cells["Illuminant"].Value != null ? row.Cells["Illuminant"].Value.ToString() : null,
                        Type = row.Cells["Type"].Value != null ? row.Cells["Type"].Value.ToString() : null,
                        L = ParseCellDouble(row.Cells["L"].Value),
                        A = ParseCellDouble(row.Cells["A"].Value),
                        B = ParseCellDouble(row.Cells["B"].Value),
                        Chroma = ParseCellDouble(row.Cells["Chroma"].Value),
                        Hue = ParseCellInt(row.Cells["Hue"].Value)
                    };
                    RowsConfirmed.Add(r);
                }
                catch
                {
                    /* ignorar filas corruptas */
                }
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // Asegurar resultados coherentes al cerrar con "X"
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.None)
                this.DialogResult = DialogResult.Cancel;

            base.OnFormClosing(e);
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
