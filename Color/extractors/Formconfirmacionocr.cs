// Formconfirmacionocr.cs
using Colorimetria;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using DrawingColor = System.Drawing.Color;

namespace Color
{
    /// <summary>
    /// Verifica datos extraídos por OCR: Mediciones + CMC(2:1) + Tolerances/PrintDate.
    /// Vista combinada: una sola pestaña con Split arriba (Mediciones ↑ / CMC ↓) y Texto abajo.
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

        // ===== Shade (receta + batch) =====
        private ShadeExtractionResult _shadeResult;

        // ===== UI =====
        private DataGridView dgvData;
        private DataGridView dgvCmc;
        private DataGridView dgvRecipe;
        private DataGridView dgvBatch;
        private TextBox txtRaw;
        private TabControl tabControl;
        private Button btnConfirmar;
        private Button btnCancelar;
        private Label lblTitulo;
        private Label lblSubtitulo;
        private Label lblCount;
        private Label lblTol;

        // Splitters
        private SplitContainer splitTop;     // Mediciones | CMC
        private SplitContainer splitBottom;  // Receta | Batch
        private SplitContainer splitMain;    // splitTop | splitBottom
        private double splitTopRatio = 0.50;
        private double splitMainRatio = 0.55;

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

        public FormConfirmacionOCR(OcrReport report,
            ShadeExtractionResult shade = null)
        {
            _report = report ?? new OcrReport();
            _rows = _report.Measures ?? new List<ColorimetricRow>();
            _shadeResult = shade;

            InitializeComponents();
            LoadData();
            HookSizingEvents();

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
            this.BackColor = DrawingColor.FromArgb(30, 30, 30);
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
                ForeColor = DrawingColor.White,
                AutoSize = true,
                Location = new Point(20, 15)
            };
            lblSubtitulo = new Label
            {
                Text = "Revisa los valores antes de continuar con los cálculos de corrección colorimétrica.",
                Font = new Font("Segoe UI", 9),
                ForeColor = DrawingColor.LightGray,
                AutoSize = true,
                Location = new Point(20, 45)
            };

            // ---- TabControl con UNA sola pestaña ("Combinado") ----
            tabControl = new TabControl
            {
                Location = new Point(20, 75),
                Font = new Font("Segoe UI", 9)
            };
            var tabCombined = new TabPage("📎 Combinado")
            {
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };

            // ── Layout: splitMain vertical (arriba: mediciones+CMC / abajo: receta+batch)
            splitTop = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };
            dgvData = BuildMeasuresGrid();
            dgvData.Dock = DockStyle.Fill;
            splitTop.Panel1.Controls.Add(dgvData);

            dgvCmc = BuildCmcGrid();
            dgvCmc.Dock = DockStyle.Fill;
            splitTop.Panel2.Controls.Add(dgvCmc);

            splitBottom = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };
            dgvRecipe = BuildRecipeGrid();
            dgvRecipe.Dock = DockStyle.Fill;
            splitBottom.Panel1.Controls.Add(dgvRecipe);

            dgvBatch = BuildBatchGrid();
            dgvBatch.Dock = DockStyle.Fill;
            splitBottom.Panel2.Controls.Add(dgvBatch);

            splitMain = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };
            splitMain.Panel1.Controls.Add(splitTop);
            splitMain.Panel2.Controls.Add(splitBottom);

            txtRaw = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10.5f),
                BackColor = DrawingColor.FromArgb(20, 20, 20),
                ForeColor = DrawingColor.LightGreen,
                ReadOnly = true,
                WordWrap = false
            };

            // TabControl con 2 pestañas: Combinado y Texto
            var tabText = new TabPage("📄 Texto OCR")
            {
                BackColor = DrawingColor.FromArgb(45, 45, 45)
            };
            tabText.Controls.Add(txtRaw);

            tabCombined.Controls.Add(splitMain);
            tabControl.TabPages.Add(tabCombined);
            tabControl.TabPages.Add(tabText);

            // ---- Botones y etiquetas inferiores ----
            btnCancelar = new Button
            {
                Text = "❌ Cancelar",
                Size = new Size(160, 40),
                BackColor = DrawingColor.FromArgb(180, 50, 50),
                ForeColor = DrawingColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.Cancel
            };
            btnCancelar.FlatAppearance.BorderSize = 0;
            btnCancelar.Click += delegate { this.DialogResult = DialogResult.Cancel; this.Close(); };

            btnConfirmar = new Button
            {
                Text = "✅ Confirmar",
                Size = new Size(160, 40),
                BackColor = DrawingColor.FromArgb(46, 125, 50),
                ForeColor = DrawingColor.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };
            btnConfirmar.FlatAppearance.BorderSize = 0;
            btnConfirmar.Click += BtnConfirmar_Click;

            lblCount = new Label
            {
                Text = "Filas detectadas (mediciones): 0",
                ForeColor = DrawingColor.LightGray,
                Font = new Font("Segoe UI", 9),
                AutoSize = true
            };
            lblTol = new Label
            {
                Text = "",
                ForeColor = DrawingColor.DarkGray,
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
                // splitMain: 55% arriba (mediciones+CMC) / 45% abajo (receta+batch)
                if (splitMain != null)
                {
                    int h = splitMain.ClientSize.Height;
                    int d = (int)(h * splitMainRatio);
                    splitMain.SplitterDistance = Math.Max(120, Math.Min(h - 120, d));
                }
                // splitTop: 50% mediciones / 50% CMC
                if (splitTop != null)
                {
                    int h = splitTop.ClientSize.Height;
                    int d = (int)(h * splitTopRatio);
                    splitTop.SplitterDistance = Math.Max(80, Math.Min(h - 80, d));
                }
                // splitBottom: 60% receta / 40% batch
                if (splitBottom != null)
                {
                    int h = splitBottom.ClientSize.Height;
                    int d = (int)(h * 0.60);
                    splitBottom.SplitterDistance = Math.Max(50, Math.Min(h - 50, d));
                }
            }
            catch { }
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
                BackgroundColor = DrawingColor.FromArgb(45, 45, 45),
                GridColor = DrawingColor.FromArgb(80, 80, 80),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells, // mide por contenido
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10.5f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 34
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = DrawingColor.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = DrawingColor.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.DefaultCellStyle.SelectionBackColor = DrawingColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = DrawingColor.FromArgb(45, 45, 45);

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
                BackgroundColor = DrawingColor.FromArgb(45, 45, 45),
                GridColor = DrawingColor.FromArgb(80, 80, 80),
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

            dgv.ColumnHeadersDefaultCellStyle.BackColor = DrawingColor.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgv.DefaultCellStyle.BackColor = DrawingColor.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.DefaultCellStyle.SelectionBackColor = DrawingColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = DrawingColor.FromArgb(45, 45, 45);

            // Columnas (cabeceras)
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Illuminant", HeaderText = "Iluminante" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaLightness", HeaderText = "ΔL (Lightness)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaChroma", HeaderText = "ΔC (Chroma)" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeltaHue", HeaderText = "ΔH (Hue)" });
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

        private DataGridView BuildRecipeGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = DrawingColor.FromArgb(45, 45, 45),
                GridColor = DrawingColor.FromArgb(80, 80, 80),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowTemplate = { Height = 24 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 30
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = DrawingColor.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.DefaultCellStyle.BackColor = DrawingColor.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.DefaultCellStyle.SelectionBackColor = DrawingColor.FromArgb(0, 90, 160);
            dgv.AlternatingRowsDefaultCellStyle.BackColor = DrawingColor.FromArgb(45, 45, 45);

            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Code", HeaderText = "Código" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Nombre" });
            dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = "Pct", HeaderText = "%" });

            dgv.Columns["Code"].FillWeight = 120f; dgv.Columns["Code"].MinimumWidth = 90;
            dgv.Columns["Name"].FillWeight = 500f; dgv.Columns["Name"].MinimumWidth = 200;
            dgv.Columns["Pct"].FillWeight = 100f; dgv.Columns["Pct"].MinimumWidth = 80;
            dgv.Columns["Pct"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            return dgv;
        }

        private DataGridView BuildBatchGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = DrawingColor.FromArgb(45, 45, 45),
                GridColor = DrawingColor.FromArgb(80, 80, 80),
                BorderStyle = BorderStyle.None,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowTemplate = { Height = 26 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Consolas", 10.5f),
                EnableHeadersVisualStyles = false,
                ColumnHeadersHeight = 30
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = DrawingColor.FromArgb(0, 120, 215);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.DefaultCellStyle.BackColor = DrawingColor.FromArgb(55, 55, 55);
            dgv.DefaultCellStyle.ForeColor = DrawingColor.White;
            dgv.DefaultCellStyle.SelectionBackColor = DrawingColor.FromArgb(0, 90, 160);
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            foreach (string col in new[] { "L", "A", "B", "dL", "dC", "dH", "dE", "P/F" })
                dgv.Columns.Add(new DataGridViewTextBoxColumn { Name = col, HeaderText = col });

            foreach (DataGridViewColumn c in dgv.Columns)
            {
                c.FillWeight = 100f;
                c.MinimumWidth = 55;
            }
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

            LoadRecipeSection();
            LoadBatchSection();

            txtRaw.Text = BuildTextView();
        }

        private void LoadRecipeSection()
        {
            if (dgvRecipe == null) return;
            dgvRecipe.Rows.Clear();
            if (_shadeResult == null || _shadeResult.Recipe == null) return;

            foreach (var item in _shadeResult.Recipe)
            {
                int idx = dgvRecipe.Rows.Add(item.Code, item.Name, item.Percentage);
                // Color por fila alternada ya configurado en AlternatingRowsDefaultCellStyle
                // Colorear % en verde/naranja según valor
                try
                {
                    string pct = (item.Percentage ?? "").Replace("%", "").Trim();
                    double pctVal;
                    if (double.TryParse(pct, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out pctVal))
                    {
                        dgvRecipe.Rows[idx].Cells["Pct"].Style.ForeColor =
                            pctVal > 1.0 ? DrawingColor.FromArgb(255, 200, 80)
                                         : DrawingColor.FromArgb(100, 200, 120);
                    }
                }
                catch { }
            }
        }

        private void LoadBatchSection()
        {
            if (dgvBatch == null) return;
            dgvBatch.Rows.Clear();

            // Preferir BatchMeasure; fallback a LabValues
            string l = "", a = "", b = "", dl = "", dc = "", dh = "", de = "", pf = "";

            if (_shadeResult != null && _shadeResult.Batch != null)
            {
                var bm = _shadeResult.Batch;
                l = bm.L ?? ""; a = bm.A ?? ""; b = bm.B ?? "";
                dl = bm.dL ?? ""; dc = bm.dC ?? ""; dh = bm.dH ?? ""; de = bm.dE ?? ""; pf = bm.PF ?? "";
            }
            else if (_shadeResult != null && _shadeResult.Lab != null)
            {
                var lv = _shadeResult.Lab;
                l = lv.L ?? ""; a = lv.A ?? ""; b = lv.B ?? "";
                dl = lv.dL ?? ""; dc = lv.da ?? ""; dh = lv.dB ?? ""; de = lv.cde ?? ""; pf = lv.PF ?? "";
            }

            if (l == "" && a == "" && b == "") return;

            int idx = dgvBatch.Rows.Add(l, a, b, dl, dc, dh, de, pf);
            // Color P/F
            DrawingColor pfColor = pf.Trim().ToUpper() == "P"
                ? DrawingColor.FromArgb(46, 125, 50)
                : pf.Trim().ToUpper() == "F"
                    ? DrawingColor.FromArgb(180, 50, 50)
                    : DrawingColor.FromArgb(55, 55, 55);
            dgvBatch.Rows[idx].Cells["P/F"].Style.BackColor = pfColor;
            dgvBatch.Rows[idx].Cells["P/F"].Style.ForeColor = DrawingColor.White;
            dgvBatch.Rows[idx].Cells["P/F"].Style.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        }

        private void LoadMeasuresSection()
        {
            dgvData.Rows.Clear();

            for (int i = 0; i < _rows.Count; i++)
            {
                ColorimetricRow r = _rows[i];
                int idx = dgvData.Rows.Add(r.Illuminant, r.Type, r.L, r.A, r.B, r.Chroma, r.Hue);

                DrawingColor rowColor = DrawingColor.FromArgb(55, 55, 55);
                if (r.Illuminant == "D65") rowColor = DrawingColor.FromArgb(40, 60, 100);
                else if (r.Illuminant == "TL84") rowColor = DrawingColor.FromArgb(40, 80, 60);
                else if (r.Illuminant == "A") rowColor = DrawingColor.FromArgb(80, 80, 40);
                else if (r.Illuminant == "CWF") rowColor = DrawingColor.FromArgb(80, 55, 40);

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

                DrawingColor rowColor = DrawingColor.FromArgb(55, 55, 55);
                if (r.Illuminant == "D65") rowColor = DrawingColor.FromArgb(40, 60, 100);
                else if (r.Illuminant == "TL84") rowColor = DrawingColor.FromArgb(40, 80, 60);
                else if (r.Illuminant == "A") rowColor = DrawingColor.FromArgb(80, 80, 40);

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
                sb.AppendLine("Tolerances:");
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