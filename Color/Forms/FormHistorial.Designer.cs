namespace Color.Forms
{
    partial class FormHistorial
    {
        /// Variable del diseñador requerida.
        private System.ComponentModel.IContainer components = null;

        /// Limpiar los recursos que se estén usando.
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        private void InitializeComponent()
        {
            this.dgvHistorial = new System.Windows.Forms.DataGridView();
            this.pnlTitulo = new System.Windows.Forms.Panel();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.txtBuscar = new System.Windows.Forms.TextBox();
            this.lblBuscar = new System.Windows.Forms.Label();
            this.pnlPie = new System.Windows.Forms.Panel();
            this.lblContador = new System.Windows.Forms.Label();
            this.btnBorrar = new System.Windows.Forms.Button();
            this.btnExportar = new System.Windows.Forms.Button();
            this.btnCerrar = new System.Windows.Forms.Button();

            // Columnas del DataGridView
            this.colFechaHora = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colShadeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colIluminante = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colLightness = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colChroma = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiagL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCorrL = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiagA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCorrA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colDiagB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCorrB = new System.Windows.Forms.DataGridViewTextBoxColumn();

            ((System.ComponentModel.ISupportInitialize)(this.dgvHistorial)).BeginInit();
            this.pnlTitulo.SuspendLayout();
            this.pnlPie.SuspendLayout();
            this.SuspendLayout();

            // ──────────────────────────────────────────────────────────────
            // pnlTitulo  (cabecera azul — igual que FormResultados)
            // ──────────────────────────────────────────────────────────────
            this.pnlTitulo.BackColor = System.Drawing.Color.FromArgb(0, 102, 204);
            this.pnlTitulo.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTitulo.Height = 56;
            this.pnlTitulo.Controls.Add(this.lblTitulo);
            this.pnlTitulo.Controls.Add(this.lblBuscar);
            this.pnlTitulo.Controls.Add(this.txtBuscar);

            // lblTitulo
            this.lblTitulo.AutoSize = false;
            this.lblTitulo.Dock = System.Windows.Forms.DockStyle.None;
            this.lblTitulo.Location = new System.Drawing.Point(14, 13);
            this.lblTitulo.Size = new System.Drawing.Size(380, 28);
            this.lblTitulo.Text = "HISTORIAL DE ANALISIS";
            this.lblTitulo.Font = new System.Drawing.Font("Segoe UI Black", 14F,
                                           System.Drawing.FontStyle.Bold,
                                           System.Drawing.GraphicsUnit.Point, 0);
            this.lblTitulo.ForeColor = System.Drawing.Color.White;

            // lblBuscar
            this.lblBuscar.AutoSize = false;
            this.lblBuscar.Location = new System.Drawing.Point(500, 18);
            this.lblBuscar.Size = new System.Drawing.Size(60, 22);
            this.lblBuscar.Text = "Buscar:";
            this.lblBuscar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblBuscar.ForeColor = System.Drawing.Color.White;
            this.lblBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;

            // txtBuscar
            this.txtBuscar.Location = new System.Drawing.Point(568, 16);
            this.txtBuscar.Size = new System.Drawing.Size(220, 24);
            this.txtBuscar.BackColor = System.Drawing.Color.White;
            this.txtBuscar.ForeColor = System.Drawing.Color.FromArgb(30, 30, 70);
            this.txtBuscar.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtBuscar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);

            // ──────────────────────────────────────────────────────────────
            // dgvHistorial  (tema claro — igual que FormResultados)
            // ──────────────────────────────────────────────────────────────
            this.dgvHistorial.AllowUserToAddRows = false;
            this.dgvHistorial.AllowUserToDeleteRows = false;
            this.dgvHistorial.AllowUserToResizeRows = false;
            this.dgvHistorial.ReadOnly = true;
            this.dgvHistorial.AutoSizeColumnsMode =
                System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvHistorial.SelectionMode =
                System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvHistorial.MultiSelect = false;
            this.dgvHistorial.RowHeadersVisible = false;
            this.dgvHistorial.BorderStyle =
                System.Windows.Forms.BorderStyle.None;
            this.dgvHistorial.EnableHeadersVisualStyles = false;
            this.dgvHistorial.ScrollBars =
                System.Windows.Forms.ScrollBars.Both;

            // Fondo general de la grilla — blanco
            this.dgvHistorial.BackgroundColor = System.Drawing.Color.White;
            this.dgvHistorial.GridColor = System.Drawing.Color.FromArgb(200, 212, 228);

            // Encabezados — azul oscuro #1F3864 igual que el sHeader del Excel
            System.Windows.Forms.DataGridViewCellStyle estiloHeader =
                new System.Windows.Forms.DataGridViewCellStyle();
            estiloHeader.BackColor = System.Drawing.Color.FromArgb(31, 56, 100);
            estiloHeader.ForeColor = System.Drawing.Color.White;
            estiloHeader.Font = new System.Drawing.Font("Segoe UI", 8.5F,
                                          System.Drawing.FontStyle.Bold);
            estiloHeader.Alignment =
                System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            estiloHeader.WrapMode =
                System.Windows.Forms.DataGridViewTriState.False;
            this.dgvHistorial.ColumnHeadersDefaultCellStyle = estiloHeader;
            this.dgvHistorial.ColumnHeadersHeight = 30;
            this.dgvHistorial.ColumnHeadersHeightSizeMode =
                System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // Celdas — fondo blanco, texto negro
            System.Windows.Forms.DataGridViewCellStyle estiloCelda =
                new System.Windows.Forms.DataGridViewCellStyle();
            estiloCelda.BackColor = System.Drawing.Color.White;
            estiloCelda.ForeColor = System.Drawing.Color.FromArgb(30, 30, 60);
            estiloCelda.SelectionBackColor = System.Drawing.Color.FromArgb(46, 117, 182);
            estiloCelda.SelectionForeColor = System.Drawing.Color.White;
            estiloCelda.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dgvHistorial.DefaultCellStyle = estiloCelda;

            // Filas alternas — azul claro #DCE6F1 (idéntico al sAlt del Excel)
            System.Windows.Forms.DataGridViewCellStyle estiloAlt =
                new System.Windows.Forms.DataGridViewCellStyle();
            estiloAlt.BackColor = System.Drawing.Color.FromArgb(220, 230, 241);
            this.dgvHistorial.AlternatingRowsDefaultCellStyle = estiloAlt;

            this.dgvHistorial.RowTemplate.Height = 24;

            // Columnas
            //  1. Fecha y Hora
            this.colFechaHora.Name = "colFechaHora";
            this.colFechaHora.HeaderText = "Fecha y Hora";
            this.colFechaHora.ToolTipText = "Fecha y hora de la medición";

            //  2. Shade Name
            this.colShadeName.Name = "colShadeName";
            this.colShadeName.HeaderText = "Shade Name";
            this.colShadeName.ToolTipText = "Nombre del color / receta estándar";

            //  3. Iluminante
            this.colIluminante.Name = "colIluminante";
            this.colIluminante.HeaderText = "Iluminante";

            //  4. Lightness
            this.colLightness.Name = "colLightness";
            this.colLightness.HeaderText = "Lightness (%)";

            //  5. Chroma
            this.colChroma.Name = "colChroma";
            this.colChroma.HeaderText = "Chroma (%)";

            //  6. DiagL
            this.colDiagL.Name = "colDiagL";
            this.colDiagL.HeaderText = "Diagnóstico L*";

            //  7. CorrL
            this.colCorrL.Name = "colCorrL";
            this.colCorrL.HeaderText = "Corrección L*";

            //  8. DiagA
            this.colDiagA.Name = "colDiagA";
            this.colDiagA.HeaderText = "Diagnóstico a*";

            //  9. CorrA
            this.colCorrA.Name = "colCorrA";
            this.colCorrA.HeaderText = "Corrección a*";

            //  10. DiagB
            this.colDiagB.Name = "colDiagB";
            this.colDiagB.HeaderText = "Diagnóstico b*";

            //  11. CorrB
            this.colCorrB.Name = "colCorrB";
            this.colCorrB.HeaderText = "Corrección b*";

            this.dgvHistorial.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[]
            {
                this.colFechaHora,
                this.colShadeName,
                this.colIluminante,
                this.colLightness,
                this.colChroma,
                this.colDiagL,
                this.colCorrL,
                this.colDiagA,
                this.colCorrA,
                this.colDiagB,
                this.colCorrB
            });

            this.dgvHistorial.Dock = System.Windows.Forms.DockStyle.Fill;

            // ──────────────────────────────────────────────────────────────
            // pnlPie  (barra inferior blanca — igual que FormResultados)
            // ──────────────────────────────────────────────────────────────
            this.pnlPie.BackColor = System.Drawing.Color.White;
            this.pnlPie.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlPie.Height = 54;
            this.pnlPie.Controls.Add(this.lblContador);
            this.pnlPie.Controls.Add(this.btnBorrar);
            this.pnlPie.Controls.Add(this.btnExportar);
            this.pnlPie.Controls.Add(this.btnCerrar);

            // lblContador
            this.lblContador.AutoSize = false;
            this.lblContador.Location = new System.Drawing.Point(14, 16);
            this.lblContador.Size = new System.Drawing.Size(260, 22);
            this.lblContador.Text = "Total de registros: 0";
            this.lblContador.Font = new System.Drawing.Font("Segoe UI", 8.5F);
            this.lblContador.ForeColor = System.Drawing.Color.FromArgb(60, 64, 70);

            // btnBorrar — Rojo igual al btnCerrar de FormResultados
            this.btnBorrar.Size = new System.Drawing.Size(130, 34);
            this.btnBorrar.Location = new System.Drawing.Point(530, 10);
            this.btnBorrar.Text = "Borrar";
            this.btnBorrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBorrar.FlatAppearance.BorderSize = 0;
            this.btnBorrar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(170, 30, 45);
            this.btnBorrar.BackColor = System.Drawing.Color.FromArgb(200, 30, 30);
            this.btnBorrar.ForeColor = System.Drawing.Color.White;
            this.btnBorrar.Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold);
            this.btnBorrar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBorrar.Click += new System.EventHandler(this.btnBorrar_Click);

            // btnExportar — Azul igual al btnHistorial de FormResultados
            this.btnExportar.Size = new System.Drawing.Size(140, 34);
            this.btnExportar.Location = new System.Drawing.Point(670, 10);
            this.btnExportar.Text = "⬇ Exportar CSV";
            this.btnExportar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExportar.FlatAppearance.BorderSize = 0;
            this.btnExportar.FlatAppearance.MouseOverBackColor =
                System.Drawing.Color.FromArgb(30, 100, 200);
            this.btnExportar.BackColor = System.Drawing.Color.FromArgb(45, 126, 247);
            this.btnExportar.ForeColor = System.Drawing.Color.White;
            this.btnExportar.Font = new System.Drawing.Font("Segoe UI", 9.5F,
                                             System.Drawing.FontStyle.Bold);
            this.btnExportar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);

            // btnCerrar — Rojo oscuro igual que FormResultados
            this.btnCerrar.Size = new System.Drawing.Size(100, 34);
            this.btnCerrar.Location = new System.Drawing.Point(820, 10);
            this.btnCerrar.Text = "Cerrar";
            this.btnCerrar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCerrar.FlatAppearance.BorderSize = 0;
            this.btnCerrar.FlatAppearance.MouseOverBackColor =
                System.Drawing.Color.FromArgb(170, 30, 45);
            this.btnCerrar.BackColor = System.Drawing.Color.FromArgb(200, 30, 30);
            this.btnCerrar.ForeColor = System.Drawing.Color.White;
            this.btnCerrar.Font = new System.Drawing.Font("Segoe UI", 9.5F,
                                           System.Drawing.FontStyle.Bold);
            this.btnCerrar.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCerrar.Click += new System.EventHandler(this.btnCerrar_Click);

            // ──────────────────────────────────────────────────────────────
            // FormHistorial  (formulario principal)
            // ──────────────────────────────────────────────────────────────
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(960, 580);
            this.MinimumSize = new System.Drawing.Size(760, 420);
            this.Name = "FormHistorial";
            this.Text = " TINT COATS CADENA";
            this.StartPosition =
                System.Windows.Forms.FormStartPosition.CenterParent;
            this.Font = new System.Drawing.Font("Segoe UI", 9F,
                                           System.Drawing.FontStyle.Regular,
                                           System.Drawing.GraphicsUnit.Point, 0);

            // Orden de capas: pnlTitulo (top) → dgvHistorial (fill) → pnlPie (bottom)
            this.Controls.Add(this.dgvHistorial);
            this.Controls.Add(this.pnlPie);
            this.Controls.Add(this.pnlTitulo);

            ((System.ComponentModel.ISupportInitialize)(this.dgvHistorial)).EndInit();
            this.pnlTitulo.ResumeLayout(false);
            this.pnlPie.ResumeLayout(false);
            this.ResumeLayout(false);
        }

        #endregion

        // ── Controles declarados ──────────────────────────────────────────
        private System.Windows.Forms.DataGridView dgvHistorial;
        private System.Windows.Forms.Panel pnlTitulo;
        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Label lblBuscar;
        private System.Windows.Forms.TextBox txtBuscar;
        private System.Windows.Forms.Panel pnlPie;
        private System.Windows.Forms.Label lblContador;
        private System.Windows.Forms.Button btnBorrar;
        private System.Windows.Forms.Button btnExportar;
        private System.Windows.Forms.Button btnCerrar;

        // Columnas
        private System.Windows.Forms.DataGridViewTextBoxColumn colFechaHora;
        private System.Windows.Forms.DataGridViewTextBoxColumn colShadeName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colIluminante;
        private System.Windows.Forms.DataGridViewTextBoxColumn colLightness;
        private System.Windows.Forms.DataGridViewTextBoxColumn colChroma;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiagL;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCorrL;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiagA;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCorrA;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDiagB;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCorrB;
    }
}