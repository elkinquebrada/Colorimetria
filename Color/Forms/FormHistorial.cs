using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using Color.Services;

namespace Color
{
    public partial class FormHistorial : Form
    {
        public FormHistorial()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.TopMost = false;

            // Mostrar elementos de control
            this.lblTitulo.Visible = true;
            this.btnBorrar.Visible = true;
            this.btnExportar.Visible = true;
            this.btnCerrar.Visible = true;
            
            this.btnCerrar.Text = "← Regresar";
            this.btnCerrar.Click += (s, e) => this.Close();

            // Asegurar posición en el panel (fuerza bruta para evitar errores del diseñador)
            this.btnCerrar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnExportar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.btnBorrar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            
            this.btnCerrar.Location = new Point(this.pnlPie.Width - this.btnCerrar.Width - 10, 10);
            this.btnExportar.Location = new Point(this.btnCerrar.Left - this.btnExportar.Width - 10, 10);
            this.btnBorrar.Location = new Point(this.btnExportar.Left - this.btnBorrar.Width - 10, 10);

            ConfigurarColumnas();
            this.Resize += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.dgvHistorial.ColumnWidthChanged += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.dgvHistorial.Scroll += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.Load += (s, e) => AjustarAnchoCabecerasAgrupadas();
            AddBrandingLogo();
        }

        private void AddBrandingLogo()
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string[] paths = {
                    Path.Combine(baseDir, "logicDocs", "Coats_logo.svg.png"),
                    Path.Combine(baseDir, "..", "..", "logicDocs", "Coats_logo.svg.png")
                };

                string finalPath = paths.FirstOrDefault(p => File.Exists(p));
                if (string.IsNullOrEmpty(finalPath)) return;

                var logo = new PictureBox
                {
                    Image = Image.FromFile(finalPath),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Width = 50,
                    Height = 50,
                    Anchor = AnchorStyles.Top | AnchorStyles.Right,
                    BackColor = System.Drawing.Color.Transparent
                };
                logo.Location = new Point(this.pnlTitulo.Width - logo.Width - 15, 3);
                this.pnlTitulo.Controls.Add(logo);
                logo.BringToFront();
            }
            catch { }
        }

        private void ConfigurarColumnas()
        {
            if (dgvHistorial.Columns.Count == 0) return;

            // Ajuste de pesos para las 13 columnas (Estructura Unificada)
            dgvHistorial.Columns["colShadeName"].FillWeight = 10;
            dgvHistorial.Columns["colFechaHora"].FillWeight = 10;
            dgvHistorial.Columns["colIluminante"].FillWeight = 5;
            dgvHistorial.Columns["colDyeName"].FillWeight = 15;
            dgvHistorial.Columns["colConcentration"].FillWeight = 8;
            dgvHistorial.Columns["colReceta1"].FillWeight = 11;
            dgvHistorial.Columns["colReceta2"].FillWeight = 11;
            dgvHistorial.Columns["colReceta3"].FillWeight = 11;
            
            // --- Nuevas Columnas de Ingeniería (Panel Izquierdo) ---
            if (!dgvHistorial.Columns.Contains("colImpactodl"))
            {
                dgvHistorial.Columns.Add("colImpactodl", "Impacto dl");
                dgvHistorial.Columns.Add("colDiagdl", "Diagnóstico dl");
                dgvHistorial.Columns.Add("colRecdl", "Recomendación dl");
                
                dgvHistorial.Columns.Add("colImpactoda", "Impacto da");
                dgvHistorial.Columns.Add("colDiagda", "Diagnóstico da");
                dgvHistorial.Columns.Add("colRecda", "Recomendación da");
                
                dgvHistorial.Columns.Add("colImpactodb", "Impacto db");
                dgvHistorial.Columns.Add("colDiagdb", "Diagnóstico db");
                dgvHistorial.Columns.Add("colRecdb", "Recomendación db");


                // Estilo suave para las columnas de texto largo
                foreach (string col in new[] { "colDiagdl", "colRecdl", "colDiagda", "colRecda", "colDiagdb", "colRecdb" })
                {
                    dgvHistorial.Columns[col].FillWeight = 20;
                    dgvHistorial.Columns[col].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }

            DataGridViewCellStyle center = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter };

            dgvHistorial.Columns["colIluminante"].DefaultCellStyle = center;
            dgvHistorial.Columns["colConcentration"].DefaultCellStyle = center;
            dgvHistorial.Columns["colConcentration"].DefaultCellStyle = center;
        }

        private void AjustarAnchoCabecerasAgrupadas()
        {
            if (dgvHistorial.Columns.Count < 9) return;

            try
            {
                // El grupo 1 abarca desde colShadeName (0) hasta colIluminante (2)
                int x1 = dgvHistorial.GetColumnDisplayRectangle(0, true).X;
                int x2 = dgvHistorial.GetColumnDisplayRectangle(2, true).X + dgvHistorial.GetColumnDisplayRectangle(2, true).Width;
                
                lblShadeHistoryHeader.Location = new Point(x1, 0);
                lblShadeHistoryHeader.Width = x2 - x1;
                lblShadeHistoryHeader.Text = "HISTORIAL DE ANALISIS (CABECERA)";

                // El grupo 2 abarca desde colDyeName (3) hasta colReceta3 (7)
                int x3 = dgvHistorial.GetColumnDisplayRectangle(3, true).X;
                int x4 = dgvHistorial.GetColumnDisplayRectangle(7, true).X + dgvHistorial.GetColumnDisplayRectangle(7, true).Width;
 
                lblCalculoRecetaHeader.Location = new Point(x3, 0);
                lblCalculoRecetaHeader.Width = x4 - x3;
                lblCalculoRecetaHeader.Text = "FORMULACIÓN Y CONCENTRACIONES";

                // El grupo 3 (NUEVO): Diagnóstico Experto
                if (dgvHistorial.Columns.Count > 8)
                {
                    int x5 = dgvHistorial.GetColumnDisplayRectangle(8, true).X;
                    int x6 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns.Count - 1, true).X + 
                             dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns.Count - 1, true).Width;

                    if (!pnlGroupHeaders.Controls.ContainsKey("lblExpertHeader"))
                    {
                        var lblExpert = new Label
                        {
                            Name = "lblExpertHeader",
                            BackColor = System.Drawing.Color.FromArgb(0, 80, 160),
                            ForeColor = System.Drawing.Color.White,
                            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                            TextAlign = ContentAlignment.MiddleCenter,
                            Text = "DIAGNÓSTICO EXPERTO (INGENIERÍA)",
                            BorderStyle = BorderStyle.FixedSingle,
                            Height = pnlGroupHeaders.Height
                        };
                        pnlGroupHeaders.Controls.Add(lblExpert);
                    }
                    
                    var lblE = pnlGroupHeaders.Controls["lblExpertHeader"];
                    lblE.Location = new Point(x5, 0);
                    lblE.Width = x6 - x5;
                    lblE.Visible = x6 > x5;
                }
                
                lblShadeHistoryHeader.Height = pnlGroupHeaders.Height;
                lblCalculoRecetaHeader.Height = pnlGroupHeaders.Height;
            }
            catch { }
        }

        public void CargarHistorial(DataTable tabla)
        {
            dgvHistorial.Rows.Clear();
            foreach (DataRow row in tabla.Rows)
            {
                int rowIndex = dgvHistorial.Rows.Add();
                var fila = dgvHistorial.Rows[rowIndex];

                if (tabla.Columns.Contains("ShadeName")) fila.Cells["colShadeName"].Value = row["ShadeName"];
                if (tabla.Columns.Contains("FechaHora")) fila.Cells["colFechaHora"].Value = row["FechaHora"];
                if (tabla.Columns.Contains("Iluminante")) fila.Cells["colIluminante"].Value = row["Iluminante"];
                if (tabla.Columns.Contains("DyeName")) fila.Cells["colDyeName"].Value = row["DyeName"];
                if (tabla.Columns.Contains("ConcOriginal")) fila.Cells["colConcentration"].Value = row["ConcOriginal"];
                if (tabla.Columns.Contains("Receta1")) fila.Cells["colReceta1"].Value = row["Receta1"];
                if (tabla.Columns.Contains("Receta2")) fila.Cells["colReceta2"].Value = row["Receta2"];
                if (tabla.Columns.Contains("Receta3")) fila.Cells["colReceta3"].Value = row["Receta3"];

                // Datos de Ingeniería
                if (tabla.Columns.Contains("Impactodl")) fila.Cells["colImpactodl"].Value = row["Impactodl"];
                if (tabla.Columns.Contains("Diagdl")) fila.Cells["colDiagdl"].Value = row["Diagdl"];
                if (tabla.Columns.Contains("Recdl")) fila.Cells["colRecdl"].Value = row["Recdl"];
                
                if (tabla.Columns.Contains("Impactoda")) fila.Cells["colImpactoda"].Value = row["Impactoda"];
                if (tabla.Columns.Contains("Diagda")) fila.Cells["colDiagda"].Value = row["Diagda"];
                if (tabla.Columns.Contains("Recda")) fila.Cells["colRecda"].Value = row["Recda"];
                
                if (tabla.Columns.Contains("Impactodb")) fila.Cells["colImpactodb"].Value = row["Impactodb"];
                if (tabla.Columns.Contains("Diagdb")) fila.Cells["colDiagdb"].Value = row["Diagdb"];
                if (tabla.Columns.Contains("Recdb")) fila.Cells["colRecdb"].Value = row["Recdb"];

            }
            lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
        }

        private void txtBuscar_TextChanged(object sender, EventArgs e)
        {
            string filtro = txtBuscar.Text.ToLower();
            foreach (DataGridViewRow row in dgvHistorial.Rows)
            {
                bool visible =
                    row.Cells["colShadeName"].Value?.ToString().ToLower().Contains(filtro) == true ||
                    row.Cells["colDyeName"].Value?.ToString().ToLower().Contains(filtro) == true;

                row.Visible = visible;
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "Excel (*.xls)|*.xls";
                dlg.FileName = "HistorialConsolidado_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xls";
                if (dlg.ShowDialog() != DialogResult.OK) return;

                try
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    sb.AppendLine("<?mso-application progid=\"Excel.Sheet\"?>");
                    sb.AppendLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
                    sb.AppendLine("<Styles>");
                    sb.AppendLine("<Style ss:ID=\"sHeader\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/><Font ss:Bold=\"1\" ss:Color=\"#FFFFFF\" ss:Size=\"10\"/><Interior ss:Color=\"#1F3864\" ss:Pattern=\"Solid\"/></Style>");
                    sb.AppendLine("<Style ss:ID=\"sRow\"><Font ss:Size=\"10\"/></Style>");
                    sb.AppendLine("</Styles>");
                    sb.AppendLine("<Worksheet ss:Name=\"Historial Consolidado\">");
                    sb.AppendLine("<Table>");
                    sb.AppendLine("<Row>");
                    foreach (DataGridViewColumn col in dgvHistorial.Columns)
                        sb.AppendLine($"<Cell ss:StyleID=\"sHeader\"><Data ss:Type=\"String\">{System.Security.SecurityElement.Escape(col.HeaderText)}</Data></Cell>");
                    sb.AppendLine("</Row>");
                    foreach (DataGridViewRow row in dgvHistorial.Rows)
                    {
                        if (!row.Visible || row.IsNewRow) continue;
                        sb.AppendLine("<Row>");
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            string val = System.Security.SecurityElement.Escape((cell.Value ?? "").ToString());
                            sb.AppendLine($"<Cell><Data ss:Type=\"String\">{val}</Data></Cell>");
                        }
                        sb.AppendLine("</Row>");
                    }
                    sb.AppendLine("</Table></Worksheet></Workbook>");
                    System.IO.File.WriteAllText(dlg.FileName, sb.ToString(), new System.Text.UTF8Encoding(true));
                    MessageBox.Show("Exportación completada exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex) { MessageBox.Show("Error al exportar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e) => Close();

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            if (dgvHistorial.SelectedRows.Count > 0)
            {
                if (MessageBox.Show("¿Borrar registro seleccionado?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dgvHistorial.SelectedRows) dgvHistorial.Rows.Remove(row);
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ShadeName"); dt.Columns.Add("FechaHora"); dt.Columns.Add("Iluminante");
                    dt.Columns.Add("DyeName"); dt.Columns.Add("ConcOriginal");
                    dt.Columns.Add("Receta1"); dt.Columns.Add("Receta2"); dt.Columns.Add("Receta3");
                    dt.Columns.Add("Impactodl"); dt.Columns.Add("Diagdl"); dt.Columns.Add("Recdl");
                    dt.Columns.Add("Impactoda"); dt.Columns.Add("Diagda"); dt.Columns.Add("Recda");
                    dt.Columns.Add("Impactodb"); dt.Columns.Add("Diagdb"); dt.Columns.Add("Recdb");

                    foreach (DataGridViewRow r in dgvHistorial.Rows)
                    {
                        dt.Rows.Add(
                            r.Cells["colShadeName"].Value, r.Cells["colFechaHora"].Value, r.Cells["colIluminante"].Value,
                            r.Cells["colDyeName"].Value, r.Cells["colConcentration"].Value,
                            r.Cells["colReceta1"].Value, r.Cells["colReceta2"].Value, r.Cells["colReceta3"].Value,
                            r.Cells["colImpactodl"].Value, r.Cells["colDiagdl"].Value, r.Cells["colRecdl"].Value,
                            r.Cells["colImpactoda"].Value, r.Cells["colDiagda"].Value, r.Cells["colRecda"].Value,
                            r.Cells["colImpactodb"].Value, r.Cells["colDiagdb"].Value, r.Cells["colRecdb"].Value
                        );
                    }
                    HistorialService.GuardarHistorialCompleto(dt);
                    lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
                }
            }
        }
    }
}
