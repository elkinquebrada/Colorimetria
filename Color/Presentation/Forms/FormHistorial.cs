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
            this.TopMost = false;
            this.ShowInTaskbar = true;
            this.MinimizeBox = true;
            this.MaximizeBox = true;
            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.TopMost = false;

            // Mostrar elementos de control
            this.lblTitulo.Visible = true;
            this.btnBorrar.Visible = true;
            this.btnExportar.Visible = true;
            this.btnCerrar.Visible = true;
            
            this.btnCerrar.Text = " Regresar";
            this.btnCerrar.Click += (s, e) => this.Close();

            // Asegurar posicion en el panel (fuerza bruta para evitar errores)
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
            // Limpiar y recrear para asegurar la estructura exacta V4
            dgvHistorial.Columns.Clear();
            
            dgvHistorial.Columns.Add("colShadeName", "Shade Name");
            dgvHistorial.Columns.Add("colLotNo", "Lot No");
            dgvHistorial.Columns.Add("colFechaHora", "Fecha/Hora");
            dgvHistorial.Columns.Add("colDyeCode", "Dye Code");
            dgvHistorial.Columns.Add("colDyeName", "Dye Name");
            dgvHistorial.Columns.Add("colConcentration", "Concentracion %");
            dgvHistorial.Columns.Add("colProportion", "Proporcion %");

            // Receta 1
            dgvHistorial.Columns.Add("colR1Con", "R1 Conc.");
            dgvHistorial.Columns.Add("colR1Part", "R1 Propor.");
            dgvHistorial.Columns.Add("colR1Ajuste", "R1 Ajuste");

            // Receta 2
            dgvHistorial.Columns.Add("colR2Con", "R2 Conc.");
            dgvHistorial.Columns.Add("colR2Part", "R2 Propor.");
            dgvHistorial.Columns.Add("colR2Ajuste", "R2 Ajuste");

            // Receta 3
            dgvHistorial.Columns.Add("colR3Con", "R3 Conc.");
            dgvHistorial.Columns.Add("colR3Part", "R3 Propor.");
            dgvHistorial.Columns.Add("colR3Ajuste", "R3 Ajuste");

            dgvHistorial.Columns.Add("colStatus", "Status");

            dgvHistorial.Columns.Add("colIdDetalle", "IdDetalle");
            dgvHistorial.Columns["colIdDetalle"].Visible = false;

            // Estilos
            foreach (DataGridViewColumn col in dgvHistorial.Columns)
            {
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            dgvHistorial.Columns["colDyeName"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgvHistorial.Columns["colDyeName"].FillWeight = 130;
        }

        private void AjustarAnchoCabecerasAgrupadas()
        {
            if (dgvHistorial.Columns.Count < 5) return;

            try
            {
                // Grupo 1: Datos de Lote
                int x1 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colShadeName"].Index, true).X;
                int x2 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colFechaHora"].Index, true).Right;
                lblShadeHistoryHeader.Location = new Point(x1, 0);
                lblShadeHistoryHeader.Width = x2 - x1;
                lblShadeHistoryHeader.Text = "DATOS DEL LOTE / REPORTE";

                // Grupo 2: FormulaciÃ³n Original
                int x3 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colDyeCode"].Index, true).X;
                int x4 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colProportion"].Index, true).Right;
                lblCalculoRecetaHeader.Location = new Point(x3, 0);
                lblCalculoRecetaHeader.Width = x4 - x3;
                lblCalculoRecetaHeader.Text = "FORMULACION ORIGINAL";

                // Grupo 3: Recetas Correctivas (V4)
                if (!pnlGroupHeaders.Controls.ContainsKey("lblV4Header"))
                {
                    var lblV4 = new Label
                    {
                        Name = "lblV4Header",
                        BackColor = System.Drawing.Color.FromArgb(0, 80, 160),
                        ForeColor = System.Drawing.Color.White,
                        Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Text = "RECETAS CORRECTIVAS (LUMINOSIDAD / CROMA / TONO)",
                        BorderStyle = BorderStyle.FixedSingle,
                        Height = pnlGroupHeaders.Height
                    };
                    pnlGroupHeaders.Controls.Add(lblV4);
                }

                int x5 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colR1Con"].Index, true).X;
                int x6 = dgvHistorial.GetColumnDisplayRectangle(dgvHistorial.Columns["colR3Ajuste"].Index, true).Right;
                var lblH = pnlGroupHeaders.Controls["lblV4Header"];
                lblH.Location = new Point(x5, 0);
                lblH.Width = x6 - x5;
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

                // Mapeo flexible para CSV y SQL
                if (tabla.Columns.Contains("ShadeName")) fila.Cells["colShadeName"].Value = row["ShadeName"];
                if (tabla.Columns.Contains("LotNo")) fila.Cells["colLotNo"].Value = row["LotNo"];
                
                string colFecha = tabla.Columns.Contains("FechaRegistro") ? "FechaRegistro" : "FechaHora";
                if (tabla.Columns.Contains(colFecha)) fila.Cells["colFechaHora"].Value = row[colFecha];

                if (tabla.Columns.Contains("DyeCode")) fila.Cells["colDyeCode"].Value = row["DyeCode"];
                if (tabla.Columns.Contains("DyeName")) fila.Cells["colDyeName"].Value = row["DyeName"];

                string colConc = tabla.Columns.Contains("Concentration_Original") ? "Concentration_Original" : "ConcOriginal";
                if (tabla.Columns.Contains(colConc)) fila.Cells["colConcentration"].Value = row[colConc];

                if (tabla.Columns.Contains("Proportion_Original")) fila.Cells["colProportion"].Value = row["Proportion_Original"];

                // Recetas V4
                if (tabla.Columns.Contains("R1_Con")) fila.Cells["colR1Con"].Value = row["R1_Con"];
                if (tabla.Columns.Contains("R1_Part")) fila.Cells["colR1Part"].Value = row["R1_Part"];
                if (tabla.Columns.Contains("R1_Ajuste")) fila.Cells["colR1Ajuste"].Value = row["R1_Ajuste"];

                if (tabla.Columns.Contains("R2_Con")) fila.Cells["colR2Con"].Value = row["R2_Con"];
                if (tabla.Columns.Contains("R2_Part")) fila.Cells["colR2Part"].Value = row["R2_Part"];
                if (tabla.Columns.Contains("R2_Ajuste")) fila.Cells["colR2Ajuste"].Value = row["R2_Ajuste"];

                if (tabla.Columns.Contains("R3_Con")) fila.Cells["colR3Con"].Value = row["R3_Con"];
                if (tabla.Columns.Contains("R3_Part")) fila.Cells["colR3Part"].Value = row["R3_Part"];
                if (tabla.Columns.Contains("R3_Ajuste")) fila.Cells["colR3Ajuste"].Value = row["R3_Ajuste"];

                string colSt = tabla.Columns.Contains("Status_TL84") ? "Status_TL84" : "Status";
                if (tabla.Columns.Contains(colSt)) fila.Cells["colStatus"].Value = row[colSt];

                if (tabla.Columns.Contains("Id_Detalle")) fila.Cells["colIdDetalle"].Value = row["Id_Detalle"];
            }
            lblContador.Text = "Analisis encontrados: " + dgvHistorial.Rows.Count;
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
                    sb.AppendLine("<Style ss:ID=\"sHeader\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\" ss:WrapText=\"1\"/><Font ss:Bold=\"1\" " +
                        "ss:Color=\"#FFFFFF\" ss:Size=\"10\"/><Interior ss:Color=\"#1F3864\" ss:Pattern=\"Solid\"/></Style>");
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
                    MessageBox.Show("Exportacion completada exitosamente.", "Exito", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    List<int> idsSQL = new List<int>();
                    
                    foreach (DataGridViewRow row in dgvHistorial.SelectedRows) 
                    {
                        if (row.Cells["colIdDetalle"]?.Value != null && int.TryParse(row.Cells["colIdDetalle"].Value.ToString(), out int idDetalle))
                        {
                            idsSQL.Add(idDetalle);
                        }
                        dgvHistorial.Rows.Remove(row);
                    }
                    
                    if (idsSQL.Count > 0)
                    {
                        HistorialService.EliminarDetallesSQL(idsSQL);
                    }
                    
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ShadeName"); dt.Columns.Add("LotNo"); dt.Columns.Add("FechaHora"); 
                    dt.Columns.Add("DyeCode"); dt.Columns.Add("DyeName"); dt.Columns.Add("ConcOriginal");
                    dt.Columns.Add("PropOriginal"); 
                    dt.Columns.Add("R1_Con"); dt.Columns.Add("R1_Part"); dt.Columns.Add("R1_Ajuste");
                    dt.Columns.Add("R2_Con"); dt.Columns.Add("R2_Part"); dt.Columns.Add("R2_Ajuste");
                    dt.Columns.Add("R3_Con"); dt.Columns.Add("R3_Part"); dt.Columns.Add("R3_Ajuste");
                    dt.Columns.Add("Status");

                    foreach (DataGridViewRow r in dgvHistorial.Rows)
                    {
                        dt.Rows.Add(
                            r.Cells["colShadeName"].Value, r.Cells["colLotNo"].Value, r.Cells["colFechaHora"].Value,
                            r.Cells["colDyeCode"].Value, r.Cells["colDyeName"].Value, r.Cells["colConcentration"].Value,
                            r.Cells["colProportion"].Value,
                            r.Cells["colR1Con"].Value, r.Cells["colR1Part"].Value, r.Cells["colR1Ajuste"].Value,
                            r.Cells["colR2Con"].Value, r.Cells["colR2Part"].Value, r.Cells["colR2Ajuste"].Value,
                            r.Cells["colR3Con"].Value, r.Cells["colR3Part"].Value, r.Cells["colR3Ajuste"].Value,
                            r.Cells["colStatus"].Value
                        );
                    }
                    HistorialService.GuardarHistorialCompleto(dt);
                    lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
                }
            }
        }
    }
}
