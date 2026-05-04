using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Color.Services;

namespace Color.Forms
{
    public partial class FormHistorial : Form
    {
        public FormHistorial()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
            ConfigurarColumnas();
            this.Resize += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.dgvHistorial.ColumnWidthChanged += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.dgvHistorial.Scroll += (s, e) => AjustarAnchoCabecerasAgrupadas();
            this.Load += (s, e) => AjustarAnchoCabecerasAgrupadas();
        }

        private void ConfigurarColumnas()
        {
            if (dgvHistorial.Columns.Count == 0) return;

            // Ajuste de pesos para las 12 columnas (Estructura Unificada)
            dgvHistorial.Columns["colShadeName"].FillWeight = 10;
            dgvHistorial.Columns["colFechaHora"].FillWeight = 10;
            dgvHistorial.Columns["colIluminante"].FillWeight = 6;
            dgvHistorial.Columns["colDLEje"].FillWeight = 8;
            dgvHistorial.Columns["colDCEje"].FillWeight = 8;
            dgvHistorial.Columns["colDHEje"].FillWeight = 8;
            dgvHistorial.Columns["colDyeCode"].FillWeight = 8;
            dgvHistorial.Columns["colDyeName"].FillWeight = 12;
            dgvHistorial.Columns["colConcentration"].FillWeight = 8;
            dgvHistorial.Columns["colAjusteDL"].FillWeight = 8;
            dgvHistorial.Columns["colAjusteDH"].FillWeight = 8;
            dgvHistorial.Columns["colNuevaReceta"].FillWeight = 10;

            DataGridViewCellStyle center = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter };

            dgvHistorial.Columns["colIluminante"].DefaultCellStyle = center;
            dgvHistorial.Columns["colDLEje"].DefaultCellStyle = center;
            dgvHistorial.Columns["colDCEje"].DefaultCellStyle = center;
            dgvHistorial.Columns["colDHEje"].DefaultCellStyle = center;
            dgvHistorial.Columns["colDyeCode"].DefaultCellStyle = center;
            dgvHistorial.Columns["colConcentration"].DefaultCellStyle = center;
        }

        private void AjustarAnchoCabecerasAgrupadas()
        {
            if (dgvHistorial.Columns.Count < 12) return;

            try
            {
                // El grupo 1 abarca desde colShadeName (0) hasta colDHEje (5)
                int x1 = dgvHistorial.GetColumnDisplayRectangle(0, true).X;
                int x2 = dgvHistorial.GetColumnDisplayRectangle(5, true).X + dgvHistorial.GetColumnDisplayRectangle(5, true).Width;
                
                lblShadeHistoryHeader.Location = new Point(x1, 0);
                lblShadeHistoryHeader.Width = x2 - x1;
                lblShadeHistoryHeader.Text = "HISTORIAL DE ANALISIS (CABECERA)";

                // El grupo 2 abarca desde colDyeCode (6) hasta colNuevaReceta (11)
                int x3 = dgvHistorial.GetColumnDisplayRectangle(6, true).X;
                int x4 = dgvHistorial.GetColumnDisplayRectangle(11, true).X + dgvHistorial.GetColumnDisplayRectangle(11, true).Width;

                lblCalculoRecetaHeader.Location = new Point(x3, 0);
                lblCalculoRecetaHeader.Width = x4 - x3;
                lblCalculoRecetaHeader.Text = "FORMULACIÓN Y CONCENTRACIONES";
                
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
                if (tabla.Columns.Contains("DLEje")) fila.Cells["colDLEje"].Value = row["DLEje"];
                if (tabla.Columns.Contains("DCEje")) fila.Cells["colDCEje"].Value = row["DCEje"];
                if (tabla.Columns.Contains("DHEje")) fila.Cells["colDHEje"].Value = row["DHEje"];
                if (tabla.Columns.Contains("DyeCode")) fila.Cells["colDyeCode"].Value = row["DyeCode"];
                if (tabla.Columns.Contains("DyeName")) fila.Cells["colDyeName"].Value = row["DyeName"];
                if (tabla.Columns.Contains("ConcOriginal")) fila.Cells["colConcentration"].Value = row["ConcOriginal"];
                if (tabla.Columns.Contains("AjusteDL")) fila.Cells["colAjusteDL"].Value = row["AjusteDL"];
                if (tabla.Columns.Contains("AjusteDH")) fila.Cells["colAjusteDH"].Value = row["AjusteDH"];
                if (tabla.Columns.Contains("NuevaReceta")) fila.Cells["colNuevaReceta"].Value = row["NuevaReceta"];
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
                    row.Cells["colDyeName"].Value?.ToString().ToLower().Contains(filtro) == true ||
                    row.Cells["colDyeCode"].Value?.ToString().ToLower().Contains(filtro) == true;

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
                    dt.Columns.Add("DLEje"); dt.Columns.Add("DCEje"); dt.Columns.Add("DHEje");
                    dt.Columns.Add("DyeCode"); dt.Columns.Add("DyeName"); dt.Columns.Add("ConcOriginal");
                    dt.Columns.Add("AjusteDL"); dt.Columns.Add("AjusteDH"); dt.Columns.Add("NuevaReceta");
                    foreach (DataGridViewRow r in dgvHistorial.Rows)
                    {
                        dt.Rows.Add(r.Cells["colShadeName"].Value, r.Cells["colFechaHora"].Value, r.Cells["colIluminante"].Value,
                                    r.Cells["colDLEje"].Value, r.Cells["colDCEje"].Value, r.Cells["colDHEje"].Value,
                                    r.Cells["colDyeCode"].Value, r.Cells["colDyeName"].Value, r.Cells["colConcentration"].Value,
                                    r.Cells["colAjusteDL"].Value, r.Cells["colAjusteDH"].Value, r.Cells["colNuevaReceta"].Value);
                    }
                    HistorialService.GuardarHistorialCompleto(dt);
                    lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
                }
            }
        }
    }
}
