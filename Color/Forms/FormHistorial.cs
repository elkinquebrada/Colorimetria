using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Color.Forms
{
    public partial class FormHistorial : Form
    {
        public FormHistorial()
        {
            InitializeComponent();
            ConfigurarColumnas();
        }

        // ─────────────────────────────────────────────
        // Configuración inicial del DataGridView
        // ─────────────────────────────────────────────
        private void dgvHistorial_ColumnHeaderMouseClick(
    object sender,
    DataGridViewCellMouseEventArgs e)
        {
            dgvHistorial.Sort(
                dgvHistorial.Columns[e.ColumnIndex],
                ListSortDirection.Ascending);
        }


        private void ConfigurarColumnas()
        {
            if (dgvHistorial.Columns.Count == 0)
                return;

            dgvHistorial.Columns["colFechaHora"].FillWeight = 12;
            dgvHistorial.Columns["colShadeName"].FillWeight = 12;
            dgvHistorial.Columns["colIluminante"].FillWeight = 8;
            dgvHistorial.Columns["colLightness"].FillWeight = 8;
            dgvHistorial.Columns["colChroma"].FillWeight = 8;
            dgvHistorial.Columns["colDiagL"].FillWeight = 10;
            dgvHistorial.Columns["colCorrL"].FillWeight = 12;
            dgvHistorial.Columns["colDiagA"].FillWeight = 10;
            dgvHistorial.Columns["colCorrA"].FillWeight = 12;
            dgvHistorial.Columns["colDiagB"].FillWeight = 10;
            dgvHistorial.Columns["colCorrB"].FillWeight = 12;

            DataGridViewCellStyle center = new DataGridViewCellStyle();
            center.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dgvHistorial.Columns["colIluminante"].DefaultCellStyle = center;
            dgvHistorial.Columns["colLightness"].DefaultCellStyle = center;
            dgvHistorial.Columns["colChroma"].DefaultCellStyle = center;
        }

        // ─────────────────────────────────────────────
        // Cargar historial (se llamará más adelante)
        // ─────────────────────────────────────────────
        public void CargarHistorial(DataTable tabla)
        {
            dgvHistorial.Rows.Clear();

            if (tabla == null)
            {
                lblContador.Text = "Total de registros: 0";
                return;
            }

            foreach (DataRow row in tabla.Rows)
            {
                int idx = dgvHistorial.Rows.Add();
                DataGridViewRow fila = dgvHistorial.Rows[idx];

                if (tabla.Columns.Contains("FechaHora")) fila.Cells["colFechaHora"].Value = row["FechaHora"];
                if (tabla.Columns.Contains("ShadeName")) fila.Cells["colShadeName"].Value = row["ShadeName"];
                if (tabla.Columns.Contains("Iluminante")) fila.Cells["colIluminante"].Value = row["Iluminante"];
                if (tabla.Columns.Contains("Lightness")) fila.Cells["colLightness"].Value = row["Lightness"];
                if (tabla.Columns.Contains("Chroma")) fila.Cells["colChroma"].Value = row["Chroma"];
                if (tabla.Columns.Contains("Diagnostico_L")) fila.Cells["colDiagL"].Value = row["Diagnostico_L"];
                if (tabla.Columns.Contains("Correccion_L")) fila.Cells["colCorrL"].Value = row["Correccion_L"];
                if (tabla.Columns.Contains("Diagnostico_a")) fila.Cells["colDiagA"].Value = row["Diagnostico_a"];
                if (tabla.Columns.Contains("Correccion_a")) fila.Cells["colCorrA"].Value = row["Correccion_a"];
                if (tabla.Columns.Contains("Diagnostico_b")) fila.Cells["colDiagB"].Value = row["Diagnostico_b"];
                if (tabla.Columns.Contains("Correccion_b")) fila.Cells["colCorrB"].Value = row["Correccion_b"];
            }

            lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
        }

        // ─────────────────────────────────────────────
        // EVENTOS (estos SON los que el Designer espera)
        // ─────────────────────────────────────────────

        private void txtBuscar_TextChanged(object sender, EventArgs e)
        {
            string filtro = txtBuscar.Text.ToLower();

            foreach (DataGridViewRow row in dgvHistorial.Rows)
            {
                bool visible =
                    row.Cells["colShadeName"].Value?.ToString().ToLower().Contains(filtro) == true ||
                    row.Cells["colIluminante"].Value?.ToString().ToLower().Contains(filtro) == true;

                row.Visible = visible;
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "CSV (*.csv)|*.csv";
                dlg.FileName = "Historial_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".csv";

                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                // Usar UTF8 con BOM para que Excel detecte automáticamente los acentos
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(dlg.FileName, false, new System.Text.UTF8Encoding(true)))
                {
                    // Encabezados
                    var headers = new System.Collections.Generic.List<string>();
                    foreach (DataGridViewColumn col in dgvHistorial.Columns)
                    {
                        headers.Add(col.HeaderText);
                    }
                    sw.WriteLine(string.Join(";", headers));

                    // Filas visibles
                    foreach (DataGridViewRow row in dgvHistorial.Rows)
                    {
                        if (!row.Visible || row.IsNewRow) continue;

                        var cells = new System.Collections.Generic.List<string>();
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            cells.Add((cell.Value ?? "").ToString());
                        }
                        sw.WriteLine(string.Join(";", cells));
                    }
                }

                MessageBox.Show(
                    "Exportación completada correctamente.",
                    "Exportar CSV",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnBorrar_Click(object sender, EventArgs e)
        {
            if (dgvHistorial.SelectedRows.Count > 0)
            {
                var result = MessageBox.Show(
                    "¿Estás seguro de que deseas borrar este registro del historial?",
                    "Confirmar borrado",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // Remover de la grilla
                    foreach (DataGridViewRow rowSelect in dgvHistorial.SelectedRows)
                    {
                        dgvHistorial.Rows.Remove(rowSelect);
                    }
                    
                    // Actualizar el CSV
                    DataTable dt = new DataTable();
                    dt.Columns.Add("FechaHora");
                    dt.Columns.Add("ShadeName");
                    dt.Columns.Add("Iluminante");
                    dt.Columns.Add("Lightness");
                    dt.Columns.Add("Chroma");
                    dt.Columns.Add("Diagnostico_L");
                    dt.Columns.Add("Correccion_L");
                    dt.Columns.Add("Diagnostico_a");
                    dt.Columns.Add("Correccion_a");
                    dt.Columns.Add("Diagnostico_b");
                    dt.Columns.Add("Correccion_b");
                    
                    foreach (DataGridViewRow r in dgvHistorial.Rows)
                    {
                        dt.Rows.Add(
                            r.Cells["colFechaHora"].Value, r.Cells["colShadeName"].Value, r.Cells["colIluminante"].Value,
                            r.Cells["colLightness"].Value, r.Cells["colChroma"].Value, r.Cells["colDiagL"].Value,
                            r.Cells["colCorrL"].Value, r.Cells["colDiagA"].Value, r.Cells["colCorrA"].Value,
                            r.Cells["colDiagB"].Value, r.Cells["colCorrB"].Value
                        );
                    }
                    
                    Color.Services.HistorialService.GuardarHistorialCompleto(dt);
                    lblContador.Text = "Total de registros: " + dgvHistorial.Rows.Count;
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecciona una fila para borrar.", "Ninguna selección", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}