namespace Color.Tolerancias
{
    partial class FormConfigTolerancias
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.lblTitulo = new System.Windows.Forms.Label();
            this.flowCards = new System.Windows.Forms.FlowLayoutPanel();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.SuspendLayout();

            // lblTitulo
            this.lblTitulo.Text = "Configuración de Tolerancias";
            this.lblTitulo.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblTitulo.Location = new System.Drawing.Point(0, 15);
            this.lblTitulo.Size = new System.Drawing.Size(600, 35);
            this.lblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblTitulo.AutoSize = false;

            // flowCards
            this.flowCards.Location = new System.Drawing.Point(20, 60);
            this.flowCards.Size = new System.Drawing.Size(560, 240);
            this.flowCards.AutoScroll = true;
            this.flowCards.WrapContents = false;
            // Center the flow layout if there's enough space
            this.flowCards.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;

            // btnGuardar
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.Location = new System.Drawing.Point(180, 310);
            this.btnGuardar.Size = new System.Drawing.Size(110, 40);
            this.btnGuardar.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);

            // btnCancelar
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.Location = new System.Drawing.Point(310, 310);
            this.btnCancelar.Size = new System.Drawing.Size(110, 40);
            this.btnCancelar.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);

            // FormConfigTolerancias
            this.Text = "Configuración de Tolerancias";
            this.ClientSize = new System.Drawing.Size(600, 380);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.lblTitulo);
            this.Controls.Add(this.flowCards);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.btnCancelar);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.FlowLayoutPanel flowCards;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnCancelar;
    }
}