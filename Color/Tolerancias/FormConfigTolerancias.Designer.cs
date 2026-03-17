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
            this.lblDL = new System.Windows.Forms.Label();
            this.txtDL = new System.Windows.Forms.TextBox();
            this.lblDC = new System.Windows.Forms.Label();
            this.txtDC = new System.Windows.Forms.TextBox();
            this.lblDH = new System.Windows.Forms.Label();
            this.txtDH = new System.Windows.Forms.TextBox();
            this.lblDE = new System.Windows.Forms.Label();
            this.txtDE = new System.Windows.Forms.TextBox();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.SuspendLayout();

            // lblTitulo
            this.lblTitulo.Text = "Configuración de Tolerancias";
            this.lblTitulo.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.lblTitulo.Location = new System.Drawing.Point(20, 15);
            this.lblTitulo.Size = new System.Drawing.Size(340, 28);
            this.lblTitulo.AutoSize = false;

            // lblDL
            this.lblDL.Text = "Tolerancia DL:";
            this.lblDL.Location = new System.Drawing.Point(20, 65);
            this.lblDL.Size = new System.Drawing.Size(130, 23);
            this.lblDL.Font = new System.Drawing.Font("Segoe UI", 9F);

            // txtDL
            this.txtDL.Location = new System.Drawing.Point(160, 62);
            this.txtDL.Size = new System.Drawing.Size(120, 23);
            this.txtDL.Font = new System.Drawing.Font("Segoe UI", 9F);

            // lblDC
            this.lblDC.Text = "Tolerancia DC:";
            this.lblDC.Location = new System.Drawing.Point(20, 105);
            this.lblDC.Size = new System.Drawing.Size(130, 23);
            this.lblDC.Font = new System.Drawing.Font("Segoe UI", 9F);

            // txtDC
            this.txtDC.Location = new System.Drawing.Point(160, 102);
            this.txtDC.Size = new System.Drawing.Size(120, 23);
            this.txtDC.Font = new System.Drawing.Font("Segoe UI", 9F);

            // lblDH
            this.lblDH.Text = "Tolerancia DH:";
            this.lblDH.Location = new System.Drawing.Point(20, 145);
            this.lblDH.Size = new System.Drawing.Size(130, 23);
            this.lblDH.Font = new System.Drawing.Font("Segoe UI", 9F);

            // txtDH
            this.txtDH.Location = new System.Drawing.Point(160, 142);
            this.txtDH.Size = new System.Drawing.Size(120, 23);
            this.txtDH.Font = new System.Drawing.Font("Segoe UI", 9F);

            // lblDE
            this.lblDE.Text = "Tolerancia DE:";
            this.lblDE.Location = new System.Drawing.Point(20, 185);
            this.lblDE.Size = new System.Drawing.Size(130, 23);
            this.lblDE.Font = new System.Drawing.Font("Segoe UI", 9F);

            // txtDE
            this.txtDE.Location = new System.Drawing.Point(160, 182);
            this.txtDE.Size = new System.Drawing.Size(120, 23);
            this.txtDE.Font = new System.Drawing.Font("Segoe UI", 9F);

            // btnGuardar
            this.btnGuardar.Text = "Guardar";
            this.btnGuardar.Location = new System.Drawing.Point(90, 235);
            this.btnGuardar.Size = new System.Drawing.Size(90, 32);
            this.btnGuardar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(0, 120, 215);
            this.btnGuardar.ForeColor = System.Drawing.Color.White;
            this.btnGuardar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);

            // btnCancelar
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.Location = new System.Drawing.Point(195, 235);
            this.btnCancelar.Size = new System.Drawing.Size(90, 32);
            this.btnCancelar.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnCancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);

            // FormConfigTolerancias
            this.Text = "Configuración de Tolerancias";
            this.ClientSize = new System.Drawing.Size(390, 295);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Controls.Add(this.lblTitulo);
            this.Controls.Add(this.lblDL);
            this.Controls.Add(this.txtDL);
            this.Controls.Add(this.lblDC);
            this.Controls.Add(this.txtDC);
            this.Controls.Add(this.lblDH);
            this.Controls.Add(this.txtDH);
            this.Controls.Add(this.lblDE);
            this.Controls.Add(this.txtDE);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.btnCancelar);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Label lblDL;
        private System.Windows.Forms.TextBox txtDL;
        private System.Windows.Forms.Label lblDC;
        private System.Windows.Forms.TextBox txtDC;
        private System.Windows.Forms.Label lblDH;
        private System.Windows.Forms.TextBox txtDH;
        private System.Windows.Forms.Label lblDE;
        private System.Windows.Forms.TextBox txtDE;
        private System.Windows.Forms.Button btnGuardar;
        private System.Windows.Forms.Button btnCancelar;
    }
}