namespace Color
{
    partial class FormMainShell
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel pnlSidebar;
        private System.Windows.Forms.Panel pnlHeader;
        private System.Windows.Forms.Panel pnlWorkspace;
        private System.Windows.Forms.Label lblLogo;
        private System.Windows.Forms.Label lblSeccionActiva;
        private System.Windows.Forms.Button btnNavOcr;
        private System.Windows.Forms.Button btnNavAnalisis;
        private System.Windows.Forms.Button btnNavHistorial;
        private System.Windows.Forms.Button btnSalir;
        private System.Windows.Forms.Label lblFooter;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.pnlSidebar = new System.Windows.Forms.Panel();
            this.btnSalir = new System.Windows.Forms.Button();
            this.btnNavHistorial = new System.Windows.Forms.Button();
            this.btnNavAnalisis = new System.Windows.Forms.Button();
            this.btnNavOcr = new System.Windows.Forms.Button();
            this.lblLogo = new System.Windows.Forms.Label();
            this.pnlHeader = new System.Windows.Forms.Panel();
            this.lblSeccionActiva = new System.Windows.Forms.Label();
            this.pnlWorkspace = new System.Windows.Forms.Panel();
            this.lblFooter = new System.Windows.Forms.Label();
            this.pnlSidebar.SuspendLayout();
            this.pnlHeader.SuspendLayout();
            this.SuspendLayout();

            // 
            // pnlSidebar
            // 
            this.pnlSidebar.BackColor = System.Drawing.Color.FromArgb(0, 51, 153); // Azul Coats Oscuro
            this.pnlSidebar.Controls.Add(this.btnSalir);
            this.pnlSidebar.Controls.Add(this.btnNavHistorial);
            this.pnlSidebar.Controls.Add(this.btnNavAnalisis);
            this.pnlSidebar.Controls.Add(this.btnNavOcr);
            this.pnlSidebar.Controls.Add(this.lblLogo);
            this.pnlSidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlSidebar.Location = new System.Drawing.Point(0, 0);
            this.pnlSidebar.Name = "pnlSidebar";
            this.pnlSidebar.Size = new System.Drawing.Size(240, 750);
            this.pnlSidebar.TabIndex = 0;

            // 
            // lblLogo
            // 
            this.lblLogo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblLogo.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            this.lblLogo.ForeColor = System.Drawing.Color.White;
            this.lblLogo.Location = new System.Drawing.Point(0, 0);
            this.lblLogo.Name = "lblLogo";
            this.lblLogo.Size = new System.Drawing.Size(240, 100);
            this.lblLogo.TabIndex = 0;
            this.lblLogo.Text = "COATS\r\nCOLORIMETRÍA";
            this.lblLogo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // 
            // btnNavOcr
            // 
            this.btnNavOcr.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavOcr.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavOcr.FlatAppearance.BorderSize = 0;
            this.btnNavOcr.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavOcr.Font = new System.Drawing.Font("Segoe UI Semibold", 10F);
            this.btnNavOcr.ForeColor = System.Drawing.Color.White;
            this.btnNavOcr.Location = new System.Drawing.Point(0, 100);
            this.btnNavOcr.Name = "btnNavOcr";
            this.btnNavOcr.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.btnNavOcr.Size = new System.Drawing.Size(240, 60);
            this.btnNavOcr.TabIndex = 1;
            this.btnNavOcr.Text = "  🔍  Lectura OCR / Inicio";
            this.btnNavOcr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavOcr.UseVisualStyleBackColor = true;
            this.btnNavOcr.Click += new System.EventHandler(this.btnNavOcr_Click);

            // 
            // btnNavAnalisis
            // 
            this.btnNavAnalisis.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavAnalisis.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavAnalisis.FlatAppearance.BorderSize = 0;
            this.btnNavAnalisis.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavAnalisis.Font = new System.Drawing.Font("Segoe UI Semibold", 10F);
            this.btnNavAnalisis.ForeColor = System.Drawing.Color.White;
            this.btnNavAnalisis.Location = new System.Drawing.Point(0, 160);
            this.btnNavAnalisis.Name = "btnNavAnalisis";
            this.btnNavAnalisis.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.btnNavAnalisis.Size = new System.Drawing.Size(240, 60);
            this.btnNavAnalisis.TabIndex = 2;
            this.btnNavAnalisis.Text = "  📊  Análisis de Color";
            this.btnNavAnalisis.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavAnalisis.UseVisualStyleBackColor = true;
            this.btnNavAnalisis.Click += new System.EventHandler(this.btnNavAnalisis_Click);

            // 
            // btnNavHistorial
            // 
            this.btnNavHistorial.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavHistorial.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavHistorial.FlatAppearance.BorderSize = 0;
            this.btnNavHistorial.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavHistorial.Font = new System.Drawing.Font("Segoe UI Semibold", 10F);
            this.btnNavHistorial.ForeColor = System.Drawing.Color.White;
            this.btnNavHistorial.Location = new System.Drawing.Point(0, 220);
            this.btnNavHistorial.Name = "btnNavHistorial";
            this.btnNavHistorial.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.btnNavHistorial.Size = new System.Drawing.Size(240, 60);
            this.btnNavHistorial.TabIndex = 3;
            this.btnNavHistorial.Text = "  🕒  Historial de Análisis";
            this.btnNavHistorial.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavHistorial.UseVisualStyleBackColor = true;
            this.btnNavHistorial.Click += new System.EventHandler(this.btnNavHistorial_Click);

            // 
            // btnSalir
            // 
            this.btnSalir.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSalir.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnSalir.FlatAppearance.BorderSize = 0;
            this.btnSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalir.Font = new System.Drawing.Font("Segoe UI Semibold", 10F);
            this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnSalir.Location = new System.Drawing.Point(0, 690);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(240, 60);
            this.btnSalir.TabIndex = 4;
            this.btnSalir.Text = "  🚪  Salir del Sistema";
            this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSalir.Padding = new System.Windows.Forms.Padding(20, 0, 0, 0);
            this.btnSalir.UseVisualStyleBackColor = true;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);

            // 
            // pnlHeader
            // 
            this.pnlHeader.BackColor = System.Drawing.Color.White;
            this.pnlHeader.Controls.Add(this.lblSeccionActiva);
            this.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlHeader.Location = new System.Drawing.Point(240, 0);
            this.pnlHeader.Name = "pnlHeader";
            this.pnlHeader.Size = new System.Drawing.Size(860, 60);
            this.pnlHeader.TabIndex = 1;

            // 
            // lblSeccionActiva
            // 
            this.lblSeccionActiva.AutoSize = true;
            this.lblSeccionActiva.Font = new System.Drawing.Font("Segoe UI Semilight", 14F);
            this.lblSeccionActiva.ForeColor = System.Drawing.Color.FromArgb(64, 64, 64);
            this.lblSeccionActiva.Location = new System.Drawing.Point(20, 16);
            this.lblSeccionActiva.Name = "lblSeccionActiva";
            this.lblSeccionActiva.Size = new System.Drawing.Size(125, 32);
            this.lblSeccionActiva.TabIndex = 0;
            this.lblSeccionActiva.Text = "Dashboard";

            // 
            // pnlWorkspace
            // 
            this.pnlWorkspace.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlWorkspace.Location = new System.Drawing.Point(240, 60);
            this.pnlWorkspace.Name = "pnlWorkspace";
            this.pnlWorkspace.Size = new System.Drawing.Size(860, 690);
            this.pnlWorkspace.TabIndex = 2;

            // 
            // FormMainShell
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1100, 750);
            this.Controls.Add(this.pnlWorkspace);
            this.Controls.Add(this.pnlHeader);
            this.Controls.Add(this.pnlSidebar);
            this.Name = "FormMainShell";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "COATS CADENA - Expert Colorimetric System";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.pnlSidebar.ResumeLayout(false);
            this.pnlHeader.ResumeLayout(false);
            this.pnlHeader.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
