namespace Color
{
    partial class FormMainShell
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Panel pnlSidebar;
        private System.Windows.Forms.Panel pnlContent;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Button btnNavScan;
        private System.Windows.Forms.Button btnNavDashboard;
        private System.Windows.Forms.Button btnNavHistory;
        private System.Windows.Forms.Button btnNavConfig;
        private System.Windows.Forms.Button btnNavCielab;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Label lblLogo;
        private System.Windows.Forms.Label lblStatusInfo;
        private System.Windows.Forms.PictureBox picCoats;

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
            this.btnExit = new System.Windows.Forms.Button();
            this.btnNavConfig = new System.Windows.Forms.Button();
            this.btnNavHistory = new System.Windows.Forms.Button();
            this.btnNavCielab = new System.Windows.Forms.Button();
            this.btnNavDashboard = new System.Windows.Forms.Button();
            this.btnNavScan = new System.Windows.Forms.Button();
            this.lblLogo = new System.Windows.Forms.Label();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblStatusInfo = new System.Windows.Forms.Label();
            this.pnlContent = new System.Windows.Forms.Panel();
            this.pnlSidebar.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.SuspendLayout();

            // 
            // pnlSidebar
            // 
            this.pnlSidebar.BackColor = System.Drawing.Color.FromArgb(20, 25, 45); // Azul Marino Profundo
            this.pnlSidebar.Controls.Add(this.btnExit);
            this.pnlSidebar.Controls.Add(this.btnNavConfig);
            this.pnlSidebar.Controls.Add(this.btnNavHistory);
            this.pnlSidebar.Controls.Add(this.btnNavCielab);
            this.pnlSidebar.Controls.Add(this.btnNavDashboard);
            this.pnlSidebar.Controls.Add(this.btnNavScan);
            this.pnlSidebar.Controls.Add(this.lblLogo);
            this.pnlSidebar.Dock = System.Windows.Forms.DockStyle.Left;
            this.pnlSidebar.Location = new System.Drawing.Point(0, 0);
            this.pnlSidebar.Name = "pnlSidebar";
            this.pnlSidebar.Size = new System.Drawing.Size(260, 800);
            this.pnlSidebar.TabIndex = 0;

            // 
            // lblLogo
            // 
            this.lblLogo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblLogo.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold);
            this.lblLogo.ForeColor = System.Drawing.Color.White;
            this.lblLogo.Location = new System.Drawing.Point(0, 0);
            this.lblLogo.Name = "lblLogo";
            this.lblLogo.Size = new System.Drawing.Size(260, 120);
            this.lblLogo.TabIndex = 0;
            this.lblLogo.Text = "COATS\r\nTINT";
            this.lblLogo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // 
            // btnNavScan
            // 
            this.btnNavScan.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavScan.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavScan.FlatAppearance.BorderSize = 0;
            this.btnNavScan.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavScan.Font = new System.Drawing.Font("Segoe UI Semibold", 11F);
            this.btnNavScan.ForeColor = System.Drawing.Color.White;
            this.btnNavScan.Location = new System.Drawing.Point(0, 120);
            this.btnNavScan.Name = "btnNavScan";
            this.btnNavScan.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnNavScan.Size = new System.Drawing.Size(260, 60);
            this.btnNavScan.TabIndex = 1;
            this.btnNavScan.Text = "  🔍  Lectura Digital";
            this.btnNavScan.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavScan.UseVisualStyleBackColor = true;
            this.btnNavScan.Click += new System.EventHandler(this.btnNavScan_Click);

            // 
            // btnNavDashboard
            // 
            this.btnNavDashboard.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavDashboard.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavDashboard.FlatAppearance.BorderSize = 0;
            this.btnNavDashboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavDashboard.Font = new System.Drawing.Font("Segoe UI Semibold", 11F);
            this.btnNavDashboard.ForeColor = System.Drawing.Color.White;
            this.btnNavDashboard.Location = new System.Drawing.Point(0, 180);
            this.btnNavDashboard.Name = "btnNavDashboard";
            this.btnNavDashboard.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnNavDashboard.Size = new System.Drawing.Size(260, 60);
            this.btnNavDashboard.TabIndex = 2;
            this.btnNavDashboard.Text = "  📊  Análisis Datos";
            this.btnNavDashboard.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavDashboard.UseVisualStyleBackColor = true;
            this.btnNavDashboard.Click += new System.EventHandler(this.btnNavDashboard_Click);

            // 
            // btnNavCielab
            // 
            this.btnNavCielab.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavCielab.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavCielab.FlatAppearance.BorderSize = 0;
            this.btnNavCielab.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavCielab.Font = new System.Drawing.Font("Segoe UI Semibold", 11F);
            this.btnNavCielab.ForeColor = System.Drawing.Color.White;
            this.btnNavCielab.Location = new System.Drawing.Point(0, 240);
            this.btnNavCielab.Name = "btnNavCielab";
            this.btnNavCielab.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnNavCielab.Size = new System.Drawing.Size(260, 60);
            this.btnNavCielab.TabIndex = 6;
            this.btnNavCielab.Text = "  🎯  Gráfico CIELAB";
            this.btnNavCielab.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavCielab.UseVisualStyleBackColor = true;
            this.btnNavCielab.Click += new System.EventHandler(this.btnNavCielab_Click);

            // 
            // btnNavHistory
            // 
            this.btnNavHistory.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavHistory.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavHistory.FlatAppearance.BorderSize = 0;
            this.btnNavHistory.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavHistory.Font = new System.Drawing.Font("Segoe UI Semibold", 11F);
            this.btnNavHistory.ForeColor = System.Drawing.Color.White;
            this.btnNavHistory.Location = new System.Drawing.Point(0, 300);
            this.btnNavHistory.Name = "btnNavHistory";
            this.btnNavHistory.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnNavHistory.Size = new System.Drawing.Size(260, 60);
            this.btnNavHistory.TabIndex = 3;
            this.btnNavHistory.Text = "  🕒  Historial";
            this.btnNavHistory.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavHistory.UseVisualStyleBackColor = true;
            this.btnNavHistory.Click += new System.EventHandler(this.btnNavHistory_Click);

            // 
            // btnNavConfig
            // 
            this.btnNavConfig.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnNavConfig.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnNavConfig.FlatAppearance.BorderSize = 0;
            this.btnNavConfig.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnNavConfig.Font = new System.Drawing.Font("Segoe UI Semibold", 11F);
            this.btnNavConfig.ForeColor = System.Drawing.Color.White;
            this.btnNavConfig.Location = new System.Drawing.Point(0, 330);
            this.btnNavConfig.Name = "btnNavConfig";
            this.btnNavConfig.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnNavConfig.Size = new System.Drawing.Size(260, 70);
            this.btnNavConfig.TabIndex = 4;
            this.btnNavConfig.Text = "  ⚙️  Tolerancias";
            this.btnNavConfig.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnNavConfig.UseVisualStyleBackColor = true;
            this.btnNavConfig.Click += new System.EventHandler(this.btnNavConfig_Click);

            // 
            // btnExit
            // 
            this.btnExit.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnExit.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnExit.FlatAppearance.BorderSize = 0;
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnExit.ForeColor = System.Drawing.Color.FromArgb(255, 100, 100);
            this.btnExit.Location = new System.Drawing.Point(0, 730);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(260, 70);
            this.btnExit.TabIndex = 5;
            this.btnExit.Text = "Cerrar Sistema";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);

            // 
            // pnlTop
            // 
            this.pnlTop.BackColor = System.Drawing.Color.White;
            this.pnlTop.Controls.Add(this.lblStatusInfo);
            this.pnlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlTop.Location = new System.Drawing.Point(260, 0);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Size = new System.Drawing.Size(940, 60);
            this.pnlTop.TabIndex = 1;

            // 
            // lblStatusInfo
            // 
            this.lblStatusInfo.AutoSize = true;
            this.lblStatusInfo.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblStatusInfo.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100);
            this.lblStatusInfo.Location = new System.Drawing.Point(20, 15);
            this.lblStatusInfo.Name = "lblStatusInfo";
            this.lblStatusInfo.Size = new System.Drawing.Size(180, 28);
            this.lblStatusInfo.TabIndex = 0;
            this.lblStatusInfo.Text = "Entorno Unificado";

            // 
            // pnlContent
            // 
            this.pnlContent.BackColor = System.Drawing.Color.FromArgb(245, 247, 250);
            this.pnlContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlContent.Location = new System.Drawing.Point(260, 60);
            this.pnlContent.Name = "pnlContent";
            this.pnlContent.Size = new System.Drawing.Size(940, 740);
            this.pnlContent.TabIndex = 2;

            // 
            // FormMainShell
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 800);
            this.Controls.Add(this.pnlContent);
            this.Controls.Add(this.pnlTop);
            this.Controls.Add(this.pnlSidebar);
            this.Name = "FormMainShell";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.pnlSidebar.ResumeLayout(false);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.ResumeLayout(false);
        }
    }
}
