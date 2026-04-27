using System;
using System.Drawing;
using System.Windows.Forms;

namespace Color
{
    partial class Form1
    {
        private System.Windows.Forms.Panel leftNav;
        private System.Windows.Forms.Label lblBrand;
        private System.Windows.Forms.Button btnTolerancias;
        private System.Windows.Forms.Button btnBaseDatos;
        private System.Windows.Forms.Button btnSalir;

        private System.Windows.Forms.Panel mainArea;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label lblSubtitle;

        private System.Windows.Forms.Panel contentBorder;
        private System.Windows.Forms.Label lblLeftTitle;
        private System.Windows.Forms.Label lblRightTitle;

        private System.Windows.Forms.Panel pnlLeftFrame;
        private System.Windows.Forms.Panel pnlRightFrame;

        private System.Windows.Forms.PictureBox picLeft;
        private System.Windows.Forms.PictureBox picRight;

        private System.Windows.Forms.Label lblLeftHint;
        private System.Windows.Forms.Label lblRightHint;

        private System.Windows.Forms.Button btnCargarLeft;
        private System.Windows.Forms.Button btnCargarRight;

        private System.Windows.Forms.Label lblStatus;

        // Botones que aparecen sólo cuando ambas imágenes están cargadas
        private System.Windows.Forms.Button btnIniciar;
        private System.Windows.Forms.Button btnCancelarAccion;

        /// Limpieza de recursos.
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.leftNav = new System.Windows.Forms.Panel();
            this.btnSalir = new System.Windows.Forms.Button();
            this.btnBaseDatos = new System.Windows.Forms.Button();
            this.btnTolerancias = new System.Windows.Forms.Button();
            this.lblBrand = new System.Windows.Forms.Label();
            this.mainArea = new System.Windows.Forms.Panel();
            this.btnCancelarAccion = new System.Windows.Forms.Button();
            this.btnIniciar = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.contentBorder = new System.Windows.Forms.Panel();
            this.btnCargarRight = new System.Windows.Forms.Button();
            this.btnCargarLeft = new System.Windows.Forms.Button();
            this.pnlRightFrame = new System.Windows.Forms.Panel();
            this.lblRightHint = new System.Windows.Forms.Label();
            this.picRight = new System.Windows.Forms.PictureBox();
            this.pnlLeftFrame = new System.Windows.Forms.Panel();
            this.lblLeftHint = new System.Windows.Forms.Label();
            this.picLeft = new System.Windows.Forms.PictureBox();
            this.lblRightTitle = new System.Windows.Forms.Label();
            this.lblLeftTitle = new System.Windows.Forms.Label();
            this.lblSubtitle = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.leftNav.SuspendLayout();
            this.mainArea.SuspendLayout();
            this.contentBorder.SuspendLayout();
            this.pnlRightFrame.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picRight)).BeginInit();
            this.pnlLeftFrame.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picLeft)).BeginInit();
            this.SuspendLayout();
            // 
            // leftNav
            // 
            this.leftNav.BackColor = System.Drawing.Color.FromArgb(10, 33, 58);
            this.leftNav.Controls.Add(this.btnSalir);
            this.leftNav.Controls.Add(this.btnBaseDatos);
            this.leftNav.Controls.Add(this.btnTolerancias);
            this.leftNav.Controls.Add(this.lblBrand);
            this.leftNav.Dock = System.Windows.Forms.DockStyle.Left;
            this.leftNav.Location = new System.Drawing.Point(0, 0);
            this.leftNav.Name = "leftNav";
            this.leftNav.Size = new System.Drawing.Size(220, 720);
            this.leftNav.TabIndex = 0;
            // 
            // btnSalir
            // 
            this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Bottom)));
            this.btnSalir.BackColor = System.Drawing.Color.FromArgb(220, 53, 69);
            this.btnSalir.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSalir.FlatAppearance.BorderSize = 0;
            this.btnSalir.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(170, 30, 45);
            this.btnSalir.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(200, 35, 51);
            this.btnSalir.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSalir.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnSalir.ForeColor = System.Drawing.Color.White;
            this.btnSalir.Location = new System.Drawing.Point(20, 664);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(180, 36);
            this.btnSalir.TabIndex = 3;
            this.btnSalir.Text = "Salir";
            this.btnSalir.UseVisualStyleBackColor = false;
            // 
            // btnBaseDatos
            // 
            this.btnBaseDatos.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnBaseDatos.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBaseDatos.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnBaseDatos.ForeColor = System.Drawing.Color.White;
            this.btnBaseDatos.Location = new System.Drawing.Point(20, 160);
            this.btnBaseDatos.Name = "btnBaseDatos";
            this.btnBaseDatos.Size = new System.Drawing.Size(180, 36);
            this.btnBaseDatos.TabIndex = 2;
            this.btnBaseDatos.Text = "Base de datos";
            this.btnBaseDatos.UseVisualStyleBackColor = true;
            // 
            // btnTolerancias
            // 
            this.btnTolerancias.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.btnTolerancias.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnTolerancias.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.btnTolerancias.ForeColor = System.Drawing.Color.White;
            this.btnTolerancias.Location = new System.Drawing.Point(20, 110);
            this.btnTolerancias.Name = "btnTolerancias";
            this.btnTolerancias.Size = new System.Drawing.Size(180, 36);
            this.btnTolerancias.TabIndex = 1;
            this.btnTolerancias.Text = "Config. de tolerancias";
            this.btnTolerancias.UseVisualStyleBackColor = true;
            // 
            // lblBrand
            // 
            this.lblBrand.AutoSize = true;
            this.lblBrand.Font = new System.Drawing.Font("Segoe UI Semibold", 16F, System.Drawing.FontStyle.Bold);
            this.lblBrand.ForeColor = System.Drawing.Color.White;
            this.lblBrand.Location = new System.Drawing.Point(20, 30);
            this.lblBrand.Name = "lblBrand";
            this.lblBrand.Size = new System.Drawing.Size(162, 30);
            this.lblBrand.TabIndex = 0;
            this.lblBrand.Text = "COLORIMETRIA";
            // 
            // mainArea
            // 
            this.mainArea.BackColor = System.Drawing.Color.White;
            this.mainArea.Controls.Add(this.btnCancelarAccion);
            this.mainArea.Controls.Add(this.btnIniciar);
            this.mainArea.Controls.Add(this.lblStatus);
            this.mainArea.Controls.Add(this.contentBorder);
            this.mainArea.Controls.Add(this.lblSubtitle);
            this.mainArea.Controls.Add(this.lblTitle);
            this.mainArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainArea.Location = new System.Drawing.Point(220, 0);
            this.mainArea.Name = "mainArea";
            this.mainArea.Size = new System.Drawing.Size(980, 720);
            this.mainArea.TabIndex = 1;
            // 
            // btnCancelarAccion
            // 
            this.btnCancelarAccion.BackColor = System.Drawing.Color.FromArgb(90, 97, 104);
            this.btnCancelarAccion.FlatAppearance.BorderSize = 0;
            this.btnCancelarAccion.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancelarAccion.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.btnCancelarAccion.ForeColor = System.Drawing.Color.White;
            this.btnCancelarAccion.Location = new System.Drawing.Point(570, 650);
            this.btnCancelarAccion.Name = "btnCancelarAccion";
            this.btnCancelarAccion.Size = new System.Drawing.Size(160, 36);
            this.btnCancelarAccion.TabIndex = 12;
            this.btnCancelarAccion.Text = "Cancelar";
            this.btnCancelarAccion.UseVisualStyleBackColor = false;
            this.btnCancelarAccion.Visible = false;
            // 
            // btnIniciar
            // 
            this.btnIniciar.BackColor = System.Drawing.Color.FromArgb(45, 126, 247);
            this.btnIniciar.FlatAppearance.BorderSize = 0;
            this.btnIniciar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnIniciar.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold);
            this.btnIniciar.ForeColor = System.Drawing.Color.White;
            this.btnIniciar.Location = new System.Drawing.Point(400, 650);
            this.btnIniciar.Name = "btnIniciar";
            this.btnIniciar.Size = new System.Drawing.Size(160, 36);
            this.btnIniciar.TabIndex = 11;
            this.btnIniciar.Text = "Iniciar escaneo";
            this.btnIniciar.UseVisualStyleBackColor = false;
            this.btnIniciar.Visible = false;
            // 
            // lblStatus
            // 
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(60, 64, 70);
            this.lblStatus.Location = new System.Drawing.Point(230, 610);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(520, 24);
            this.lblStatus.TabIndex = 10;
            this.lblStatus.Text = "Cargue ambas imágenes para continuar";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // contentBorder
            // 
            this.contentBorder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.contentBorder.BackColor = System.Drawing.Color.White;
            this.contentBorder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.contentBorder.Controls.Add(this.btnCargarRight);
            this.contentBorder.Controls.Add(this.btnCargarLeft);
            this.contentBorder.Controls.Add(this.pnlRightFrame);
            this.contentBorder.Controls.Add(this.pnlLeftFrame);
            this.contentBorder.Controls.Add(this.lblRightTitle);
            this.contentBorder.Controls.Add(this.lblLeftTitle);
            this.contentBorder.Location = new System.Drawing.Point(60, 160);
            this.contentBorder.Name = "contentBorder";
            this.contentBorder.Size = new System.Drawing.Size(860, 430);
            this.contentBorder.TabIndex = 9;
            // 
            // btnCargarRight
            // 
            this.btnCargarRight.BackColor = System.Drawing.Color.FromArgb(45, 126, 247);
            this.btnCargarRight.FlatAppearance.BorderSize = 0;
            this.btnCargarRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCargarRight.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Bold);
            this.btnCargarRight.ForeColor = System.Drawing.Color.White;
            this.btnCargarRight.Location = new System.Drawing.Point(548, 340);
            this.btnCargarRight.Name = "btnCargarRight";
            this.btnCargarRight.Size = new System.Drawing.Size(180, 36);
            this.btnCargarRight.TabIndex = 8;
            this.btnCargarRight.Text = "CARGAR IMAGEN";
            this.btnCargarRight.UseVisualStyleBackColor = false;
            // 
            // btnCargarLeft
            // 
            this.btnCargarLeft.BackColor = System.Drawing.Color.FromArgb(45, 126, 247);
            this.btnCargarLeft.FlatAppearance.BorderSize = 0;
            this.btnCargarLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCargarLeft.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Bold);
            this.btnCargarLeft.ForeColor = System.Drawing.Color.White;
            this.btnCargarLeft.Location = new System.Drawing.Point(132, 340);
            this.btnCargarLeft.Name = "btnCargarLeft";
            this.btnCargarLeft.Size = new System.Drawing.Size(180, 36);
            this.btnCargarLeft.TabIndex = 7;
            this.btnCargarLeft.Text = "CARGAR IMAGEN";
            this.btnCargarLeft.UseVisualStyleBackColor = false;
            // 
            // pnlRightFrame
            // 
            this.pnlRightFrame.BackColor = System.Drawing.Color.White;
            this.pnlRightFrame.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlRightFrame.Controls.Add(this.lblRightHint);
            this.pnlRightFrame.Controls.Add(this.picRight);
            this.pnlRightFrame.Location = new System.Drawing.Point(490, 80);
            this.pnlRightFrame.Name = "pnlRightFrame";
            this.pnlRightFrame.Size = new System.Drawing.Size(300, 240);
            this.pnlRightFrame.TabIndex = 6;
            // 
            // lblRightHint
            // 
            this.lblRightHint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblRightHint.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblRightHint.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            this.lblRightHint.Location = new System.Drawing.Point(0, 0);
            this.lblRightHint.Name = "lblRightHint";
            this.lblRightHint.Size = new System.Drawing.Size(298, 238);
            this.lblRightHint.TabIndex = 1;
            this.lblRightHint.Text = "Cargar Shade History Report";
            this.lblRightHint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // picRight
            // 
            this.picRight.BackColor = System.Drawing.Color.White;
            this.picRight.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picRight.Location = new System.Drawing.Point(0, 0);
            this.picRight.Name = "picRight";
            this.picRight.Size = new System.Drawing.Size(298, 238);
            this.picRight.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picRight.TabIndex = 0;
            this.picRight.TabStop = false;
            // 
            // pnlLeftFrame
            // 
            this.pnlLeftFrame.BackColor = System.Drawing.Color.White;
            this.pnlLeftFrame.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnlLeftFrame.Controls.Add(this.lblLeftHint);
            this.pnlLeftFrame.Controls.Add(this.picLeft);
            this.pnlLeftFrame.Location = new System.Drawing.Point(74, 80);
            this.pnlLeftFrame.Name = "pnlLeftFrame";
            this.pnlLeftFrame.Size = new System.Drawing.Size(300, 240);
            this.pnlLeftFrame.TabIndex = 5;
            // 
            // lblLeftHint
            // 
            this.lblLeftHint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblLeftHint.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblLeftHint.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            this.lblLeftHint.Location = new System.Drawing.Point(0, 0);
            this.lblLeftHint.Name = "lblLeftHint";
            this.lblLeftHint.Size = new System.Drawing.Size(298, 238);
            this.lblLeftHint.TabIndex = 1;
            this.lblLeftHint.Text = "Cargar Sample Comparison";
            this.lblLeftHint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // picLeft
            // 
            this.picLeft.BackColor = System.Drawing.Color.White;
            this.picLeft.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picLeft.Location = new System.Drawing.Point(0, 0);
            this.picLeft.Name = "picLeft";
            this.picLeft.Size = new System.Drawing.Size(298, 238);
            this.picLeft.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picLeft.TabIndex = 0;
            this.picLeft.TabStop = false;
            // 
            // lblRightTitle
            // 
            this.lblRightTitle.AutoSize = true;
            this.lblRightTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 14F, System.Drawing.FontStyle.Bold);
            this.lblRightTitle.Location = new System.Drawing.Point(560, 45);
            this.lblRightTitle.Name = "lblRightTitle";
            this.lblRightTitle.Size = new System.Drawing.Size(69, 25);
            this.lblRightTitle.TabIndex = 4;
            this.lblRightTitle.Text = "Shade History Report";
            // 
            // lblLeftTitle
            // 
            this.lblLeftTitle.AutoSize = true;
            this.lblLeftTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 14F, System.Drawing.FontStyle.Bold);
            this.lblLeftTitle.Location = new System.Drawing.Point(130, 45);
            this.lblLeftTitle.Name = "lblLeftTitle";
            this.lblLeftTitle.Size = new System.Drawing.Size(111, 25);
            this.lblLeftTitle.TabIndex = 3;
            this.lblLeftTitle.Text = "Sample Comparison";
            // 
            // lblSubtitle
            // 
            this.lblSubtitle.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblSubtitle.ForeColor = System.Drawing.Color.FromArgb(60, 64, 70);
            this.lblSubtitle.Location = new System.Drawing.Point(60, 110);
            this.lblSubtitle.Name = "lblSubtitle";
            this.lblSubtitle.Size = new System.Drawing.Size(860, 26);
            this.lblSubtitle.TabIndex = 8;
            this.lblSubtitle.Text = "Analisis de Colorimetria";
            this.lblSubtitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblTitle
            // 
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Black", 20F, System.Drawing.FontStyle.Bold);
            this.lblTitle.Location = new System.Drawing.Point(60, 60);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(860, 38);
            this.lblTitle.TabIndex = 7;
            this.lblTitle.Text = "COATS CADENA";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form1
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1200, 720);
            this.Controls.Add(this.mainArea);
            this.Controls.Add(this.leftNav);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.MinimumSize = new System.Drawing.Size(1100, 680);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TINT COATS CADENA";
            this.leftNav.ResumeLayout(false);
            this.leftNav.PerformLayout();
            this.mainArea.ResumeLayout(false);
            this.contentBorder.ResumeLayout(false);
            this.contentBorder.PerformLayout();
            this.pnlRightFrame.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picRight)).EndInit();
            this.pnlLeftFrame.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picLeft)).EndInit();
            this.ResumeLayout(false);

        }
    }
}