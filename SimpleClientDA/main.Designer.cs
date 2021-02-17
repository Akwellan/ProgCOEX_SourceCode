namespace Dashboard
{
    partial class ProgCOEX
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgCOEX));
            this.panel1 = new System.Windows.Forms.Panel();
            this.pnlNav = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnParametres = new System.Windows.Forms.Button();
            this.btnRecettes = new System.Windows.Forms.Button();
            this.btnCorrection = new System.Windows.Forms.Button();
            this.btnDashboard = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.nomPage = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.notifyProgCOEX = new System.Windows.Forms.NotifyIcon(this.components);
            this.btnHide = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.RecetteControlleur = new Siemens.Opc.DaClient.Controller.Recette();
            this.DashControlleur = new Siemens.Opc.DaClient.Controller.Dash();
            this.CorrectionControlleur = new Dashboard.Correction();
            this.panel1.SuspendLayout();
            this.pnlNav.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(24)))), ((int)(((byte)(30)))), ((int)(((byte)(54)))));
            this.panel1.Controls.Add(this.pnlNav);
            this.panel1.Controls.Add(this.btnParametres);
            this.panel1.Controls.Add(this.btnRecettes);
            this.panel1.Controls.Add(this.btnCorrection);
            this.panel1.Controls.Add(this.btnDashboard);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(247, 738);
            this.panel1.TabIndex = 0;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // pnlNav
            // 
            this.pnlNav.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(126)))), ((int)(((byte)(249)))));
            this.pnlNav.Controls.Add(this.flowLayoutPanel1);
            this.pnlNav.Location = new System.Drawing.Point(0, 238);
            this.pnlNav.Margin = new System.Windows.Forms.Padding(4);
            this.pnlNav.Name = "pnlNav";
            this.pnlNav.Size = new System.Drawing.Size(4, 123);
            this.pnlNav.TabIndex = 2;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Location = new System.Drawing.Point(3, 15);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(267, 123);
            this.flowLayoutPanel1.TabIndex = 0;
            // 
            // btnParametres
            // 
            this.btnParametres.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.btnParametres.FlatAppearance.BorderSize = 0;
            this.btnParametres.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnParametres.Font = new System.Drawing.Font("Nirmala UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnParametres.ForeColor = System.Drawing.Color.White;
            this.btnParametres.Image = ((System.Drawing.Image)(resources.GetObject("btnParametres.Image")));
            this.btnParametres.Location = new System.Drawing.Point(0, 676);
            this.btnParametres.Margin = new System.Windows.Forms.Padding(4);
            this.btnParametres.Name = "btnParametres";
            this.btnParametres.Size = new System.Drawing.Size(247, 62);
            this.btnParametres.TabIndex = 1;
            this.btnParametres.Text = "Paramètres";
            this.btnParametres.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnParametres.UseVisualStyleBackColor = true;
            this.btnParametres.Click += new System.EventHandler(this.btnParametres_Click);
            this.btnParametres.Leave += new System.EventHandler(this.btnParametres_Leave);
            // 
            // btnRecettes
            // 
            this.btnRecettes.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnRecettes.FlatAppearance.BorderSize = 0;
            this.btnRecettes.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRecettes.Font = new System.Drawing.Font("Nirmala UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRecettes.ForeColor = System.Drawing.Color.White;
            this.btnRecettes.Image = ((System.Drawing.Image)(resources.GetObject("btnRecettes.Image")));
            this.btnRecettes.Location = new System.Drawing.Point(0, 290);
            this.btnRecettes.Margin = new System.Windows.Forms.Padding(4);
            this.btnRecettes.Name = "btnRecettes";
            this.btnRecettes.Size = new System.Drawing.Size(247, 62);
            this.btnRecettes.TabIndex = 1;
            this.btnRecettes.Text = "Recettes   ";
            this.btnRecettes.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnRecettes.UseVisualStyleBackColor = true;
            this.btnRecettes.Click += new System.EventHandler(this.btnRecettes_Click);
            this.btnRecettes.Leave += new System.EventHandler(this.btnRecettes_Leave);
            // 
            // btnCorrection
            // 
            this.btnCorrection.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnCorrection.FlatAppearance.BorderSize = 0;
            this.btnCorrection.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCorrection.Font = new System.Drawing.Font("Nirmala UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCorrection.ForeColor = System.Drawing.Color.White;
            this.btnCorrection.Image = ((System.Drawing.Image)(resources.GetObject("btnCorrection.Image")));
            this.btnCorrection.Location = new System.Drawing.Point(0, 228);
            this.btnCorrection.Margin = new System.Windows.Forms.Padding(4);
            this.btnCorrection.Name = "btnCorrection";
            this.btnCorrection.Size = new System.Drawing.Size(247, 62);
            this.btnCorrection.TabIndex = 1;
            this.btnCorrection.Text = "Correction";
            this.btnCorrection.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnCorrection.UseVisualStyleBackColor = true;
            this.btnCorrection.Click += new System.EventHandler(this.btnCorrection_Click);
            this.btnCorrection.Leave += new System.EventHandler(this.btnCorrection_Leave);
            // 
            // btnDashboard
            // 
            this.btnDashboard.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnDashboard.FlatAppearance.BorderSize = 0;
            this.btnDashboard.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnDashboard.Font = new System.Drawing.Font("Nirmala UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDashboard.ForeColor = System.Drawing.Color.White;
            this.btnDashboard.Image = ((System.Drawing.Image)(resources.GetObject("btnDashboard.Image")));
            this.btnDashboard.Location = new System.Drawing.Point(0, 166);
            this.btnDashboard.Margin = new System.Windows.Forms.Padding(4);
            this.btnDashboard.Name = "btnDashboard";
            this.btnDashboard.Size = new System.Drawing.Size(247, 62);
            this.btnDashboard.TabIndex = 1;
            this.btnDashboard.Text = "Dashboard";
            this.btnDashboard.TextImageRelation = System.Windows.Forms.TextImageRelation.TextBeforeImage;
            this.btnDashboard.UseVisualStyleBackColor = true;
            this.btnDashboard.Click += new System.EventHandler(this.btnDashboard_Click);
            this.btnDashboard.Leave += new System.EventHandler(this.btnDashboard_Leave);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(247, 166);
            this.panel2.TabIndex = 0;
            this.panel2.Paint += new System.Windows.Forms.PaintEventHandler(this.panel2_Paint);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("STEPS Outline", 14.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(65, 118);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 27);
            this.label1.TabIndex = 1;
            this.label1.Text = "COEX";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(71, 15);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(107, 100);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // nomPage
            // 
            this.nomPage.AutoSize = true;
            this.nomPage.Font = new System.Drawing.Font("Nirmala UI", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nomPage.ForeColor = System.Drawing.Color.White;
            this.nomPage.Location = new System.Drawing.Point(267, 30);
            this.nomPage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.nomPage.Name = "nomPage";
            this.nomPage.Size = new System.Drawing.Size(96, 37);
            this.nomPage.TabIndex = 1;
            this.nomPage.Text = "label2";
            // 
            // btnClose
            // 
            this.btnClose.FlatAppearance.BorderSize = 0;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.Location = new System.Drawing.Point(1289, 15);
            this.btnClose.Margin = new System.Windows.Forms.Padding(4);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(28, 26);
            this.btnClose.TabIndex = 3;
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // notifyProgCOEX
            // 
            this.notifyProgCOEX.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.notifyProgCOEX.BalloonTipText = "Programme COEX\r\n\r\nFonctionnalité :\r\n- Correction diametre par le vide,\r\n- Mini-Mo" +
    "nitoring,\r\n- Gestion des recettes et parametre.\r\n\r\nDévelopper par :\r\n- THIERRY B" +
    "enjamin,\r\n- SERANGE Bastien.";
            this.notifyProgCOEX.BalloonTipTitle = "ProgCOEX";
            this.notifyProgCOEX.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyProgCOEX.Icon")));
            this.notifyProgCOEX.Text = "ProgCOEX";
            this.notifyProgCOEX.Visible = true;
            this.notifyProgCOEX.MouseClick += new System.Windows.Forms.MouseEventHandler(this.ProgCOEX_MouseDown);
            // 
            // btnHide
            // 
            this.btnHide.FlatAppearance.BorderSize = 0;
            this.btnHide.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnHide.Image = ((System.Drawing.Image)(resources.GetObject("btnHide.Image")));
            this.btnHide.Location = new System.Drawing.Point(1237, 15);
            this.btnHide.Margin = new System.Windows.Forms.Padding(4);
            this.btnHide.Name = "btnHide";
            this.btnHide.Size = new System.Drawing.Size(28, 26);
            this.btnHide.TabIndex = 5;
            this.btnHide.UseVisualStyleBackColor = true;
            this.btnHide.Click += new System.EventHandler(this.btnHide_Click);
            // 
            // button1
            // 
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(1188, 15);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(28, 26);
            this.button1.TabIndex = 6;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // RecetteControlleur
            // 
            this.RecetteControlleur.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.RecetteControlleur.Location = new System.Drawing.Point(246, 71);
            this.RecetteControlleur.Margin = new System.Windows.Forms.Padding(4);
            this.RecetteControlleur.Name = "RecetteControlleur";
            this.RecetteControlleur.Size = new System.Drawing.Size(1079, 654);
            this.RecetteControlleur.TabIndex = 8;
            // 
            // DashControlleur
            // 
            this.DashControlleur.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.DashControlleur.Location = new System.Drawing.Point(246, 71);
            this.DashControlleur.Margin = new System.Windows.Forms.Padding(4);
            this.DashControlleur.Name = "DashControlleur";
            this.DashControlleur.Size = new System.Drawing.Size(1088, 654);
            this.DashControlleur.TabIndex = 7;
            // 
            // CorrectionControlleur
            // 
            this.CorrectionControlleur.AllowDrop = true;
            this.CorrectionControlleur.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.CorrectionControlleur.Location = new System.Drawing.Point(246, 71);
            this.CorrectionControlleur.Margin = new System.Windows.Forms.Padding(4);
            this.CorrectionControlleur.Name = "CorrectionControlleur";
            this.CorrectionControlleur.Size = new System.Drawing.Size(1079, 654);
            this.CorrectionControlleur.TabIndex = 9;
            // 
            // ProgCOEX
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(46)))), ((int)(((byte)(51)))), ((int)(((byte)(73)))));
            this.ClientSize = new System.Drawing.Size(1394, 738);
            this.Controls.Add(this.CorrectionControlleur);
            this.Controls.Add(this.RecetteControlleur);
            this.Controls.Add(this.DashControlleur);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnHide);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.nomPage);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ProgCOEX";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ProgCOEX";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.Leave += new System.EventHandler(this.btnRecettes_Leave);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ProgCOEX_MouseDown);
            this.panel1.ResumeLayout(false);
            this.pnlNav.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnDashboard;
        private System.Windows.Forms.Button btnParametres;
        private System.Windows.Forms.Button btnRecettes;
        private System.Windows.Forms.Button btnCorrection;
        private System.Windows.Forms.Panel pnlNav;
        private System.Windows.Forms.Label nomPage;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnClose;
        private Siemens.Opc.DaClient.Controller.Dash dash1;
        public System.Windows.Forms.NotifyIcon notifyProgCOEX;
        private System.Windows.Forms.Button btnHide;
        private System.Windows.Forms.Button button1;
        private Siemens.Opc.DaClient.Controller.Dash DashControlleur;
        private Siemens.Opc.DaClient.Controller.Recette RecetteControlleur;
        private Correction CorrectionControlleur;
    }
}

