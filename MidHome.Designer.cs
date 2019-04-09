namespace cost_management
{
    partial class MidHome
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MidHome));
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.partSummeryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.kPIToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TSMLNG = new System.Windows.Forms.ToolStripMenuItem();
            this.TSMRefresh = new System.Windows.Forms.ToolStripMenuItem();
            this.extraPolationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.imgstatus = new System.Windows.Forms.ImageList(this.components);
            this.menuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip
            // 
            this.menuStrip.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            resources.ApplyResources(this.menuStrip, "menuStrip");
            this.menuStrip.GripMargin = new System.Windows.Forms.Padding(0);
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.partSummeryToolStripMenuItem,
            this.kPIToolStripMenuItem,
            this.TSMLNG,
            this.TSMRefresh,
            this.extraPolationToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.menuStrip.Name = "menuStrip";
            // 
            // partSummeryToolStripMenuItem
            // 
            this.partSummeryToolStripMenuItem.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            resources.ApplyResources(this.partSummeryToolStripMenuItem, "partSummeryToolStripMenuItem");
            this.partSummeryToolStripMenuItem.ForeColor = System.Drawing.Color.White;
            this.partSummeryToolStripMenuItem.Name = "partSummeryToolStripMenuItem";
            this.partSummeryToolStripMenuItem.Click += new System.EventHandler(this.partSummeryToolStripMenuItem_Click);
            // 
            // kPIToolStripMenuItem
            // 
            this.kPIToolStripMenuItem.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            resources.ApplyResources(this.kPIToolStripMenuItem, "kPIToolStripMenuItem");
            this.kPIToolStripMenuItem.ForeColor = System.Drawing.Color.White;
            this.kPIToolStripMenuItem.Name = "kPIToolStripMenuItem";
            this.kPIToolStripMenuItem.Click += new System.EventHandler(this.kPIToolStripMenuItem_Click);
            // 
            // TSMLNG
            // 
            this.TSMLNG.Checked = true;
            this.TSMLNG.CheckState = System.Windows.Forms.CheckState.Indeterminate;
            resources.ApplyResources(this.TSMLNG, "TSMLNG");
            this.TSMLNG.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.TSMLNG.Name = "TSMLNG";
            this.TSMLNG.Click += new System.EventHandler(this.TSMLNG_Click);
            // 
            // TSMRefresh
            // 
            resources.ApplyResources(this.TSMRefresh, "TSMRefresh");
            this.TSMRefresh.ForeColor = System.Drawing.Color.White;
            this.TSMRefresh.Name = "TSMRefresh";
            this.TSMRefresh.Click += new System.EventHandler(this.TSMRefresh_Click);
            // 
            // extraPolationToolStripMenuItem
            // 
            this.extraPolationToolStripMenuItem.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.extraPolationToolStripMenuItem.Name = "extraPolationToolStripMenuItem";
            resources.ApplyResources(this.extraPolationToolStripMenuItem, "extraPolationToolStripMenuItem");
            this.extraPolationToolStripMenuItem.Click += new System.EventHandler(this.extraPolationToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            resources.ApplyResources(this.exitToolStripMenuItem, "exitToolStripMenuItem");
            this.exitToolStripMenuItem.ForeColor = System.Drawing.Color.White;
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // imgstatus
            // 
            this.imgstatus.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgstatus.ImageStream")));
            this.imgstatus.TransparentColor = System.Drawing.Color.Transparent;
            this.imgstatus.Images.SetKeyName(0, "DarkYello.png");
            this.imgstatus.Images.SetKeyName(1, "Green.png");
            this.imgstatus.Images.SetKeyName(2, "red.png");
            this.imgstatus.Images.SetKeyName(3, "yellow.png");
            // 
            // MidHome
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.Controls.Add(this.menuStrip);
            this.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.IsMdiContainer = true;
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip;
            this.Name = "MidHome";
            this.TransparencyKey = System.Drawing.Color.Black;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MidHome_FormClosing);
            this.Load += new System.EventHandler(this.MidHome_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MidHome_KeyDown);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem kPIToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem TSMLNG;
        public System.Windows.Forms.ToolStripMenuItem partSummeryToolStripMenuItem;
        private System.Windows.Forms.ImageList imgstatus;
        private System.Windows.Forms.ToolStripMenuItem TSMRefresh;
        private System.Windows.Forms.ToolStripMenuItem extraPolationToolStripMenuItem;
    }
}



