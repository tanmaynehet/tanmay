namespace cost_management
{
    partial class frmuploadExpo
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            this.txtuploadfilepath = new System.Windows.Forms.TextBox();
            this.btnselectfile = new System.Windows.Forms.Button();
            this.btnupload = new System.Windows.Forms.Button();
            this.btnclose = new System.Windows.Forms.Button();
            this.ADGDATA = new atcs.ADGV.AdvancedDataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.ComSheetName = new System.Windows.Forms.ComboBox();
            this.btnsheetName = new System.Windows.Forms.Button();
            this.pbUpload = new System.Windows.Forms.ProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.ADGDATA)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtuploadfilepath
            // 
            this.txtuploadfilepath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.txtuploadfilepath.Location = new System.Drawing.Point(7, 14);
            this.txtuploadfilepath.Name = "txtuploadfilepath";
            this.txtuploadfilepath.ReadOnly = true;
            this.txtuploadfilepath.Size = new System.Drawing.Size(212, 26);
            this.txtuploadfilepath.TabIndex = 0;
            // 
            // btnselectfile
            // 
            this.btnselectfile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnselectfile.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnselectfile.FlatAppearance.BorderSize = 0;
            this.btnselectfile.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnselectfile.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnselectfile.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnselectfile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnselectfile.Font = new System.Drawing.Font("CorpoS", 9F);
            this.btnselectfile.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnselectfile.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnselectfile.Location = new System.Drawing.Point(222, 12);
            this.btnselectfile.Name = "btnselectfile";
            this.btnselectfile.Size = new System.Drawing.Size(78, 30);
            this.btnselectfile.TabIndex = 13;
            this.btnselectfile.Text = "Select Excel ";
            this.btnselectfile.UseVisualStyleBackColor = false;
            this.btnselectfile.Click += new System.EventHandler(this.btnselectfile_Click);
            // 
            // btnupload
            // 
            this.btnupload.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnupload.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnupload.FlatAppearance.BorderSize = 0;
            this.btnupload.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnupload.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnupload.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnupload.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnupload.Font = new System.Drawing.Font("CorpoS", 9F);
            this.btnupload.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnupload.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnupload.Location = new System.Drawing.Point(301, 12);
            this.btnupload.Name = "btnupload";
            this.btnupload.Size = new System.Drawing.Size(73, 30);
            this.btnupload.TabIndex = 14;
            this.btnupload.Text = "Upload";
            this.btnupload.UseVisualStyleBackColor = false;
            this.btnupload.Click += new System.EventHandler(this.btnupload_Click);
            // 
            // btnclose
            // 
            this.btnclose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnclose.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnclose.FlatAppearance.BorderSize = 0;
            this.btnclose.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnclose.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnclose.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnclose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnclose.Font = new System.Drawing.Font("CorpoS", 9F);
            this.btnclose.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnclose.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnclose.Location = new System.Drawing.Point(375, 12);
            this.btnclose.Name = "btnclose";
            this.btnclose.Size = new System.Drawing.Size(73, 30);
            this.btnclose.TabIndex = 16;
            this.btnclose.Text = "Close";
            this.btnclose.UseVisualStyleBackColor = false;
            this.btnclose.Click += new System.EventHandler(this.btnclose_Click);
            // 
            // ADGDATA
            // 
            this.ADGDATA.AllowUserToAddRows = false;
            this.ADGDATA.AllowUserToDeleteRows = false;
            this.ADGDATA.AllowUserToResizeRows = false;
            dataGridViewCellStyle17.BackColor = System.Drawing.Color.DarkGray;
            this.ADGDATA.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle17;
            this.ADGDATA.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ADGDATA.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.ADGDATA.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText;
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle18.Font = new System.Drawing.Font("CorpoS", 10F);
            dataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ADGDATA.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle18;
            this.ADGDATA.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle19.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle19.Font = new System.Drawing.Font("CorpoS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle19.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle19.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle19.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle19.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.ADGDATA.DefaultCellStyle = dataGridViewCellStyle19;
            this.ADGDATA.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.ADGDATA.FilterAndSortEnabled = true;
            this.ADGDATA.Location = new System.Drawing.Point(30, 48);
            this.ADGDATA.Name = "ADGDATA";
            dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle20.Font = new System.Drawing.Font("CorpoS", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ADGDATA.RowHeadersDefaultCellStyle = dataGridViewCellStyle20;
            this.ADGDATA.RowHeadersVisible = false;
            this.ADGDATA.Size = new System.Drawing.Size(394, 453);
            this.ADGDATA.TabIndex = 17;
            this.ADGDATA.SortStringChanged += new System.EventHandler(this.ADGDATA_SortStringChanged);
            this.ADGDATA.FilterStringChanged += new System.EventHandler(this.ADGDATA_FilterStringChanged);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.pbUpload);
            this.panel1.Controls.Add(this.btnsheetName);
            this.panel1.Controls.Add(this.ComSheetName);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(100, 168);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(227, 123);
            this.panel1.TabIndex = 18;
            this.panel1.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(24, 17);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(99, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Sheet Name";
            // 
            // ComSheetName
            // 
            this.ComSheetName.Font = new System.Drawing.Font("CorpoS", 12F);
            this.ComSheetName.FormattingEnabled = true;
            this.ComSheetName.Location = new System.Drawing.Point(27, 44);
            this.ComSheetName.Name = "ComSheetName";
            this.ComSheetName.Size = new System.Drawing.Size(178, 26);
            this.ComSheetName.TabIndex = 1;
            // 
            // btnsheetName
            // 
            this.btnsheetName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.btnsheetName.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnsheetName.FlatAppearance.BorderSize = 0;
            this.btnsheetName.FlatAppearance.CheckedBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnsheetName.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnsheetName.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(161)))), ((int)(((byte)(226)))));
            this.btnsheetName.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnsheetName.Font = new System.Drawing.Font("CorpoS", 9F);
            this.btnsheetName.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnsheetName.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnsheetName.Location = new System.Drawing.Point(27, 76);
            this.btnsheetName.Name = "btnsheetName";
            this.btnsheetName.Size = new System.Drawing.Size(82, 30);
            this.btnsheetName.TabIndex = 19;
            this.btnsheetName.Text = "Select Sheet";
            this.btnsheetName.UseVisualStyleBackColor = false;
            this.btnsheetName.Click += new System.EventHandler(this.btnsheetName_Click);
            // 
            // pbUpload
            // 
            this.pbUpload.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.pbUpload.Location = new System.Drawing.Point(3, 49);
            this.pbUpload.Name = "pbUpload";
            this.pbUpload.Size = new System.Drawing.Size(219, 16);
            this.pbUpload.TabIndex = 20;
            this.pbUpload.Visible = false;
            // 
            // frmuploadExpo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Desktop;
            this.ClientSize = new System.Drawing.Size(465, 511);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.ADGDATA);
            this.Controls.Add(this.btnclose);
            this.Controls.Add(this.btnupload);
            this.Controls.Add(this.btnselectfile);
            this.Controls.Add(this.txtuploadfilepath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmuploadExpo";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Load += new System.EventHandler(this.frmuploadExpo_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ADGDATA)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtuploadfilepath;
        private System.Windows.Forms.Button btnselectfile;
        private System.Windows.Forms.Button btnupload;
        private System.Windows.Forms.Button btnclose;
        public atcs.ADGV.AdvancedDataGridView ADGDATA;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnsheetName;
        private System.Windows.Forms.ComboBox ComSheetName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar pbUpload;
    }
}