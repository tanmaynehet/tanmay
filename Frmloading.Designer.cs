﻿namespace cost_management
{
    partial class Frmloading
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
            this.PBWait = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.PBWait)).BeginInit();
            this.SuspendLayout();
            // 
            // PBWait
            // 
            this.PBWait.Image = global::cost_management.Properties.Resources.SneakyImpishGaur_max_1mb;
            this.PBWait.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.PBWait.Location = new System.Drawing.Point(-4, 5);
            this.PBWait.Name = "PBWait";
            this.PBWait.Size = new System.Drawing.Size(226, 238);
            this.PBWait.TabIndex = 14;
            this.PBWait.TabStop = false;
            // 
            // Frmloading
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(228, 244);
            this.ControlBox = false;
            this.Controls.Add(this.PBWait);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Frmloading";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Frmloading";
            this.Load += new System.EventHandler(this.Frmloading_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PBWait)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.PictureBox PBWait;
    }
}