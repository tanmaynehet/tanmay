using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Data.OleDb;
using INI;
using System.Globalization;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace cost_management
{
    public partial class MidHome : Form
    {
        ClsCommonFunction clsFun = new ClsCommonFunction();
        ComboBox comboBox = new ComboBox();
        IniFile ini = new IniFile();
        public MidHome()
        {
            Application.DoEvents();
            InitializeComponent();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;
            Resolution objFormResizer = new Resolution();
            objFormResizer.ResizeForm(this, screenHeight, screenWidth);
        }
        MdiClient ctlMDI;
        private void MidHome_Load(object sender, EventArgs e)
        {

            foreach (Control ctl in this.Controls)
            {
                try
                {

                    // Attempt to cast the control to type MdiClient.
                    ctlMDI = (MdiClient)ctl;

                    // Set the BackColor of the MdiClient control. 
                    ctlMDI.BackColor = this.BackColor;
                }
                catch (InvalidCastException exc)
                {
                    string smsg = exc.Message.ToString();
                }
            }
            if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
            {
                TSMLNG.Text = ini.IniReadValue("Literls", "TSMLNG.Text");
                exitToolStripMenuItem.Text = ini.IniReadValue("Literls", "exitToolStripMenuItem.Text");
                kPIToolStripMenuItem.Text = ini.IniReadValue("Literls", "kPIToolStripMenuItem.Text");
                partSummeryToolStripMenuItem.Text = ini.IniReadValue("Literls", "partSummeryToolStripMenuItem.Text ");
                TSMRefresh.Text = ini.IniReadValue("Literls", "refresh");
            }
            else
            {
                TSMLNG.Text = "Language";
                exitToolStripMenuItem.Text = "Exit";
                kPIToolStripMenuItem.Text = "KPI";
                partSummeryToolStripMenuItem.Text = "Part Summary";
                TSMRefresh.Text = "Refresh";
            }

            kPIToolStripMenuItem_Click(sender, e);
            this.Cursor = Cursors.Default;
        }
        public void DisposeAllInActiveForms()
        {
            foreach (Form frm in this.MdiChildren)
            {
                if (!frm.Focused)
                {
                    frm.Visible = false;
                    frm.Dispose();
                }
            }

        }
        private void displayDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisposeAllInActiveForms();
            Application.DoEvents();
            frmDisplayReports frmd = new frmDisplayReports();
            frmd.MdiParent = this;
            frmd.Dock = DockStyle.Fill;
            frmd.FormBorderStyle = FormBorderStyle.None;
            frmd.Show();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void partSummeryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            if (GlobalVariable.dtCompleteData.Rows.Count > 0)
            {
                GlobalVariable.iTotalCOunt = GlobalVariable.dtCompleteData.Rows.Count;
            }
            else
            {
                try
                {
                    Loaddata();
                    GlobalVariable.iTotalCOunt = GlobalVariable.dtCompleteData.Rows.Count;
                
                }
                catch (Exception ex)
                {
                    GlobalVariable.iTotalCOunt = 0;
                }
            }
            if (GlobalVariable.iTotalCOunt > 0)
            {
                DisposeAllInActiveForms();
                Application.DoEvents();
                frmDisplayReportsTest frmd = new frmDisplayReportsTest();
                frmd.MdiParent = this;
                frmd.FormBorderStyle = FormBorderStyle.None;
                frmd.Dock = DockStyle.Fill;

                frmd.Show();
            }
            else
            {
                MessageBox.Show("Please click on refresh button to create data.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            this.Cursor = Cursors.Default;
        }

        private void kPIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                GlobalVariable.iTotalCOunt = Convert.ToInt32(clsFun.dtGetData("Select  * from tblRecoredCount").Rows[0][0].ToString());
            }
            catch (Exception ex)
            {
                GlobalVariable.iTotalCOunt = 0;
            }
            if (GlobalVariable.iTotalCOunt > 0)
            {
                DisposeAllInActiveForms();
                Application.DoEvents();
                FrmdataForcasting frmd = new FrmdataForcasting();
                frmd.MdiParent = this;
                frmd.FormBorderStyle = FormBorderStyle.None;
                frmd.Dock = DockStyle.Fill;

                frmd.Show();
            }
            else
            {
                MessageBox.Show("Please click on refresh button to create data.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }


        private void TSMRefresh_Click(object sender, EventArgs e)
        {
            string sDisplayMsg = "";
            if (ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
                sDisplayMsg = ini.IniReadValue("Messages", "en");
            else
                sDisplayMsg = ini.IniReadValue("Messages", "de");

            DialogResult result = MessageBox.Show(sDisplayMsg, "Cost Management", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result.Equals(DialogResult.No))
            {
                return;
            }
            else
            {
                DisposeAllInActiveForms();
                Application.DoEvents();
                frmWait frmd = new frmWait();
                frmd.ShowIcon = false;
                frmd.ShowInTaskbar = false;
                frmd.ShowDialog();
            }
        }

        private void MidHome_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void TSMLNG_Click(object sender, EventArgs e)
        {
            frmcHangeLanguage frmd = new frmcHangeLanguage();
            frmd.ShowIcon = false;
            frmd.ShowInTaskbar = false;
            if (frmd.ShowDialog().Equals(DialogResult.OK))
            {
                if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
                {
                    TSMLNG.Text = ini.IniReadValue("Literls", "TSMLNG.Text");
                    exitToolStripMenuItem.Text = ini.IniReadValue("Literls", "exitToolStripMenuItem.Text");
                    kPIToolStripMenuItem.Text = ini.IniReadValue("Literls", "kPIToolStripMenuItem.Text");
                    partSummeryToolStripMenuItem.Text = ini.IniReadValue("Literls", "partSummeryToolStripMenuItem.Text ");
                    TSMRefresh.Text = ini.IniReadValue("Literls", "refresh");
                }
                else
                {
                    TSMLNG.Text = "Language";
                    exitToolStripMenuItem.Text = "Exit";
                    kPIToolStripMenuItem.Text = "KPI";
                    partSummeryToolStripMenuItem.Text = "Part Summary";
                    TSMRefresh.Text = "Refresh";
                }

                foreach (Form frm in this.MdiChildren)
                {
                    if (frm.Name.Equals("frmDisplayReports"))
                    {
                        partSummeryToolStripMenuItem_Click(sender, e);
                    }
                }
            }
        }

        private void MidHome_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.F7))
            {
                DisposeAllInActiveForms();
                Application.DoEvents();
                frmuploadExpo frmd = new frmuploadExpo();
                frmd.ShowIcon = false;
                frmd.ShowInTaskbar = false;
                frmd.ShowDialog();
            }
        }
        void Loaddata()
        {
            GlobalVariable.dtCompleteData = clsFun.dtGetData("SELECT g.* from GRID g inner Join TblDiscSNR ds on g.[Date id]=ds.[date_id] and g.[Part No]=ds.[Part_No] order by g.[Date id] desc ;");
        }

        private void extraPolationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DisposeAllInActiveForms();
            Application.DoEvents();
            frmuploadExpo frmd = new frmuploadExpo();
            frmd.ShowIcon = false;
            frmd.ShowInTaskbar = false;
            frmd.ShowDialog();
        }
    }
    public class FlatCombo : ComboBox
    {
        private const int WM_PAINT = 0xF;
        private int buttonWidth = SystemInformation.HorizontalScrollBarArrowWidth;
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_PAINT)
            {
                using (var g = Graphics.FromHwnd(Handle))
                {
                    using (var p = new Pen(this.ForeColor))
                    {
                        g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
                        g.DrawLine(p, Width - buttonWidth, 0, Width - buttonWidth, Height);
                    }
                }
            }
        }
    }
}
