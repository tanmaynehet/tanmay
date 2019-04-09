using INI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Collections;
using System.Windows.Forms;
using System.IO;

namespace cost_management
{
    public partial class frmuploadExpo : Form
    {
        public frmuploadExpo()
        {
            InitializeComponent();
        }
        string strDbpath = "";
        string str_Dbname = "";
        BindingSource bindingSource_main = null;
        INI.IniFile ini = new IniFile();
        DataTable dtdata = new DataTable();
        OleDbConnection dbCon;
        ClsCommonFunction clscon = new ClsCommonFunction();
        private void frmuploadExpo_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
            bindingSource_main = new BindingSource();
            bindingSource_main.DataSource = clscon.dtGetData("select * from tblextrapolationfactor");
            ADGDATA.DataSource = bindingSource_main;
            ADGDATA.DoubleBuffered(true);
            ADGDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            ADGDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        }

        private void btnselectfile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Title = "Please select a Excel file";
            ofd.FilterIndex = 2;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            ofd.Filter = "Excel 2007 (*.xlsx)|*xlsx";
            if (ofd.ShowDialog().Equals(DialogResult.OK))
            {
                txtuploadfilepath.Text = ofd.FileName.ToString();
                strDbpath = txtuploadfilepath.Text.Substring(0, txtuploadfilepath.Text.LastIndexOf("\\"));
                str_Dbname = Path.GetFileName(txtuploadfilepath.Text);
            }
            List<string> listSheet = clscon.ListSheetInExcel(strDbpath + "\\" + str_Dbname);
            if (listSheet.Count > 1)
            {
                ComSheetName.DataSource = listSheet;
                panel1.Visible = true;
                return;
            }
            else
            {
                DataTable dt = clscon.ImportDataUsingExcel(strDbpath + "\\" + str_Dbname, listSheet[0].ToString());
                bindingSource_main = new BindingSource();
                bindingSource_main.DataSource = dt;
                ADGDATA.DataSource = bindingSource_main;
            }
        }

        private void btnclose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void ADGDATA_FilterStringChanged(object sender, EventArgs e)
        {
            bindingSource_main.Filter = ADGDATA.FilterString;
            ADGDATA.DataSource = bindingSource_main;
            ADGDATA.DoubleBuffered(true);
        }

        private void ADGDATA_SortStringChanged(object sender, EventArgs e)
        {
            bindingSource_main.Sort = ADGDATA.SortString;
            ADGDATA.DataSource = bindingSource_main;
            ADGDATA.DoubleBuffered(true);
        }

        private void btnsheetName_Click(object sender, EventArgs e)
        {
            if (!ComSheetName.Text.Equals(""))
            {
                DataTable dt = clscon.ImportDataUsingExcel(strDbpath + "\\" + str_Dbname, ComSheetName.Text.Trim());
                bindingSource_main = new BindingSource();
                bindingSource_main.DataSource = dt;
                ADGDATA.DataSource = bindingSource_main;
            }
            else
            {

            }
            panel1.Visible = false;
        }

        private void btnupload_Click(object sender, EventArgs e)
        {
            int i = 0;
            panel1.Visible = true;
            btnsheetName.Visible = false;
            ComSheetName.Visible = false;
            label1.Visible = false;
            pbUpload.Visible = true;
            pbUpload.Minimum = 0;
            try
            {
                pbUpload.Minimum = 0;
                pbUpload.Maximum = ADGDATA.RowCount;
                StringBuilder sbSQL = new StringBuilder();
                dbCon = new OleDbConnection();
                dbCon = clscon.Openconnection();
                OleDbCommand oleCmd;
                if (ADGDATA.RowCount > 0)
                {
                    sbSQL = new StringBuilder();
                    try
                    {
                        sbSQL.Append("Drop table tblextrapolationfactor");
                        oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                        oleCmd.ExecuteNonQuery();
                        oleCmd.Dispose();
                    }
                    catch (Exception ex)
                    {

                    }
                    try
                    {
                        sbSQL = new StringBuilder();
                        sbSQL.Append(" Create Table tblextrapolationfactor ([Sparte] Text(10),[Monat] Integer,[Index] Text(15),HR_Faktor DECIMAL(18,2))");
                        oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                        oleCmd.ExecuteNonQuery();
                        oleCmd.Dispose();
                    }
                    catch (Exception ex) { }

                    foreach (DataGridViewRow drs in ADGDATA.Rows)
                    {
                        sbSQL = new StringBuilder();
                        sbSQL.Append("Insert Into tblextrapolationfactor Values ");
                        sbSQL.Append("('" + drs.Cells[0].Value.ToString() + "','" + drs.Cells[1].Value.ToString() + "',");
                        sbSQL.Append("'" + drs.Cells[2].Value.ToString() + "','" + Math.Round(Convert.ToDecimal(drs.Cells[3].Value.ToString()), 3) + "')");
                        dbCon = clscon.Openconnection();
                        oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                        oleCmd.ExecuteNonQuery();
                        oleCmd.Dispose();
                        i += 1;
                        pbUpload.Value = i;
                    }
                    MessageBox.Show("Data Uploaded sucessfully...!!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
            }
            catch (Exception ex) { }
            panel1.Visible = false;
        }
    }
}
