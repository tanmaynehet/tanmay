using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using RDotNet;
using INI;
namespace cost_management
{
    public partial class frmDisplay : Form
    {
        INI.IniFile ini = new IniFile();
        DataTable dtdata = new DataTable();
        ClsCommonFunction clscon = new ClsCommonFunction();
        string sMarketCode = "";
        string sPartNo = "";
        string sSOBLc = "";
        string sSupplierName = "";
        string sSupplierCode = "";
        public frmDisplay(string PartNo)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            sPartNo = PartNo;
            InitializeComponent();
        }

        public void frmDisplay_Load(object sender, EventArgs e)
        {
            string sSQL = "Select * from GRID where [Part No]='" + sPartNo + "' order by [Date id] DESC";
            BindAdvData(clscon.dtGetData(sSQL));
            Application.DoEvents();
            ADGDATA.DoubleBuffered(true);
            ADGDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            ADGDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        }
        private void BindAdvData(DataTable dt)
        {
            try
            {
                if (dt.Rows.Count == 0)
                {
                    return;
                }
              
                dt.Columns.Add(new DataColumn("Status", typeof(Bitmap)));
                dt.Columns["Status"].SetOrdinal(0);
                setdata(dt);
                if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
                {
                    foreach (DataColumn drcol in dt.Columns)
                    {
                        string sNewLan = ini.IniReadValue("GridLables", drcol.ColumnName.ToString());
                        if (!sNewLan.Equals(""))
                            drcol.ColumnName = sNewLan;
                    }
                }
                bindingSource_main.DataSource = null;
                bindingSource_main.DataSource = dt;
                ADGDATA.DataSource = bindingSource_main;
                ADGDATA.AllowUserToOrderColumns = true;

                ADGDATA.AllowUserToOrderColumns = true;
                ADGDATA.Columns["sId"].Visible = false;
                ADGDATA.Columns["delta"].Visible = false;
                ADGDATA.Columns["delta1"].Visible = false;
                ADGDATA.Columns["delta2"].Visible = false;
                ADGDATA.Columns["delta3"].Visible = false;
                //ADGDATA.Columns["delta4"].Visible = false;
                //ADGDATA.Columns["delta5"].Visible = false;
                ADGDATA.Columns["FMCT1"].Visible = false;
                ADGDATA.Columns["FMCT2"].Visible = false;
                ADGDATA.Columns["FMCT3"].Visible = false;
                //ADGDATA.Columns["FMCT4"].Visible = false;
                //ADGDATA.Columns["FMCT5"].Visible = false;
                ADGDATA.Columns["delta-1"].Visible = false;
                ADGDATA.Columns["delta-2"].Visible = false;
                ADGDATA.Columns["Color_code"].Visible = false;
                 
                ADGDATA.Columns[0].Frozen = true;
                ADGDATA.Columns[1].Frozen = true;
                ADGDATA.Columns[2].Frozen = true;
                ADGDATA.Columns[3].Frozen = true;
                ADGDATA.Columns[4].Frozen = true;
                ADGDATA.Columns.GetFirstColumn(DataGridViewElementStates.Frozen);

                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {

            }
        }
        public void ChangeGridLangeuage()
        {
            if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
            {
                foreach (DataGridViewColumn drcol in ADGDATA.Columns)
                {
                    string sNewLan = ini.IniReadValue("GridLables", drcol.HeaderText.ToString());
                    if (!sNewLan.Equals(""))
                        drcol.HeaderText = sNewLan;
                }
            }
        }
        void setdata(DataTable dt)
        {
            Random r = new Random();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i]["Status"] = imgstatus.Images[r.Next(0, 4)];
            }
        }
        private void ADGDATA_SortStringChanged(object sender, EventArgs e)
        {
            bindingSource_main.Sort = ADGDATA.SortString;
            ADGDATA.DataSource = bindingSource_main;
            ADGDATA.DoubleBuffered(true);
        }

        private void ADVsearchbar_Search(object sender, atcs.ADGV.AdvancedDataGridViewSearchToolBarSearchEventArgs e)
        {
            bool restartsearch = true;
            int startColumn = 0;
            int startRow = 0;
            if (!e.FromBegin)
            {
                bool endcol = ADGDATA.CurrentCell.ColumnIndex + 1 >= ADGDATA.ColumnCount;
                bool endrow = ADGDATA.CurrentCell.RowIndex + 1 >= ADGDATA.RowCount;

                if (endcol && endrow)
                {
                    startColumn = ADGDATA.CurrentCell.ColumnIndex;
                    startRow = ADGDATA.CurrentCell.RowIndex;
                }
                else
                {
                    startColumn = endcol ? 0 : ADGDATA.CurrentCell.ColumnIndex + 1;
                    startRow = ADGDATA.CurrentCell.RowIndex + (endcol ? 1 : 0);
                }
            }
            DataGridViewCell c = ADGDATA.FindCell(
                e.ValueToSearch,
                e.ColumnToSearch != null ? e.ColumnToSearch.Name : null,
                startRow,
                startColumn,
                e.WholeWord,
                e.CaseSensitive);
            if (c == null && restartsearch)
                c = ADGDATA.FindCell(
                    e.ValueToSearch,
                    e.ColumnToSearch != null ? e.ColumnToSearch.Name : null,
                    0,
                    0,
                    e.WholeWord,
                    e.CaseSensitive);
            if (c != null)
            {
                ADGDATA.CurrentCell = c;
                ADGDATA.CurrentCell.Selected = true;
            }
        }

        private string CreateWhereClause()
        {
            StringBuilder sWherCaluse = new StringBuilder();

            if (!sMarketCode.Equals("All Marketing Code"))
            {
                sWherCaluse.Append(" GRID.[Marketing Code] ='" + sMarketCode.Trim() + "'");
            }

            if (!sPartNo.Equals("Part Number"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and GRID.[Part No] ='" + sPartNo.Trim() + "'");
                else
                    sWherCaluse.Append(" GRID.[Part No] ='" + sPartNo.Trim() + "'");
            }
            if (!sSupplierName.Equals("Supplier Number"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and  GRID.[Supplier Code] =" + sSupplierName.Trim() + "");
                else
                    sWherCaluse.Append(" GRID.[Supplier Code] =" + sSupplierName.Trim() + "");
            }
            if (!sSOBLc.Equals("SOBSL-Selection"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and GRID.[Part Sobsl] ='" + sSOBLc.Trim() + "'");
                else
                    sWherCaluse.Append(" GRID.[Part Sobsl] ='" + sSOBLc.Trim() + "'");
            }
            if (!sSupplierName.Equals("Supplier Name"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and  GRID.[Supplier Name] ='" + sSupplierName.Trim() + "'");
                else
                    sWherCaluse.Append(" GRID.[Supplier Name]='" + sSupplierName.Trim() + "'");
            }

            return sWherCaluse.ToString();
        }

        private DataTable GetsqlqueryPagination1(string sWherclause)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("Select  * from GRID ");
            sbSQL.Append(sWherclause.ToString().Equals("") ? " " : " where " + sWherclause);
            sbSQL.Append(" order by GRID.sId desc,GRID.[Date ID] , GRID.[Marketing Code],GRID.[Part No]");
            DataTable dtdata = new DataTable();
            dtdata = clscon.dtGetData(sbSQL.ToString());
            return dtdata;
        }

        private void ADVsearchbar_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void ADVsearchbar_Search_1(object sender, atcs.ADGV.AdvancedDataGridViewSearchToolBarSearchEventArgs e)
        {
            ADVsearchbar_Search(sender, e);
        }

        private void btnexport_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            Btnclose.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(27, 161, 226);
            DataTable dt = new DataTable();
            if (ADGDATA.FilterString.Equals(""))
            {
                dt = (DataTable)bindingSource_main.DataSource;
            }
            else
            {
                dt = (DataTable)bindingSource_main.DataSource;
                DataView view = new DataView(dtdata);
                view.RowFilter = ADGDATA.FilterString;
                dt = view.ToTable();
            }
            clscon.export(dt);

            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            this.Cursor = Cursors.Default;
        }

        private void Btnclose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
    public class Page
    {
        public string Text { get; set; }
        public string Value { get; set; }
        public bool Selected { get; set; }
    }
    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
