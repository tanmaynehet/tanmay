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
    public partial class frmDisplayReports : Form
    {
        INI.IniFile ini = new IniFile();
        DataTable dtdata = new DataTable();
        ClsCommonFunction clscon = new ClsCommonFunction();
        int iCurrentButtonInddex = 0;
        int PageSize = 500;
        string sBtnIdex = "";
        string sBtnName = "";
        string sSearchvalue = "";
        string sSortingOrder = "";
        string sColumnToSort = "";
        bool bFlag = false;
        private SortedDictionary<int, string> _filtersaved = new SortedDictionary<int, string>();
        private SortedDictionary<int, string> _sortsaved = new SortedDictionary<int, string>();
        public frmDisplayReports()
        {
            InitializeComponent();
            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            int screenHeight = Screen.PrimaryScreen.Bounds.Height;
            Resolution objFormResizer = new Resolution();
            objFormResizer.ResizeForm(this, screenHeight, screenWidth);
            ComSortingFields.Items.Add("Prognosed Series End Date");
            ComSortingFields.Items.Add("Delta Manufacturing cost current year vs " + clscon.GetYearName(1, "-") + "");
            ComSortingFields.Items.Add("Delta manufacturing cost forecast " + clscon.GetYearName(1, "+") + " vs " + clscon.GetYearName(1, "") + "");
            ComSortingFields.Items.Add("Delta manufacturing cost forecast " + clscon.GetYearName(2, "+") + " vs " + clscon.GetYearName(1, "+") + "");
            ComSortingFields.Items.Add("Delta manufacturing cost forecast " + clscon.GetYearName(3, "+") + " vs " + clscon.GetYearName(2, "+") + "");
            ComSortingFields.Items.Add("Delta manufacturing cost forecast " + clscon.GetYearName(4, "+") + " vs " + clscon.GetYearName(3, "+") + "");
            ComSortingFields.Items.Add("Delta manufacturing cost forecast " + clscon.GetYearName(5, "+") + " vs " + clscon.GetYearName(4, "+") + "");
        }

        public void frmDisplay_Load(object sender, EventArgs e)
        {
            Application.DoEvents();

            int iVariable = (GlobalVariable.iTotalCOunt - 500) + 1;
            PopulatePager(GlobalVariable.iTotalCOunt, 1);
            dtdata = (GetsqlqueryPagination1(iVariable, GlobalVariable.iTotalCOunt));
            BindAdvData(dtdata);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            this.WindowState = FormWindowState.Maximized;
            lbltoolver.Text = "Tool v " + clscon.GetAPPVersion();
            Btnall.Focus();
            Application.DoEvents();
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.Visible = true;
            btnpc.Visible = true;
            btnvans.Visible = true;
            ADGDATA.DoubleBuffered(true);
            ADGDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
            ADGDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        }

        private void PopulatePager(int recordCount, int currentPage)
        {
            List<Page> pages = new List<Page>();
            int startIndex, endIndex;
            int pagerSpan = 500;
            //Calculate the Start and End Index of pages to be displayed.
            double dblPageCount = (double)((decimal)recordCount / Convert.ToDecimal(PageSize));
            int pageCount = (int)Math.Ceiling(dblPageCount);
            startIndex = currentPage > 1 && currentPage + pagerSpan - 1 < pagerSpan ? currentPage : 1;
            endIndex = pageCount > pagerSpan ? pagerSpan : pageCount;
            if (currentPage > pagerSpan % 2)
            {
                if (currentPage == 2)
                {
                    endIndex = pageCount;
                }
            }
            else
            {
                endIndex = (pagerSpan - currentPage) + 1;
            }

            if (endIndex - (pagerSpan - 1) > startIndex)
            {
                startIndex = endIndex - (pagerSpan - 1);
            }

            if (endIndex > pageCount)
            {
                endIndex = pageCount;
                startIndex = ((endIndex - pagerSpan) + 1) > 0 ? (endIndex - pagerSpan) + 1 : 1;
            }

            for (int i = startIndex; i <= endIndex; i++)
            {
                pages.Add(new Page { Text = i.ToString(), Value = i.ToString(), Selected = i == currentPage });
            }

            //Clear existing Pager Buttons.
            pnlPager.Controls.Clear();

            //Loop and add Buttons for Pager.
            int count = 0;
            foreach (Page page in pages)
            {
                Button btnPage = new Button();
                btnPage.Location = new System.Drawing.Point(38 * count, 5);
                btnPage.Size = new System.Drawing.Size(35, 25);
                btnPage.ForeColor = Color.Wheat;

                btnPage.Name = page.Value;
                btnPage.Text = page.Text;
                //btnPage.Enabled = !page.Selected;
                btnPage.Click += new System.EventHandler(this.Page_Click);
                pnlPager.Controls.Add(btnPage);
                count++;
            }
        }


        private void BindAdvData(DataTable dt)
        {
            try
            {
                if (dt.Rows.Count.Equals(0))
                {
                    return;
                }
                if (!dt.Columns[0].ColumnName.Equals("Status"))
                {
                    dt.Columns.Add(new DataColumn("Status", typeof(Bitmap)));
                    dt.Columns["Status"].SetOrdinal(0);
                    setdata(dt);
                }
                if (!ini.IniReadValue("CultureInfo", "Language").ToString().Equals("en"))
                {
                    foreach (DataColumn drcol in dt.Columns)
                    {
                        string sNewLan = ini.IniReadValue("GridLables", drcol.ColumnName.ToString());
                        if (!sNewLan.Equals(""))
                            drcol.ColumnName = sNewLan;
                    }
                }
                bindingSource_main = new BindingSource();
                bindingSource_main.DataSource = dt;
                ADGDATA.DataSource = bindingSource_main;
                ADGDATA.AllowUserToOrderColumns = true;
                ADGDATA.Columns["sId"].Visible = false;
                ADGDATA.Columns["delta"].Visible = false;
                ADGDATA.Columns["delta1"].Visible = false;
                ADGDATA.Columns["delta2"].Visible = false;
                ADGDATA.Columns["delta3"].Visible = false;
                ADGDATA.Columns["delta4"].Visible = false;
                ADGDATA.Columns["delta5"].Visible = false;
                ADGDATA.Columns["FMCT1"].Visible = false;
                ADGDATA.Columns["FMCT2"].Visible = false;
                ADGDATA.Columns["FMCT3"].Visible = false;
                ADGDATA.Columns["FMCT4"].Visible = false;
                ADGDATA.Columns["FMCT5"].Visible = false;
                ADGDATA.Columns["delta-1"].Visible = false;
                ADGDATA.Columns["delta-2"].Visible = false;
                ADGDATA.Columns["Color_code"].Visible = false;
                ADVsearchbar.SetColumns(ADGDATA.Columns);

                ADGDATA.Columns[0].Frozen = true;
                ADGDATA.Columns[1].Frozen = true;
                ADGDATA.Columns[2].Frozen = true;
                ADGDATA.Columns[3].Frozen = true;
                ADGDATA.Columns[4].Frozen = true;
                ADGDATA.Columns.GetFirstColumn(DataGridViewElementStates.Frozen);
            }
            catch (Exception ex)
            {

            }
            this.Cursor = Cursors.Default;
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
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (!dt.Rows[i]["Color_code"].ToString().Equals(""))
                    dt.Rows[i]["Status"] = imgstatus.Images[Convert.ToInt32(dt.Rows[i]["Color_code"].ToString())];
                else
                    dt.Rows[i]["Status"] = imgstatus.Images[3];
            }
        }

        private void btnpc_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            ClearsearchFilter();
            btnpc.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            bFlag = true;
            sBtnName = btnpc.Name.ToString();
            BindAdvData(GetPagingData(Convert.ToInt32(sBtnIdex.Equals("") ? "1" : sBtnIdex), sBtnName));
            this.Cursor = Cursors.Default;
        }

        private void btnvans_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            ClearsearchFilter();
            btnvans.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            bFlag = true;
            sBtnName = btnvans.Name.ToString();
            BindAdvData(GetPagingData(Convert.ToInt32(sBtnIdex.Equals("") ? "1" : sBtnIdex), sBtnName));
            this.Cursor = Cursors.Default;
        }

        private void btntruks_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            ClearsearchFilter();
            btntruks.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            bFlag = true;
            sBtnName = btntruks.Name.ToString();
            BindAdvData(GetPagingData(Convert.ToInt32(sBtnIdex.Equals("") ? "1" : sBtnIdex), sBtnName));
            this.Cursor = Cursors.Default;
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

        private void ADVsearchbar_Search(object sender, atcs.ADGV.AdvancedDataGridViewSearchToolBarSearchEventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            try
            {
                sSearchvalue = "[" + getSearchColumnValueForSoring(e.sSearchfieldName) + "] like '%" + e.ValueToSearch + "%'";
                if (!e.sSearchfieldName.Equals("(Select columns)"))
                {
                    StringBuilder sbSQL = new StringBuilder();
                    sbSQL.Append("Select * from GRID where [");
                    if (Btnall.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sbSQL.Append(getSearchColumnValueForSoring(e.sSearchfieldName) + "] like '%" + e.ValueToSearch + "%'");
                    }
                    else if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sbSQL.Append(getSearchColumnValueForSoring(e.sSearchfieldName) + "] like '%" + e.ValueToSearch + "%' and (GRID.[Marketing Code] like '1%' or GRID.[Marketing Code] like 'SP%')");
                    }
                    else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sbSQL.Append(getSearchColumnValueForSoring(e.sSearchfieldName) + "] like '%" + e.ValueToSearch + "%' and GRID.[Marketing Code] like '2T%'");
                    }
                    else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sbSQL.Append(e.sSearchfieldName + "] like '%" + e.ValueToSearch + "%' and (GRID.[Marketing Code] like '2U%'  or  GRID.[Marketing Code] like '2L%')");
                    }
                    else
                    {
                        sbSQL.Append(getSearchColumnValueForSoring(e.sSearchfieldName) + "] like '%" + e.ValueToSearch + "%'");
                    }

                    if (!sSortingOrder.Equals(""))
                    {
                        if (!sColumnToSort.Equals(""))
                        {
                            sbSQL.Append(" order By GRID." + sColumnToSort + " " + sSortingOrder);
                        }
                        else
                        {
                            sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                        }
                    }
                    DataTable dtdata = clscon.dtGetData(sbSQL.ToString());
                    if (dtdata.Rows.Count > 0)
                    {
                        if (dtdata.Rows.Count > 500)
                        {
                            PopulatePager(dtdata.Rows.Count, 1);
                        }
                        else
                        {
                            PopulatePager(0, 1);
                        }
                        BindAdvData((dtdata));
                    }
                    else
                    {
                        MessageBox.Show("No data found for search..!!", "Cost Management ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show("Please select column to search..!!", "Cost Management ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                sSearchvalue = "";
            }
            this.Cursor = Cursors.Default;
        }



        private void Btnall_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            ClearsearchFilter();
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            bFlag = false;
            sBtnName = "Btnall";
            int iVariable = (GlobalVariable.iTotalCOunt - 500) + 1;
            dtdata = (GetsqlqueryPagination(iVariable, GlobalVariable.iTotalCOunt, 1));
            if (dtdata.Rows.Count >= 500)
            {
                PopulatePager(GlobalVariable.iTotalCOunt, 1);
            }
            else
            {
                PopulatePager(0, 1);
            }
            BindAdvData((dtdata));
            this.Cursor = Cursors.Default;
        }

        private DataTable GetPagingData(int iSatringIdex, string btnname)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            StringBuilder sSQL = new StringBuilder();
            if (btnname.Equals("btntruks"))
            {
                sSQL.Append("Select  * from GRID   Where (GRID.[Marketing Code] like '2U%'  or  GRID.[Marketing Code] like '2L%') ");
            }
            else if (btnname.Equals("btnvans"))
            {
                sSQL.Append("Select  * from GRID  Where GRID.[Marketing Code] like '2T%'");
            }
            else if (btnname.Equals("btnpc"))
            {
                sSQL.Append("Select  * from GRID  Where (GRID.[Marketing Code] like '1%' or GRID.[Marketing Code] like 'SP%')  ");
            }
            if (!sSearchvalue.Equals(""))
            {
                sSQL.Append(sSearchvalue.Equals("") ? "" : " and " + sSearchvalue);
                if (!sSortingOrder.Equals(""))
                {
                    if (!sColumnToSort.Equals(""))
                    {
                        sSQL.Append(" order By  Grid." + sColumnToSort + " " + sSortingOrder);
                    }
                    else
                    {
                        sSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
                else
                {
                    sSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                }
            }
            else
            {
                if (!sSortingOrder.Equals(""))
                {
                    if (!sColumnToSort.Equals(""))
                    {
                        sSQL.Append(" order By  Grid." + sColumnToSort + " " + sSortingOrder);
                    }
                    else
                    {
                        sSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
                else
                {
                    sSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                }
            }
            dtdata = clscon.dtGetData(sSQL.ToString());
            DataTable dt = new DataTable();
            if (dtdata.Rows.Count > 0)
            {
                try
                {
                    if (btnname.Equals("btntruks"))
                    {
                        int ival = dtdata.Rows.Count / iSatringIdex;
                        if (ival < 500)
                            dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((iSatringIdex - 1) * ival).Take(ival).CopyToDataTable();
                        else
                            dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((iSatringIdex - 1) * 500).Take(500).CopyToDataTable();
                    }
                    else
                    {
                        int ival = dtdata.Rows.Count / iSatringIdex;
                        if (ival < 500)
                            dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((iSatringIdex - 1) * ival).Take(ival).CopyToDataTable();
                        else
                            dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((iSatringIdex - 1) * 500).Take(500).CopyToDataTable();
                    }
                }
                catch (Exception ex) { }

                PopulatePager(dtdata.Rows.Count, Convert.ToInt32(iSatringIdex));
                return (dt);
            }
            else
            {
                MessageBox.Show("No data found for Above search Criteria", "Cost Management ", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return dt;
            }
        }

        private DataTable GetsqlqueryPagination(int iSIndex, int iEIndex, int iPageNo)
        {
            StringBuilder sbSQL = new StringBuilder();
            int iIndex = GlobalVariable.iTotalCOunt - iSIndex;
            int iEnIndex = iIndex + 500;
            sbSQL.Append("Select * from GRID ");
            if (!sSearchvalue.Equals(""))
            {
                sbSQL.Append(sSearchvalue.Equals("") ? "" : " WHERE " + sSearchvalue);
                if (!sSortingOrder.Equals(""))
                {
                    if (!sColumnToSort.Equals(""))
                    {
                        sbSQL.Append(" order By  GRID." + sColumnToSort + " " + sSortingOrder);
                    }
                    else
                    {
                        sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
                else
                {
                    if (!sSortingOrder.Equals(""))
                    {
                        if (!sColumnToSort.Equals(""))
                        {
                            sbSQL.Append(" order By Grid." + sColumnToSort + " " + sSortingOrder);
                        }
                        else
                        {
                            sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                        }
                    }
                    else
                    {
                        sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
            }
            else
            {
                //
                if (!sSortingOrder.Equals(""))
                {
                    if (!sColumnToSort.Equals(""))
                    {
                        sbSQL.Append(" order By Grid." + sColumnToSort + " " + sSortingOrder);
                    }
                    else
                    {
                        sbSQL.Append(" WHERE GRID.sid >= " + iIndex + " and GRID.sId <= " + iEnIndex + "");
                        sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
                else
                {
                    sbSQL.Append(" WHERE GRID.sid >= " + iIndex + " and GRID.sId <= " + iEnIndex + "");
                    sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                }

            }
            DataTable dtdata = new DataTable();
            dtdata = clscon.dtGetData(sbSQL.ToString());
            DataTable dt = new DataTable();
            int ival = dtdata.Rows.Count;
            if (ival < 500)
                dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((iPageNo - 1) * ival).Take(ival).CopyToDataTable();
            else
                dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((1 - 1) * 500).Take(500).CopyToDataTable();
            return dt;
        }
        private DataTable GetsqlqueryPagination1(int iSatringIdex, int iEndIndex)
        {
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("Select * from GRID  WHERE GRID.sid>=" + (iSatringIdex) + " and GRID.sId<=" + iEndIndex + "  ");
            if (!sSearchvalue.Equals(""))
            {
                sbSQL.Append(sSearchvalue.Equals("") ? "" : " and " + sSearchvalue);
                if (!sSortingOrder.Equals(""))
                {
                    if (!sColumnToSort.Equals(""))
                    {
                        sbSQL.Append(" order By " + sColumnToSort + " " + sSortingOrder);
                    }
                    else
                    {
                        sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                    }
                }
                else
                {
                    sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
                }
            }
            else
            {
                sbSQL.Append(" order by GRID.sid desc,GRID.[Date ID], GRID.[Marketing Code],GRID.[Part No] desc");
            }
            DataTable dtdata = new DataTable();
            dtdata = clscon.dtGetData(sbSQL.ToString());
            DataTable dt = dtdata.Rows.Cast<System.Data.DataRow>().Skip((1 - 1) * 500).Take(500).CopyToDataTable();
            return dt;
        }
        private string getSearchColumnValueForSoring(string sVal)
        {
            string sValue = "";
            try
            {
                switch (sVal.Trim())
                {
                    case "Delta Manufacturing cost current year vs 2017":
                        sValue = "delta";
                        break;
                    case "Forecast Manufacturing cost total 2019":
                        sValue = "FMCT1";
                        break;
                    case "Forecast Manufacturing cost total 2020":
                        sValue = "FMCT2";
                        break;
                    case "Forecast Manufacturing cost total 2021":
                        sValue = "FMCT3";
                        break;
                    case "Forecast Manufacturing cost total 2022":
                        sValue = "FMCT4";
                        break;
                    case "Forecast Manufacturing cost total 2023":
                        sValue = "FMCT5";
                        break;
                    case "Delta manufacturing cost forecast 2019 vs 2018":
                        sValue = "Delta1";
                        break;
                    case "Delta manufacturing cost forecast 2020 vs 2019":
                        sValue = "Delta2";
                        break;
                    case "Delta manufacturing cost forecast 2021 vs 2020":
                        sValue = "Delta3";
                        break;
                    case "Delta manufacturing cost forecast 2022 vs 2021":
                        sValue = "Delta4";
                        break;
                    case "Delta manufacturing cost forecast 2023 vs 2022":
                        sValue = "Delta5";
                        break;
                    default:
                        sValue = sVal;
                        break;
                }
            }
            catch (Exception ex)
            {
            }

            return sValue;
        }

        private void ADVsearchbar_Search_1(object sender, atcs.ADGV.AdvancedDataGridViewSearchToolBarSearchEventArgs e)
        {
            ADVsearchbar_Search(sender, e);
        }

        private void btnexport_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(27, 161, 226);
            DataTable dt = new DataTable();
            if (ADGDATA.FilterString.Equals(""))
            {
                dt = (DataTable)bindingSource_main.DataSource;
                clscon.export(dt);
            }
            else
            {
                dt = (DataTable)bindingSource_main.DataSource;
                //dt.DefaultView.RowFilter = ADGDATA.FilterString;
                DataView view = new DataView(dt);
                view.RowFilter = ADGDATA.FilterString;
                dt = view.ToTable();
                clscon.export(dt);
            }

            if (sBtnName.Equals("btntruks"))
            {
                btntruks.BackColor = Color.FromArgb(27, 161, 226);
            }
            else if (sBtnName.Equals("btnvans"))
            {
                btnvans.BackColor = Color.FromArgb(27, 161, 226);
            }
            else if (sBtnName.Equals("btnpc"))
            {
                btnpc.BackColor = Color.FromArgb(27, 161, 226);
            }
            else
            {
                Btnall.BackColor = Color.FromArgb(27, 161, 226);
            }
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            this.Cursor = Cursors.Default;
        }
        private void Page_Click(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            Button btnPager = (sender as Button);
            sBtnIdex = btnPager.Text;
            if (!bFlag)
            {
                if (!btnPager.Text.Equals("1"))
                {
                    int iLastpage = Convert.ToInt32(btnPager.Text);
                    int iCurrentPage = Convert.ToInt32(btnPager.Text);
                    iCurrentButtonInddex = iCurrentPage;
                    int iLastIndex = iCurrentPage * 500;
                    int iStaringIndex = iCurrentPage * 500;
                    dtdata = GetsqlqueryPagination(iStaringIndex, 0, Convert.ToInt32(btnPager.Text));
                    BindAdvData(GETSortedData(dtdata));
                }
                else
                {
                    int iVariable = (GlobalVariable.iTotalCOunt - 500) + 1;
                    dtdata = (GetsqlqueryPagination1(iVariable, GlobalVariable.iTotalCOunt));
                    BindAdvData(GETSortedData(dtdata));
                }
            }
            else
            {
                int iPageNo = Convert.ToInt32(btnPager.Text);
                BindAdvData(GetPagingData(Convert.ToInt32(sBtnIdex.Equals("") ? "1" : sBtnIdex), sBtnName));
            }
            this.Cursor = Cursors.Default;
        }

        private void ADGDATA_SortStringChanged_1(object sender, EventArgs e)
        {
            ADGDATA_SortStringChanged(sender, e);
        }

        private void ADVsearchbar_Reset(object sender, atcs.ADGV.AdvancedDataGridViewResetEventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            ComSortingFields.SelectedIndex = -1;
            ComSortingOrder.SelectedIndex = -1;
            sSearchvalue = "";
            ADGDATA.CleanFilterAndSort();
            ClearsearchFilter();
            sColumnToSort = "";
            sSortingOrder = "";
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            bFlag = false;
            sBtnName = "Btnall";

            PopulatePager(GlobalVariable.iTotalCOunt, 1);
            int iVariable = (GlobalVariable.iTotalCOunt - 500) + 1;
            dtdata = (GetsqlqueryPagination1(iVariable, GlobalVariable.iTotalCOunt));
            BindAdvData(dtdata);
            this.Cursor = Cursors.Default;
        }

        private void ComSortingOrder_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            try
            {
                if (!ComSortingOrder.SelectedIndex.Equals(-1))
                {
                    if (!ComSortingFields.Text.Equals(""))
                    {
                        ClearsearchFilter();
                        sSortingOrder = ComSortingOrder.Text.Equals("Ascending") ? "Asc" : "Desc";
                        sColumnToSort = "[" + getSearchColumnValueForSoring(ComSortingFields.Text) + "]";
                        if (Btnall.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            Btnall_Click(sender, e);
                        }
                        else if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btnpc_Click(sender, e);
                        }
                        else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btnvans_Click(sender, e);
                        }
                        else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btntruks_Click(sender, e);
                        }
                        else
                        {
                            Btnall_Click(sender, e);
                        }
                    }
                }
            }
            catch (Exception ex) { }
            this.Cursor = Cursors.Default;
        }

        private void ComSortingFields_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            try
            {
                if (!ComSortingOrder.SelectedIndex.Equals(-1))
                {
                    if (!ComSortingFields.Text.Equals(""))
                    {
                        ClearsearchFilter();
                        sSortingOrder = ComSortingOrder.Text.Equals("Ascending") ? "Asc" : "Desc";
                        sColumnToSort = "[" + getSearchColumnValueForSoring(ComSortingFields.Text) + "]";
                        if (Btnall.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            Btnall_Click(sender, e);
                        }
                        else if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btnpc_Click(sender, e);
                        }
                        else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btnvans_Click(sender, e);
                        }
                        else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        {
                            btntruks_Click(sender, e);
                        }
                        else
                        {
                            Btnall_Click(sender, e);
                        }
                    }
                }
            }
            catch (Exception ex) { }
            this.Cursor = Cursors.Default;
        }
        void ClearsearchFilter()
        {
            if (ADVsearchbar.textBox_search.Text.Equals(""))
            {
                sSearchvalue = "";
            }

            if (ADVsearchbar.textBox_search.Text.Equals("Value for Search"))
            {
                sSearchvalue = "";
            }
        }
        // Direcly sortthe GRID Data
        private void ChnagetheOrderingFoData()
        {
            DataTable dt = new DataTable();
            try
            {
                if (!ComSortingOrder.SelectedIndex.Equals(-1))
                {
                    if (!ComSortingFields.Text.Equals(""))
                    {
                        string sSortingOrder = ComSortingOrder.Text.Equals("Ascending") ? "Asc" : "Desc";
                        dt = (DataTable)bindingSource_main.DataSource;
                        int sColumnToSort = dt.Columns[ComSortingFields.Text].Ordinal;
                        if (sSortingOrder.Equals("Desc"))
                        {
                            if (ComSortingFields.Text.Equals("Prognosed Series End Date"))
                            {
                                var Rows = (from row in dt.AsEnumerable()
                                            orderby row[sColumnToSort] descending
                                            select row);
                                dt = Rows.AsDataView().ToTable();
                            }
                            else
                            {
                                var Rows = (from row in dt.AsEnumerable()
                                            orderby Convert.ToDecimal(row[sColumnToSort]) descending
                                            select row);
                                dt = Rows.AsDataView().ToTable();
                            }
                            BindAdvData(dt);
                        }
                        else
                        {
                            if (ComSortingFields.Text.Equals("Prognosed Series End Date"))
                            {
                                var Rows = (from row in dt.AsEnumerable()
                                            orderby row[sColumnToSort] ascending
                                            select row);
                                dt = Rows.AsDataView().ToTable();
                            }
                            else
                            {
                                var Rows = (from row in dt.AsEnumerable()
                                            orderby Convert.ToDecimal(row[sColumnToSort]) ascending
                                            select row);
                                dt = Rows.AsDataView().ToTable();
                            }
                            BindAdvData(dt);
                        }
                    }
                }
            }
            catch (Exception ex) { }
        }
        //Sorting data while binding to Gridview
        private DataTable GETSortedData(DataTable dtdataforsorting)
        {
            DataTable dtsortedData = new DataTable();
            DataView view = null;
            if (!ComSortingOrder.SelectedIndex.Equals(-1))
            {
                if (!ComSortingFields.Text.Equals(""))
                {
                    string sSortingOrder = ComSortingOrder.Text.Equals("Ascending") ? "Asc" : "Desc";
                    int sColumnToSort = dtdataforsorting.Columns[ComSortingFields.Text].Ordinal;
                    if (sSortingOrder.Equals("Desc"))
                    {
                        if (ComSortingFields.Text.Equals("Prognosed Series End Date"))
                        {
                            var Rows = (from row in dtdataforsorting.AsEnumerable()
                                        orderby row[sColumnToSort] descending
                                        select row);
                            dtsortedData = Rows.AsDataView().ToTable();
                        }
                        else
                        {
                            var Rows = (from row in dtdataforsorting.AsEnumerable()
                                        orderby Convert.ToDecimal(row[sColumnToSort]) descending
                                        select row);
                            dtsortedData = Rows.AsDataView().ToTable();
                        }
                    }
                    else
                    {
                        if (ComSortingFields.Text.Equals("Prognosed Series End Date"))
                        {
                            var Rows = (from row in dtdataforsorting.AsEnumerable()
                                        orderby row[sColumnToSort] ascending
                                        select row);
                            dtsortedData = Rows.AsDataView().ToTable();
                        }
                        else
                        {
                            var Rows = (from row in dtdataforsorting.AsEnumerable()
                                        orderby Convert.ToDecimal(row[sColumnToSort]) ascending
                                        select row);
                            dtsortedData = Rows.AsDataView().ToTable();
                        }
                    }
                }
                else
                {
                    view = new DataView(dtdataforsorting);
                    dtsortedData = view.ToTable();
                }
            }
            else
            {
                view = new DataView(dtdataforsorting);
                dtsortedData = view.ToTable();
            }
            return dtsortedData;
        }

        private void ComSortingOrder_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            ComSortingOrder_SelectedIndexChanged(sender, e);
        }

        private void ComSortingFields_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            ComSortingFields_SelectedIndexChanged(sender, e);
        }
    }
    #region Comment
    //dt = (DataTable)bindingSource_main.DataSource;
    ////dt.DefaultView.RowFilter = ADGDATA.FilterString;
    //DataView view = new DataView(dt);
    //view.Sort = sColumnToSort + " " + sSortingOrder;
    //dt = view.ToTable(); 
    #endregion
}

