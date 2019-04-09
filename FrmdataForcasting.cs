using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ZedGraph;

namespace cost_management
{
    public partial class FrmdataForcasting : Form
    {
        ToolTip tooltip = new ToolTip();
        string sMonat = "";
        ToolTip buttonToolTip = new ToolTip();

        ClsCommonFunction clscon = new ClsCommonFunction();
        public FrmdataForcasting()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;

            COMfiltering.Items.Add("Filtering based on " + clscon.GetYearName(1, "+") + " vs " + clscon.GetYearName(0, ""));
            COMfiltering.Items.Add("Filtering based on " + clscon.GetYearName(2, "+") + " vs " + clscon.GetYearName(0, ""));
            COMfiltering.Items.Add("Filtering based on " + clscon.GetYearName(3, "+") + " vs " + clscon.GetYearName(0, ""));

            this.chart1.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart1_GetToolTipText);
            this.chart2.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart2_GetToolTipText);

            buttonToolTip.UseFading = true;
            buttonToolTip.UseAnimation = true;
            buttonToolTip.IsBalloon = true;
            COMfiltering.Visible = false;
            //buttonToolTip.ToolTipIcon = ToolTipIcon.Info;
            buttonToolTip.ShowAlways = true;
            buttonToolTip.AutoPopDelay = 5000;
            buttonToolTip.InitialDelay = 10000;
            buttonToolTip.ReshowDelay = 10000;

            //panel10.Visible = false;
            //panel9.Visible = false;
            SetPanelLocation(true);
            btnexport.Enabled = false;
        }

        private void FrmdataForcasting_Load(object sender, EventArgs e)
        {
            try
            {
                lblper.Text = "Increase in %";
                lbltotal.Text = "Increase in Total";
                Application.DoEvents();
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("select [Date Id] from Cost_Summary where sid=(select Max(sid) from Cost_Summary)");
                sMonat = clscon.dtGetData(sbSQl.ToString()).Rows[0][0].ToString();
                Btnall.BackColor = Color.FromArgb(27, 161, 226);
                btnincrease.BackColor = Color.FromArgb(27, 161, 226);
                btnalldata.BackColor = Color.FromArgb(27, 161, 226);
                bindCombobox();
                this.WindowState = FormWindowState.Maximized;
                lbltoolver.Text = "Tool v" + clscon.GetAPPVersion();
                comForcastYtd.SelectedIndex = 0;
                comForcastYtd_SelectedIndexChanged(sender, e);
            }
            catch (Exception ex)
            {

            }
        }

        public void bindCombobox()
        {
            ComPartNo.SelectedIndexChanged -= new EventHandler(ComPartNo_SelectedIndexChanged);
            string str_SqlQuery = "select distinct [Part No] as SNR from [Cost_summary]";
            getcombox(ComPartNo, clscon.dtGetData(str_SqlQuery), "Part Number");
            ComPartNo.SelectedIndexChanged += new EventHandler(ComPartNo_SelectedIndexChanged);

            ComSupplierNo.SelectedIndexChanged -= new EventHandler(ComSupplierNo_SelectedIndexChanged);
            str_SqlQuery = "select distinct Cstr([Supplier Code]) as LN from [Cost_summary]";
            getcombox(ComSupplierNo, clscon.dtGetData(str_SqlQuery), "Supplier Number");
            ComSupplierNo.SelectedIndexChanged += new EventHandler(ComSupplierNo_SelectedIndexChanged);

            comsupplier.SelectedIndexChanged -= new EventHandler(comsupplier_SelectedIndexChanged);
            str_SqlQuery = "select distinct [Supplier Name] as Benennung from [Cost_summary] where [Supplier Name]<>''";
            getcombox(comsupplier, clscon.dtGetData(str_SqlQuery), "Supplier Name");
            comsupplier.SelectedIndexChanged += new EventHandler(comsupplier_SelectedIndexChanged);

            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [Cost_summary]";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);

            ComSOBLS.SelectedIndexChanged -= new EventHandler(ComSOBLS_SelectedIndexChanged);
            str_SqlQuery = "select distinct [Part Sobsl] as Mcode from [Cost_summary] where [Part Sobsl]<>''";
            getcombox(ComSOBLS, clscon.dtGetData(str_SqlQuery), "SOBSL-Selection");
            ComSOBLS.SelectedIndexChanged += new EventHandler(ComSOBLS_SelectedIndexChanged);

        }
        void getcombox(ComboBox cmb, DataTable dt, string strMsg)
        {
            try
            {
                DataRow dr = dt.NewRow();
                dr[0] = strMsg;
                dt.Rows.InsertAt(dr, 0);
                cmb.DataSource = dt;
                cmb.DisplayMember = dt.Columns[0].ColumnName.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }


        private void btnalldata_Click(object sender, EventArgs e)
        {
            btnalldata.BackColor = Color.FromArgb(27, 161, 226);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [Cost_summary]";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
            comForcastYtd_SelectedIndexChanged(sender, e);
        }

        private void btnpc_Click(object sender, EventArgs e)
        {
            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [Cost_summary] where ([Marketing Code] like '1%' or [Marketing Code] like 'SP%')";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
            string sWhers = " s.[Marketing Code] like '1%'";
            btnalldata.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(27, 161, 226);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData(sWhers);
        }

        private void btnvans_Click(object sender, EventArgs e)
        {
            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [Cost_summary] where [Marketing Code] like '2T%'";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
            string sWhers = " s.[Marketing Code]  like '2T%'";
            btnalldata.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(27, 161, 226);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(64, 64, 64);
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData(sWhers);
        }

        private void btntruks_Click(object sender, EventArgs e)
        {
            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [Cost_summary] where ([Marketing Code] like '2U%' or [Marketing Code] like '2L%')";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
            string sWhers = " (s.[Marketing Code] like '2U%' or s.[Marketing Code] like '2L%')";
            btnalldata.BackColor = Color.FromArgb(64, 64, 64);
            btnvans.BackColor = Color.FromArgb(64, 64, 64);
            btnpc.BackColor = Color.FromArgb(64, 64, 64);
            btntruks.BackColor = Color.FromArgb(27, 161, 226);
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData(sWhers);
        }

        private void Btnall_Click(object sender, EventArgs e)
        {
            DGIncDATA.DataSource = null;
            COMfiltering.Visible = false;
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.Enabled = false;
            comForcastYtd_SelectedIndexChanged(sender, e);
            SetPanelLocation(true);
            //panel10.Visible = false;
            //panel9.Visible = false;
        }

        private void btntb5_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.Enabled = true;
            // panel10.Visible = true;
            //panel9.Visible = true;
            SetPanelLocation(false);
            BindTop5();

        }

        public void BindTop5()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sASCPArtNo = "";
            string sDESCPartNo = "";
            string sWHerC = GetDynamicWhereClause();

            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2L%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]   like '2T%'";
                }
            }
            bindTopBottomData(sWHerC, "Top 5");
        }
        public void BindTop10()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();

            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2L%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]   like '2T%'";
                }
            }
            bindTopBottomData(sWHerC, "Top 10");
        }

        public void BindTop30()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();

            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2L%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]   like '2T%'";
                }
            }

            bindTopBottomData(sWHerC, "Top 30");
        }

        private void btntb10_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(27, 161, 226);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            //panel10.Visible = true;
            btnexport.Enabled = true;
            //panel9.Visible = true;
            SetPanelLocation(false);
            BindTop10();
        }

        private void ComMCSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!ComMCSelection.SelectedIndex.Equals(0))
            {
                string str_SqlQuery = "select distinct [Part No] from [Cost_summary] where [Marketing Code]='" + ComMCSelection.Text + "'";
                getcombox(ComPartNo, clscon.dtGetData(str_SqlQuery), "Part Number");
                string sValue = ComMCSelection.Text.Trim();
                if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop10();
                }
                else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop5();
                }
                else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop30();
                }
                else
                    GetChartData("");
            }
            else
            {
                string str_SqlQuery = "select distinct [Part No] from [Cost_summary] ";
                getcombox(ComPartNo, clscon.dtGetData(str_SqlQuery), "Part Number");
                if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop10();
                }
                else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop5();
                }
                else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    BindTop30();
                }
                else
                    GetChartData("");
            }
        }

        private void ComComCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void ComPartNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void ComSOBLS_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                BindTop10();
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                BindTop5();
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void comsupplier_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void ComSupplierNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void comForcastYtd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Btnall.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (comForcastYtd.SelectedIndex.Equals(1))
                {
                    COMfiltering.Visible = true;
                    COMfiltering.SelectedIndex = 0;
                }
                else
                {
                    COMfiltering.Visible = false;
                    COMfiltering.SelectedIndex = -1;
                }
            }
            else
            {
                COMfiltering.Visible = false;
                COMfiltering.SelectedIndex = -1;
            }

            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");

        }


        void GetChartData(string sWhereCaluses)
        {
            StringBuilder sWherCaluse = new StringBuilder();
            sWherCaluse.Append(GetDynamicWhereClause());
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("SELECT   SUM(g.[delta-1]) as [2017],SUM(g.[delta-2]) as [2018], ");
                sbSQl.Append("Round(SUM(g.[Delta]),2) as [2019] ");
                sbSQl.Append("From PART_YRLY_VALUES s inner join GRID g on s.[Part No]=g.[Part No]");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + " ");
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1 > s.AVG_HKSP_2 ");
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1 < AVG_HKSP_2");
                }

                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }

                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_3] - s.[AVG_HKSP_4]) /  s.[AVG_HKSP_4] ),2) AS [2017],");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2] - s.[AVG_HKSP_3]) /  s.[AVG_HKSP_3]), 2) AS[2018], ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_1] - s.[AVG_HKSP_2]) /  s.[AVG_HKSP_2]),2) AS [2019] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1>=s.AVG_HKSP_2");
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append("  s.AVG_HKSP_1<s.AVG_HKSP_2 ");
                }
                else
                {
                    sbSQl.Append(" s.AVG_HKSP_1 >= s.AVG_HKSP_2 ");
                }

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
            }
            else if (comForcastYtd.Text.Equals("Current To Future (3 Year)"))
            {
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(SUM(g.[Delta]),2) as [2019],Round(SUM(g.delta1),2) as [2020],Round(SUM(g.delta2),2) as [2021], ");
                sbSQl.Append(" Round(SUM(g.delta3),2) as [2022] ");
                sbSQl.Append("From PART_YRLY_VALUES s inner join GRID g on s.[Part No]=g.[Part No] ");
                sbSQl.Append(Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : " Where " + sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                //sbSQl.Append(FilterForCurrToFut(false));
                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }

                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(SUM(( s.[AVG_HKSP_1] - s.[AVG_HKSP_2]) /  s.[AVG_HKSP_2] ),2) AS[2019],");
                sbSQl.Append("ROUND(SUM(( s.[Forecast_1] - s.[AVG_HKSP_1]) /  s.[AVG_HKSP_1] ),2) AS[2020],");
                sbSQl.Append("ROUND(SUM(( s.[Forecast_2] - s.[Forecast_1]) / iif( s.[Forecast_1]=0,1, s.[Forecast_1])),2) AS  [2021],");
                sbSQl.Append("ROUND(SUM(( s.[Forecast_3] - s.[Forecast_2]) / iif( s.[Forecast_2]=0,1, s.[Forecast_2])),2) AS  [2022] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s inner join GRID g on s.[Part No]=g.[Part No] ");
                sbSQl.Append(Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : " Where " + sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                //sbSQl.Append(FilterForCurrToFut(false));
                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }

            }
            else if (comForcastYtd.Text.Equals("Previous To Future (3 Year)"))
            {
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(SUM(g.[delta-2]),2) as [2017],Round(SUM(g.[delta-1]),2) as [2018], ");
                sbSQl.Append("Round(SUM(g.[Delta]),2) as [2019]");
                sbSQl.Append(" ,Round(SUM(g.delta1),2) as [2020],Round(SUM(g.delta2),2) as [2021], ");
                sbSQl.Append("Round(SUM(g.delta3),2) as [2022] ");
                sbSQl.Append("From PART_YRLY_VALUES s inner join GRID g on s.[Part No]=g.[Part No] ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1>=s.AVG_HKSP_2");
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1<s.AVG_HKSP_2 ");
                }
                else
                {
                    sbSQl.Append(" s.AVG_HKSP_1 >= s.AVG_HKSP_2 ");
                }
                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }

                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(SUM((s.[AVG_HKSP_3] -s.[AVG_HKSP_4]) / s.[AVG_HKSP_4]), 2) AS[2017],");
                sbSQl.Append("ROUND(SUM((s.[AVG_HKSP_2] - s.[AVG_HKSP_3]) / s.[AVG_HKSP_3]), 2) AS[2018], ");
                sbSQl.Append("ROUND(SUM((s.[AVG_HKSP_1] -s.[AVG_HKSP_2]) / s.[AVG_HKSP_2]),2) AS[2019],");
                sbSQl.Append("ROUND(SUM((s.[Forecast_1] - s.[AVG_HKSP_1]) / s.[AVG_HKSP_1]),2) AS[2020],");
                sbSQl.Append("ROUND(SUM((s.[Forecast_2] -s.[Forecast_1]) / iif(s.[Forecast_1]=0,1,s.[Forecast_1])),2) AS[20210],");
                sbSQl.Append("ROUND(SUM((s.[Forecast_3] -s.[Forecast_2]) / iif(s.[Forecast_2]=0,1,s.[Forecast_2])),2) AS[2022]");
                sbSQl.Append(" FROM PART_YRLY_VALUES s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append(" s.AVG_HKSP_1>=s.AVG_HKSP_2");
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sbSQl.Append("  s.AVG_HKSP_1<s.AVG_HKSP_2 ");
                }
                else
                {
                    sbSQl.Append(" s.AVG_HKSP_1 >= s.AVG_HKSP_2 ");
                }
                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
            }
        }

        public string FilterForCurrToFut(bool bTopBottom)
        {
            string sFilterstring = "";
            try
            {
                if (COMfiltering.SelectedIndex.Equals(0))
                {
                    if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = " s.Forecast_1 >=s.AVG_HKSP_1";
                    }
                    else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = "  s.Forecast_1 < s.AVG_HKSP_1 ";
                    }
                }
                else if (COMfiltering.SelectedIndex.Equals(1))
                {
                    if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = " s.Forecast_2 >=s.AVG_HKSP_1";
                    }
                    else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = "  s.Forecast_2 < s.AVG_HKSP_1 ";
                    }
                }
                else if (COMfiltering.SelectedIndex.Equals(2))
                {
                    if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = " s.Forecast_3 >= s.AVG_HKSP_1";
                    }
                    else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = "  s.Forecast_3 < s.AVG_HKSP_1 ";
                    }
                }
                else
                {
                    if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = " s.Forecast_1 >=s.AVG_HKSP_1";
                    }
                    else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sFilterstring = "  s.Forecast_1 < s.AVG_HKSP_1 ";
                    }
                }
                sFilterstring = bTopBottom.Equals(true) ? sFilterstring.Replace("s.", "") : sFilterstring;
            }
            catch (Exception ex) { }

            return sFilterstring;
        }

        public string OrderForCurrToFut()
        {
            string sOrderString = "";
            if (COMfiltering.SelectedIndex.Equals(0))
            {
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " Order by (Delta_Total_Part_1) Desc,s.[Part No]";
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " and Delta_Total_Part_1 < 0  Order by (Delta_Total_Part_1),s.[Part No]";
                }
            }
            else if (COMfiltering.SelectedIndex.Equals(1))
            {
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " Order by (Delta_Total_Part_2) Desc,s.[Part No]";
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " and Delta_Total_Part_2 < 0  Order by (Delta_Total_Part_2),s.[Part No]";
                }
            }
            else if (COMfiltering.SelectedIndex.Equals(2))
            {
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " Order by (Delta_Total_Part_3) Desc,s.[Part No]";
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " and Delta_Total_Part_3 < 0  Order by (Delta_Total_Part_3),s.[Part No]";
                }
            }
            else
            {
                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " Order by (s.Delta) Desc,s.[Part No]";
                }
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sOrderString = " Order by (s.Delta),s.[Part No]";
                }
            }

            return sOrderString;
        }

        public string GetDynamicWhereClause()
        {
            StringBuilder sWherCaluse = new StringBuilder();

            if (!ComMCSelection.Text.Equals("All Marketing Code"))
            {
                sWherCaluse.Append(" s.[Marketing Code] ='" + ComMCSelection.Text.Trim() + "'");
            }

            if (!ComPartNo.Text.Equals("Part Number"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and s.[Part No] ='" + ComPartNo.Text.Trim() + "'");
                else
                    sWherCaluse.Append(" s.[Part No] ='" + ComPartNo.Text.Trim() + "'");
            }
            if (!ComSupplierNo.Text.Equals("Supplier Number"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and  s.[Supplier Code] =" + ComSupplierNo.Text.Trim() + "");
                else
                    sWherCaluse.Append(" s.[Supplier Code] =" + ComSupplierNo.Text.Trim() + "");
            }
            if (!ComSOBLS.Text.Equals("SOBSL-Selection"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and s.[Part Sobsl] ='" + ComSOBLS.Text.Trim() + "'");
                else
                    sWherCaluse.Append(" s.[Part Sobsl] ='" + ComSOBLS.Text.Trim() + "'");
            }
            if (!comsupplier.Text.Equals("Supplier Name"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and  s.[Supplier Name] ='" + comsupplier.Text.Trim() + "'");
                else
                    sWherCaluse.Append(" s.[Supplier Name]='" + comsupplier.Text.Trim() + "'");
            }
            if (!ComComCode.Text.Equals("Commodity Code"))
            {
                if (sWherCaluse.Length > 0)
                    sWherCaluse.Append(" and  ='" + ComComCode.Text.Trim() + "'");
                else
                    sWherCaluse.Append(" ='" + ComComCode.Text.Trim() + "' ");
            }
            return sWherCaluse.ToString();
        }


        public void BindTop5INCDEC(string sSign, decimal dValue, string sType)
        {
            if (dValue.Equals(0))
            {
                BindTop5();
                return;
            }
        }
        public void BindTop10INCDEC(string sSign, decimal dValue, string sType)
        {
            if (dValue.Equals(0))
            {
                BindTop10();
                return;
            }
        }
        public void BindTop30INCDEC(string sSign, decimal dValue, string sType)
        {
            if (dValue.Equals(0))
            {
                BindTop30();
                return;
            }
        }

        public string CreateInCalsue(DataTable dtIndata)
        {
            string sClasue = "";
            try
            {
                for (int i = 1; i <= dtIndata.Columns.Count; i++)
                {
                    if (i.Equals(dtIndata.Columns.Count - 1))
                    {
                        sClasue = sClasue + "'" + dtIndata.Columns[i].ColumnName.ToString() + "'";
                    }
                    else
                    {
                        sClasue = sClasue + "'" + dtIndata.Columns[i].ColumnName.ToString() + "',";
                    }
                }
            }
            catch { }
            return sClasue;
        }
        public DataTable ConvertRowsToColForChart(DataTable dtd)
        {
            DataTable dtdata = new DataTable();
            try
            {
                dtdata.Columns.Add(new DataColumn("Years", typeof(string)));
                dtdata.Columns.Add(new DataColumn("All Parts", typeof(string)));
                for (int i = 0; i < dtd.Columns.Count; i++)
                {
                    dtdata.Rows.Add();
                    dtdata.Rows[i][0] = dtd.Columns[i].ColumnName;
                    dtdata.Rows[i][1] = (dtd.Rows[0][i].ToString().Equals("") ? "0" : dtd.Rows[0][i].ToString());
                }
            }
            catch { }
            return dtdata;
        }

        public DataTable convertRowsToColFotTOPBottom(DataTable dtd)
        {
            DataTable dtdata = new DataTable();
            try
            {
                if (dtd.Rows.Count > 0)
                {
                    for (int i = 0; i < dtd.Rows.Count; i++)
                    {
                        try
                        {
                            dtdata.Columns.Add(new DataColumn(dtd.Rows[i][0].ToString(), typeof(string)));
                        }
                        catch { }
                    }
                    for (int x = 0; x < dtd.Columns.Count - 1; x++)
                    {
                        dtdata.Rows.Add();
                    }
                    for (int j = 0; j < dtd.Rows.Count; j++)
                    {
                        for (int i = 0; i < dtd.Columns.Count - 1; i++)
                        {
                            string sValue = Convert.ToString(dtd.Rows[j][i + 1].ToString().Equals("") ? "0" : dtd.Rows[j][i + 1].ToString());
                            dtdata.Rows[i][j] = Math.Truncate(Convert.ToDouble(sValue));
                        }
                    }
                    dtdata.Columns.Add(new DataColumn("Year", typeof(string)));
                    dtdata.Columns["Year"].SetOrdinal(0);
                    for (int i = 1; i < dtd.Columns.Count; i++)
                    {
                        dtdata.Rows[i - 1][0] = dtd.Columns[i].ColumnName;
                    }
                    for (int i = dtdata.Rows.Count - 1; i > 0; i--)
                    {
                        if (dtdata.Rows[i][0].ToString().Equals(""))
                        {
                            dtdata.Rows.RemoveAt(i);
                        }
                    }
                }
                else
                {
                    dtdata.Columns.Add(new DataColumn("All Parts", typeof(string)));
                    for (int x = 0; x < dtd.Columns.Count - 1; x++)
                    {
                        dtdata.Rows.Add();
                    }

                    for (int i = 0; i < dtd.Columns.Count - 1; i++)
                    {
                        dtdata.Rows[i][0] = "0";
                    }
                    dtdata.Columns.Add(new DataColumn("Year", typeof(string)));
                    dtdata.Columns["Year"].SetOrdinal(0);
                    for (int i = 1; i < dtd.Columns.Count; i++)
                    {
                        dtdata.Rows[i - 1][0] = dtd.Columns[i].ColumnName;
                    }
                }
            }
            catch (Exception ex) { }

            return dtdata;
        }

        private void ComPartNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Enter))
            {
                if (ComPartNo.Text.Length.Equals(0))
                {
                    ComPartNo.SelectedIndexChanged -= new EventHandler(ComPartNo_SelectedIndexChanged);
                    ComPartNo.SelectedIndex = 0;
                    ComPartNo.SelectedIndexChanged += new EventHandler(ComPartNo_SelectedIndexChanged);
                    ComPartNo_SelectedIndexChanged(sender, e);
                }
            }
        }



        private void chart1_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            // Check selected chart element and set tooltip text for it
            switch (e.HitTestResult.ChartElementType)
            {
                case ChartElementType.DataPoint:
                    var dataPoint = e.HitTestResult.Series.Points[e.HitTestResult.PointIndex];
                    string sSNR = e.HitTestResult.Series.Name;
                    string sText = string.Format("Part No:  {0} \nKost:\t{1}%", sSNR, dataPoint.YValues[0]);
                    buttonToolTip.SetToolTip(chart1, sText);
                    break;
            }
        }
        private void chart2_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            // Check selected chart element and set tooltip text for it
            switch (e.HitTestResult.ChartElementType)
            {
                case ChartElementType.DataPoint:
                    var dataPoint = e.HitTestResult.Series.Points[e.HitTestResult.PointIndex];
                    string sSNR = e.HitTestResult.Series.Name;
                    string sText = string.Format("Part No:  {0} \nKost:\t€ {1}", sSNR, dataPoint.YValues[0].ToString("#,##0"));
                    buttonToolTip.SetToolTip(chart2, sText);
                    break;
            }
        }

        private void ComMCSelection_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Enter))
            {
                if (ComMCSelection.Text.Length.Equals(0))
                {
                    ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
                    ComMCSelection.SelectedIndex = 0;
                    ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
                    ComMCSelection_SelectedIndexChanged(sender, e);
                }
            }
        }
        private void BindChartwithdata(string sChartTitle, DataTable dtdata, System.Windows.Forms.DataVisualization.Charting.Chart cChart, string sChartArea,
                                       string Xvalue)
        {
            try
            {

                if (dtdata.Rows.Count.Equals(0))
                {
                    cChart.Series.Clear();
                    return;
                }

                cChart.Series.Clear();
                System.Windows.Forms.DataVisualization.Charting.Legend legend = cChart.Legends[0];
                foreach (DataColumn dcol in dtdata.Columns)
                {
                    if (!dcol.ColumnName.ToString().Equals(Xvalue))
                    {
                        cChart.Series.Add(dcol.ColumnName.ToString());
                        cChart.Series[dcol.ColumnName.ToString()].XValueMember = Xvalue;
                        cChart.Series[dcol.ColumnName.ToString()].YValueMembers = dcol.ColumnName.ToString();
                    }
                }
                foreach (Series s in cChart.Series)
                {
                    s.ChartType = SeriesChartType.RangeColumn;
                    s.IsValueShownAsLabel = false;
                    s.LegendText = "";

                }
                cChart.ChartAreas[sChartArea].AxisX.Interval = 1;

                cChart.ChartAreas[sChartArea].AxisX.MajorGrid.LineWidth = 0;
                cChart.ChartAreas[sChartArea].AxisY.MajorGrid.LineWidth = 0;
                if (cChart.Name.Equals("chart1") || cChart.Name.Equals("chart4"))
                {
                    cChart.ChartAreas[sChartArea].AxisY.LabelStyle.Format = "{0.00} %";
                }
                cChart.DataSource = dtdata;
            }
            catch { }
        }

        private void Validation(object sender, KeyPressEventArgs e)
        {

        }

        private void comtotal_SelectedIndexChanged(object sender, EventArgs e)
        {
            //txtincper_TextChanged(sender, e);
        }

        private void btntb30_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            btntb30.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            // panel10.Visible = true;
            btnexport.Enabled = true;
            //panel9.Visible = true;
            SetPanelLocation(false);
            BindTop30();
        }

        private void DATAGRIDBinding(string sASCPartNo, string sDESCPartNo, string sWHerC)
        {
            //panel9.Visible = true;
            //panel10.Visible = true;
            // panel9.BringToFront();
            //panel10.BringToFront();
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            StringBuilder sbSQL = new StringBuilder();
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            try
            {
                if (!sDESCPartNo.Equals(""))
                {
                    if (!comForcastYtd.SelectedIndex.Equals(1))
                    {
                        sbSQL.Append("select distinct [Marketing Code],[Part No],[Part Name],[Supplier Name],[Manufacturing cost " + clscon.GetYear(1, "-") + " Year per Part],[Current month Manufacturing cost per part],");
                        sbSQL.Append("[Delta Value per Part Current Year vs " + clscon.GetYear(1, "-") + "],[Delta Manufacturing cost current year vs " + clscon.GetYear(1, "-") + "], [Delta manufacturing cost forecast " + clscon.GetYear(1, "+") + " vs " + clscon.GetYear(0, "") + "],");
                        sbSQL.Append("[Delta manufacturing cost forecast " + clscon.GetYear(2, "+") + " vs " + clscon.GetYear(1, "+") + "],(s.delta) as Deltas FROM Grid s where s.[Part No] in (" + sDESCPartNo + ")");
                        sbSQL.Append(" and [date Id]='" + sMoath + "'");
                    }
                    else
                    {
                        sbSQL.Append("select distinct s.[Marketing Code],s.[Part No],[Part Name],s.[Supplier Name],Format(s.[Manufacturing cost 2017 Year per Part],'#,##0.00') as [Manufacturing cost 2017 Year per Part]");
                        sbSQL.Append(" ,Format(s.[Current month Manufacturing cost per part],'#,##0.00') as [Current month Manufacturing cost per part],");
                        sbSQL.Append("s.[Forecast Part Quantity Current Year],");
                        sbSQL.Append("Format(s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(1, "+") + "],'#,##0.00') as [Forcasted Manufacturing Cost Year " + clscon.GetYear(1, "+") + "],");
                        sbSQL.Append("Format(s.[Forecast Quntity " + clscon.GetYear(1, "+") + "],'#,##0.00') as [Forecast Quntity " + clscon.GetYear(1, "+") + "],");
                        sbSQL.Append("Format(s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(2, "+") + "],'#,##0.00') as [Forcasted Manufacturing Cost Year " + clscon.GetYear(2, "+") + "],");
                        sbSQL.Append("Format(s.[Forecast Quntity " + clscon.GetYear(2, "+") + "],'#,##0.00') as [Forecast Quntity " + clscon.GetYear(2, "+") + "],");
                        sbSQL.Append("Format(s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(3, "+") + "],'#,##0.00')as [Forcasted Manufacturing Cost Year " + clscon.GetYear(3, "+") + "],");
                        sbSQL.Append("Format(s.[Forecast Quntity " + clscon.GetYear(3, "+") + "],'#,##0.00')as [Forecast Quntity " + clscon.GetYear(3, "+") + "],");
                        sbSQL.Append("Format(ROUND((s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(1, "+") + "]-PART_YRLY_VALUES.AVG_HKSP_1),2),'#,##0.00') as [Delta Per part " + clscon.GetYear(1, "+") + " vs " + clscon.GetYear(0, "") + "]");
                        sbSQL.Append(",Format(ROUND(PART_YRLY_VALUES.Delta_Total_Part_1,2),'#,##0.00') as [Delta Total part " + clscon.GetYear(1, "+") + " vs " + clscon.GetYear(0, "") + " ],");
                        sbSQL.Append("Format(ROUND((s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(2, "+") + "]-PART_YRLY_VALUES.AVG_HKSP_1),2),'#,##0.00') as [Delta Per part " + clscon.GetYear(2, "+") + " vs " + clscon.GetYear(0, "") + "]");
                        sbSQL.Append(",Format(ROUND(PART_YRLY_VALUES.Delta_Total_Part_2,2),'#,##0.00') as [Delta Total part " + clscon.GetYear(2, "+") + " vs " + clscon.GetYear(0, "") + " ],");
                        sbSQL.Append("Format(ROUND((s.[Forcasted Manufacturing Cost Year " + clscon.GetYear(3, "+") + "]-PART_YRLY_VALUES.AVG_HKSP_1),2),'#,##0.00') as [Delta Per part " + clscon.GetYear(3, "+") + " vs " + clscon.GetYear(0, "") + "]");
                        sbSQL.Append(",Format(ROUND(PART_YRLY_VALUES.Delta_Total_Part_3,2),'#,##0.00') as [Delta Total part " + clscon.GetYear(3, "+") + " vs " + clscon.GetYear(0, "") + " ]");
                        sbSQL.Append(",Delta_Total_Part_1,Delta_Total_Part_2,Delta_Total_Part_3 FROM Grid s inner join PART_YRLY_VALUES on s.[Part No]=PART_YRLY_VALUES.[Part No] where s.[Part No] in (" + sDESCPartNo + ")");
                        sbSQL.Append(" and [date Id]='" + sMoath + "'");
                    }
                    sbSQL.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
                    if (comForcastYtd.SelectedIndex.Equals(1))
                        sbSQL.Append(OrderForCurrToFut());
                    else
                        sbSQL.Append(" order By (s.[Delta]) desc");
                    DataTable dtdats = clscon.dtGetData(sbSQL.ToString());
                    try
                    {
                        dtdats.Columns.Remove("Delta_Total_Part_1");
                        dtdats.Columns.Remove("Delta_Total_Part_2");
                        dtdats.Columns.Remove("Delta_Total_Part_3");
                    }
                    catch { }
                    DGIncDATA.DataSource = dtdats;
                    DGIncDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                    DGIncDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                    DGIncDATA.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlDark;

                    if (!comForcastYtd.SelectedIndex.Equals(1))

                        DGIncDATA.Columns["Deltas"].Visible = false;
                }
            }
            catch (Exception ex)
            { }

        }
        private void SetPanelLocation(bool isLocationChange)
        {

        }



        private void btnexport_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(27, 161, 226);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            clscon.ExportData((DataTable)DGIncDATA.DataSource, null);
        }

        private void comIncrease_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        public void bindTopBottomData(string sWHerC, string sTopNo)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            if (comForcastYtd.SelectedIndex.Equals(1))
            {
                COMfiltering.Visible = true;
                if (COMfiltering.SelectedIndex == 0)
                    COMfiltering.SelectedIndex = 0;
            }
            else
                COMfiltering.Visible = false;

            StringBuilder sbSQl = new StringBuilder();
            string sASCPArtNo = "";
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            string sDESCPartNo = "";
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append(" SELECT  distinct " + sTopNo + " s.[Part No],s.[Delta-2] as [2017], s.[Delta-1] as [2018], ");
                sbSQl.Append(" (s.delta)  AS[2019]  from GRID s ");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (3 Year)"))
            {
                sbSQl.Append(" SELECT  distinct  " + sTopNo + " s.[Part No],s.Delta as [2019],Delta_Total_Part_1 as [2020],Delta_Total_Part_2 as [2021],Delta_Total_Part_3 as [2022] ");
                sbSQl.Append(" from GRID s inner Join PART_YRLY_VALUES on s.[Part no]=PART_YRLY_VALUES.[part no] ");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (3 Year)"))
            {
                sbSQl.Append(" SELECT distinct " + sTopNo + " s.[Part No],Round((s.[Delta-2]), 2) as [2017], Round((s.[Delta-1]), 2) as [2018],");
                sbSQl.Append(" (s.[Delta]) AS[2019], Round((delta1), 2) as [2020], Round((delta2), 2) as [2021], ");
                sbSQl.Append(" Round((delta3), 2) AS [2022] from GRID s");
            }

            sbSQl.Append(" where [date Id]='" + sMoath + "'");
            sbSQl.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
            sbSQl.Append(OrderForCurrToFut());
            DataTable dtd = clscon.dtGetData(sbSQl.ToString());
            DataTable dtdata = new DataTable();
            try
            {
                dtd.Columns.Remove("Delta_Total_Part_1");
                dtd.Columns.Remove("Delta_Total_Part_2");
                dtd.Columns.Remove("Delta_Total_Part_3");
            }
            catch (Exception ex) { }
            dtdata = convertRowsToColFotTOPBottom(dtd);
            sDESCPartNo = CreateInCalsue(dtdata);

            BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);
            sbSQl = new StringBuilder();


            sbSQl = new StringBuilder();
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2019]");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (3 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2021],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2022]");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (3 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2021],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2022]");
            }
            sbSQl.Append(" FROM PART_YRLY_CHANGE s ");
            sbSQl.Append(" where s.snr in(" + sDESCPartNo + ") ");
            sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);

            if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                sbSQl.Append(" order by ROUND(s.Change_2018,2)  DESC,s.snr ");
            }
            else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                sbSQl.Append("  order by Round(s.Change_2018,2)  ASC,s.snr");
            }
            else
            {
                sbSQl.Append(" order by ROUND(s.Change_2018,2)  DESC,s.snr ");
            }

            dtd = clscon.dtGetData(sbSQl.ToString());

            dtdata = new DataTable();
            dtdata = convertRowsToColFotTOPBottom(dtd);
            BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName);

            DATAGRIDBinding(sASCPArtNo, sDESCPartNo, sWHerC);
            this.Cursor = Cursors.Default;
        }

        public void bindTopBottomDataICDEC(string sWHerC, string sTopNo, bool isTotal)
        {
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            StringBuilder sbSQl = new StringBuilder();
            string sASCPArtNo = "";
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            string sDESCPartNo = "";
            if (isTotal)
            {
                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append(" SELECT  distinct " + sTopNo + " s.[Part No],s.[Delta-2] as [2016], s.[Delta-1] as [2017], ");
                    sbSQl.Append(" (s.delta)  AS[2018]  from GRID s ");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (3 Year)"))
                {
                    sbSQl.Append(" SELECT  distinct  " + sTopNo + " s.[Part No],s.Delta as [2018],Delta_Total_Part_1 as [2019],Delta_Total_Part_2 as [2020],Delta_Total_Part_3 as [2021] ");
                    sbSQl.Append(" from GRID s inner Join PART_YRLY_VALUES on s.[Part no]=PART_YRLY_VALUES.[part no] ");
                }
                else if (comForcastYtd.Text.Equals("Previous To Future (3 Year)"))
                {
                    sbSQl.Append(" SELECT distinct " + sTopNo + " s.[Part No],Round((s.[Delta-2]), 2) as [2016], Round((s.[Delta-1]), 2) as [2017],");
                    sbSQl.Append(" (s.[Delta]) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                    sbSQl.Append(" Round((delta3), 2) AS[2021]  from GRID s");
                }
                sbSQl.Append(" where [date Id]='" + sMoath + "'");
                if (comForcastYtd.SelectedIndex.Equals(1))
                {
                    if (COMfiltering.SelectedIndex.Equals(0))
                    {
                        if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_1] >= " + txttotal.Text.Trim() + "");
                        else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_1] <= " + txttotal.Text.Trim() + "");
                    }
                    else if (COMfiltering.SelectedIndex.Equals(1))
                    {
                        if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_2] >= " + txttotal.Text.Trim() + "");
                        else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_2] <= " + txttotal.Text.Trim() + "");
                    }
                    else if (COMfiltering.SelectedIndex.Equals(2))
                    {
                        if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_3] >= " + txttotal.Text.Trim() + "");
                        else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                            sbSQl.Append(" and [Delta_Total_Part_3] <= " + txttotal.Text.Trim() + "");
                    }
                }
                else
                {
                    if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        sbSQl.Append(" and [Delta] >= " + txttotal.Text.Trim() + "");
                    else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                        sbSQl.Append(" and [Delta] <= " + txttotal.Text.Trim() + "");
                }
                sbSQl.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
                sbSQl.Append(OrderForCurrToFut());
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                DataTable dtdata = new DataTable();
                try
                {
                    dtd.Columns.Remove("Delta_Total_Part_1");
                    dtd.Columns.Remove("Delta_Total_Part_2");
                    dtd.Columns.Remove("Delta_Total_Part_3");
                }
                catch (Exception ex) { }
                dtdata = convertRowsToColFotTOPBottom(dtd);

                BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);
                sbSQl = new StringBuilder();
            }
            else
            {
                sbSQl = new StringBuilder();
                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append("SELECT DISTINCT " + sTopNo + " s.SNR,");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (3 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT " + sTopNo + " s.SNR,");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021]");
                }
                else if (comForcastYtd.Text.Equals("Previous To Future (3 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT " + sTopNo + " s.SNR,");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021]");
                }
                sbSQl.Append(" FROM PART_YRLY_CHANGE s");

                if (btnincrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    sbSQl.Append(" where  s.[Delta_all_part]>= " + txtincper.Text + "");
                else if (btndecrease.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    sbSQl.Append(" where  s.[Delta_all_part]<=" + txtincper.Text + "");

                sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                DataTable dtdata = new DataTable();
                dtdata = convertRowsToColFotTOPBottom(dtd);
                BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName);
                //sDESCPartNo = CreateInCalsue(dtdata);
            }
            this.Cursor = Cursors.Default;
        }

        private void groupBox1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnincrease_Click(object sender, EventArgs e)
        {
            lblper.Text = "Increase in %";
            lbltotal.Text = "Increase in Total";
            btnincrease.BackColor = Color.FromArgb(27, 161, 226);
            btndecrease.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void btndecrease_Click(object sender, EventArgs e)
        {
            lblper.Text = "Decrease in %";
            lbltotal.Text = "Decrease in Total";
            btnincrease.BackColor = Color.FromArgb(64, 64, 64);
            btndecrease.BackColor = Color.FromArgb(27, 161, 226);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void txttotal_TextChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txttotal.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 10", true);
                else
                    BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txttotal.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 5", true);
                else
                    BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txttotal.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 30", true);
                else
                    BindTop30();
            }
        }

        private void txtincper_TextChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txtincper.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 10", false);
                else
                    BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txtincper.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 5", false);
                else
                    BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                if (!txtincper.Text.Length.Equals(0))
                    bindTopBottomDataICDEC(GetDynamicWhereClause(), "Top 30", false);
                else
                    BindTop30();
            }
        }

        private void COMfiltering_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10();
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5();
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30();
            }
            else
                GetChartData("");
        }

        private void COMfiltering_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            COMfiltering_SelectedIndexChanged(sender, e);
        }
    }
}