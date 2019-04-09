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
            //this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            //int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            //int screenHeight = Screen.PrimaryScreen.Bounds.Height;
            //Resolution objFormResizer = new Resolution();
            //objFormResizer.ResizeForm(this, screenHeight, screenWidth);
            this.chart1.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart1_GetToolTipText);
            this.chart2.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart2_GetToolTipText);
            this.chart3.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart3_GetToolTipText);
            this.chart4.GetToolTipText += new EventHandler<ToolTipEventArgs>(chart4_GetToolTipText);

            buttonToolTip.UseFading = true;
            buttonToolTip.UseAnimation = true;
            buttonToolTip.IsBalloon = true;
            //buttonToolTip.ToolTipIcon = ToolTipIcon.Info;
            buttonToolTip.ShowAlways = true;
            buttonToolTip.AutoPopDelay = 5000;
            buttonToolTip.InitialDelay = 10000;
            buttonToolTip.ReshowDelay = 10000;

            panel10.Visible = false;
            panel9.Visible = false;
            SetPanelLocation(true);
            btnexport.Enabled = false;
        }

        private void FrmdataForcasting_Load(object sender, EventArgs e)
        {
            Application.DoEvents();
            StringBuilder sbSQl = new StringBuilder();
            sbSQl.Append("select [Date Id] from Cost_Summary where sid=(select Max(sid) from Cost_Summary)");
            sMonat = clscon.dtGetData(sbSQl.ToString()).Rows[0][0].ToString();
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            btnalldata.BackColor = Color.FromArgb(27, 161, 226);
            bindCombobox();
            this.WindowState = FormWindowState.Maximized;
            lbltoolver.Text = "Tool v" + clscon.GetAPPVersion();
            comForcastYtd.SelectedIndex = 0;
            comForcastYtd_SelectedIndexChanged(sender, e);
        }

        public void bindCombobox()
        {
            ComPartNo.SelectedIndexChanged -= new EventHandler(ComPartNo_SelectedIndexChanged);
            string str_SqlQuery = "select distinct [Part No] as SNR from [PART_YRLY_VALUES]";
            getcombox(ComPartNo, clscon.dtGetData(str_SqlQuery), "Part Number");
            ComPartNo.SelectedIndexChanged += new EventHandler(ComPartNo_SelectedIndexChanged);

            ComSupplierNo.SelectedIndexChanged -= new EventHandler(ComSupplierNo_SelectedIndexChanged);
            str_SqlQuery = "select distinct Cstr([Supplier Code]) as LN from [PART_YRLY_VALUES]";
            getcombox(ComSupplierNo, clscon.dtGetData(str_SqlQuery), "Supplier Number");
            ComSupplierNo.SelectedIndexChanged += new EventHandler(ComSupplierNo_SelectedIndexChanged);

            comsupplier.SelectedIndexChanged -= new EventHandler(comsupplier_SelectedIndexChanged);
            str_SqlQuery = "select distinct [Supplier Name] as Benennung from [PART_YRLY_VALUES] where [Supplier Name]<>''";
            getcombox(comsupplier, clscon.dtGetData(str_SqlQuery), "Supplier Name");
            comsupplier.SelectedIndexChanged += new EventHandler(comsupplier_SelectedIndexChanged);

            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [PART_YRLY_VALUES]";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);

            ComSOBLS.SelectedIndexChanged -= new EventHandler(ComSOBLS_SelectedIndexChanged);
            str_SqlQuery = "select distinct [Part Sobsl] as Mcode from [PART_YRLY_VALUES] where [Part Sobsl]<>''";
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
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [PART_YRLY_VALUES]";
            getcombox(ComMCSelection, clscon.dtGetData(str_SqlQuery), "All Marketing Code");
            ComMCSelection.SelectedIndexChanged += new EventHandler(ComMCSelection_SelectedIndexChanged);
            comForcastYtd_SelectedIndexChanged(sender, e);
        }

        private void btnpc_Click(object sender, EventArgs e)
        {
            ComMCSelection.SelectedIndexChanged -= new EventHandler(ComMCSelection_SelectedIndexChanged);
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [PART_YRLY_VALUES] where ([Marketing Code] like '1%' or [Marketing Code] like 'SP%')";
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
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [PART_YRLY_VALUES] where [Marketing Code] like '2T%'";
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
            string str_SqlQuery = "select distinct Cstr([Marketing Code]) as Mcode from [PART_YRLY_VALUES] where ([Marketing Code] like '2U%' or [Marketing Code] like '2L%')";
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
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(27, 161, 226);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.Enabled = false;
            comForcastYtd_SelectedIndexChanged(sender, e);
            SetPanelLocation(true);
            panel10.Visible = false;
            panel9.Visible = false;
        }

        private void btntb5_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(27, 161, 226);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.Enabled = true;
            panel10.Visible = true;
            panel9.Visible = true;
            SetPanelLocation(false);
            BindTop5();

        }

        public void BindTop5()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sASCPArtNo = "";
            string sDESCPartNo = "";
            string sWHerC = GetDynamicWhereClause();
            if (!txttotal.Text.Equals(""))
            {
                BindTop5INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (!txtincper.Text.Equals(""))
            {
                BindTop5INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "");
            }
            else
            {
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
                        sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                    }
                    else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                        sWHerC = sWHerC + " s.[Marketing Code]  like '2T%'";
                    }
                }
                bindTopBottomData(sWHerC, "Top 5");
            }
        }
        public void BindTop10()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();
            if (!txttotal.Text.Equals(""))
            {
                BindTop10INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (!txtincper.Text.Equals(""))
            {
                BindTop10INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "");
            }
            else
            {
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
                        sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                    }
                    else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                        sWHerC = sWHerC + " s.[Marketing Code]   like '2T%'";
                    }
                }
                bindTopBottomData(sWHerC, "Top 10");
            }
        }

        public void BindTop30()
        {
            StringBuilder sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();
            if (!txttotal.Text.Equals(""))
            {
                BindTop30INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (!txtincper.Text.Equals(""))
            {
                BindTop30INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "");
            }
            else
            {
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
                        sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                    }
                    else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                    {
                        sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                        sWHerC = sWHerC + " s.[Marketing Code]  like '2T%'";
                    }
                }

                bindTopBottomData(sWHerC, "Top 30");
            }
        }

        private void btntb10_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(27, 161, 226);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(64, 64, 64);
            panel10.Visible = true;
            btnexport.Enabled = true;
            panel9.Visible = true;
            SetPanelLocation(false);
            BindTop10();
        }

        private void ComMCSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!ComMCSelection.SelectedIndex.Equals(0))
            {
                string str_SqlQuery = "select distinct [Part No] from [PART_YRLY_VALUES] where [Marketing Code]='" + ComMCSelection.Text + "'";
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
                string str_SqlQuery = "select distinct [Part No] from [PART_YRLY_VALUES] ";
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
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2015),2) as [2015],Round(AVG(s.AVG_HKSP_2016),2) as [2016],Round(AVG(s.AVG_HKSP_2017),2) as [2017], ");
                sbSQl.Append("Round(AVG(s.AVG_HKSP_2018),2) as [2018] ");
                sbSQl.Append("From Part_Yrly_Values s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + " s.AVG_HKSP_2018 >= s.AVG_HKSP_2017");
                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2015),2) as [2015],Round(AVG(s.AVG_HKSP_2016),2) as [2016],Round(AVG(s.AVG_HKSP_2017),2) as [2017], ");
                sbSQl.Append("Round(AVG(s.AVG_HKSP_2018),2) as [2018] ");
                sbSQl.Append("From Part_Yrly_Values s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + " s.AVG_HKSP_2018 < AVG_HKSP_2017");


                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart3, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2016] - s.[AVG_HKSP_2015]) /  s.[AVG_HKSP_2015] ),2) AS[2016],");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2017] - s.[AVG_HKSP_2016]) /  s.[AVG_HKSP_2016]), 2) AS[2017], ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2018] - s.[AVG_HKSP_2017]) /  s.[AVG_HKSP_2017]),2) AS[2018] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018>=s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2016] - s.[AVG_HKSP_2015]) /  s.[AVG_HKSP_2015]), 2) AS[2016],");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2017] - s.[AVG_HKSP_2016]) /  s.[AVG_HKSP_2016]), 2) AS[2017], ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2018] - s.[AVG_HKSP_2017]) /  s.[AVG_HKSP_2017]), 2) AS[2018] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018<s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart4, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
            }
            else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
            {
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2018),2) as [2018],Round(AVG(s.Forecast_1),2) as [2019],Round(AVG(s.Forecast_2),2) as [2020], ");
                sbSQl.Append(" Round(AVG(s.Forecast_3),2) as [2021],Round(AVG(s.Forecast_4),2) as [2022],Round(AVG(s.Forecast_5),2) as [2023] ");
                sbSQl.Append("From Part_Yrly_Values s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018>=s.AVG_HKSP_2017");
                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2018),2) as [2018],Round(AVG(s.Forecast_1),2) as [2019],Round(AVG(s.Forecast_2),2) as [2020], ");
                sbSQl.Append("Round(AVG(s.Forecast_3),2) as [2021],Round(AVG(s.Forecast_4),2) as [2022],Round(AVG(s.Forecast_5),2) as [2023] ");
                sbSQl.Append("From Part_Yrly_Values s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018< s.AVG_HKSP_2017");


                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart3, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2018] - s.[AVG_HKSP_2017]) /  s.[AVG_HKSP_2017] ),2) AS[2018],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_1] - s.[AVG_HKSP_2018]) /  s.[AVG_HKSP_2018] ),2) AS[2019],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_2] - s.[Forecast_1]) / iif( s.[Forecast_1]=0,1, s.[Forecast_1])),2) AS  [2020],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_3] - s.[Forecast_2]) / iif( s.[Forecast_2]=0,1, s.[Forecast_2])),2) AS  [2021],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_4] - s.[Forecast_3]) / iif( s.[Forecast_3]=0,1, s.[Forecast_3])), 2) AS [2022],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_5] - s.[Forecast_4]) / iif( s.[Forecast_4]=0,1, s.[Forecast_3])), 2) AS [2023] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018>=s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2018] - s.[AVG_HKSP_2017]) /  s.[AVG_HKSP_2017] ),2) AS[2018],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_1] - s.[AVG_HKSP_2018]) /  s.[AVG_HKSP_2018] ),2) AS[2019],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_2] - s.[Forecast_1]) / iif( s.[Forecast_1]=0,1, s.[Forecast_1])),2) AS[2020],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_3] - s.[Forecast_2]) / iif( s.[Forecast_2]=0,1, s.[Forecast_2])),2) AS[2021],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_4] - s.[Forecast_3]) / iif( s.[Forecast_3]=0,1, s.[Forecast_3])), 2) AS [2022], ");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_5] - s.[Forecast_4]) / iif( s.[Forecast_4]=0,1, s.[Forecast_4])), 2) AS [2023] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018<s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart4, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
            {
                StringBuilder sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2015),2) as [2015],Round(AVG(s.AVG_HKSP_2016),2) as [2016],Round(AVG(s.AVG_HKSP_2017),2) as [2017], ");
                sbSQl.Append("Round(AVG(s.AVG_HKSP_2018),2) as [2018],Round(AVG(s.Forecast_1),2) as [2019],Round(AVG(s.Forecast_2),2) as [2020], ");
                sbSQl.Append("Round(AVG(s.Forecast_3),2) as [2021],Round(AVG(s.Forecast_4),2) as [2022],Round(AVG(s.Forecast_5),2) as [2023] ");
                sbSQl.Append("From Part_Yrly_Values s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018 >= s.AVG_HKSP_2017");
                DataTable dtdata = new DataTable();
                DataTable dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  Round(AVG(s.AVG_HKSP_2015),2) as [2015],Round(AVG(s.AVG_HKSP_2016),2) as [2016],Round(AVG(s.AVG_HKSP_2017),2) as [2017], ");
                sbSQl.Append("Round(AVG(s.AVG_HKSP_2018),2) as [2018],Round(AVG(s.Forecast_1),2) as [2019],Round(AVG(s.Forecast_2),2) as [2020], ");
                sbSQl.Append("Round(AVG(s.Forecast_3),2) as [2021],Round(AVG(s.Forecast_4),2) as [2022],Round(AVG(s.Forecast_5),2) as [2023] ");
                sbSQl.Append("From Part_Yrly_Values s");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018<s.AVG_HKSP_2017");


                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart3, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG((s.[AVG_HKSP_2016] -s.[AVG_HKSP_2015]) / s.[AVG_HKSP_2015]), 2) AS[2016],");
                sbSQl.Append("ROUND(AVG((s.[AVG_HKSP_2017] - s.[AVG_HKSP_2016]) / s.[AVG_HKSP_2016]), 2) AS[2017], ");
                sbSQl.Append("ROUND(AVG((s.[AVG_HKSP_2018] -s.[AVG_HKSP_2017]) / s.[AVG_HKSP_2017]),2) AS[2018],");
                sbSQl.Append("ROUND(AVG((s.[Forecast_1] - s.[AVG_HKSP_2018]) / s.[AVG_HKSP_2018]),2) AS[2019],");
                sbSQl.Append("ROUND(AVG((s.[Forecast_2] -s.[Forecast_1]) / iif(s.[Forecast_1]=0,1,s.[Forecast_1])),2) AS[2020],");
                sbSQl.Append("ROUND(AVG((s.[Forecast_3] -s.[Forecast_2]) / iif(s.[Forecast_2]=0,1,s.[Forecast_2])),2) AS[2021],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_4] - s.[Forecast_3]) / iif( s.[Forecast_3]=0,1, s.[Forecast_3])), 2) AS [2022],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_5] - s.[Forecast_4]) / iif( s.[Forecast_4]=0,1, s.[Forecast_4])), 2) AS [2023] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018 >= s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
                sbSQl = new StringBuilder();
                sbSQl.Append("SELECT  ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2016] - s.[AVG_HKSP_2015]) /  s.[AVG_HKSP_2015]), 2) AS[2016],");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2017] - s.[AVG_HKSP_2016]) /  s.[AVG_HKSP_2016]), 2) AS[2017], ");
                sbSQl.Append("ROUND(AVG(( s.[AVG_HKSP_2018] - s.[AVG_HKSP_2017]) /  s.[AVG_HKSP_2017]),2) AS[2018],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_1] - s.[AVG_HKSP_2018]) /  s.[AVG_HKSP_2018]),2) AS[2019],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_2] - s.[Forecast_1]) /iif( s.[Forecast_1]=0,1, s.[Forecast_1])),2) AS[2020],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_3] - s.[Forecast_2]) /iif( s.[Forecast_2]=0,1, s.[Forecast_2])),2) AS[2021],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_4] - s.[Forecast_3]) / iif( s.[Forecast_3]=0,1, s.[Forecast_3])), 2) AS [2022],");
                sbSQl.Append("ROUND(AVG(( s.[Forecast_5] - s.[Forecast_4]) / iif( s.[Forecast_4]=0,1, s.[Forecast_4])), 2) AS [2023] ");
                sbSQl.Append(" FROM PART_YRLY_VALUES s ");
                sbSQl.Append(" Where " + Convert.ToString(sWherCaluse.ToString().Equals("") ? "" : sWherCaluse.ToString() + " and ") + "");
                sbSQl.Append(" " + Convert.ToString(sWhereCaluses.ToString().Equals("") ? "" : sWhereCaluses.ToString() + " and ") + "");
                sbSQl.Append(" s.AVG_HKSP_2018 < s.AVG_HKSP_2017");

                dtd = clscon.dtGetData(sbSQl.ToString());
                if (dtd.Rows.Count > 0)
                {
                    dtdata = new DataTable();
                    dtdata = ConvertRowsToColForChart(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart4, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                }
            }
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


        private void txtincper_TextChanged(object sender, EventArgs e)
        {
            if (txtincper.Text.Length <= 0)
            {
                return;
            }
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            StringBuilder sbSQl = new StringBuilder();
            string sASCPartNo = "";
            string sDESCPartNo = "";
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            string sWherClause = GetDynamicWhereClause();
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "Per");
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "Per");
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30INCDEC(comIncrease.Text, Convert.ToDecimal(txtincper.Text.Equals("") ? "0" : txtincper.Text), "Per");
            }
            else if (txtincper.Text.Equals(""))
            {
                GetChartData("");
                return;
            }
            else
            {
                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                    sbSQl.Append("  (s.delta)  AS[2018]  from GRID s ");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],Round((delta), 2) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                    sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");

                }
                else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                    sbSQl.Append(" Round((s.delta), 2) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                    sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");

                }
                sbSQl.Append(" where s.[date Id]='" + sMoath + "'");
                sbSQl.Append(" and s.[Delta] ");
                sbSQl.Append("" + Convert.ToString(comtotal.Text.Equals("") ? " " : comIncrease.Text + txtincper.Text) + "");
                sbSQl.Append(sWherClause.Equals("") ? "" : " and " + sWherClause);
                dtd = clscon.dtGetData(sbSQl.ToString());
                dtdata = new DataTable();
                dtdata = ConvertRowsToColForChart(dtd);
                BindChartwithdata("KPI Chart 2", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);

                sDESCPartNo = CreateInCalsue(dtdata);
                sbSQl = new StringBuilder();
                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append("SELECT DISTINCT  ");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT ");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                    sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
                }
                else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT  ");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                    sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");

                }
                sbSQl.Append(" FROM PART_YRLY_CHANGE s");
                sbSQl.Append(" where s.[Delta_all_part] ");
                sbSQl.Append("" + Convert.ToString(comtotal.Text.Equals("") ? " " : comtotal.Text + txttotal.Text) + "");
                sbSQl.Append(sWherClause.Equals("") ? "" : " and " + sWherClause);
                dtd = clscon.dtGetData(sbSQl.ToString());
                dtdata = new DataTable();
                dtdata = ConvertRowsToColForChart(dtd);
                BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());

                //DATAGRIDBinding(sASCPartNo, sDESCPartNo); 
            }

        }

        private void txttotal_TextChanged(object sender, EventArgs e)
        {
            string sASCPartNo = "";
            string sDESCPartNo = "";
            if (txttotal.Text.Length <= 0)
            {
                return;
            }
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            StringBuilder sbSQl = new StringBuilder();
            string sWherClause = GetDynamicWhereClause();
            if (btntb10.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop10INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (btntb5.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop5INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (btntb30.BackColor.Equals(Color.FromArgb(27, 161, 226)))
            {
                BindTop30INCDEC(comtotal.Text, Convert.ToDecimal(txttotal.Text.Equals("") ? "0" : txttotal.Text), "");
            }
            else if (txttotal.Text.Equals(""))
            {
                GetChartData("");
                return;
            }
            else
            {

                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                    sbSQl.Append("  (s.delta)  AS[2018]  from GRID s ");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],Round((delta), 2) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                    sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");

                }
                else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                {
                    sbSQl.Append(" SELECT  s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                    sbSQl.Append(" Round((s.delta), 2) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                    sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");

                }
                sbSQl.Append(" where s.[date Id]='" + sMoath + "'");
                sbSQl.Append(" and s.[Delta] ");
                sbSQl.Append("" + Convert.ToString(comtotal.Text.Equals("") ? " " : comtotal.Text + txttotal.Text) + "");
                sbSQl.Append(sWherClause.Equals("") ? "" : " and " + sWherClause);
                dtd = clscon.dtGetData(sbSQl.ToString());
                dtdata = new DataTable();
                dtdata = ConvertRowsToColForChart(dtd);
                BindChartwithdata("KPI Chart 2", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);
                sbSQl = new StringBuilder();

                if (comForcastYtd.Text.Equals("Previous To Current"))
                {
                    sbSQl.Append("SELECT DISTINCT ");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
                }
                else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT ");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                    sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
                }
                else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                {
                    sbSQl.Append("SELECT DISTINCT  ");
                    sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                    sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                    sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                    sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                    sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                    sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");

                }
                sbSQl.Append(" FROM PART_YRLY_CHANGE s");
                sbSQl.Append(" where  s.[Delta_all_part] ");
                sbSQl.Append("" + Convert.ToString(comtotal.Text.Equals("") ? " " : comtotal.Text + txttotal.Text) + "");
                sbSQl.Append(sWherClause.Equals("") ? "" : " and " + sWherClause);
                dtd = clscon.dtGetData(sbSQl.ToString());
                dtdata = new DataTable();
                dtdata = ConvertRowsToColForChart(dtd);
                BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());

                //DATAGRIDBinding(sASCPartNo, sDESCPartNo); 
            }

        }

        public void BindTop5INCDEC(string sSign, decimal dValue, string sType)
        {

            StringBuilder sbSQl = new StringBuilder();
            string sASCPartNo = "";
            string sDESCPartNo = "";
            sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            if (dValue.Equals(0))
            {
                BindTop5();
                return;
            }
            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + "(s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]  like '2T%'";
                }
            }
            if (txtincper.Text.Equals(""))
            {
                BindChartsINCDSC(comtotal.Text, "TOP 5", false);
            }
            else
            {
                BindChartsINCDSC(comIncrease.Text, "TOP 5", true);
            }
        }
        public void BindTop10INCDEC(string sSign, decimal dValue, string sType)
        {

            StringBuilder sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            if (dValue.Equals(0))
            {
                BindTop10();
                return;
            }
            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + "(s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]  like '2T%'";
                }
            }

            if (txtincper.Text.Equals(""))
            {
                BindChartsINCDSC(comtotal.Text, "TOP 10", false);
            }
            else
            {
                BindChartsINCDSC(comIncrease.Text, "TOP 10", true);
            }
        }
        public void BindTop30INCDEC(string sSign, decimal dValue, string sType)
        {

            StringBuilder sbSQl = new StringBuilder();
            sbSQl = new StringBuilder();
            string sWHerC = GetDynamicWhereClause();
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            if (dValue.Equals(0))
            {
                BindTop30();
                return;
            }
            if (ComMCSelection.Text.Equals("All Marketing Code"))
            {
                if (btnpc.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + "(s.[Marketing Code] like '1%' or s.[Marketing Code] like 'SP%')";
                }
                else if (btntruks.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " (s.[Marketing Code] like '2T%' or s.[Marketing Code] like '2L%')";
                }
                else if (btnvans.BackColor.Equals(Color.FromArgb(27, 161, 226)))
                {
                    sWHerC = sWHerC.Equals("") ? "" : sWHerC + " and";
                    sWHerC = sWHerC + " s.[Marketing Code]  like '2T%'";
                }
            }
            if (txtincper.Text.Equals(""))
            {
                BindChartsINCDSC(comtotal.Text, "TOP 30", false);
            }
            else
            {
                BindChartsINCDSC(comIncrease.Text, "TOP 30", true);
            }
        }
        public void BindChartsINCDSC(string sSingn, string sTopcount, bool isIncinPer)
        {
            StringBuilder sbSQl = new StringBuilder();
            string sASCPartNo = "";
            string sDESCPartNo = "";
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            string sWHerC = GetDynamicWhereClause();
            DataTable dtd = new DataTable();
            DataTable dtdata = new DataTable();
            if (sSingn.Equals(""))
            {
                bindTopBottomData(sWHerC, sTopcount);
            }
            else
            {
                if (!isIncinPer)
                {
                    if (comForcastYtd.Text.Equals("Previous To Current"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                        sbSQl.Append("  (s.delta)  AS[2018]  from GRID s ");
                    }
                    else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],Round((delta), 2) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                        sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");

                    }
                    else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                        sbSQl.Append(" Round((s.delta), 2) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                        sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");

                    }
                    sbSQl.Append(" where s.[date Id]='" + sMoath + "'");
                    sbSQl.Append(" and s.[Delta] ");
                    sbSQl.Append("" + Convert.ToString(comtotal.Text.Equals("") ? " " : comtotal.Text + txttotal.Text) + "");
                    sbSQl.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
                    sbSQl.Append("  order by s.[Delta]  DESC,s.[Part No] ");
                    dtd = clscon.dtGetData(sbSQl.ToString());
                    dtdata = new DataTable();
                    dtdata = convertRowsToColFotTOPBottom(dtd);
                    BindChartwithdata("KPI Chart 2", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);

                    sDESCPartNo = CreateInCalsue(dtdata);
                    sbSQl = new StringBuilder();
                    if (comForcastYtd.Text.Equals("Previous To Current"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
                    }
                    else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                        sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                        sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                        sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                        sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
                    }
                    else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                        sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                        sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                        sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                        sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");

                    }
                    sbSQl.Append(" FROM PART_YRLY_CHANGE s");
                    sbSQl.Append(" where s.snr in(" + sDESCPartNo + ") ");
                    sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);
                    sbSQl.Append(" order by ROUND(s.Change_2018,2)  DESC,s.snr ");
                    dtd = clscon.dtGetData(sbSQl.ToString());
                    dtdata = new DataTable();
                    dtdata = convertRowsToColFotTOPBottom(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                    panel10.Visible = true;
                    panel9.Visible = true;
                    DATAGRIDBinding(sASCPartNo, sDESCPartNo, sWHerC);

                }
                else if (isIncinPer)
                {
                    if (comForcastYtd.Text.Equals("Previous To Current"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                        sbSQl.Append("  (s.delta)  AS[2018]  from GRID s ");
                    }
                    else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],Round((delta), 2) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                        sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");

                    }
                    else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                    {
                        sbSQl.Append(" SELECT " + sTopcount + " s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                        sbSQl.Append(" Round((s.delta), 2) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                        sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");

                    }
                    sbSQl.Append(" where [date Id]='" + sMoath + "'");
                    sbSQl.Append(" and s.[Delta]");
                    sbSQl.Append("" + Convert.ToString(comIncrease.Text.Equals("") ? " " : comIncrease.Text + txtincper.Text) + "");
                    sbSQl.Append("  order by s.[Delta] ,s.[Part No] ");
                    dtd = clscon.dtGetData(sbSQl.ToString());
                    dtdata = new DataTable();
                    dtdata = convertRowsToColFotTOPBottom(dtd);
                    BindChartwithdata("KPI Chart 2", dtdata, chart3, "ChartArea1", dtdata.Columns[0].ColumnName);

                    sASCPartNo = CreateInCalsue(dtdata);
                    sbSQl = new StringBuilder();
                    if (comForcastYtd.Text.Equals("Previous To Current"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
                    }
                    else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                        sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                        sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                        sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                        sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
                    }
                    else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
                    {
                        sbSQl.Append("SELECT DISTINCT s.SNR,");
                        sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                        sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                        sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                        sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                        sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                        sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");

                    }
                    sbSQl.Append(" FROM PART_YRLY_CHANGE s");
                    sbSQl.Append(" where s.snr in(" + sASCPartNo + ") ");
                    sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);
                    sbSQl.Append(" order by ROUND(s.Change_2018,2) ,s.snr ");
                    dtd = clscon.dtGetData(sbSQl.ToString());
                    dtdata = new DataTable();
                    dtdata = convertRowsToColFotTOPBottom(dtd);
                    BindChartwithdata("KPI Chart 1", dtdata, chart4, "ChartArea1", dtdata.Columns[0].ColumnName.ToString());
                    panel10.Visible = true;
                    panel9.Visible = true;
                    DATAGRIDBinding(sASCPartNo, sDESCPartNo, sWHerC);
                }
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
                        dtdata.Columns.Add(new DataColumn(dtd.Rows[i][0].ToString(), typeof(string)));
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
                    string sText = string.Format("Part No:  {0} \nKost:\t€ {1}", sSNR, dataPoint.YValues[0]);
                    buttonToolTip.SetToolTip(chart2, sText);
                    break;
            }
        }
        private void chart3_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            // Check selected chart element and set tooltip text for it
            switch (e.HitTestResult.ChartElementType)
            {
                case ChartElementType.DataPoint:
                    var dataPoint = e.HitTestResult.Series.Points[e.HitTestResult.PointIndex];
                    string sSNR = e.HitTestResult.Series.Name;
                    string sText = string.Format("Part No:  {0} \nKost:\t{1}", sSNR, dataPoint.YValues[0]);
                    buttonToolTip.SetToolTip(chart3, sText);
                    break;
            }
        }
        private void chart4_GetToolTipText(object sender, ToolTipEventArgs e)
        {
            // Check selected chart element and set tooltip text for it
            switch (e.HitTestResult.ChartElementType)
            {
                case ChartElementType.DataPoint:
                    string sSNR = e.HitTestResult.Series.Name;
                    var dataPoint = e.HitTestResult.Series.Points[e.HitTestResult.PointIndex];
                    string sText = string.Format("Part No:  {0} \nKost:\t€ {1}%", sSNR, dataPoint.YValues[0]);
                    buttonToolTip.SetToolTip(chart4, sText);
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
            //if (!(char.IsDigit(e.KeyChar) || e.KeyChar == (char)Keys.Back || e.KeyChar == '.'))
            //{ e.Handled = true; }
            //TextBox txtDecimal = sender as TextBox;
            //if (e.KeyChar == '.' && txtDecimal.Text.Contains("."))
            //{
            //    e.Handled = true;
            //}
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
            panel10.Visible = true;
            btnexport.Enabled = true;
            panel9.Visible = true;
            SetPanelLocation(false);
            BindTop30();
        }

        private void DATAGRIDBinding(string sASCPartNo, string sDESCPartNo, string sWHerC)
        {
            panel9.Visible = true;
            panel10.Visible = true;
            panel9.BringToFront();
            panel10.BringToFront();
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
                    sbSQL.Append("select distinct [Marketing Code],[Part No],[Part Name],[Supplier Name],[Manufacturing cost 2017 Year per Part],[Current month Manufacturing cost per part],");
                    sbSQL.Append("[Delta Value per Part Current Year vs 2017],[Delta Manufacturing cost current year vs 2017], [Delta manufacturing cost forecast 2019 vs 2018],");
                    sbSQL.Append("[Delta manufacturing cost forecast 2020 vs 2019],[Delta] FROM Grid s where s.[Part No] in (" + sDESCPartNo + ")");
                    sbSQL.Append(" and [date Id]='" + sMoath + "'");
                    sbSQL.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
                    sbSQL.Append(" order By s.[Delta] desc");
                    DGIncDATA.DataSource = clscon.dtGetData(sbSQL.ToString());
                    DGIncDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                    DGIncDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    DGIncDATA.Columns["Delta"].Visible = false;
                    DGIncDATA.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlDark;
                }
            }
            catch (Exception ex)
            { }
            try
            {
                if (!sASCPartNo.Equals(""))
                {
                    sbSQL = new StringBuilder();

                    sbSQL.Append("select distinct [Marketing Code],[Part No],[Part Name],[Supplier Name],[Manufacturing cost 2017 Year per Part],[Current month Manufacturing cost per part],");
                    sbSQL.Append("[Delta Value per Part Current Year vs 2017],[Delta Manufacturing cost current year vs 2017], [Delta manufacturing cost forecast 2019 vs 2018],");
                    sbSQL.Append("[Delta manufacturing cost forecast 2020 vs 2019],[Delta] FROM Grid s where s.[Part No] in (" + sASCPartNo + ") ");
                    sbSQL.Append(" and s.[date Id]='" + sMoath + "'");
                    sbSQL.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
                    sbSQL.Append(" order By s.[Delta] ASC");
                    DGDESCDATA.DataSource = clscon.dtGetData(sbSQL.ToString());
                    DGDESCDATA.Columns["Delta"].Visible = false;
                    DGDESCDATA.ColumnHeadersDefaultCellStyle.BackColor = Color.Black;
                    DGDESCDATA.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    DGDESCDATA.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.ControlDark;
                }
            }
            catch (Exception ex)
            { }

        }
        private void SetPanelLocation(bool isLocationChange)
        {
            if (isLocationChange)
            {
                panel4.Location = new Point(704, 488);
                panel2.Location = new Point(4, 487);
                panel7.Location = new Point(4, 485);
            }
            else
            {
                panel4.Location = new Point(703, 629);
                panel2.Location = new Point(5, 631);
                panel7.Location = new Point(5, 630);
                panel10.Location = new Point(7, 923);
                panel9.Location = new Point(7, 476);
            }
        }



        private void btnexport_Click(object sender, EventArgs e)
        {
            btntb10.BackColor = Color.FromArgb(64, 64, 64);
            btntb5.BackColor = Color.FromArgb(64, 64, 64);
            btnexport.BackColor = Color.FromArgb(27, 161, 226);
            btntb30.BackColor = Color.FromArgb(64, 64, 64);
            Btnall.BackColor = Color.FromArgb(64, 64, 64);
            clscon.ExportData((DataTable)DGIncDATA.DataSource, (DataTable)DGDESCDATA.DataSource);
        }

        private void comIncrease_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        public void bindTopBottomData(string sWHerC, string sTopNo)
        {
            StringBuilder sbSQl = new StringBuilder();
            string sASCPArtNo = "";
            string[] sdate = sMonat.Split('-');
            DateTime dt = new DateTime(Convert.ToInt32(sdate[1]), Convert.ToInt32(sdate[0]), 1);
            string sMoath = dt.ToString("MMM-yyyy");
            string sDESCPartNo = "";
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append(" SELECT  distinct " + sTopNo + " s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                sbSQl.Append("  VAL(s.delta)  AS[2018]  from GRID s ");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
            {
                sbSQl.Append(" SELECT  distinct  " + sTopNo + " s.[Part No], VAL(s.[Delta]) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
            {
                sbSQl.Append(" SELECT distinct " + sTopNo + " s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                sbSQl.Append(" VAL(s.[Delta]) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");
            }

            sbSQl.Append(" where [date Id]='" + sMoath + "'");
            sbSQl.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
            sbSQl.Append("  order by VAL(s.[Delta]) DESC ,s.[Part No] ");
            DataTable dtd = clscon.dtGetData(sbSQl.ToString());
            DataTable dtdata = new DataTable();
            dtdata = convertRowsToColFotTOPBottom(dtd);
            sDESCPartNo = CreateInCalsue(dtdata);

            BindChartwithdata("KPI Chart 1", dtdata, chart2, "ChartArea1", dtdata.Columns[0].ColumnName);
            sbSQl = new StringBuilder();

            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append(" SELECT distinct " + sTopNo + " s.[Part No],s.[Delta all Parts 2016 Vs 2015] as [2016], s.[Delta all Parts 2017 Vs 2016] as [2017], ");
                sbSQl.Append("  VAL(s.[Delta])  AS[2018]  from GRID s ");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
            {
                sbSQl.Append(" SELECT distinct  " + sTopNo + " s.[Part No], VAL(s.[Delta]) as [2018], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                sbSQl.Append(" Round((delta5), 2) AS[2021],Round((delta4), 2) AS[2022],Round((delta5), 2) AS[2023]   from GRID s ");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
            {
                sbSQl.Append(" SELECT  distinct " + sTopNo + " s.[Part No],Round((s.[Delta all Parts 2016 Vs 2015]), 2) as [2016], Round((s.[Delta all Parts 2017 Vs 2016]), 2) as [2017],");
                sbSQl.Append("  VAL(s.[Delta]) AS[2018],Round((s.delta4), 2) AS[2022], Round((delta1), 2) as [2019], Round((delta2), 2) as [2020], ");
                sbSQl.Append(" Round((delta3), 2) AS[2021],Round((delta5), 2) AS[2023]  from GRID s");

            }
            sbSQl.Append(" where [date Id]='" + sMoath + "'");
            sbSQl.Append(sWHerC.Equals("") ? "" : " and " + sWHerC);
            sbSQl.Append("  order by  VAL(s.[Delta]) ,s.[Part No] ");
            dtd = clscon.dtGetData(sbSQl.ToString());
            dtdata = convertRowsToColFotTOPBottom(dtd);
            sASCPArtNo = CreateInCalsue(dtdata);
            BindChartwithdata("KPI Chart 1", dtdata, chart3, "ChartArea1", dtdata.Columns[0].ColumnName);

            sbSQl = new StringBuilder();
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
            }
            sbSQl.Append(" FROM PART_YRLY_CHANGE s");
            sbSQl.Append(" where s.snr in(" + sDESCPartNo + ") ");
            sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);
            sbSQl.Append(" order by ROUND(s.Change_2018,2)  DESC,s.snr ");
            dtd = clscon.dtGetData(sbSQl.ToString());

            dtdata = new DataTable();
            dtdata = convertRowsToColFotTOPBottom(dtd);
            BindChartwithdata("KPI Chart 1", dtdata, chart1, "ChartArea1", dtdata.Columns[0].ColumnName);

            sbSQl = new StringBuilder();
            if (comForcastYtd.Text.Equals("Previous To Current"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018]");
            }
            else if (comForcastYtd.Text.Equals("Current To Future (5 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
            }
            else if (comForcastYtd.Text.Equals("Previous To Future (5 Year)"))
            {
                sbSQl.Append("SELECT DISTINCT s.SNR,");
                sbSQl.Append(" ROUND(s.Change_2017, 2) AS[2017],");
                sbSQl.Append(" ROUND(s.Change_2018, 2) AS[2018],");
                sbSQl.Append(" ROUND(s.Change_2019, 2) AS[2019],");
                sbSQl.Append(" ROUND(s.Change_2020, 2) AS[2020],");
                sbSQl.Append(" ROUND(s.Change_2021, 2) AS[2021],ROUND(s.Change_2022, 2) AS[2022],");
                sbSQl.Append(" ROUND(s.Change_2023, 2) AS[2023] ");
            }
            sbSQl.Append("  FROM  PART_YRLY_CHANGE s where s.Change_2018 < 0 ");
            sbSQl.Append(" and s.snr in(" + sASCPArtNo + ") ");
            sbSQl.Append(sWHerC.Equals("") ? "" : " And " + sWHerC);
            sbSQl.Append("  order by Round(s.Change_2018,2)  ASC,s.snr");
            dtd = clscon.dtGetData(sbSQl.ToString());
            dtdata = new DataTable();
            dtdata = convertRowsToColFotTOPBottom(dtd);
            BindChartwithdata("KPI Chart 1", dtdata, chart4, "ChartArea1", dtdata.Columns[0].ColumnName);
            DATAGRIDBinding(sASCPArtNo, sDESCPartNo, sWHerC);
        }
    }
}