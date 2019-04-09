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

namespace cost_management
{
    public partial class frmWait : Form
    {
        DataTable dtexpo = new DataTable();
        public static Thread th;
        ClsCommonFunction clsFun = new ClsCommonFunction();
        ComboBox comboBox = new ComboBox();
        public OleDbConnection dbCon = new OleDbConnection();

        IniFile ini = new IniFile();
        public frmWait()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            dbCon = new ClsCommonFunction().Openconnection();
        }

        private void frmWait_Load(object sender, EventArgs e)
        {
            th = new Thread(new ThreadStart(GenrateData));
            th.Start();
        }

        public void GenrateData()
        {
            string time = DateTime.Now.ToString("hh:mm:ss tt");
            DataTable dtsummdt = new DataTable();
            DataTable dtsummrise = new DataTable();
            new MVSPMDBInfo().checkdatatabe();
            dtexpo = clsFun.dtGetData("Select * from tblextrapolationfactor");
            if (dtexpo.Rows.Count.Equals(0))
            {
                MessageBox.Show("Please upload extrapolation factor", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.Cancel;
                th = null;
                return;
            }
            CreateTable();

            this.label1.BeginInvoke((MethodInvoker)delegate () { this.label1.Text = "Populating Data ...!!!"; ; });
            DataTable dt = clsFun.dtGetData("select distinct (Mid(monat,4)) as YEARS from  Cost_Management_Datenbasis where (Mid(monat,4))<>2017");
            this.Cursor = NativeMethods.LoadCustomCursor(Application.StartupPath + "\\mb.ico");
            //Calculate PROG_END_DATE
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT DISTINCT Cost_Management_datenbasis.snr, Serie_von, Mcode, Switch( (Mcode like '1%' or  Mcode like 'SP%'), ");
            sbSQL.Append("Format(DateAdd('yyyy', 6, CDate(Serie_von)), 'yyyy-MM-dd'),");
            sbSQL.Append(" Mcode like '2T%', Format(DateAdd('yyyy', 10, CDate(Serie_von)), 'yyyy-MM-dd'),");
            sbSQL.Append("(Mcode like '2U%' or Mcode like '2L%') ,Format(DateAdd('yyyy', 14, CDate(Serie_von)), 'yyyy-MM-dd')) AS PROG_END_DATE ");
            sbSQL.Append(" FROM Cost_Management_datenbasis WHERE((([Serie_bis]) Is Null) and Serie_von<> null )ORDER BY SNR;");

            DataTable dtProgDate = clsFun.dtGetData(sbSQL.ToString());
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    for (int iMonth = 1; iMonth <= 12; iMonth++)
                    {
                        try
                        {
                            dtsummdt = clsFun.GetData(Convert.ToInt32(dr[0].ToString()), iMonth);
                            GlobalVariable.iTotalCOunt = Convert.ToInt32(dtsummdt.Rows.Count);
                            foreach (DataRow drs in dtsummdt.Rows)
                            {
                                StringBuilder sbsSQL = new StringBuilder();
                                DataView view = new DataView(dtProgDate);
                                string sFilter = dtProgDate.Columns[0].ColumnName + " = '" + drs["Part No"].ToString() + "'";
                                view.RowFilter = sFilter;
                                DataTable dtval = view.ToTable();
                                string sProgDate = "";
                                if (dtval.Rows.Count > 0)
                                {
                                    if (!dtval.Rows[0]["PROG_END_DATE"].ToString().Equals(""))
                                        sProgDate = dtval.Rows[0]["PROG_END_DATE"].ToString().Substring(0, 10);
                                    else
                                        sProgDate = "";
                                }
                                else
                                    sProgDate = "";
                                sbsSQL = new StringBuilder();
                                sbsSQL.Append("Insert Into Cost_summary ([Date id],[Part No]	,[Part Name],[Marketing Code],[Series Production Start],[Series Production End],[Manufacturing cost  YTD-2Y Year per Part],");
                                sbsSQL.Append(" [Manufacturing total cost YTD-2Y]	, [Delta value per part YTD Vs YTD-2Y],[% change YTD vs YTD-2Y],[Delta all Parts YTD Vs YTD-2],");
                                sbsSQL.Append(" [Manufacturing cost YTD-1Y Year per Part],[Manufacturing total cost YTD-1Y]	,[Delta value per part YTD Vs YTD-1Y]	,[% change YTD vs YTD-1Y]	,");
                                sbsSQL.Append(" [Part Quantity Current Month],[Part Quantity YTD],[Current Month Manufacturing Cost Total],[YTD Manufacturing Cost per Year]	,[Part Sobsl],[Creation date],[Dispo Code],[Supplier Code],");
                                sbsSQL.Append(" [Supplier Name]	,[Gross List Price  BLP / VLP]	,[Return Value (RW)],[Trade Value],[Price Date],[Gross Revenue Current Month],[Gross Revenue Current YTD],[Factory Cost SKW],[tax value approach],");
                                sbsSQL.Append(" [SKW without  SWA],[Total Cost] , [Total Cost excl Return Cost] , [Total Cost YTD] ,[Material Cost],[Return Cost],[GuK],[Package Absolute],[VTKP abs],[VTKF abs],[Result Factory],");
                                sbsSQL.Append(" [CORE value Supplior],[Credit Voucher Old Part],");
                                sbsSQL.Append(" [Discount Group],[NET Revenue per VLP] ,[NET Revenue YTD] ,[Net revenue without return value current Month] ,");
                                sbsSQL.Append(" [Net Revenu RETURN YTD],[Result],[Delta all Parts YTD Vs YTD-1],[Current Month Manufacturing Cost per Part],[Dispo Name],[Part Technique],");
                                sbsSQL.Append(" [Total Cost per Part],[Model Type Baureihe],[Prognosed Series End Date],[Delta all part Current Year],");
                                sbsSQL.Append(" [Delta Value per Part Current Year]) Values( ");
                                sbsSQL.Append("'" + drs["MONAT"].ToString().Replace(".", "-") + "','" + drs["Part No"].ToString() + "', '" + drs["Part Name"].ToString() + "',");
                                sbsSQL.Append("'" + drs["Marketing Code"].ToString() + "','" + drs["Serie_von"].ToString() + "', '" + drs["Serie_bis"].ToString() + "', ");
                                sbsSQL.Append("'" + Convert.ToString(drs["Manufacturing cost YTD-2Y Year per Part"].ToString().Equals("") ? "0" : drs["Manufacturing cost YTD-2Y Year per Part"].ToString()) + "', ");
                                sbsSQL.Append("'" + Convert.ToString(drs["Manufacturing total cost YTD-2Y"].ToString().Equals("") ? "0" : drs["Manufacturing total cost YTD-2Y"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Delta value per part YTD Vs YTD-2Y"].ToString().Equals("") ? "0" : drs["Delta value per part YTD Vs YTD-2Y"].ToString()) + "',");
                                sbsSQL.Append("'" + drs["% change YTD vs YTD-2Y"].ToString() + "','" + drs["Delta all Parts YTD Vs YTD-2"].ToString() + "',");
                                sbsSQL.Append("'" + drs["Manufacturing cost YTD-1Y Year per Part"].ToString() + "',");
                                sbsSQL.Append(" '" + Convert.ToString(drs["Manufacturing total cost YTD-1Y"].ToString().Equals("") ? "0" : drs["Manufacturing total cost YTD-1Y"].ToString()) + "',");
                                sbsSQL.Append("'" + drs["Delta value per part YTD Vs YTD-1Y"].ToString() + "','" + drs["% change YTD vs YTD-1Y"].ToString() + "',");
                                sbsSQL.Append("'" + drs["Part Quantity Current Month"].ToString() + "','" + drs["YTD"].ToString() + "', ");
                                sbsSQL.Append("'" + Convert.ToString(drs["Current Month Manufacturing Cost Total"].ToString().Equals("") ? "0" : drs["Current Month Manufacturing Cost Total"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["YTD Manufacturing Cost per Year"].ToString().Equals("") ? "0" : drs["YTD Manufacturing Cost per Year"].ToString()) + "',");
                                sbsSQL.Append("'" + drs["Part Sobsl"].ToString() + "','" + drs["Creation date"].ToString() + "','" + drs["Dispo Code"].ToString() + "',");
                                sbsSQL.Append("'" + drs["Supplier Code"].ToString() + "','" + drs["Supplier Name"].ToString() + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Gross List Price  BLP / VLP"].ToString().Equals("") ? "0" : drs["Gross List Price  BLP / VLP"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Return Value(RW)"].ToString().Equals("") ? "0" : drs["Return Value(RW)"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Trade Value"].ToString().Equals("") ? "0" : drs["Trade Value"].ToString()) + "',");
                                sbsSQL.Append("'" + drs["Price Date"].ToString() + "', ");
                                sbsSQL.Append("'" + Convert.ToString(drs["Gross Revenue Current Month"].ToString().Equals("") ? "0" : drs["Gross Revenue Current Month"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Gross Revenue Current YTD"].ToString().Equals("") ? "0" : drs["Gross Revenue Current YTD"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Factory Cost SKW"].ToString().Equals("") ? "0" : drs["Factory Cost SKW"].ToString()) + "', ");
                                sbsSQL.Append("'" + Convert.ToString(drs["tax value approach"].ToString().Equals("") ? "0" : drs["tax value approach"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["SKW without  SWA"].ToString().Equals("") ? "0" : drs["SKW without  SWA"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Total Cost"].ToString().Equals("") ? "0" : drs["Total Cost"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Total Cost excl Return Cost"].ToString().Equals("") ? "0" : drs["Total Cost excl Return Cost"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Total YTD cost"].ToString().Equals("") ? "0" : drs["Total YTD cost"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Material Cost"].ToString().Equals("") ? "0" : drs["Material Cost"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Return Cost"].ToString().Equals("") ? "0" : drs["Return Cost"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["GuK"].ToString().Equals("") ? "0" : drs["GuK"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Package Absolute"].ToString().Equals("") ? "0" : drs["Package Absolute"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["VTKP abs"].ToString().Equals("") ? "0" : drs["VTKP abs"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["VTKF abs"].ToString().Equals("") ? "0" : drs["VTKF abs"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Result Factory"].ToString().Equals("") ? "0" : drs["Result Factory"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["CORE value Supplior"].ToString().Equals("") ? "0" : drs["Result Factory"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Credit Voucher Old Part"].ToString().Equals("") ? "0" : drs["Credit Voucher Old Part"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Discount Group"].ToString().Equals("") ? "0" : drs["Discount Group"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["NET Revenue per VLP"].ToString().Equals("") ? "0" : drs["NET Revenue per VLP"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Net_Revenu_YTD"].ToString().Equals("") ? "0" : drs["Net_Revenu_YTD"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Net revenue without return value current Month"].ToString().Equals("") ? "0" : drs["Net revenue without return value current Month"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Net Revenu Without RETURN YTD"].ToString().Equals("") ? "0" : drs["Net Revenu Without RETURN YTD"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Result"].ToString().Equals("") ? "0" : drs["Result"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Delta all Parts YTD Vs YTD-1"].ToString().Equals("") ? "0" : drs["Delta all Parts YTD Vs YTD-1"].ToString()) + "'");
                                sbsSQL.Append(",'" + Convert.ToString(drs["Current Month Manufacturing Cost per Part"].ToString().Equals("") ? "0" : drs["Current Month Manufacturing Cost per Part"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Dispo Name"].ToString().Equals("") ? "0" : drs["Dispo Name"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Part Technique"].ToString().Equals("") ? "0" : drs["Part Technique"].ToString()) + "'");
                                sbsSQL.Append(",'" + Convert.ToString(drs["Total Cost per Part"].ToString().Equals("") ? "0" : drs["Total Cost per Part"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Model Type Baureihe"].ToString().Equals("") ? "0" : drs["Model Type Baureihe"].ToString()) + "',");
                                sbsSQL.Append("'" + sProgDate + "','" + Convert.ToString(drs["Delta all part Current Year"].ToString().Equals("") ? "0" : drs["Delta all part Current Year"].ToString()) + "',");
                                sbsSQL.Append("'" + Convert.ToString(drs["Delta Value per Part Current Year"].ToString().Equals("") ? "0" : drs["Delta Value per Part Current Year"].ToString()) + "'");
                                sbsSQL.Append(" );");

                                clsFun.AlterDatabase(sbsSQL.ToString());
                            }
                        }
                        catch (Exception EX)
                        {

                        }
                    }
                }

            }
            //delete Duplicate data
            //KeepUniqueRecords();
            //Genrate Forcasting Data
            GenrateForcasting();
            //Genrate Final Data sets for graph
            creteFinalDataset();
            string sendtime = DateTime.Now.ToString("hh:mm:ss tt");
            //lbltimer.Text = time + "  " + sendtime;
            MessageBox.Show("Process Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();

            this.Cursor = Cursors.Default;
            th = null;
        }

        public void KeepUniqueRecords()
        {
            DataTable dt = clsFun.dtGetData("select distinct [Part No] from  Cost_Summary");
            foreach (DataRow dr in dt.Rows)
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append("Select Top 1 a.[Part No] as snr, Max(YR) as yrs, MAX(Mo) as mon  from(SELECT max(Mid([Date Id], 4)) as YR,max(Mid([date Id], 1, 2)) as Mo, [Part No] ");
                sbSQL.Append(" from Cost_Summary  where [Part No] ='" + dr[0].ToString() + "' group by [date Id], [Part No] ");
                sbSQL.Append(" order BY max(Mid([date Id], 4)) Desc, max(Mid([date Id], 1, 2)) desc) as A where a.Yr = (SELECT distinct max(Mid([date Id], 4)) as YR");
                sbSQL.Append(" from Cost_Summary where [Part No] ='" + dr[0].ToString() + "')  Group BY a.[Part No],a.MO,a.YR Order BY a.MO Desc,a.YR Desc");
                DataTable dts = clsFun.dtGetData(sbSQL.ToString());
                if (dts.Rows.Count > 0)
                {
                    string sSQL = "";
                    string sMonth = dts.Rows[0]["mon"].ToString();
                    string sYear = dts.Rows[0]["yrs"].ToString();
                    string sSNR = dts.Rows[0]["Snr"].ToString();
                    if (sYear.Equals(clsFun.GetYearName(1, "-")))
                        sSQL = "delete from Cost_summary where [Part No]='" + sSNR + "' and Mid([Date Id], 1, 2)<>'" + sMonth + "'";
                    else
                        sSQL = "delete from Cost_summary where [Part No]='" + sSNR + "' and Mid([Date Id], 1, 2)<>'" + sMonth + "' and Mid([Date Id], 4)<>'" + sYear + "'";
                    clsFun.AlterDatabase(sSQL);
                }
            }

            string sSQLs = "ALTER TABLE Cost_summary Drop COLUMN sID ";
            clsFun.AlterDatabase(sSQLs);

            sSQLs = "ALTER TABLE Cost_summary Add COLUMN sID COUNTER(1,1)";
            clsFun.AlterDatabase(sSQLs);
        }

        public void CreateTable()
        {
            OleDbConnection dbCon = new OleDbConnection();
            dbCon = new ClsCommonFunction().Openconnection();

            try
            {
                string sSQL = "Drop Table Cost_summary";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append(" Create Table Cost_summary (sId counter,[Date Id]	Text(100),[Marketing Code]	Text(100),[Part No]	Text(100),[Part Name]	Text(100),");
                sbSQL.Append(" [Series Production Start]	Text(100),[Series Production End] Text(100),[Prognosed Series End Date]	Text(100),");
                sbSQL.Append("  [Manufacturing cost  YTD-2Y Year per Part] DECIMAL(18,2),");
                sbSQL.Append(" [Manufacturing total cost YTD-2Y]	DECIMAL(18,2),[Delta value per part YTD Vs YTD-2Y]	DECIMAL(18,2),");
                sbSQL.Append(" [% change YTD vs YTD-2Y]	DECIMAL(18,4),");
                sbSQL.Append(" [Delta all Parts YTD Vs YTD-2] DECIMAL(18,2),[Manufacturing cost YTD-1Y Year per Part] DECIMAL(18,2),");
                sbSQL.Append("[Manufacturing total cost YTD-1Y]	DECIMAL(18,2),[Delta value per part YTD Vs YTD-1Y] DECIMAL(18,2), ");
                sbSQL.Append(" [% change YTD vs YTD-1Y]	DECIMAL(18,4),[Delta all Parts YTD Vs YTD-1] DECIMAL(18,2),");
                sbSQL.Append(" [Current Month Manufacturing Cost per Part]  DECIMAL(18,2),[Current Month Manufacturing Cost Total] DECIMAL(18,2),");
                sbSQL.Append(" [Part Quantity Current Month] Long,[Part Quantity YTD]	long, [YTD Manufacturing Cost per Year] DECIMAL(18,2),");
                sbSQL.Append(" [Delta Value per Part Current Year] DECIMAL(18,2),");
                sbSQL.Append(" [Delta all part Current Year] DECIMAL(18,2),");
                sbSQL.Append(" [Part Sobsl]	Text(100),[Creation date]	Text(100),[Dispo Code] Text(100),[Dispo Name] Text(100),[Part Technique] Text(100),");
                sbSQL.Append(" [Model Type Baureihe] Text(100),[Supplier Code]	Text(100), ");
                sbSQL.Append(" [Supplier Name]	Text(100),[Gross List Price  BLP / VLP]	DECIMAL(18,2),[Return Value (RW)] DECIMAL(18,2),");
                sbSQL.Append(" [Trade Value]	DECIMAL(18,2),[Price Date]	Text(100),	 ");
                sbSQL.Append(" [Gross Revenue Current Month]	DECIMAL(18,2),[Gross Revenue Current YTD]	DECIMAL(18,2),[NET Revenue per VLP] DECIMAL(18,2),");
                sbSQL.Append(" [NET Revenue YTD] DECIMAL(18,2),[Net revenue without return value current Month] DECIMAL(18,2),");
                sbSQL.Append(" [Net Revenu RETURN YTD] DECIMAL(18,2),[Factory Cost SKW]	DECIMAL(18,2),[tax value approach] DECIMAL(18,2),");
                sbSQL.Append(" [SKW without  SWA]	DECIMAL(18,2),[Total Cost per Part]	DECIMAL(18,2),[Total Cost] DECIMAL(18,2),");
                sbSQL.Append(" [Total Cost excl Return Cost] DECIMAL(18,2), [Total Cost YTD]DECIMAL(18,2),[Material Cost] DECIMAL(18,2),[Return Cost] DECIMAL(18,2),	 ");
                sbSQL.Append(" [GuK]	DECIMAL(18,2),[Package Absolute]	DECIMAL(18,2),[VTKP abs]	DECIMAL(18,2),[VTKF abs] DECIMAL(18,2),");
                sbSQL.Append("[Result Factory]	DECIMAL(18,2),[CORE value Supplior]	DECIMAL(18,2),	 ");
                sbSQL.Append(" [Credit Voucher Old Part]	DECIMAL(18,2),[Result] DECIMAL(18,2),[Discount Group]	DECIMAL(18,2)) ");
                OleDbCommand oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //Create Forecasting Data
            try
            {
                string sSQL = "Drop Table TblForcasting";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append(" Create Table TblForcasting ([Part No] Text(100),");
                sbSQL.Append("[Forcasted Manufacturing Cost Year 1] DECIMAL(18,2),[Forcasted Manufacturing Cost Year 2] DECIMAL(18,2)");
                sbSQL.Append(" ,[Forcasted Manufacturing Cost Year 3] DECIMAL(18,2),");
                sbSQL.Append(" [Forecast Manufacturing cost total 1] DECIMAL(18,2)");
                sbSQL.Append(",[Forecast Manufacturing cost total 2] DECIMAL(18,2),[Forecast Manufacturing cost total 3] DECIMAL(18,2),");
                sbSQL.Append(" [Delta manufacturing cost 1] DECIMAL(18,2)");
                sbSQL.Append(",[Delta manufacturing cost 2] DECIMAL(18,2),[Delta manufacturing cost 3] DECIMAL(18,2)");
                sbSQL.Append(" ,[Forecast Quntity 1] DECIMAL(18,2)");
                sbSQL.Append(",[Forecast Quntity 2] DECIMAL(18,2),[Forecast Quntity 3] DECIMAL(18,2),Color_Code DECIMAL(18,0),[Part YTD extrapol] Long,[Manufacturing Material cost Current Year Estimation] DECIMAL(18,2)");
                sbSQL.Append(")");
                OleDbCommand oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { }

            try
            {
                string sSQL = "Drop Table TblDiscSNR";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append(" Create Table TblDiscSNR ([Part_No] Text(100),");
                sbSQL.Append("[date_id] Text(12)");
                sbSQL.Append(")");
                OleDbCommand oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex) { }
        }

        void creteFinalDataset()
        {
            //Create Procedure GRID
            this.label1.BeginInvoke((MethodInvoker)delegate () { this.label1.Text = "Calculating Final Datasets...!!!"; });
            try
            {
                OleDbCommand oleCmd = new OleDbCommand();
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE GRID");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }

                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE GRID  as ");
                sbSqlQuery.Append("SELECT cs.sId, Format(cs.[Date Id],'MMM-yyyy') as [Date Id], cs.[Marketing Code], cs.[Part No], cs.[Part Name], ");
                sbSqlQuery.Append("cs.[Series Production Start], cs.[Series Production End], Format(cs.[Prognosed Series End Date],'yyyy-MM-dd') as [Prognosed Series End Date], ");
                sbSqlQuery.Append("Format(cs.[Manufacturing cost  YTD-2Y Year per Part],'#,##0.00') as [Manufacturing cost  " + clsFun.GetYear(2, "-") + " Year per Part]");
                sbSqlQuery.Append(",Format(cs.[Manufacturing total cost YTD-2Y],'#,##0.00')as [Manufacturing total cost " + clsFun.GetYear(2, "-") + "], ");
                sbSqlQuery.Append(" Format(cs.[Delta value per part YTD Vs YTD-2Y],'#,##0.00')as [Delta value per part " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "], ");
                sbSqlQuery.Append("Format(cs.[% change YTD vs YTD-2Y],'Percent') as [% change " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "], ");
                sbSqlQuery.Append("Format( cs.[Delta all Parts YTD Vs YTD-2],'#,##0.00') as [Delta all Parts " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "],");
                sbSqlQuery.Append("Format(cs.[Manufacturing cost YTD-1Y Year per Part],'#,##0.00') as [Manufacturing cost " + clsFun.GetYear(1, "-") + " Year per Part] ");
                sbSqlQuery.Append(",Format(cs.[Manufacturing total cost YTD-1Y],'#,##0.00') as [Manufacturing total cost " + clsFun.GetYear(1, "-") + "], ");
                sbSqlQuery.Append("Format(cs.[Delta value per part YTD Vs YTD-1Y],'#,##0.00') as [Delta value per part " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "],");
                sbSqlQuery.Append("Format(cs.[% change YTD vs YTD-1Y],'Percent') as [% change " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "], ");
                sbSqlQuery.Append("Format(cs.[Delta all Parts YTD Vs YTD-1],'#,##0.00') as [Delta all Parts " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "], ");
                sbSqlQuery.Append("Format(cs.[Current Month Manufacturing Cost per Part],'#,##0.00')as [Current Month Manufacturing Cost per Part] , ");
                sbSqlQuery.Append("Format(cs.[Current Month Manufacturing Cost Total],'#,##0.00') as [Current Month Manufacturing Cost Total], ");
                sbSqlQuery.Append("Format(cs.[Part Quantity Current Month],'#,##0.00') as [Part Quantity Current Month], cs.[Part Quantity YTD],TblForcasting.[Part YTD Extrapol] as [Forecast Part Quantity Current Year], ");
                sbSqlQuery.Append("Format(cs.[YTD Manufacturing Cost per Year],'#,##0.00') as [YTD Manufacturing Cost per Year],");
                sbSqlQuery.Append("Format(cs.[Delta Value per Part Current Year] ,'#,##0.00') as [Delta Value per Part Current Year vs " + clsFun.GetYear(1, "-") + "]");
                sbSqlQuery.Append(",Format(TblForcasting.[Manufacturing Material cost Current Year Estimation],'#,##0.00') as [Current year estimation vs " + clsFun.GetYear(1, "-") + "],");
                sbSqlQuery.Append("Format(cs.[Delta Value per Part Current Year] * TblForcasting.[Part YTD Extrapol],'#,##0.00') as [Delta Manufacturing cost current year vs " + clsFun.GetYear(1, "-") + "],");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forcasted Manufacturing Cost Year 1] is Null,0,TblForcasting.[Forcasted Manufacturing Cost Year 1]),'#,##0.00') as  [Forcasted Manufacturing Cost Year " + clsFun.GetYear(1, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Manufacturing cost total 1] is null,0,TblForcasting.[Forecast Manufacturing cost total 1]),'#,##0.00') as  [Forecast Manufacturing cost total " + clsFun.GetYear(1, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Quntity 1] is Null,0,TblForcasting.[Forecast Quntity 1]),'#,##0.00') as  [Forecast Quntity " + clsFun.GetYear(1, "+") + "], ");
                sbSqlQuery.Append("Format(Iif(TblForcasting.[Delta manufacturing cost 1] is Null,0,TblForcasting.[Delta manufacturing cost 1]),'#,##0.00') as  [Delta manufacturing cost forecast " + clsFun.GetYear(1, "+") + " vs " + clsFun.GetYear(0, "") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forcasted Manufacturing Cost Year 2] is null,0,TblForcasting.[Forcasted Manufacturing Cost Year 2]),'#,##0.00')  as [Forcasted Manufacturing Cost Year " + clsFun.GetYear(2, "+") + "],");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Manufacturing cost total 2]is null,0,TblForcasting.[Forecast Manufacturing cost total 2]),'#,##0.00') as  [Forecast Manufacturing cost total " + clsFun.GetYear(2, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Quntity 2] is Null,0,TblForcasting.[Forecast Quntity 2]),'#,##0.00') as  [Forecast Quntity " + clsFun.GetYear(2, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Delta manufacturing cost 2]is Null,0,TblForcasting.[Delta manufacturing cost 2]),'#,##0.00') as  [Delta manufacturing cost forecast " + clsFun.GetYear(2, "+") + " vs " + clsFun.GetYear(1, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forcasted Manufacturing Cost Year 3] Is Null,0,TblForcasting.[Forcasted Manufacturing Cost Year 3]),'#,##0.00')  as [Forcasted Manufacturing Cost Year " + clsFun.GetYear(3, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Manufacturing cost total 3]Is Null,0,TblForcasting.[Forecast Manufacturing cost total 3]),'#,##0.00') as  [Forecast Manufacturing cost total " + clsFun.GetYear(3, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Forecast Quntity 3] is Null,0,TblForcasting.[Forecast Quntity 3]),'#,##0.00') as  [Forecast Quntity " + clsFun.GetYear(3, "+") + "], ");
                sbSqlQuery.Append("Format(iif(TblForcasting.[Delta manufacturing cost 3] Is Null,0,TblForcasting.[Delta manufacturing cost 3]),'#,##0.00') as  [Delta manufacturing cost forecast " + clsFun.GetYear(3, "+") + " vs " + clsFun.GetYear(2, "+") + "], ");
                sbSqlQuery.Append("cs.[Part Sobsl], cs.[Creation date], cs.[Dispo Code], cs.[Dispo Name], cs.[Part Technique],");
                sbSqlQuery.Append("cs.[Model Type Baureihe], cs.[Supplier Code], cs.[Supplier Name],Format(cs.[Gross List Price  BLP / VLP],'#,##0.00') as [Gross List Price  BLP / VLP], ");
                sbSqlQuery.Append("Format(cs.[Return Value (RW)],'#,##0.00') as [Return Value (RW)],  Format(cs.[Trade Value],'#,##0.00')as [Trade Value], ");
                sbSqlQuery.Append(" cs.[Price Date], Format(cs.[Gross Revenue Current Month],'#,##0.00') as [Gross Revenue Current Month], ");
                sbSqlQuery.Append("Format(cs.[Gross Revenue Current YTD],'#,##0.00') as [Gross Revenue Current YTD], Format(cs.[NET Revenue per VLP],'#,##0.00') as  [NET Revenue per VLP], ");
                sbSqlQuery.Append("Format( cs.[NET Revenue YTD],'#,##0.00') as [NET Revenue YTD], ");
                sbSqlQuery.Append("Format(cs.[Net revenue without return value current Month],'#,##0.00') as [Net revenue without return value current Month] , ");
                sbSqlQuery.Append(" Format(cs.[Net Revenu RETURN YTD],'#,##0.00') as [Net Revenu RETURN YTD],Format(cs.[Factory Cost SKW],'#,##0.00') as [Factory Cost SKW], ");
                sbSqlQuery.Append("Format(cs.[tax value approach],'#,##0.00') as [tax value approach], Format(cs.[SKW without  SWA],'#,##0.00') as [SKW without  SWA], ");
                sbSqlQuery.Append(" Format(cs.[Total Cost per Part],'#,##0.00') as [Total Cost per Part],Format( cs.[Total Cost],'#,##0.00') as [Total Cost], ");
                sbSqlQuery.Append("Format(cs.[Total Cost excl Return Cost],'#,##0.00') as [Total Cost excl Return Cost],Format( cs.[Total Cost YTD],'#,##0.00') as [Total Cost YTD], ");
                sbSqlQuery.Append("Format(cs.[Material Cost],'#,##0.00') as [Material Cost], cs.[Return Cost], ");
                sbSqlQuery.Append("Format(cs.GuK,'#,##0.00') as GuK, Format(cs.[Package Absolute],'#,##0.00') as [Package Absolute], ");
                sbSqlQuery.Append(" Format( cs.[VTKP abs],'#,##0.00') as [VTKP abs], Format( cs.[Result Factory],'#,##0.00') as [Result Factory], Format( cs.[CORE value Supplior],'#,##0.00') as [CORE value Supplior], ");
                sbSqlQuery.Append(" Format( cs.[Credit Voucher Old Part],'#,##0.00') as [Credit Voucher Old Part], Format(cs.Result,'#,##0.00') as Result,Format( cs.[Discount Group],'#,##0.00') as [Discount Group]");
                sbSqlQuery.Append(" ,(cs.[Delta Value per Part Current Year] * TblForcasting.[Part YTD Extrapol]) as [delta],TblForcasting.[Forecast Manufacturing cost total 1] as [FMCT1],TblForcasting.[Forecast Manufacturing cost total 2] as [FMCT2] ");
                sbSqlQuery.Append(",TblForcasting.[Forecast Manufacturing cost total 3] as [FMCT3]");
                sbSqlQuery.Append(" ,TblForcasting.[Delta manufacturing cost 1] as [Delta1],TblForcasting.[Delta manufacturing cost 2] as [Delta2],TblForcasting.[Delta manufacturing cost 3] as [Delta3] ");
                sbSqlQuery.Append(" ,cs.[Delta all Parts YTD Vs YTD-1] as [delta-1],cs.[Delta all Parts YTD Vs YTD-2] as [delta-2],TblForcasting.Color_Code,");
                sbSqlQuery.Append("cs.[Manufacturing cost  YTD-2Y Year per Part] as [Manufacturing cost  " + clsFun.GetYear(2, "-") + " Year per Part_1]");
                sbSqlQuery.Append(",cs.[Manufacturing total cost YTD-2Y]as [Manufacturing total cost " + clsFun.GetYear(2, "-") + "_1], ");
                sbSqlQuery.Append(" cs.[Delta value per part YTD Vs YTD-2Y]as [Delta value per part " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "_1], ");
                sbSqlQuery.Append("cs.[% change YTD vs YTD-2Y] as [% change " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "_1], ");
                sbSqlQuery.Append(" cs.[Delta all Parts YTD Vs YTD-2] as [Delta all Parts " + clsFun.GetYear(2, "-") + " Vs " + clsFun.GetYear(3, "-") + "_1],");
                sbSqlQuery.Append("cs.[Manufacturing cost YTD-1Y Year per Part] as [Manufacturing cost " + clsFun.GetYear(1, "-") + " Year per Part_1] ");
                sbSqlQuery.Append(",cs.[Manufacturing total cost YTD-1Y] as [Manufacturing total cost " + clsFun.GetYear(1, "-") + "_1], ");
                sbSqlQuery.Append("cs.[Delta value per part YTD Vs YTD-1Y] as [Delta value per part " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "_1],");
                sbSqlQuery.Append("cs.[% change YTD vs YTD-1Y] as [% change " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "_1], ");
                sbSqlQuery.Append("cs.[Delta all Parts YTD Vs YTD-1] as [Delta all Parts " + clsFun.GetYear(1, "-") + " Vs " + clsFun.GetYear(2, "-") + "_1], ");
                sbSqlQuery.Append("cs.[Current Month Manufacturing Cost per Part]as [Current Month Manufacturing Cost per Part_1] , ");
                sbSqlQuery.Append("cs.[Current Month Manufacturing Cost Total] as [Current Month Manufacturing Cost Total_1], ");
                sbSqlQuery.Append("cs.[Part Quantity Current Month] as [Part Quantity Current Month_1], cs.[Part Quantity YTD] as [Part Quantity YTD_1],TblForcasting.[Part YTD Extrapol] as [Forecast Part Quantity Current Year_1], ");
                sbSqlQuery.Append("cs.[YTD Manufacturing Cost per Year] as [YTD Manufacturing Cost per Year_1],");
                sbSqlQuery.Append("cs.[Delta Value per Part Current Year]  as [Delta Value per Part Current Year vs " + clsFun.GetYear(1, "-") + "_1]");
                sbSqlQuery.Append(",TblForcasting.[Manufacturing Material cost Current Year Estimation] as [Current year estimation vs " + clsFun.GetYear(1, "-") + "_1],");
                sbSqlQuery.Append("cs.[Delta Value per Part Current Year] * TblForcasting.[Part YTD Extrapol] as [Delta Manufacturing cost current year vs " + clsFun.GetYear(1, "-") + "_1],");
                sbSqlQuery.Append("iif(TblForcasting.[Forcasted Manufacturing Cost Year 1] is Null,0,TblForcasting.[Forcasted Manufacturing Cost Year 1]) as  [Forcasted Manufacturing Cost Year " + clsFun.GetYear(1, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Manufacturing cost total 1] is null,0,TblForcasting.[Forecast Manufacturing cost total 1]) as  [Forecast Manufacturing cost total " + clsFun.GetYear(1, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Quntity 1] is Null,0,TblForcasting.[Forecast Quntity 1]) as  [Forecast Quntity " + clsFun.GetYear(1, "+") + "_1], ");
                sbSqlQuery.Append("Iif(TblForcasting.[Delta manufacturing cost 1] is Null,0,TblForcasting.[Delta manufacturing cost 1]) as  [Delta manufacturing cost forecast " + clsFun.GetYear(1, "+") + " vs " + clsFun.GetYear(0, "") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forcasted Manufacturing Cost Year 2] is null,0,TblForcasting.[Forcasted Manufacturing Cost Year 2])  as [Forcasted Manufacturing Cost Year " + clsFun.GetYear(2, "+") + "_1],");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Manufacturing cost total 2]is null,0,TblForcasting.[Forecast Manufacturing cost total 2]) as  [Forecast Manufacturing cost total " + clsFun.GetYear(2, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Quntity 2] is Null,0,TblForcasting.[Forecast Quntity 2]) as  [Forecast Quntity " + clsFun.GetYear(2, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Delta manufacturing cost 2]is Null,0,TblForcasting.[Delta manufacturing cost 2]) as  [Delta manufacturing cost forecast " + clsFun.GetYear(2, "+") + " vs " + clsFun.GetYear(1, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forcasted Manufacturing Cost Year 3] Is Null,0,TblForcasting.[Forcasted Manufacturing Cost Year 3])  as [Forcasted Manufacturing Cost Year " + clsFun.GetYear(3, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Manufacturing cost total 3]Is Null,0,TblForcasting.[Forecast Manufacturing cost total 3]) as  [Forecast Manufacturing cost total " + clsFun.GetYear(3, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Forecast Quntity 3] is Null,0,TblForcasting.[Forecast Quntity 3]) as  [Forecast Quntity " + clsFun.GetYear(3, "+") + "_1], ");
                sbSqlQuery.Append("iif(TblForcasting.[Delta manufacturing cost 3] Is Null,0,TblForcasting.[Delta manufacturing cost 3]) as  [Delta manufacturing cost forecast " + clsFun.GetYear(3, "+") + " vs " + clsFun.GetYear(2, "+") + "_1], ");
                sbSqlQuery.Append("cs.[Part Sobsl] as [Part Sobsl_1], cs.[Creation date] as [Creation date_1], cs.[Dispo Code] as [Dispo Code_1], cs.[Dispo Name] as [Dispo Name_1], cs.[Part Technique] as [Part Technique_1],");
                sbSqlQuery.Append("cs.[Model Type Baureihe] as [Model Type Baureihe_1], cs.[Supplier Code] as [Supplier Code_1], cs.[Supplier Name] as [Supplier Name_1],cs.[Gross List Price  BLP / VLP] as [Gross List Price  BLP / VLP_1], ");
                sbSqlQuery.Append("cs.[Return Value (RW)] as [Return Value (RW)_1],  cs.[Trade Value]as [Trade Value_1], ");
                sbSqlQuery.Append(" cs.[Price Date] as [Price Date_1], cs.[Gross Revenue Current Month] as [Gross Revenue Current Month_1], ");
                sbSqlQuery.Append("cs.[Gross Revenue Current YTD] as [Gross Revenue Current YTD_1], cs.[NET Revenue per VLP] as  [NET Revenue per VLP_1], ");
                sbSqlQuery.Append(" cs.[NET Revenue YTD] as [NET Revenue YTD_1], ");
                sbSqlQuery.Append("cs.[Net revenue without return value current Month] as [Net revenue without return value current Month_1] , ");
                sbSqlQuery.Append(" cs.[Net Revenu RETURN YTD] as [Net Revenu RETURN YTD_1],cs.[Factory Cost SKW] as [Factory Cost SKW_1], ");
                sbSqlQuery.Append("cs.[tax value approach] as [tax value approach_1], cs.[SKW without  SWA] as [SKW without  SWA_1], ");
                sbSqlQuery.Append(" cs.[Total Cost per Part] as [Total Cost per Part_1], cs.[Total Cost] as [Total Cost_1], ");
                sbSqlQuery.Append("cs.[Total Cost excl Return Cost] as [Total Cost excl Return Cost_1], cs.[Total Cost YTD] as [Total Cost YTD_1], ");
                sbSqlQuery.Append("cs.[Material Cost] as [Material Cost_1], cs.[Return Cost]as [Return Cost_1], ");
                sbSqlQuery.Append("cs.GuK as GuK_1, cs.[Package Absolute] as [Package Absolute_1], ");
                sbSqlQuery.Append("  cs.[VTKP abs] as [VTKP abs_1],  cs.[Result Factory] as [Result Factory_1],  cs.[CORE value Supplior] as [CORE value Supplior_1], ");
                sbSqlQuery.Append("  cs.[Credit Voucher Old Part] as [Credit Voucher Old Part_1], cs.Result as Result_1, cs.[Discount Group] as [Discount Group_1]");
                sbSqlQuery.Append(", Format(cs.[Date Id],'MMM-yyyy') as [Date Id_1], cs.[Marketing Code]as[Marketing Code_1], cs.[Part No] as [Part No_1], cs.[Part Name]as[Part Name_1], ");
                sbSqlQuery.Append("cs.[Series Production Start]as [Series Production Start_1], cs.[Series Production End]as[Series Production End_1], Format(cs.[Prognosed Series End Date],'yyyy-MM-dd') as [Prognosed Series End Date_1] ");
                sbSqlQuery.Append(" FROM Cost_summary cs LEFT JOIN TblForcasting ON cs.[Part No] = TblForcasting.[Part No]");
                sbSqlQuery.Append("  ORDER BY sId DESC;");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            try
            {
                string sSQL = "Drop Table PART_YRLY_VALUES";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //PART_YRLY_VALUES
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" SELECT Cost_Management_datenbasis.SNR as [Part No], Avg(Cost_Management_datenbasis.[aktueller Monat HKSP]) AS AVG_HKSP_1, ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.[2018 HKSP] AS AVG_HKSP_2,");
                sbSqlQuery.Append(" Cost_Management_datenbasis.[2017 HKSP] ");
                sbSqlQuery.Append(" AS AVG_HKSP_3, Cost_Management_datenbasis.[2016 HKSP] AS AVG_HKSP_4,");
                sbSqlQuery.Append("  Cost_Management_datenbasis.Mcode as [Marketing Code], Cost_Management_datenbasis.[Teile Sobsl]  as [Part Sobsl] ,");
                sbSqlQuery.Append(" Cost_Management_datenbasis.LN as [Supplier Code],  ");
                sbSqlQuery.Append(" Round( ([2018 HKSP] * ptr.Abgang2018)-(Cost_Management_datenbasis.[aktueller Monat HKSP] * ptr.Abgang2018),2) as [Delta_all_part],");
                sbSqlQuery.Append(" Cost_Management_datenbasis.Benennung as [Supplier Name] ");
                sbSqlQuery.Append(", TblForcasting.[Forcasted Manufacturing Cost Year 1] as Forecast_1, TblForcasting.[Forcasted Manufacturing Cost Year 2] as Forecast_2,");
                sbSqlQuery.Append(" TblForcasting.[Forcasted Manufacturing Cost Year 3] as Forecast_3 , ");
                sbSqlQuery.Append(" (TblForcasting.[Forcasted Manufacturing Cost Year 1]- AVG_HKSP_1) * TblForcasting.[Forecast Quntity 1] as [Delta_Total_Part_1],");
                sbSqlQuery.Append(" (TblForcasting.[Forcasted Manufacturing Cost Year 2]- AVG_HKSP_1)* TblForcasting.[Forecast Quntity 2] as [Delta_Total_Part_2],");
                sbSqlQuery.Append(" (TblForcasting.[Forcasted Manufacturing Cost Year 3]- AVG_HKSP_1) * TblForcasting.[Forecast Quntity 3] as [Delta_Total_Part_3] ");
                sbSqlQuery.Append("  INTO PART_YRLY_VALUES ");
                sbSqlQuery.Append(" FROM (Cost_Management_datenbasis Left JOIN TblForcasting ON Cost_Management_datenbasis.SNR = TblForcasting.[Part No]) ");
                sbSqlQuery.Append("Left Join Abgang2018 ptr on(Cost_Management_datenbasis.snr = ptr.SNR)");
                sbSqlQuery.Append(" GROUP BY Cost_Management_datenbasis.SNR, Cost_Management_datenbasis.[2018 HKSP], Cost_Management_datenbasis.[2017 HKSP], ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.[2016 HKSP], Cost_Management_datenbasis.Mcode , Cost_Management_datenbasis.[Teile Sobsl], ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.LN , Cost_Management_datenbasis.Benennung , Right([Cost_Management_datenbasis].[monat],4), ");
                sbSqlQuery.Append(" TblForcasting.[Forcasted Manufacturing Cost Year 1], TblForcasting.[Forcasted Manufacturing Cost Year 2], ");
                sbSqlQuery.Append(" TblForcasting.[Forcasted Manufacturing Cost Year 3], ");
                sbSqlQuery.Append(" TblForcasting.[Delta manufacturing cost 1], TblForcasting.[Delta manufacturing cost 2], ");
                sbSqlQuery.Append(" TblForcasting.[Delta manufacturing cost 3],");
                sbSqlQuery.Append("[2018 HKSP] , ptr.Abgang2018 ,Cost_Management_datenbasis.[aktueller Monat HKSP],");
                sbSqlQuery.Append("TblForcasting.[Forcasted Manufacturing Cost Year 1],TblForcasting.[Forcasted Manufacturing Cost Year 2],TblForcasting.[Forcasted Manufacturing Cost Year 3]");
                sbSqlQuery.Append(",TblForcasting.[Forecast Quntity 1],TblForcasting.[Forecast Quntity 2],TblForcasting.[Forecast Quntity 3]");
                sbSqlQuery.Append(" HAVING(((Right([monat], 4))= 2019));");

                OleDbCommand oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //PART_YRLY_CHANGE
            try
            {
                string sSQL = "Drop Table PART_YRLY_CHANGE";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("SELECT distinct  [Part No] as snr, ");
                sbSqlQuery.Append("Round(iif([Delta2]=0,0,([Delta3] -[Delta2]) / [Delta2] ) * 100, 2)  AS Change_2021, ");
                sbSqlQuery.Append("Round(iif([Delta1]=0,0,([Delta2] -[Delta1]) / [Delta1] ) * 100, 2)  AS Change_2020, ");
                sbSqlQuery.Append("Round(iif([Delta]=0,0,([Delta1] -[Delta]) / [Delta] ) * 100, 2) AS Change_2019, ");
                sbSqlQuery.Append("Round(iif([Delta-1]=0,0,([Delta] -[Delta-1]) / [Delta-1] ) * 100, 2) AS Change_2018, ");
                sbSqlQuery.Append(" Round(iif([Delta-2]=0,0,([Delta-1] -[Delta-2]) / [Delta-2] ) * 100, 2) AS Change_2017, ");
                sbSqlQuery.Append("[Marketing Code],  [Part Sobsl], [Supplier Code],[Supplier Name] ");
                sbSqlQuery.Append(",Delta as [Delta_all_part] INTO PART_YRLY_CHANGE ");
                sbSqlQuery.Append("FROM Grid ");

                OleDbCommand oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            try
            {
                string sSQL = "Drop Table tblRecoredCount";
                OleDbCommand oleCmd = new OleDbCommand(sSQL, dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("SELECT Count(*) as Recored_Count into tblRecoredCount from GRID g inner Join TblDiscSNR ds on g.[Date id]=ds.[date_id] and g.[Part No]=ds.[Part_No] ");

                OleDbCommand oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
        }
        private void GenrateForcasting()
        {
            this.label1.BeginInvoke((MethodInvoker)delegate () { this.label1.Text = "Calculating Forecasting Data...!!!"; });
            ClsForcasting clsforcast;
            DataTable dtforcast = new DataTable();
            DataTable dtSnrdata = new DataTable();
            StringBuilder sbSQL = new StringBuilder();
            sbSQL.Append("SELECT distinct Cs.[Part No],cs.[Part Sobsl],cs.[Supplier Name],");
            sbSQL.Append("iif(cs.[Prognosed Series End Date]='',cs.[Series Production End],cs.[Prognosed Series End Date])  as Dates,ab.Abgang2018 ");
            sbSQL.Append(",ab6.Abgang2017,ab5.Abgang2016 from (((Cost_summary cs");
            sbSQL.Append(" Left Join Abgang2018 ab on cs.[part No]=ab.snr) Left Join Abgang2017 ab6 on cs.[part No]=ab6.snr)");
            sbSQL.Append(" left Join Abgang2016 ab5 on cs.[part No]=ab5.snr)");
            sbSQL.Append(" Where cs.[Part Sobsl]='HT'");
            this.label1.BeginInvoke((MethodInvoker)delegate () { this.label1.Text = "Calculating Forecasting data...!!!"; });
            DataTable dtSNR = clsFun.dtGetData(sbSQL.ToString());
            if (dtSNR.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSNR.Rows)
                {
                    try
                    {
                        try
                        {
                            string sSNR = dr[0].ToString();
                            decimal dAbang = Convert.ToDecimal(dr["Abgang2018"].ToString().Equals("") ? "0" : dr["Abgang2018"].ToString());
                            clsforcast = new ClsForcasting();
                            sbSQL = new StringBuilder();

                            sbSQL.Append(" SELECT top 1 Right([Cost_Management_datenbasis].[monat],4) AS [year], Left([Cost_Management_datenbasis].[monat],2) AS[month], ");
                            sbSQL.Append(" Cost_Management_datenbasis.SNR, Cost_Management_datenbasis.[aktueller Monat HKSP], Cost_Management_datenbasis.[2018 HKSP],");
                            sbSQL.Append(" Cost_Management_datenbasis.[2017 HKSP], Cost_Management_datenbasis.[2016 HKSP] ");
                            sbSQL.Append(" ,iif([aktueller Monat HKSP] > [2018 HKSP],'incr','') as [chkforecast],");
                            sbSQL.Append("cm.[Part Quantity YTD],Format(cm.[Date Id],'mmm-yyyy') AS [Date Id], ");
                            sbSQL.Append("switch((cm.[Marketing Code] like '1%' or cm.[Marketing Code] like 'SP%'),'MBC',");
                            sbSQL.Append("cm.[Marketing Code] like '2T%','Van',(cm.[Marketing Code] like '2U%' or cm.[Marketing Code] like '2L%'),'Truck',True,'') as Sparte");
                            sbSQL.Append(",cm.[Current Month Manufacturing Cost per Part]");
                            sbSQL.Append(" FROM Cost_Management_datenbasis inner Join cost_summary cm on Cost_Management_datenbasis.snr=cm.[Part No]");
                            sbSQL.Append(" and Right([Cost_Management_datenbasis].[monat],4)=Right(cm.[Date id],4) and Left([Cost_Management_datenbasis].[monat],2)=Left(cm.[Date Id],2)");
                            sbSQL.Append(" WHERE(((Cost_Management_datenbasis.SNR) = '" + sSNR + "'))");
                            sbSQL.Append(" ORDER BY Right([Cost_Management_datenbasis].[monat],4) DESC , Left([Cost_Management_datenbasis].[monat],2) DESC;");

                            dtSnrdata = clsFun.dtGetData(sbSQL.ToString());
                            if (dtSnrdata.Rows.Count > 0)
                            {
                                decimal dval1 = Convert.ToDecimal(dtSnrdata.Rows[0][6].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][6].ToString());
                                decimal dval2 = Convert.ToDecimal(dtSnrdata.Rows[0][5].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][5].ToString());
                                decimal dval3 = Convert.ToDecimal(dtSnrdata.Rows[0][4].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][4].ToString());
                                decimal dval4 = Convert.ToDecimal(dtSnrdata.Rows[0][3].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][3].ToString());
                                decimal dCurrentvalue = dval4;
                                dtforcast = clsforcast.simpleMovingAverageForHT(new decimal[4] { dval1, dval2, dval3, dval4 }, 3, 3, 0);

                                decimal dForecast2 = 0;
                                decimal dforecastCost2 = 0;
                                decimal dDelata2 = 0;
                                decimal dForecastqty2 = 0;

                                decimal dForecast3 = 0;
                                decimal dforecastCost3 = 0;
                                decimal dDelata3 = 0;
                                decimal dForecastqty3 = 0;

                                DataTable dtforcastqty = new DataTable();
                                decimal dExpofactor = 0;
                                if (dtSnrdata.Rows[0]["year"].ToString().Equals(clsFun.GetYear(0, "")))
                                {
                                    DataRow[] drexpo = dtexpo.Select("Sparte='" + dtSnrdata.Rows[0]["Sparte"].ToString() + "' and Monat='" + dtSnrdata.Rows[0]["Month"].ToString() + "'");
                                    if (drexpo.Length > 0)
                                        dExpofactor = Convert.ToDecimal(drexpo[0][3].ToString().Equals("") ? "1" : drexpo[0][3].ToString());
                                    else
                                        dExpofactor = 1;
                                }
                                else
                                {
                                    dExpofactor = 0;
                                }

                                if (dExpofactor.Equals(0))
                                    dval4 = 0;
                                else
                                    dval4 = Convert.ToDecimal(dtSnrdata.Rows[0]["Part Quantity YTD"].ToString().Equals("") ? "0" : dtSnrdata.Rows[0]["Part Quantity YTD"].ToString()) * dExpofactor;

                                dval3 = Convert.ToDecimal(dr["Abgang2018"].ToString().Equals("") ? "0" : dr["Abgang2018"].ToString());
                                dval2 = Convert.ToDecimal(dr["Abgang2017"].ToString().Equals("") ? "0" : dr["Abgang2017"].ToString());
                                dval1 = Convert.ToDecimal(dr["Abgang2016"].ToString().Equals("") ? "0" : dr["Abgang2016"].ToString());
                                decimal dEstimation = 0;
                                dtforcastqty = clsforcast.simpleMovingAverageForPartQTY(new decimal[4] { dval1, dval2, dval3, dval4 }, 3, 2, 0);
                                if (!dval4.Equals(0))
                                    dEstimation = Math.Round((Math.Round(Convert.ToDecimal(dtSnrdata.Rows[0]["Current Month Manufacturing Cost per Part"].ToString().Equals("") ? "0" : dtSnrdata.Rows[0]["Current Month Manufacturing Cost per Part"].ToString()), 2) * Math.Round(dval4, MidpointRounding.AwayFromZero)), 2);
                                else
                                    dEstimation = 0;

                                decimal dForecastqty = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[4][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[4][4].ToString()), MidpointRounding.AwayFromZero);
                                dForecastqty2 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[5][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[5][4].ToString()), MidpointRounding.AwayFromZero);
                                dForecastqty3 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[6][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[6][4].ToString()), MidpointRounding.AwayFromZero);

                                decimal dForecast1 = Convert.ToDecimal(dtforcast.Rows[4][4].ToString());
                                decimal dforecastCost1 = dForecast1 * Math.Round(dForecastqty, MidpointRounding.AwayFromZero);
                                decimal dDelata1 = (dForecast1- dCurrentvalue) * Math.Round(dForecastqty, MidpointRounding.AwayFromZero);

                                dForecast2 = Convert.ToDecimal(dtforcast.Rows[5][4].ToString());
                                dforecastCost2 = dForecast2 * Math.Round(dForecastqty2, MidpointRounding.AwayFromZero);
                                dDelata2 = (dForecast2 - dForecast1) * Math.Round(dForecastqty2, MidpointRounding.AwayFromZero);

                                dForecast3 = Convert.ToDecimal(dtforcast.Rows[6][4].ToString());
                                dforecastCost3 = dForecast3 * Math.Round(dForecastqty3, MidpointRounding.AwayFromZero);
                                dDelata3 = (dForecast3 - dForecast2) * Math.Round(dForecastqty3, MidpointRounding.AwayFromZero);


                                int iColorCode = 0;
                                if (Math.Round(dCurrentvalue, 2) > dForecast1)
                                {
                                    iColorCode = 1;
                                }
                                else if (dForecast1 > Math.Round(dCurrentvalue, 2))
                                {
                                    iColorCode = 2;
                                }
                                else if (Math.Round(dCurrentvalue, 2) == dForecast1)
                                {
                                    iColorCode = 0;
                                }


                                sbSQL = new StringBuilder();

                                sbSQL = new StringBuilder();
                                sbSQL.Append("Insert into TblForcasting( [Forcasted Manufacturing Cost Year 1],[Forcasted Manufacturing Cost Year 2],");
                                sbSQL.Append("[Forcasted Manufacturing Cost Year 3],");
                                sbSQL.Append("[Forecast Manufacturing cost total 1],[Forecast Manufacturing cost total 2],[Forecast Manufacturing cost total 3],");
                                sbSQL.Append("[Delta manufacturing cost 1],[Delta manufacturing cost 2],[Delta manufacturing cost 3]");
                                sbSQL.Append(",[Part No],Color_code,[Forecast Quntity 1],[Forecast Quntity 2],[Forecast Quntity 3],[Part YTD extrapol]");
                                sbSQL.Append(" ,[Manufacturing Material cost Current Year Estimation] )Values (");
                                sbSQL.Append("'" + dForecast1 + "' ,'" + dForecast2 + "',");
                                sbSQL.Append("'" + dForecast3 + "',");
                                sbSQL.Append("'" + dforecastCost1 + "','" + dforecastCost2 + "','" + dforecastCost3 + "',");
                                sbSQL.Append("'" + dDelata1 + "','" + dDelata2 + "',");
                                sbSQL.Append("'" + dDelata3 + "','" + sSNR + "','" + iColorCode + "'");
                                sbSQL.Append(",'" + Math.Round(dForecastqty, MidpointRounding.AwayFromZero) + "','" + Math.Round(dForecastqty2, MidpointRounding.AwayFromZero) + "','" + Math.Round(dForecastqty3, MidpointRounding.AwayFromZero) + "'");
                                sbSQL.Append("," + Math.Round(dval4, 0) + ",'" + Math.Round(dEstimation, 2) + "')");
                                OleDbCommand oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                                oleCmd.ExecuteNonQuery();

                                sbSQL = new StringBuilder();
                                sbSQL.Append("Insert into TblDiscSNR( [Part_No],[date_id]) Values (");
                                sbSQL.Append("'" + sSNR + "','" + dtSnrdata.Rows[0]["Date Id"].ToString() + "')");
                                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                                oleCmd.ExecuteNonQuery();
                            }
                            else
                            {
                                sSNR = "";
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            //[Part Sobsl]!='HT' Calculation 
            sbSQL = new StringBuilder();
            sbSQL.Append("SELECT distinct Cs.[Part No],cs.[Part Sobsl],cs.[Supplier Name],");
            sbSQL.Append("iif(cs.[Prognosed Series End Date]='',cs.[Series Production End],cs.[Prognosed Series End Date])  as Dates,ab.Abgang2018 ");
            sbSQL.Append(",ab6.Abgang2017,ab5.Abgang2016  from (((Cost_summary cs");
            sbSQL.Append(" left Join Abgang2018 ab on cs.[part No]=ab.snr) Left Join Abgang2017 ab6 on cs.[part No]=ab6.snr)");
            sbSQL.Append(" left Join Abgang2016 ab5 on cs.[part No]=ab5.snr)");
            sbSQL.Append(" Where cs.[Part Sobsl]<>'HT'");
            dtSNR = new DataTable();
            dtSNR = clsFun.dtGetData(sbSQL.ToString());
            if (dtSNR.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSNR.Rows)
                {
                    try
                    {
                        try
                        {
                            string sSNR = dr[0].ToString();
                            decimal dAbang = Convert.ToDecimal(dr["Abgang2017"].ToString().Equals("") ? "0" : dr["Abgang2017"].ToString());
                            string sSupplier = dr["Supplier Name"].ToString().Equals("") ? "" : dr["Supplier Name"].ToString();
                            int iYear = Convert.ToInt32(dr["Dates"].ToString().Equals("") ? "1900" : dr["Dates"].ToString().Substring(0, 4));
                            sbSQL = new StringBuilder();
                            sbSQL.Append(" SELECT top 1 Right([Cost_Management_datenbasis].[monat],4) AS [year], Left([Cost_Management_datenbasis].[monat],2) AS[month], ");
                            sbSQL.Append(" Cost_Management_datenbasis.SNR, Cost_Management_datenbasis.[aktueller Monat HKSP], Cost_Management_datenbasis.[2018 HKSP],");
                            sbSQL.Append(" Cost_Management_datenbasis.[2017 HKSP], Cost_Management_datenbasis.[2016 HKSP] ");
                            sbSQL.Append(" ,iif([aktueller Monat HKSP] > [2018 HKSP],'incr','') as [chkforecast],");
                            sbSQL.Append("cm.[Part Quantity YTD],Format(cm.[Date Id],'mmm-yyyy') AS [Date Id], ");
                            sbSQL.Append("switch((cm.[Marketing Code] like '1%' or cm.[Marketing Code] like 'SP%'),'MBC',");
                            sbSQL.Append("cm.[Marketing Code] like '2T%','Van',(cm.[Marketing Code] like '2U%' or cm.[Marketing Code] like '2L%'),'Truck',True,'') as Sparte");
                            sbSQL.Append(",cm.[Current Month Manufacturing Cost per Part]");
                            sbSQL.Append(" FROM Cost_Management_datenbasis inner Join cost_summary cm on Cost_Management_datenbasis.snr=cm.[Part No]");
                            sbSQL.Append(" and Right([Cost_Management_datenbasis].[monat],4)=Right(cm.[Date id],4) and Left([Cost_Management_datenbasis].[monat],2)=Left(cm.[Date Id],2)");
                            sbSQL.Append(" WHERE(((Cost_Management_datenbasis.SNR) = '" + sSNR + "'))");
                            sbSQL.Append(" ORDER BY Right([Cost_Management_datenbasis].[monat],4) DESC , Left([Cost_Management_datenbasis].[monat],2) DESC;");

                            dtSnrdata = clsFun.dtGetData(sbSQL.ToString());
                            if (dtSnrdata.Rows.Count > 0)
                            {
                                dtforcast = new DataTable();
                                decimal dval1 = Convert.ToDecimal(dtSnrdata.Rows[0][6].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][6].ToString());
                                decimal dval2 = Convert.ToDecimal(dtSnrdata.Rows[0][5].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][5].ToString());
                                decimal dval3 = Convert.ToDecimal(dtSnrdata.Rows[0][4].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][4].ToString());
                                decimal dval4 = Convert.ToDecimal(dtSnrdata.Rows[0][3].ToString().Equals("") ? "0" : dtSnrdata.Rows[0][3].ToString());
                                decimal dCurrentvalue = dval4;
                                clsforcast = new ClsForcasting();
                                dtforcast = clsforcast.simpleMovingAverageDAG(new decimal[4] { dval1, dval2, dval3, dval4 }, 3, 3, 0, iYear, sSupplier);

                                decimal dForecast2 = 0;
                                decimal dforecastCost2 = 0;
                                decimal dDelata2 = 0;

                                decimal dForecast3 = 0;
                                decimal dforecastCost3 = 0;
                                decimal dDelata3 = 0;


                                DataTable dtforcastqty = new DataTable();
                                decimal dExpofactor = 0;
                                if (dtSnrdata.Rows[0]["year"].ToString().Equals(clsFun.GetYear(0, "")))
                                {
                                    DataRow[] drexpo = dtexpo.Select("Sparte='" + dtSnrdata.Rows[0]["Sparte"].ToString() + "' and Monat='" + dtSnrdata.Rows[0]["Month"].ToString() + "'");
                                    if (drexpo.Length > 0)
                                        dExpofactor = Convert.ToDecimal(drexpo[0][3].ToString().Equals("") ? "1" : drexpo[0][3].ToString());
                                    else
                                        dExpofactor = 1;
                                }
                                else
                                {
                                    dExpofactor = 0;
                                }

                                if (dExpofactor.Equals(0))
                                    dval4 = 0;
                                else
                                    dval4 = Convert.ToDecimal(dtSnrdata.Rows[0]["Part Quantity YTD"].ToString().Equals("") ? "0" : dtSnrdata.Rows[0]["Part Quantity YTD"].ToString()) * dExpofactor;

                                dval3 = Convert.ToDecimal(dr["Abgang2018"].ToString().Equals("") ? "0" : dr["Abgang2018"].ToString());
                                dval2 = Convert.ToDecimal(dr["Abgang2017"].ToString().Equals("") ? "0" : dr["Abgang2017"].ToString());
                                dval1 = Convert.ToDecimal(dr["Abgang2016"].ToString().Equals("") ? "0" : dr["Abgang2016"].ToString());
                                decimal dEstimation = 0;
                                dtforcastqty = clsforcast.simpleMovingAverageForPartQTY(new decimal[4] { dval1, dval2, dval3, dval4 }, 3, 2, 0);
                                if (!dval4.Equals(0))
                                    dEstimation = Math.Round((Math.Round(Convert.ToDecimal(dtSnrdata.Rows[0]["Current Month Manufacturing Cost per Part"].ToString().Equals("") ? "0" : dtSnrdata.Rows[0]["Current Month Manufacturing Cost per Part"].ToString()), 2) * Math.Round(dval4, MidpointRounding.AwayFromZero)), 2);
                                else
                                    dEstimation = 0;

                                decimal dForecastqty = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[4][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[4][4].ToString()), MidpointRounding.AwayFromZero);
                                decimal dForecastqty2 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[5][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[5][4].ToString()), MidpointRounding.AwayFromZero);
                                decimal dForecastqty3 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[6][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[6][4].ToString()), MidpointRounding.AwayFromZero);

                                dForecastqty = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[4][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[4][4].ToString()), MidpointRounding.AwayFromZero);
                                dForecastqty2 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[5][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[5][4].ToString()), MidpointRounding.AwayFromZero);
                                dForecastqty3 = Math.Round(Convert.ToDecimal(dtforcastqty.Rows[6][4].ToString().Equals("") ? "0" : dtforcastqty.Rows[6][4].ToString()), MidpointRounding.AwayFromZero);

                                decimal dForecast1 = Convert.ToDecimal(dtforcast.Rows[4][4].ToString());
                                decimal dforecastCost1 = dForecast1 * Math.Round(dForecastqty, MidpointRounding.AwayFromZero);
                                decimal dDelata1 = (dForecast1 - dCurrentvalue) * Math.Round(dForecastqty, MidpointRounding.AwayFromZero);

                                dForecast2 = Convert.ToDecimal(dtforcast.Rows[5][4].ToString());
                                dforecastCost2 = dForecast2 * Math.Round(dForecastqty2, MidpointRounding.AwayFromZero);
                                dDelata2 = (dForecast2 - dForecast1) * Math.Round(dForecastqty2, MidpointRounding.AwayFromZero);

                                dForecast3 = Convert.ToDecimal(dtforcast.Rows[6][4].ToString());
                                dforecastCost3 = dForecast3 * Math.Round(dForecastqty3, MidpointRounding.AwayFromZero);
                                dDelata3 = (dForecast3 - dForecast2) * Math.Round(dForecastqty3, MidpointRounding.AwayFromZero);



                                int iColorCode = 0;
                                if (Math.Round(dCurrentvalue, 2) > dForecast1)
                                {
                                    iColorCode = 1;
                                }
                                else if (dForecast1 > Math.Round(dCurrentvalue, 2))
                                {
                                    iColorCode = 2;
                                }
                                else if (Math.Round(dCurrentvalue, 2) == dForecast1)
                                {
                                    iColorCode = 0;
                                }


                                sbSQL = new StringBuilder();
                                sbSQL = new StringBuilder();
                                sbSQL.Append("Insert into TblForcasting( [Forcasted Manufacturing Cost Year 1],[Forcasted Manufacturing Cost Year 2],");
                                sbSQL.Append("[Forcasted Manufacturing Cost Year 3],");
                                sbSQL.Append("[Forecast Manufacturing cost total 1],[Forecast Manufacturing cost total 2],[Forecast Manufacturing cost total 3],");
                                sbSQL.Append("[Delta manufacturing cost 1],[Delta manufacturing cost 2],[Delta manufacturing cost 3]");
                                sbSQL.Append(",[Part No],Color_code,[Forecast Quntity 1],[Forecast Quntity 2],[Forecast Quntity 3],[Part YTD extrapol]");
                                sbSQL.Append(" ,[Manufacturing Material cost Current Year Estimation] )Values (");
                                sbSQL.Append("'" + dForecast1 + "' ,'" + dForecast2 + "',");
                                sbSQL.Append("'" + dForecast3 + "',");
                                sbSQL.Append("'" + dforecastCost1 + "','" + dforecastCost2 + "','" + dforecastCost3 + "',");
                                sbSQL.Append("'" + dDelata1 + "','" + dDelata2 + "',");
                                sbSQL.Append("'" + dDelata3 + "','" + sSNR + "','" + iColorCode + "'");
                                sbSQL.Append(",'" + Math.Round(dForecastqty, MidpointRounding.AwayFromZero) + "','" + Math.Round(dForecastqty2, MidpointRounding.AwayFromZero) + "','" + Math.Round(dForecastqty3, MidpointRounding.AwayFromZero) + "'");
                                sbSQL.Append("," + Math.Round(dval4, MidpointRounding.AwayFromZero) + ",'" + Math.Round(dEstimation, 2) + "')");
                                OleDbCommand oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                                oleCmd.ExecuteNonQuery();
                                sbSQL = new StringBuilder();
                                sbSQL.Append("Insert into TblDiscSNR( [Part_No],[date_id]) Values (");
                                sbSQL.Append("'" + sSNR + "','" + dtSnrdata.Rows[0]["Date Id"].ToString() + "')");
                                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                                oleCmd.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }
    }
}
