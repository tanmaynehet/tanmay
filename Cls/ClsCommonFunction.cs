using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using INI;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Microsoft.CSharp;
using System.Runtime.InteropServices;

namespace cost_management
{
    public class ClsCommonFunction
    {
        IniFile ini = new IniFile();

        public string getConstr()
        {
            string str_Con = "";
            string str_Dbname = "";
            string strDbpath = "";
            if (!ini.IniReadValue("Database", "Name").ToString().Equals(""))
            {
                str_Dbname = ini.IniReadValue("Database", "Name").ToString();
                strDbpath = ini.IniReadValue("Database", "Path").ToString();
                if (!File.Exists(strDbpath + "\\" + str_Dbname))
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    ofd.Multiselect = false;
                    ofd.Title = "Please select a database";
                    ofd.FilterIndex = 2;
                    ofd.CheckFileExists = true;
                    ofd.CheckPathExists = true;

                    ofd.Filter = "Access 2007 (*.accdb)|*accdb";
                    if (ofd.ShowDialog().Equals(DialogResult.OK))
                    {
                        string fname = ofd.FileName.ToString();
                        string spaths = fname.Substring(0, fname.LastIndexOf("\\"));
                        str_Dbname = Path.GetFileName(fname);
                        ini.IniWriteValue("Database", "Name", str_Dbname);
                        ini.IniWriteValue("Database", "Path", spaths);
                        str_Dbname = ini.IniReadValue("Database", "Name").ToString();
                        strDbpath = ini.IniReadValue("Database", "Path").ToString();
                    }

                    str_Con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strDbpath + "\\" + str_Dbname + ";Persist Security Info=False;";
                }
                else
                    str_Con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strDbpath + "\\" + str_Dbname + ";Persist Security Info=False;";
            }
            else
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = false;
                ofd.Title = "Please select a database";
                ofd.FilterIndex = 2;
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;

                ofd.Filter = "Access 2007 (*.accdb)|*accdb";
                if (ofd.ShowDialog().Equals(DialogResult.OK))
                {
                    string fname = ofd.FileName.ToString();
                    strDbpath = fname.Substring(0, fname.LastIndexOf("\\"));
                    str_Dbname = Path.GetFileName(fname);
                    ini.IniWriteValue("Database", "Path", str_Dbname);
                    ini.IniWriteValue("Database", "Name", strDbpath);
                    str_Dbname = ini.IniReadValue("Database", "Name").ToString();
                    strDbpath = ini.IniReadValue("Database", "Path").ToString();
                }

                str_Con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strDbpath + "\\" + str_Dbname + ";Persist Security Info=False;";
            }
            return str_Con;
        }

        public OleDbConnection Openconnection()
        {
            OleDbConnection olecon = new OleDbConnection(getConstr());
            try
            {

                if (olecon.State.Equals(ConnectionState.Open))
                {
                    olecon.Close();
                    olecon.Open();
                }
                else
                {
                    olecon.Open();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            return olecon;
        }
        public OleDbConnection Closeconnection()
        {
            OleDbConnection olecon = new OleDbConnection(getConstr());
            try
            {
                olecon.Close();
            }
            catch
            {

            }
            return olecon;
        }
        public int AlterDatabase(string sSQL)
        {
            int iInsaertion = 0;
            OleDbCommand olecmd = new OleDbCommand();
            try
            {
                olecmd.Connection = Openconnection();
                olecmd.CommandText = sSQL;
                olecmd.CommandType = CommandType.Text;
                olecmd.ExecuteNonQuery();
                Closeconnection();

                olecmd.Connection.Close();
                olecmd.Dispose();
            }
            catch (Exception ex)
            {

            }
            return iInsaertion;
        }

        public DataTable dtGetData(string sSQL)
        {
            DataTable dt = new DataTable();
            OleDbCommand olecmd = new OleDbCommand();
            try
            {
                olecmd.Connection = Openconnection();
                olecmd.CommandText = sSQL;
                olecmd.CommandType = CommandType.Text;
                OleDbDataAdapter oleda = new OleDbDataAdapter(olecmd);
                oleda.Fill(dt);
                Closeconnection();
            }
            catch (Exception ex)
            {

            }
            return dt;
        }

        public DataTable GetData(int sYear, int sMonth)
        {
            DataTable dt = new DataTable();
            OleDbCommand olecmd;
            try
            {
                olecmd = new OleDbCommand("GET_Summeri_Data", Openconnection());
                olecmd.CommandType = CommandType.StoredProcedure;
                olecmd.Parameters.AddWithValue("sYear", sYear);
                olecmd.Parameters.AddWithValue("sMonth", sMonth);
                OleDbDataAdapter oleda = new OleDbDataAdapter(olecmd);
                oleda.Fill(dt);
                Closeconnection();
            }
            catch (Exception ex)
            {

            }
            return dt;
        }

        public string GetYearName(int iYear, string sSign)
        {
            int sYearValue = 0;
            if (sSign.Equals("+"))
                sYearValue = DateTime.Now.Year + iYear;
            else if (sSign.Equals("-"))
                sYearValue = DateTime.Now.Year - iYear;
            else
                sYearValue = DateTime.Now.Year;

            return sYearValue.ToString();
        }
        public string GetAPPVersion()
        {
            IniFile ini = new IniFile();
            return ini.IniReadValue("Version", "ToolVer").ToString();

        }
        public void exportToExport(DataTable dtdata)
        {
            string str_Dbname = "";
            string path = "";
            try
            {
                path = "Exported" + DateTime.Now.ToString("mmm-dd-yyyy hh:tt").Replace("-", "_").Replace(":", "_").Replace(" ", "_") + ".xls";
                if (dtdata.Rows.Count > 0)
                {
                    try
                    {
                        if (!ini.IniReadValue("Excel", "Path").ToString().Equals(""))
                        {
                            str_Dbname = ini.IniReadValue("Excel", "Path").ToString() + path;
                        }
                        else
                        {
                            FolderBrowserDialog ofd = new FolderBrowserDialog();

                            if (ofd.ShowDialog().Equals(DialogResult.OK))
                            {
                                ini.IniWriteValue("Excel", "Path", ofd.SelectedPath);
                                str_Dbname = ini.IniReadValue("Excel", "Path").ToString();
                            }
                        }

                        StreamWriter wr = new StreamWriter(str_Dbname + path);
                        // Write Columns to excel file
                        for (int i = 0; i < dtdata.Columns.Count; i++)
                        {
                            wr.Write(dtdata.Columns[i].ToString().ToUpper() + "\t");
                        }
                        wr.WriteLine();
                        //write rows to excel file
                        for (int i = 0; i < (dtdata.Rows.Count); i++)
                        {
                            for (int j = 0; j < dtdata.Columns.Count; j++)
                            {
                                if (dtdata.Rows[i][j] != null)
                                {
                                    wr.Write(Convert.ToString(dtdata.Rows[i][j]) + "\t");
                                }
                                else
                                {
                                    wr.Write("\t");
                                }
                            }
                            wr.WriteLine();
                        }
                        wr.Close();
                        MessageBox.Show("File is exported sucessfully at " + str_Dbname);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void export(DataTable dtdata)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;
                DataTable dtexcel = dtdata.Copy();
                int iOrdinal = dtexcel.Columns["delta"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["delta1"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["delta2"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["delta3"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);

                iOrdinal = dtexcel.Columns["FMCT1"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["FMCT2"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["FMCT3"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["delta-1"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["delta-2"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                iOrdinal = dtexcel.Columns["Color_code"].Ordinal;
                dtexcel.Columns.RemoveAt(iOrdinal);
                dtexcel.Columns.RemoveAt(1);
                dtexcel.Columns.RemoveAt(0);

                for (int iloop = dtexcel.Columns.Count - 1; iloop > 0; iloop--)
                {
                    if (dtexcel.Columns[iloop].ColumnName.Contains("_1"))
                    {
                        dtexcel.Columns.RemoveAt(iloop);
                    }

                }


                xlApp.Visible = true;
                for (j = 0; j <= dtexcel.Columns.Count - 1; j++)
                {
                    xlWorkSheet.Cells[1, j + 1] = dtexcel.Columns[j].ColumnName.ToString();
                }
                for (i = 0; i <= dtexcel.Rows.Count - 1; i++)
                {
                    for (j = 0; j <= dtexcel.Columns.Count - 1; j++)
                    {
                        if (dtexcel.Columns[j].DataType.Equals(typeof(decimal)))
                            xlWorkSheet.Cells[i + 2, j + 1] = Convert.ToDecimal(dtexcel.Rows[i][j].ToString());
                        else
                            xlWorkSheet.Cells[i + 2, j + 1] = dtexcel.Rows[i][j].ToString();
                    }

                }
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange.Columns;
                xlRange.Worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, xlRange,
                        System.Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = "Table 1";
                xlRange.Select();
                xlRange.Worksheet.ListObjects["Table 1"].TableStyle = "TableStyleMedium15";

                xlWorkSheet.Name = "Exported Data";
                xlRange = null;
                xlWorkSheet = null;
                xlApp = null;
            }
            catch (Exception ex)
            { }
        }

        public void ExportData(DataTable dtdatainc, DataTable dtDESC)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                int i = 0;
                int j = 0;
                DataTable dtexcel = new DataTable();
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorkSheet.UsedRange.Columns;

                i = 0;
                j = 0;
                dtexcel = new DataTable();
                dtexcel = dtdatainc.Copy();
                xlApp.Visible = true;
                for (j = 0; j <= dtexcel.Columns.Count - 1; j++)
                {
                    xlWorkSheet.Cells[1, j + 1] = dtexcel.Columns[j].ColumnName.ToString();
                }
                for (i = 0; i <= dtexcel.Rows.Count - 1; i++)
                {
                    for (j = 0; j <= dtexcel.Columns.Count - 1; j++)
                    {
                        if (dtexcel.Columns[j].DataType.Equals(typeof(decimal)))
                            xlWorkSheet.Cells[i + 2, j + 1] = Convert.ToDecimal(dtexcel.Rows[i][j].ToString());
                        else
                            xlWorkSheet.Cells[i + 2, j + 1] = dtexcel.Rows[i][j].ToString();
                    }

                }
                xlRange = xlWorkSheet.UsedRange.Columns;
                xlRange.Worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, xlRange,
                        System.Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = "Table 1";
                xlRange.Select();
                xlRange.Worksheet.ListObjects["Table 1"].TableStyle = "TableStyleMedium15";

                xlWorkSheet.Name = "Excel Export";
                xlRange = null;
                xlWorkSheet = null;
                xlApp = null;
            }
            catch (Exception ex)
            { }
        }
        public string GetYear(int iNYear, string sSign)
        {
            string sYear = "";
            try
            {
                if (sSign.Equals("+"))
                    sYear = Convert.ToString(Convert.ToInt32(DateTime.Now.Year) + iNYear);
                else if (sSign.Equals("-"))
                    sYear = Convert.ToString(Convert.ToInt32(DateTime.Now.Year) - iNYear);
                else
                    sYear = Convert.ToString(DateTime.Now.Year);
            }
            catch { }
            return sYear;
        }

        public DataTable ImportDataUsingExcel(string filePath, string sSheetName)
        {
            DataTable excelDataSet = new DataTable();
            try
            {
                string sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filePath + "';Extended Properties=\"Excel 12.0;HDR=YES;\"";
                using (OleDbConnection conn = new OleDbConnection(sConnectionString))
                {
                    conn.Open();
                    string sSQLQuery = "select * from [" + sSheetName + "]";
                    OleDbDataAdapter objDA = new OleDbDataAdapter(sSQLQuery, conn);
                    objDA.Fill(excelDataSet);
                }
            }
            catch (Exception ex) { }
            return excelDataSet;
        }
        public List<string> ListSheetInExcel(string filePath)
        {
            List<string> listSheet = new List<string>();
            string sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + filePath + "';Extended Properties=\"Excel 12.0;HDR=YES;\"";
            using (OleDbConnection conn = new OleDbConnection(sConnectionString.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        listSheet.Add(drSheet["TABLE_NAME"].ToString());
                    }
                }
            }
            return listSheet;
        }
    }
}
