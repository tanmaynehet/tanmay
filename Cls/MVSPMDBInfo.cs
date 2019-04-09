using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace cost_management
{
    public class MVSPMDBInfo
    {
        OleDbConnection dbCon;
        //----------- MOD WIN7 --------------
        ClsCommonFunction clsCOn = new ClsCommonFunction();

        public bool checkdatatabe()
        {
            dbCon = new OleDbConnection();
            dbCon = clsCOn.Openconnection();
            OleDbCommand oleCmd;
            DataTable dtTables = new DataTable();

            //Create Procedure PART_SSTART_SEND_PROGEND_DATE
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {

                    sbSqlQuery.Append("Drop PROCEDURE PART_SSTART_SEND_PROGEND_DATE");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch
                { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("CREATE PROCEDURE PART_SSTART_SEND_PROGEND_DATE as ");
                sbSqlQuery.Append("select  Monat,snr,Serie_von,DateAdd('yyyy',6,CDATE(Serie_von)) as [PROG_END_DATE] from Cost_Management_datenbasis");
                sbSqlQuery.Append(" WHERE (((Cost_Management_datenbasis.[Serie_bis]) Is Null) and Serie_von<>null );");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                oleCmd.Dispose();
            }
            catch (Exception ex)
            {

            }
            //Update NUll value
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();

                sbSqlQuery.Append("update Cost_Management_datenbasis set [Abgang_aktuell_2016]=1 where [Abgang_aktuell_2016] is null;");
                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("update Cost_Management_datenbasis set [Abgang_aktuell_2017]=1 where [Abgang_aktuell_2017] is null;");
                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            //Create Procedure Total_YTD_Cost_per_Part
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE Total_YTD_Cost_per_Part");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE Total_YTD_Cost_per_Part (sYear Long, sMonth Long) as ");
                sbSqlQuery.Append(" SELECT distinct CInt(Mid(MONAT, 4)) AS YEARS, SNR, ");
                sbSqlQuery.Append("iif(SUM(Round((Materialkosten + Rückführkosten + GuK +[Verpack absolut] + [VTKP abs] +[VTKF abs]), 2)) is Null, 0, ");
                sbSqlQuery.Append("SUM(Round((Materialkosten + Rückführkosten + GuK +[Verpack absolut] + [VTKP abs] +[VTKF abs]), 2))) AS[Total YTD Cost per Part]");
                sbSqlQuery.Append("  FROM COST_MANAGEMENT_DATENBASIS ");
                sbSqlQuery.Append("  WHERE   CInt(Mid(MONAT, 1, 2)) <= sMonth  and CInt(Mid(MONAT, 4)) = sYear ");
                sbSqlQuery.Append("  GROUP BY  CInt(Mid(MONAT, 4)), SNR ");
                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            //Create Procedure YTD_Calculations
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE YTD_Calculations");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE YTD_Calculations (sYear Long, sMonth Long) AS");
                sbSqlQuery.Append(" SELECT CInt(Mid(MONAT, 4)) AS YEARSS, SNR, SUM(ABGANG) AS YTD");
                sbSqlQuery.Append(" FROM COST_MANAGEMENT_DATENBASIS");
                sbSqlQuery.Append(" WHERE  CInt(Mid(MONAT, 1, 2)) <= sMonth  and CInt(Mid(MONAT, 4)) = sYear");
                sbSqlQuery.Append(" GROUP BY   CInt(Mid(MONAT, 4)), SNR");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            //Create Procedure Net Revenu YTD
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE Net_Revenu_YTD");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE Net_Revenu_YTD (sYear Long, sMonth Long) AS");
                sbSqlQuery.Append(" SELECT CInt(Mid(MONAT, 4)) AS YEARSS, SNR, SUM(NE) AS  [Net_Revenu_YTD]");
                sbSqlQuery.Append(" FROM COST_MANAGEMENT_DATENBASIS");
                sbSqlQuery.Append(" WHERE  CInt(Mid(MONAT, 1, 2)) <= sMonth  and CInt(Mid(MONAT, 4)) = sYear");
                sbSqlQuery.Append(" GROUP BY   CInt(Mid(MONAT, 4)), SNR");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            //Create Procedure Current_part_Value
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE Current_part_Value");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE Current_part_Value as SELECT Cost_Management_Datenbasis.Monat, ");
                sbSqlQuery.Append("   Cost_Management_Datenbasis.Snr, Cost_Management_Datenbasis.ABGANG AS ABA ");
                sbSqlQuery.Append(" FROM Cost_Management_Datenbasis;");


                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //Create Procedure Net_Revenu_WR_YTD
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE Net_Revenu_WR_YTD");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE Net_Revenu_WR_YTD (sYear Long, sMonth Long) AS");
                sbSqlQuery.Append(" SELECT CInt(Mid(MONAT, 4)) AS YEARSS, SNR, SUM([NE ohne RW]) AS  [Net_Revenu_WR_YTD]");
                sbSqlQuery.Append(" FROM COST_MANAGEMENT_DATENBASIS");
                sbSqlQuery.Append(" WHERE  CInt(Mid(MONAT, 1, 2)) <= sMonth  and CInt(Mid(MONAT, 4)) = sYear");
                sbSqlQuery.Append(" GROUP BY   CInt(Mid(MONAT, 4)), SNR");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //PART_YRLY_VALUE
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop Table PART_YRLY_VALUES");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" SELECT Cost_Management_datenbasis.SNR, Avg(Cost_Management_datenbasis.[aktueller Monat HKSP]) AS AVG_HKSP_1, Cost_Management_datenbasis.[2018 HKSP] ");
                sbSqlQuery.Append("AS AVG_HKSP_2, Cost_Management_datenbasis.[2017 HKSP] ");
                sbSqlQuery.Append(" AS AVG_HKSP_3, Cost_Management_datenbasis.[2016 HKSP] ");
                sbSqlQuery.Append(" AS AVG_HKSP_4, Cost_Management_datenbasis.Mcode, Cost_Management_datenbasis.[Teile Sobsl], Cost_Management_datenbasis.LN,  ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.Benennung ");
                sbSqlQuery.Append(", Cost_summary.[Forcasted Manufacturing Cost Year 1] as Forecast_1, Cost_summary.[Forcasted Manufacturing Cost Year 2] as Forecast_2,");
                sbSqlQuery.Append(" Cost_summary.[Forcasted Manufacturing Cost Year 3] as Forecast_3, Cost_summary.[Forcasted Manufacturing Cost Year 5] as Forecast_5, ");
                sbSqlQuery.Append(" Cost_summary.[Forcasted Manufacturing Cost Year 10] as Forecast_10 INTO PART_YRLY_VALUES ");
                sbSqlQuery.Append(" FROM Cost_Management_datenbasis Left JOIN Cost_summary ON Cost_Management_datenbasis.SNR = Cost_summary.[Part No] ");
                sbSqlQuery.Append(" GROUP BY Cost_Management_datenbasis.SNR, Cost_Management_datenbasis.[2017 HKSP], Cost_Management_datenbasis.[2016 HKSP], ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.[2018 HKSP], Cost_Management_datenbasis.Mcode, Cost_Management_datenbasis.[Teile Sobsl], ");
                sbSqlQuery.Append(" Cost_Management_datenbasis.LN, Cost_Management_datenbasis.Benennung, Right([Cost_Management_datenbasis].[monat],4), ");
                sbSqlQuery.Append(" Cost_summary.[Forcasted Manufacturing Cost Year 1], Cost_summary.[Forcasted Manufacturing Cost Year 2], ");
                sbSqlQuery.Append(" Cost_summary.[Forcasted Manufacturing Cost Year 3], Cost_summary.[Forcasted Manufacturing Cost Year 5], ");
                sbSqlQuery.Append(" Cost_summary.[Forcasted Manufacturing Cost Year 10] HAVING(((Right([monat], 4))= 2019));");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //PART_YRLY_CHANGE
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop Table PART_YRLY_CHANGE");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("SELECT PART_YRLY_VALUES.SNR, ");
                sbSqlQuery.Append("Round(([Forecast_10] -[Forecast_5]) /[Forecast_5] * 100, 2)  AS Change_2028, ");
                sbSqlQuery.Append("Round(([Forecast_5] -[Forecast_3]) /[Forecast_3] * 100, 2)  AS Change_2023, ");
                sbSqlQuery.Append("Round(([Forecast_3] -[Forecast_2]) /[Forecast_2] * 100, 2)  AS Change_2021, ");
                sbSqlQuery.Append("Round(([Forecast_2] -[Forecast_1]) /[Forecast_1] * 100, 2)  AS Change_2020, ");
                sbSqlQuery.Append("Round(([Forecast_1] -[PART_YRLY_VALUES.AVG_HKSP_2]) /[PART_YRLY_VALUES.AVG_HKSP_2] * 100, 2) AS Change_2019, ");
                sbSqlQuery.Append("Round(([PART_YRLY_VALUES.AVG_HKSP_2] -[PART_YRLY_VALUES.AVG_HKSP_2017]) /[PART_YRLY_VALUES.AVG_HKSP_2017] * 100, 2) ");
                sbSqlQuery.Append("AS Change_2018, Round(([PART_YRLY_VALUES.AVG_HKSP_3] -[PART_YRLY_VALUES.AVG_HKSP_4]) /[PART_YRLY_VALUES.AVG_HKSP_4] * 100, 2)  ");
                sbSqlQuery.Append("AS Change_2017, Round(([PART_YRLY_VALUES.AVG_HKSP_4] -[PART_YRLY_VALUES.AVG_HKSP_5]) /[PART_YRLY_VALUES.AVG_HKSP_5] * 100, 2)  ");
                sbSqlQuery.Append("AS Change_2016, PART_YRLY_VALUES.Mcode, PART_YRLY_VALUES.[Teile Sobsl], PART_YRLY_VALUES.LN, PART_YRLY_VALUES.Benennung ");
                sbSqlQuery.Append(" INTO PART_YRLY_CHANGE ");
                sbSqlQuery.Append("FROM PART_YRLY_VALUES; ");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //Previous_Year_Ytd
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop Table Previous_Year_Ytd");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" SELECT Right([monat], 4) AS[year], SNR, [2018 HKSP] AS PART_COST, Sum(Abgang)AS PART_QNTY_TOTAL");
                sbSqlQuery.Append(" into Previous_Year_Ytd FROM Cost_Management_datenbasis");
                sbSqlQuery.Append("  GROUP BY Right([monat], 4), SNR, [2018 HKSP] HAVING(((Right([Cost_Management_datenbasis].[monat], 4))= 2018))");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //Create Procedure aktueller_Monat_HKSP
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE aktueller_Monat_HKSP");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }
                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append(" CREATE PROCEDURE aktueller_Monat_HKSP (sYear Long, sMonth Long) AS");
                sbSqlQuery.Append(" SELECT CInt(Mid(MONAT, 4)) AS YEARSS, SNR, Round(SUM([aktueller Monat HKSP]),2) AS  [AM_YTD]");
                sbSqlQuery.Append(" FROM COST_MANAGEMENT_DATENBASIS");
                sbSqlQuery.Append(" WHERE  CInt(Mid(MONAT, 1, 2)) <= sMonth  and CInt(Mid(MONAT, 4)) = sYear");
                sbSqlQuery.Append(" GROUP BY   CInt(Mid(MONAT, 4)), SNR");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            //Create Procedure TOTAL_VALUE_QUERY
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE TOTAL_VALUE_QUERY");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }

                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("CREATE PROCEDURE TOTAL_VALUE_QUERY as ");
                sbSqlQuery.Append(" SELECT 'INCREASE' AS GRAPH, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2018), 2) AS AVG_HKSP_2018, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2017), 2) AS AVG_HKSP_2017, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2016), 2)  ");
                sbSqlQuery.Append(" AS AVG_HKSP_2016, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2015), 2) AS AVG_HKSP_2015 ");
                sbSqlQuery.Append(",Round(AVG(Forecast_1),2) as AVG_HKSP_2019,Round(AVG(Forecast_2),2)as AVG_HKSP_2020,Round(AVG(Forecast_3),2)as AVG_HKSP_2021,");
                sbSqlQuery.Append(" Round(AVG(Forecast_5),2) as AVG_HKSP_2023,Round(AVG(Forecast_10),2) as AVG_HKSP_2028 FROM PART_YRLY_VALUES ");
                sbSqlQuery.Append(" HAVING(([PART_YRLY_VALUES].[AVG_HKSP_2018] >[PART_YRLY_VALUES].[AVG_HKSP_2017])); ");
                sbSqlQuery.Append("  UNION ALL ");
                sbSqlQuery.Append(" SELECT 'DECREASE' AS GRAPH, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2018), 2) AS AVG_HKSP_2018, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2017), 2) AS AVG_HKSP_2017, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2016), 2)  ");
                sbSqlQuery.Append("  AS AVG_HKSP_2016, ROUND(Avg(PART_YRLY_VALUES.AVG_HKSP_2015), 2) AS AVG_HKSP_2015 ");
                sbSqlQuery.Append(",Round(AVG(Forecast_1),2) as AVG_HKSP_2019,Round(AVG(Forecast_2),2)as AVG_HKSP_2020,Round(AVG(Forecast_3),2)as AVG_HKSP_2021,");
                sbSqlQuery.Append(" Round(AVG(Forecast_5),2) as AVG_HKSP_2023,Round(AVG(Forecast_10),2) as AVG_HKSP_2028 FROM PART_YRLY_VALUES ");
                sbSqlQuery.Append(" HAVING(([PART_YRLY_VALUES].[AVG_HKSP_2018] < [PART_YRLY_VALUES].[AVG_HKSP_2017])); ");

                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }


            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop Table teletechnic");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }

                sbSqlQuery = new StringBuilder();
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.Append(" Create Table teletechnic ([Technik_AP] Text(100),[Name] Text(100))");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6775','Marco Lichtenfeld')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3633','Bernd Kaufmann')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3679','Ralf Jackl')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6793','Tobias Scherrmann')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6776','Klaus Frueh')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K2222','Stefan Leistner')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N2456','Bernd Mueller')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K1111','Ivan Sliskovic')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3659','Lothar Hirsch')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('F4053','Peter Schietinger')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3683','Thorsten Dummer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6778','Jaee Palande')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3650','Reinhard Begoin')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3684','Dennis Bouche')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3615','Thomas Reinhold')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6743','Andreas Blubacher')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K3333','Sascha Hiert')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3605','Dieter Werling')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6764','Stefan Weis')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N2452','Guido Heid')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3682','Gerd Jaeger')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6724','Wolfgang Schmid')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3723','Gerold Laubscher')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4950','Franz Weber')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6788','Frank Lehrer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3619','Peter Bretschneider')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('F4051','Dieter Beck')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3730','Markus Scherrer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3536','Michael Schaaf')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('I2223','Neha Pereira')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3635','Andreas Haria')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N2469','Peter Buegel')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K8345','Markus Pallhon')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('F4059','Thorsten Grüninger')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K4444','Jens Rudat')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4959','Marco Arndt')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3716','Klaus Weckbart')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3722','Andreas Roller')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3656','Juergen Staudenmaier')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3711','Martin Laur')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3654','Miroslaw Kazimierek')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4951','Marc Hilsinger')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4958','Andreas Danner')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3627','Timo Gsell')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4961','Markus Terner')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6787','Thomas Steiner')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3681','Klaus Meister')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3634','Marco Pauls')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('I2221','Rohit Srivastava')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6705','Karl-Heinz Asel')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N2887','Andreas Zaucker')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3686','Frank Kronemayer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6799','Peter Dausch')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('I2230','Keerthi Gangdadhar')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3504','Timo Burkhart')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('I2229','Puneet Bellubbi')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3718','Christian Kuehn')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3616','Klaus Zimmer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3704','Raphael Korn')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3695','Michael Roether-Zupec')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6732','Andreas Uebele')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3677','Andreas Pfirrmann')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3617','Tom Frech')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6791','Gerhard Keller')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6735','Bekir Cetintas')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3641','Olaf Sigmund')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6785','Frank Marschner')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3623','Alexander Krieg')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N2425','Detlef Dewitz')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6781','Kurt Gomringer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3731','Thorsten Biegard')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4954','Christian Wolf')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6760','Harald Zeyfang')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6789','Rafael Marin')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6755','Ralph Eichhorn')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6786','Alexander Kientsch')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('F3221','Kai Felix Altmeier')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3693','Kemal Taskin')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3745','Joerg Schnaufer')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3289','Thomas Winter')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N4956','Ralf Jaeger')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('I2231','Shweta Kadam')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3743','Rainer Sturm')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('K6770','Ali Ayhan')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                sbSQL = new StringBuilder();
                sbSQL.Append("Insert Into teletechnic (Technik_AP,Name) values('N3678','Thomas Wetzka')");
                oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            //Create table for extrapolation factor 
            //try
            //{

            //    StringBuilder sbSQL = new StringBuilder();
            //    dbCon = new OleDbConnection();
            //    dbCon = clsCOn.Openconnection();
            //    oleCmd = new OleDbCommand();
            //    string strDbpath = AppDomain.CurrentDomain.BaseDirectory + "\\Hochrechnungsfaktor.xlsx";
            //    string str_Dbname = Path.GetFileName(strDbpath);
            //    DataTable dtdata = clsCOn.ImportDataUsingExcel(strDbpath, "sheet1$");
            //    if (dtdata.Rows.Count > 0)
            //    {
            //        sbSQL = new StringBuilder();
            //        try
            //        {
            //            sbSQL.Append("Drop table tblextrapolationfactor");
            //            oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
            //            oleCmd.ExecuteNonQuery();
            //            oleCmd.Dispose();
            //        }
            //        catch (Exception ex)
            //        {

            //        }
            //        try
            //        {
            //            sbSQL = new StringBuilder();
            //            sbSQL.Append(" Create Table tblextrapolationfactor ([Sparte] Text(10),[Monat] Integer,[Index] Text(15),HR_Faktor DECIMAL(18,2))");
            //            oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
            //            oleCmd.ExecuteNonQuery();
            //            oleCmd.Dispose();
            //        }
            //        catch (Exception ex) { }

            //        foreach (DataRow drs in dtdata.Rows)
            //        {
            //            sbSQL = new StringBuilder();
            //            sbSQL.Append("Insert Into tblextrapolationfactor Values ");
            //            sbSQL.Append("('" + drs[0].ToString() + "','" + drs[1].ToString() + "',");
            //            sbSQL.Append("'" + drs[2].ToString() + "','" + Math.Round(Convert.ToDecimal(drs[3].ToString()), 3) + "')");
            //            dbCon = clsCOn.Openconnection();
            //            oleCmd = new OleDbCommand(sbSQL.ToString(), dbCon);
            //            oleCmd.ExecuteNonQuery();
            //            oleCmd.Dispose();
            //        }
            //    }
            //}
            //catch (Exception ex) { }

            //Create Procedure for GetSummary Table
            try
            {
                StringBuilder sbSqlQuery = new StringBuilder();
                try
                {
                    sbSqlQuery.Append("Drop PROCEDURE GET_Summeri_Data ");
                    oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                    oleCmd.ExecuteNonQuery();
                    oleCmd.Dispose();
                }
                catch (Exception ex) { }

                sbSqlQuery = new StringBuilder();
                sbSqlQuery.Append("CREATE PROCEDURE GET_Summeri_Data  (sYear Long, sMonth Long) as ");
                sbSqlQuery.Append(" SELECT DISTINCT  a.MONAT, y.YTD, a.snr AS [Part No], a.Bezeichnung AS [Part Name], a.Mcode AS [Marketing Code], a.Serie_von, a.Serie_bis,");
                sbSqlQuery.Append(" IIF(a.[2016 HKSP] Is Null, 0, a.[2016 HKSP]) AS [2016a],IIF(ab15.Abgang2016 is Null,1,ab15.Abgang2016) as [Abgang2016],");
                sbSqlQuery.Append(" IIF(a.[2017 HKSP] Is Null, 0, a.[2017 HKSP]) AS [Manufacturing cost YTD-2Y Year per Part],  ");
                sbSqlQuery.Append(" iif(([Manufacturing cost YTD-2Y Year per Part] * ab16.Abgang2017) is Null,0,([Manufacturing cost YTD-2Y Year per Part] * ab16.Abgang2017)) AS [Manufacturing total cost YTD-2Y],");
                sbSqlQuery.Append(" IIf(((a.[2017 HKSP]) - [2016a]) Is Null, 0, Round(((a.[2017 HKSP]) -[2016a]), 2)) as  [Delta value per part YTD Vs YTD-2Y], ");
                sbSqlQuery.Append(" IIf(([2016a])=0, 0,");
                sbSqlQuery.Append(" IIF((((a.[2017 HKSP]) - ([2016a])) / ([2016a] )) Is Null,0,(((a.[2017 HKSP] ) - ([2016a])) / ([2016a]))))");
                sbSqlQuery.Append(" AS [% change YTD vs YTD-2Y], ");
                sbSqlQuery.Append(" IIF(([Delta value per part YTD Vs YTD-2Y]* ab16.Abgang2017)is Null,0, ");
                sbSqlQuery.Append("([Delta value per part YTD Vs YTD-2Y] * ab16.Abgang2017)) AS [Delta all Parts YTD Vs YTD-2], ");
                sbSqlQuery.Append(" IIF(a.[2018 HKSP] Is Null, 0, a.[2018 HKSP]) AS [Manufacturing cost YTD-1Y Year per Part], ");
                sbSqlQuery.Append(" Round([Manufacturing cost YTD-1Y Year per Part] * ptr.Abgang2018,2) AS [Manufacturing total cost YTD-1Y], ");
                sbSqlQuery.Append(" IIf((a.[2018 HKSP] - a.[2017 HKSP]) Is Null, 0, Round(a.[2018 HKSP] - a.[2017 HKSP], 2)) AS [Delta value per part YTD Vs YTD-1Y],");
                sbSqlQuery.Append(" IIf(a.[2017 HKSP] is Null , 0, ");
                sbSqlQuery.Append(" IIF((((a.[2018 HKSP]) - a.[2017 HKSP]) / a.[2017 HKSP]) Is Null,0,");
                sbSqlQuery.Append("(((a.[2018 HKSP]) -a.[2017 HKSP]) / a.[2017 HKSP]))) AS [% change YTD vs YTD-1Y], ");
                sbSqlQuery.Append(" (([Delta value per part YTD Vs YTD-1Y])*(ptr.Abgang2018)) AS [Delta all Parts YTD Vs YTD-1], ");
                sbSqlQuery.Append(" Round(([aktueller Monat HKSP]), 2) as [Current Month Manufacturing Cost per Part],a.[ABGANG] AS [Part Quantity Current Month], ");
                sbSqlQuery.Append(" y.YTD, Round(a.[aktueller Monat HKSP] * a.ABGANG, 2) AS [Current Month Manufacturing Cost Total], ");
                sbSqlQuery.Append(" (([Manufacturing Material cost Current Year Estimation])-([Manufacturing total cost YTD-1Y])) as [Delta all part Current Year], ");
                sbSqlQuery.Append(" (a.[aktueller Monat HKSP] * ptr.Abgang2018) as [Manufacturing Material cost Current Year Estimation],");
                sbSqlQuery.Append(" Round(([aktueller Monat HKSP]-[2018 HKSP]), 2) as [Delta Value per Part Current Year], ");
                sbSqlQuery.Append(" IIf(Round([Current Month Manufacturing Cost per Part] * y.YTD, 2) Is Null, 0, Round([Current Month Manufacturing Cost per Part] *  y.YTD, 2)) AS [YTD Manufacturing Cost per Year],");
                sbSqlQuery.Append(" a.[Teile Sobsl] AS [Part Sobsl], a.Ersterstellung AS [Creation date],d.[SCA Beschaffung Name] as [Dispo Name], a.Disponent_AP AS [Dispo Code],t.Name as [Part technique],");
                sbSqlQuery.Append(" a.LN AS [Supplier Code], a.Benennung AS [Supplier Name], a.[Blp alt] AS [Gross List Price  BLP / VLP],");
                sbSqlQuery.Append(" Round(a.[Rw neu],2) AS [Return Value(RW)], Round(([Gross List Price  BLP / VLP])-([Return Value(RW)]),2) AS [Trade Value], Round(a.NE,2) AS [NET Revenue per VLP],");
                sbSqlQuery.Append(" Round(NR.Net_Revenu_YTD,2) as Net_Revenu_YTD, Round(a.[NE ohne RW],2) AS [Net revenue without return value current Month],Round(w.Net_Revenu_WR_YTD,2) AS [Net Revenu Without Return YTD],");
                sbSqlQuery.Append(" a.SPM_BLP_Gilt_Ab AS [Price Date], Round([Blp alt]*ABGANG,2) AS [Gross Revenue Current Month], Round([Blp alt]*y.YTD,2) AS [Gross Revenue Current YTD],");
                sbSqlQuery.Append(" a.SKW AS [Factory Cost SKW], a.[steuerlicher Wertansatz] AS [tax value approach], Round(a.[SKW ohne SWA],2) AS [SKW without  SWA],");
                sbSqlQuery.Append(" Round((Materialkosten+Rückführkosten+GuK+iif([Verpack absolut] is Null,0,[Verpack absolut])+[VTKP abs]+[VTKF abs]/a.[ABGANG]),2)  AS [Total Cost per Part],");
                sbSqlQuery.Append(" Round((Materialkosten+Rückführkosten+GuK+iif([Verpack absolut] is Null,0,[Verpack absolut])+[VTKP abs]+[VTKF abs]),2) AS [Total Cost],");
                sbSqlQuery.Append(" Round((Materialkosten+GuK+iif([Verpack absolut] is Null,0,[Verpack absolut])+[VTKP abs]+[VTKF abs]),2) AS [Total Cost excl Return Cost], tp.[Total YTD Cost per Part] AS [Total YTD Cost],");
                sbSqlQuery.Append(" a.Materialkosten AS [Material Cost], Round(a.Rückführkosten,2) AS [Return Cost], Round(a.GuK,2) as GUK, Round(iif([Verpack absolut] is Null,0,[Verpack absolut]),2) AS [Package Absolute], ");
                sbSqlQuery.Append(" Round(a.[VTKP abs],2) as [VTKP abs], Round(a.[VTKF abs],2) as [VTKF abs],");
                sbSqlQuery.Append(" Round(a.[Ergebnis Werk],2) AS [Result Factory], IIF(a.[CORE Wert Lieferant] is Null,0,a.[CORE Wert Lieferant]) AS [CORE value Supplior], a.[Gutschriften Altteile] AS [Credit Voucher Old Part],");
                sbSqlQuery.Append(" Round(([Net Revenu Without Return YTD])-([Total Cost]+[Result Factory]+[Credit Voucher Old Part]),2) AS Result, a.Rabattgruppe AS [Discount Group],a.Baureihe as [Model Type Baureihe]");
                sbSqlQuery.Append(" FROM (((((((((((COST_MANAGEMENT_DATENBASIS AS a  LEFT JOIN YTD_Calculations AS y ON a.snr = y.snr)  ");
                sbSqlQuery.Append(" LEFT JOIN  PART_SSTART_SEND_PROGEND_DATE AS q ON a.snr = q.snr and a.Monat = q.Monat)  ");
                sbSqlQuery.Append(" LEFT JOIN Total_YTD_Cost_per_Part AS tp ON a.snr = tp.snr) LEFT JOIN Net_Revenu_Ytd AS NR ON a.SNR = NR.SNR)  ");
                sbSqlQuery.Append(" LEFT JOIN Net_Revenu_WR_YTD AS w ON a.snr = w.SNR ) Left Join Current_part_Value cp on a.snr = cp.snr and a.Monat = CP.Monat) ");
                sbSqlQuery.Append(" Left join   aktueller_Monat_HKSP am on a.snr = am.snr) ");
                sbSqlQuery.Append(" Left Join Dispo d on a.Disponent_AP = d.Dispokürzel) Left Join teletechnic t on a.Technik_AP = t.Technik_AP) ");
                sbSqlQuery.Append(" Left Join Abgang2018 ptr on(a.snr = ptr.SNR) )");
                sbSqlQuery.Append(" Left Join Abgang2017 ab16 on(a.snr = ab16.SNR) )");
                sbSqlQuery.Append(" Left Join Abgang2016 ab15 on(a.snr = ab15.SNR)");
                sbSqlQuery.Append(" WHERE (((Mid(a.MONAT,4))=[sYear]) And ((CInt(Mid(a.MONAT,1,2)))=[sMonth]))");
                sbSqlQuery.Append(" GROUP BY a.MONAT, a.snr, a.Bezeichnung, a.Mcode, a.Serie_von, a.Serie_bis,  a.[2017 HKSP], a.[2018 HKSP], y.YTD,");
                sbSqlQuery.Append(" a.[Teile Sobsl], a.Ersterstellung, a.Disponent_AP, a.LN, a.Benennung, a.[BLP neu], a.[Rw neu], a.NE, NR.Net_Revenu_YTD, a.[NE ohne RW],");
                sbSqlQuery.Append(" w.Net_Revenu_WR_YTD, a.SPM_BLP_Gilt_Ab, a.SKW, a.[steuerlicher Wertansatz], a.[SKW ohne SWA], tp.[Total YTD Cost per Part], a.Materialkosten,");
                sbSqlQuery.Append(" a.Rückführkosten, a.GuK, a.[Verpack absolut], a.[VTKP abs], a.[VTKF abs], a.[Ergebnis Werk], a.[CORE Wert Lieferant], a.[Gutschriften Altteile],");
                sbSqlQuery.Append(" a.Rabattgruppe, a.[aktueller Monat HKSP], a.ABGANG, a.[Blp alt],cp.ABA,am.AM_YTD,d.[SCA Beschaffung Name],t.Name,a.Baureihe,");
                sbSqlQuery.Append(" ptr.Abgang2018, [2016 HKSP],ab16.Abgang2017,ab15.Abgang2016");
                sbSqlQuery.Append(" ORDER BY a.MONAT , a.snr, a.Mcode;");


                oleCmd = new OleDbCommand(sbSqlQuery.ToString(), dbCon);
                oleCmd.ExecuteNonQuery();
                oleCmd.Dispose();
            }
            catch (Exception ex)
            {

            }

            return true;
        }
    }
}
