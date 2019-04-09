using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace cost_management
{
    public class ClsForcasting
    {

        //
        //Simple Moving Average
        //
        //            ( Dt + D(t-1) + D(t-2) + ... + D(t-n+1) )
        //  F(t+1) =  -----------------------------------------
        //                              n

        //Atcs Forecasting Model For HT parts
        public ForecastTable simpleMovingAverageForHT(decimal[] values, int Extension, int Periods, int Holdout)
        {
            ForecastTable dt = new ForecastTable();
            bool bErrorInValue = false;
            decimal dForecast1 = 0;
            decimal dPeravge = 0;
            decimal dChangeValue = 0;
            decimal dNewavg = 0;
            decimal dPerchange = 0;
            ClsCommonFunction clsObj = new ClsCommonFunction();
            decimal dValue = 0;
            int iyear = Convert.ToInt32(clsObj.GetYear(3, "-"));
            decimal dError = 0;
            try
            {
                for (Int32 i = 0; i < values.Length + Extension; i++)
                {
                    //Insert a row for each value in set
                    DataRow row = dt.NewRow();
                    dt.Rows.Add(row);

                    row.BeginEdit();
                    //assign its sequence number
                    row["Instance"] = i;
                    row["Monat"] = iyear + i;
                    if (i < values.Length)
                    {//processing values which actually occurred
                        row["Value"] = values[i];
                    }

                    //Indicate if this is a holdout row
                    row["Holdout"] = (i > (values.Length - Holdout)) && (i < values.Length);
                    //Forecast calcluation 
                    if (i == 0)
                    {
                        //Initialize first row with its own value
                        row["Forecast"] = values[i];
                        row["PercentageChange"] = 0;
                    }
                    else if (i <= values.Length - Holdout)
                    {
                        //processing values which actually occurred, but not in holdout set
                        decimal avg = 0;
                        dError = 0;
                        dPeravge = 0;
                        dNewavg = 0;
                        DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                        foreach (DataRow priorRow in rows)
                        {
                            avg += (Decimal)priorRow["Value"];
                            dPeravge += (Decimal)(Math.Round(Convert.ToDecimal(priorRow["PercentError"].ToString().Equals("") ? "0" : priorRow["PercentError"].ToString()), 2));
                        }
                        avg /= rows.Length;
                        dError = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["PercentageChange"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentageChange"].ToString()), 2);
                        dPeravge /= rows.Length;
                        dNewavg = avg + (avg * dPeravge);//
                        dChangeValue = values[i - 1] * (dError); ;//finding decrease in value

                        dPerchange = 0;
                        if (!i.Equals(4))
                        {
                            if ((decimal)values[i - 1] > 0)
                                dPerchange = ((decimal)values[i] - (decimal)values[i - 1]) / (decimal)values[i - 1];// Pencetage Change
                            else
                                dPeravge = 0;
                        }
                        if (dError < 0)
                        {
                            if (avg < (decimal)values[i - 1])//if percentage change value is in decrease 
                            {
                                //row["Forecast"] = Math.Round(avg, 2);
                                row["Forecast"] = Math.Round(avg > 0 ? avg : (decimal)values[i - 1] + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg > 0 ? avg : (decimal)values[i - 1] + dChangeValue, 2);
                            }
                            else
                            {
                                //row["Forecast"] = Math.Round(dNewavg, 2);
                                row["Forecast"] = Math.Round(dNewavg > 0 ? dNewavg : (decimal)values[i - 1] + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(dNewavg > 0 ? dNewavg : (decimal)values[i - 1] + dChangeValue, 2);
                            }
                        }
                        else
                        {
                            // row["Forecast"] = Math.Round(avg, 2);
                            if (avg < (decimal)values[i - 1])//if forecasting is Increaseing
                            {
                                row["Forecast"] = Math.Round((decimal)values[i - 1] + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round((decimal)values[i - 1] + dChangeValue, 2);
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                    }
                    else
                    {//must be in the holdout set or the extension9
                        decimal dvals = 0;
                        decimal avg = 0;
                        dNewavg = 0;
                        //get the Periods-prior rows and calculate an average actual value for last 3 years
                        DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                        foreach (DataRow priorRow in rows)
                        {
                            if ((Int32)priorRow["Instance"] < values.Length)
                            {//in the test or holdout set
                                avg += (Decimal)priorRow["Value"];
                            }
                            else
                            {//extension, use forecast since we don't have an actual value
                                avg += (Decimal)priorRow["Forecast"];
                            }
                        }
                        avg /= rows.Length;
                        if ((decimal)dt.Rows[i - 1]["Forecast"] > 0)
                            dPerchange = Math.Round(((decimal)avg - (decimal)dt.Rows[i - 1]["Forecast"]) / (decimal)dt.Rows[i - 1]["Forecast"], 2);// Pencetage Change
                        else
                            dPeravge = 0;

                        dNewavg = Math.Round(avg + (avg * dPeravge));

                        dChangeValue = Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) * (dError >= 2 ? 1 : dError);

                        //set the forecasted value
                        if (dError < 0)
                        {
                            if (avg < Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()))//if percentage chnage value is in decrease 
                            {
                                //row["Forecast"] = Math.Round(avg, 2);
                                row["Forecast"] = Math.Round(avg > 0 ? avg : Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                            }
                            else
                            {
                                //row["Forecast"] = Math.Round(dNewavg, 2);
                                row["Forecast"] = Math.Round(dNewavg > 0 ? dNewavg : Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                            }
                        }
                        else
                        {
                            if (avg < (decimal)Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()))//if forecasting is Increaseing
                            {
                                row["Forecast"] = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                    }
                    row.EndEdit();
                }
                dt.AcceptChanges();
            }
            catch (Exception ex)
            {

            }
            return dt;
        }
        public ForecastTable simpleMovingAverageForPartQTY(decimal[] values, int Extension, int Periods, int Holdout)
        {
            ForecastTable dt = new ForecastTable();
            bool bErrorInValue = false;
            decimal dForecast1 = 0;
            decimal dPeravge = 0;
            decimal dChangeValue = 0;
            decimal dNewavg = 0;
            decimal dPerchange = 0;
            ClsCommonFunction clsObj = new ClsCommonFunction();
            decimal dValue = 0;
            int iyear = Convert.ToInt32(clsObj.GetYear(3, "-"));
            decimal dError = 0;
            try
            {
                for (Int32 i = 0; i < values.Length + Extension; i++)
                {
                    //Insert a row for each value in set
                    DataRow row = dt.NewRow();
                    dt.Rows.Add(row);

                    row.BeginEdit();
                    //assign its sequence number
                    row["Instance"] = i;
                    row["Monat"] = iyear + i;
                    if (i < values.Length)
                    {//processing values which actually occurred
                        row["Value"] = values[i];
                    }

                    //Indicate if this is a holdout row
                    row["Holdout"] = (i > (values.Length - Holdout)) && (i < values.Length);
                    //Forecast calcluation 
                    if (i == 0)
                    {
                        //Initialize first row with its own value
                        row["Forecast"] = values[i];
                        row["PercentageChange"] = 0;
                    }
                    else if (i <= values.Length - Holdout)
                    {
                        //processing values which actually occurred, but not in holdout set
                        decimal avg = 0;
                        dError = 0;
                        dPeravge = 0;
                        dNewavg = 0;
                        DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                        foreach (DataRow priorRow in rows)
                        {
                            avg += (Decimal)priorRow["Value"];
                            //dError += (Decimal)(Math.Round(Convert.ToDecimal(priorRow["PercentageChange"].ToString().Equals("") ? "0" : priorRow["PercentageChange"].ToString()), 2));
                            dPeravge += (Decimal)(Math.Round(Convert.ToDecimal(priorRow["PercentError"].ToString().Equals("") ? "0" : priorRow["PercentError"].ToString()), 2));
                        }
                        avg /= rows.Length;
                        dError = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["PercentageChange"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentageChange"].ToString()), 2);
                        dPeravge /= rows.Length;
                        dNewavg = avg + (avg * dPeravge);//
                        dChangeValue = values[i - 1] * (dError >= 2 ? 1 : dError);//finding decrease in value

                        dPerchange = 0;
                        if (!i.Equals(4))
                        {
                            if ((decimal)values[i - 1] > 0)
                                dPerchange = ((decimal)values[i] - (decimal)values[i - 1]) / (decimal)values[i - 1];// Pencetage Change
                            else
                                dPeravge = 0;
                        }
                        if (dError < 0)
                        {
                            if (avg < (decimal)values[i - 1])//if percentage chnage value is in decrease 
                            {
                                //row["Forecast"] = Math.Round(avg, 2);
                                row["Forecast"] = Math.Round(avg > 0 ? avg : (decimal)values[i - 1] + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg > 0 ? avg : (decimal)values[i - 1] + dChangeValue, 2);
                            }
                            else
                            {
                                //row["Forecast"] = Math.Round(dNewavg, 2);
                                row["Forecast"] = Math.Round(dNewavg > 0 ? dNewavg : (decimal)values[i - 1] + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(dNewavg > 0 ? dNewavg : (decimal)values[i - 1] + dChangeValue, 2);
                            }
                        }
                        else
                        {
                            // row["Forecast"] = Math.Round(avg, 2);
                            if (avg < (decimal)values[i - 1])//if forecasting is Increaseing
                            {
                                row["Forecast"] = Math.Round(avg + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round((decimal)values[i - 1] + dChangeValue, 2);
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                    }
                    else
                    {//must be in the holdout set or the extension9
                        decimal dvals = 0;
                        decimal avg = 0;
                        dNewavg = 0;
                        //get the Periods-prior rows and calculate an average actual value for last 3 years
                        DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                        foreach (DataRow priorRow in rows)
                        {
                            if ((Int32)priorRow["Instance"] < values.Length)
                            {//in the test or holdout set
                                avg += (Decimal)priorRow["Value"];
                            }
                            else
                            {//extension, use forecast since we don't have an actual value
                                avg += (Decimal)priorRow["Forecast"];
                            }
                        }

                        avg /= rows.Length;

                        if ((decimal)dt.Rows[i - 1]["Forecast"] > 0)
                            dPerchange = ((decimal)avg - (decimal)dt.Rows[i - 1]["Forecast"]) / (decimal)dt.Rows[i - 1]["Forecast"];// Pencetage Change
                        else
                            dPeravge = 0;
                        dNewavg = avg + (avg * dPeravge);
                        dChangeValue = Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) * (dError >= 2 ? 1 : dError);

                        //set the forecasted value
                        if (dError < 0)
                        {
                            if (avg < Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()))//if percentage chnage value is in decrease 
                            {
                                //row["Forecast"] = Math.Round(avg, 2);
                                row["Forecast"] = Math.Round(avg > 0 ? avg : Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                            }
                            else
                            {
                                //row["Forecast"] = Math.Round(dNewavg, 2);
                                row["Forecast"] = Math.Round(dNewavg > 0 ? dNewavg : Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                            }
                        }
                        else
                        {
                            if (avg < (decimal)Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()))//if forecasting is Increaseing
                            {
                                row["Forecast"] = Math.Round(avg + dChangeValue, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["Forecast"].ToString()) + dChangeValue, 2);
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                row["PercentageChange"] = Math.Round(dPerchange, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                    }
                    row.EndEdit();
                }
                dt.AcceptChanges();
            }
            catch (Exception ex)
            {

            }
            return dt;
        }
        public ForecastTable simpleMovingAverageDAG(decimal[] values, int Extension, int Periods, int Holdout, int iEndYear, string sSupplierName)
        {
            ForecastTable dt = new ForecastTable();
            ClsCommonFunction clsObj = new ClsCommonFunction();
            decimal dValue = 0;
            int iCurrentValue = 0;
            bool bisCheck = false;
            decimal dForecast1 = 0;
            decimal dError = 0;
            decimal dCurrentValue = 0;
            //DataTable dtforecasting = simpleMovingAverageForHT(values, Extension, Periods, Holdout);
            try
            {
                int iyear = Convert.ToInt32(clsObj.GetYear(3, "-"));
                for (Int32 i = 0; i < values.Length + Extension; i++)
                {
                    //Insert a row for each value in set
                    DataRow row = dt.NewRow();
                    dt.Rows.Add(row);

                    row.BeginEdit();
                    //assign its sequence number
                    row["Instance"] = i;
                    //adding Year of forecasting
                    row["Monat"] = iyear + i;
                    //processing values which actually occurred
                    if (i < values.Length)
                    {
                        row["Value"] = values[i];
                    }

                    //Indicate if this is a holdout row
                    row["Holdout"] = (i > (values.Length - Holdout)) && (i < values.Length);
                    //Forecast calcluation 
                    if (i == 0)
                    {
                        //Initialize first row with its own value
                        row["Forecast"] = values[i];
                    }
                    else if (i < values.Length - Holdout)
                    {
                        //processing values which actually occurred, but not in holdout set
                        decimal avg = 0;
                        dError = 0;
                        DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                        foreach (DataRow priorRow in rows)
                        {
                            avg += (Decimal)priorRow["Value"];
                            //dError += (Decimal)(Convert.ToDecimal(priorRow["PercentError"].ToString().Equals("") ? "0" : priorRow["PercentError"].ToString()));
                        }
                        avg /= rows.Length;
                        dError = Math.Round(Convert.ToDecimal(dt.Rows[i - 1]["PercentageChange"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentageChange"].ToString()), 2); ;
                        dValue = avg * (1 + Math.Round(dError, 2));
                        //(Convert.ToDecimal(dt.Rows[i - 1]["PercentError"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentError"].ToString()));
                        if (!avg.Equals(dValue))
                            avg = avg - dValue;
                        if (i.Equals(4))
                        {
                            if (Math.Round(avg, 2) < 0) // Cheking nigavative value
                            {
                                row["Forecast"] = Math.Round(values[i - 1], 2);
                                dForecast1 = (decimal)values[i - 1];
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                        else
                            row["Forecast"] = Math.Round(avg, 2);
                        iCurrentValue = i;
                    }
                    else
                    {//must be in the holdout set or the extension9

                        if (Convert.ToInt32(row["Monat"].ToString()) <= iEndYear)
                        {
                            //bisCheck = true;
                            decimal avg = 0;
                            //get the Periods-prior rows and calculate an average actual value
                            DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                            foreach (DataRow priorRow in rows)
                            {
                                if ((Int32)priorRow["Instance"] < values.Length)
                                {//in the test or holdout set
                                    avg += (Decimal)priorRow["Value"];
                                }
                                else
                                {//extension, use forecast since we don't have an actual value
                                    avg += (Decimal)priorRow["Forecast"];
                                }
                            }
                            avg /= rows.Length;
                            dValue = avg * (dError > 0 ? Math.Round(dError, 2) : 1 + Math.Round(dError, 2));
                            //(Convert.ToDecimal(dt.Rows[i - 1]["PercentError"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentError"].ToString()));
                            if (!avg.Equals(dValue))
                                avg = avg + dValue;

                            //set the forecasted value
                            if (i.Equals(4))
                            {
                                if (Math.Round(avg, 2) < 0)
                                {
                                    row["Forecast"] = Math.Round(values[i - 1], 2);
                                    dForecast1 = (decimal)values[i - 1];
                                }
                                else
                                {
                                    row["Forecast"] = Math.Round(avg, 2);
                                    dForecast1 = Math.Round(avg, 2);
                                }
                            }
                            else
                                row["Forecast"] = Math.Round(avg, 2);
                            iCurrentValue = i;
                        }
                        else if ((Convert.ToInt32(row["Monat"].ToString()) - iEndYear) <= 5)// curretn month -part end date is <=5
                        {
                            int ival = (Convert.ToInt32(row["Monat"].ToString()) - iEndYear);
                            if (bisCheck.Equals(true))
                            {
                                dCurrentValue = Convert.ToDecimal(dt.Rows[iCurrentValue]["Forecast"].ToString().Equals("") ? dt.Rows[iCurrentValue]["Value"].ToString() :
                                                          dt.Rows[iCurrentValue]["Forecast"].ToString());
                            }
                            else
                            {
                                if (iEndYear > 2014 && iEndYear < Convert.ToInt32(clsObj.GetYear(1, "+")))
                                {
                                    if (values[0].ToString().Equals("0") && iEndYear == 2015)
                                    {
                                        dCurrentValue = values[3];
                                    }
                                    else if (values[1].ToString().Equals("0") && iEndYear == 2016)
                                    {
                                        dCurrentValue = values[3];
                                    }
                                    else if (values[2].ToString().Equals("0") && iEndYear == 2017)
                                    {
                                        dCurrentValue = values[3];
                                    }
                                    else
                                    {
                                        int ivalues = 0;
                                        switch (iEndYear)
                                        {
                                            case 2015:
                                                ivalues = 0;
                                                break;
                                            case 2016:
                                                ivalues = 1;
                                                break;
                                            case 2017:
                                                ivalues = 2;
                                                break;
                                            case 2018:
                                                ivalues = 3;
                                                break;
                                            default:
                                                break;
                                        }
                                        dCurrentValue = values[ivalues];
                                    }
                                }
                                else
                                {
                                    dCurrentValue = values[3];
                                }
                            }
                            decimal avgs = ForcastingWithDAG(sSupplierName, ival, dCurrentValue); //GetCurrentValueForForecasting(iEndYear, dtforecasting));

                            row["Forecast"] = Math.Round(avgs, 2);
                        }
                        else
                        {

                            decimal avg = 0;

                            DataRow[] rows = dt.Select("Instance>=" + (i - Periods).ToString() + " AND Instance < " + i.ToString(), "Instance");
                            foreach (DataRow priorRow in rows)
                            {
                                if ((Int32)priorRow["Instance"] < values.Length)
                                {//in the test or holdout set
                                    avg += (Decimal)priorRow["Value"];
                                }
                                else
                                {//extension, use forecast since we don't have an actual value
                                    avg += (Decimal)priorRow["Forecast"];
                                }
                            }
                            avg /= rows.Length;
                            dValue = avg * (dError > 0 ? Math.Round(dError, 2) : 1 + Math.Round(dError, 2));
                            //(Convert.ToDecimal(dt.Rows[i - 1]["PercentError"].ToString().Equals("") ? "0" : dt.Rows[i - 1]["PercentError"].ToString()));
                            if (!avg.Equals(dValue))
                                avg = avg + dValue;
                            if (Math.Round(avg, 2) < 0)
                            {
                                row["Forecast"] = dForecast1;
                            }
                            else
                            {
                                row["Forecast"] = Math.Round(avg, 2);
                                dForecast1 = Math.Round(avg, 2);
                            }
                        }
                    }
                    row.EndEdit();
                }

                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                for (Int32 i = 0; i < values.Length + Extension; i++)
                {
                    DataRow row = dt.NewRow();
                    dt.Rows.Add(row);

                    row.BeginEdit();
                    //assign its sequence number
                    row["Instance"] = i;
                    row["value"] = 0;
                    row["Forecast"] = 0;
                    row.EndEdit();
                    dt.AcceptChanges();
                }
            }
            return dt;
        }
        public decimal GetCurrentValueForForecasting(int iYear, DataTable dtforecast)
        {
            decimal dValuetobeForecast = 0;
            try
            {
                if (iYear > Convert.ToInt32(new ClsCommonFunction().GetYear(0, "")))
                {
                    DataRow[] drs = dtforecast.Select("Monat=" + iYear);
                    foreach (DataRow dr in drs)
                    {
                        dValuetobeForecast = Convert.ToDecimal(dr["Forecast"].ToString().Equals("") ? "0" : dr["Forecast"].ToString());
                    }
                }
                else
                {
                    DataRow[] drs = dtforecast.Select("[Monat]='" + new ClsCommonFunction().GetYear(0, "").ToString() + "'");
                    foreach (DataRow dr in drs)
                    {
                        dValuetobeForecast = Convert.ToDecimal(dr["Forecast"].ToString().Equals("") ? "0" : dr["Forecast"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return dValuetobeForecast;
        }
        public decimal ForcastingWithDAG(string sName, int iYear, decimal dvalue)
        {
            decimal dForcasValue = 0;
            try
            {
                switch (sName)
                {
                    case "Robert Bosch GmbH":
                        switch (iYear)
                        {
                            case 1:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)25 / (decimal)100));
                                break;
                            case 2:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)45 / (decimal)100));
                                break;
                            case 3:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)65 / (decimal)100));
                                break;
                            case 4:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)200 / (decimal)100));
                                break;
                            case 5:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)200 / (decimal)100));
                                break;
                        }
                        break;
                    case "Continental":
                        switch (iYear)
                        {
                            case 1:
                                dForcasValue = dvalue + (dvalue * (decimal)(100 / (decimal)100));
                                break;
                            case 2:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)100 / (decimal)100));
                                break;
                            case 3:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)100 / (decimal)100));
                                break;
                            case 4:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)250 / (decimal)100));
                                break;
                            case 5:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)250 / (decimal)100));
                                break;
                        }
                        break;
                    default:
                        switch (iYear)
                        {
                            case 1:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)50 / (decimal)100));
                                break;
                            case 2:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)75 / (decimal)100));
                                break;
                            case 3:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)100 / (decimal)100));
                                break;
                            case 4:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)200 / (decimal)100));
                                break;
                            case 5:
                                dForcasValue = dvalue + (dvalue * (decimal)((decimal)200 / (decimal)100));
                                break;
                        }
                        break;
                }
            }
            catch
            {
                dForcasValue = 0;
            }
            return dForcasValue;
        }
    }
    public class ForecastTable : DataTable
    {
        // An instance of a DataTable with some default columns.  Expressions help quickly calculate E
        public ForecastTable()
        {
            this.Columns.Add("Monat", typeof(string));    //The position in which this value occurred in the time-series
            this.Columns.Add("PartNo", typeof(string));     //The value which actually occurred
            this.Columns.Add("Instance", typeof(Int32));    //The position in which this value occurred in the time-series
            this.Columns.Add("Value", typeof(Decimal));     //The value which actually occurred
            this.Columns.Add("Forecast", typeof(Decimal));  //The forecasted value for this instance
            this.Columns.Add("Holdout", typeof(Boolean));   //Identifies a holdout actual value row, for testing after err is calculated
            this.Columns.Add("PercentageChange", typeof(Decimal));

            //E(t) = D(t) - F(t)
            this.Columns.Add("Error", typeof(Decimal), "Value-Forecast");
            //Absolute Error = |E(t)|
            this.Columns.Add("AbsoluteError", typeof(Decimal), "IIF(Error>=0, Error, Error * -1)");
            //Percent Error = E(t) / D(t)
            this.Columns.Add("PercentError", typeof(Decimal), "IIF(Value<>0, Error / Value, Null)");
            //Absolute Percent Error = |E(t)| / D(t)
            this.Columns.Add("AbsolutePercentError", typeof(Decimal), "IIF(Value <> 0, AbsoluteError / Value, Null)");
        }
    }
}
