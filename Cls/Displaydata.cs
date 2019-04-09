using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
namespace cost_management
{
    public class Displaydata
    {
        public string PART_NUMBER { get; set; }
        public string PART_NAME { get; set; }
        public string SERIES_PROD_SDATE { get; set; }
        public string SERIES_PROD_EDATE { get; set; }
        public string PROG_SERIES_EDATE { get; set; }
        public string MANUFACTR_PART_COST_YTD_2Y { get; set; }
        public string PART_QUANTITY_YTD_2Y { get; set; }
        public string MANUFACTR_TOTAL_COST_YTD_2Y { get; set; }
        public string DELTA_VALUE_PART_YTD_2Y { get; set; }
        public string DELTA_CHANGE_PART_YTD_2Y { get; set; }
        public string DELTA_VALUE_ALL_PART_YTD_2Y { get; set; }
        public string MANUFACTR_PART_COST_YTD_1Y { get; set; }
        public string PART_QUANTITY_YTD_1Y { get; set; }
        public string MANUFACTR_TOTAL_COST_YTD_1Y { get; set; }
        public string DELTA_VALUE_PART_YTD_1Y { get; set; }
        public string DELTA_CHANGE_PART_YTD_1Y { get; set; }
        public string DELTA_VALUE_ALL_PART_YTD_1Y { get; set; }
        public string CURR_MONTH_COST { get; set; }
        public string PART_QNTY_YTD { get; set; }
        public string TOTAL_COST_YTD { get; set; }
        public string FORECASTED_COST_1Y { get; set; }
        public string FORECASTED_COST_2Y { get; set; }
        public string FORECASTED_COST_3Y { get; set; }
        public string FORECASTED_COST_5Y { get; set; }
        public string FORECASTED_COST_10Y { get; set; }
        public string PART_SOBSL { get; set; }
        public string PROVISIONING_TYPE { get; set; }
        public string DATE_CREATED { get; set; }
        public string DISPO_CODE { get; set; }
        public string DISPO_NAME { get; set; }
        public string PART_TECHNIQUE { get; set; }
        public string MODEL_TYPE { get; set; }
        public string MARKETING_CODE { get; set; }
        public string DC_COMMODITY_CODE { get; set; }
        public string SUPPLIER_CODE { get; set; }
        public string SUPPLIER_NAME { get; set; }
        public string GROSS_LIST_PRICE { get; set; }
        public string TRADE_PRICE { get; set; }
        public string RETURN_VALUE { get; set; }
        public string GROSS_REVENUE { get; set; }
        public string NET_REVENUE_PER_VLP { get; set; }
        public string NET_REVENUE_PER_TP { get; set; }
        public string ACTUAL_PAID_PRICE { get; set; }
        public string NET_REVENUE_WITHOUT_RETURN_VALUE { get; set; }
        public string FACTORY_COST { get; set; }
        public string TAV_VALUE_APPROACH { get; set; }
        public string SKW_WITHOUT_SWA { get; set; }
        public string TOTAL_COST { get; set; }
        public string MATERIAL_COST { get; set; }
        public string RETURN_COST { get; set; }
        public string GUK { get; set; }
        public string PACKAGE { get; set; }
        public string PACKAGE_ABSOLUTE { get; set; }
        public string VTKP_ABSOLUTE { get; set; }
        public string VTKF_ABSOLUTE { get; set; }
        public string RESULT_FACTORY { get; set; }
        public string CORE_VALUE_SUPP { get; set; }
        public string CREDIT_VOUCHR_OLD_PART { get; set; }
        public string RESULT { get; set; }
        public string DISCOUNT_GROUP { get; set; }

    }
}
