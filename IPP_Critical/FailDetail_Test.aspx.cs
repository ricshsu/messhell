using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Drawing;
using Dundas.Charting.WebControl;


public partial class FailDetail_Test : System.Web.UI.Page
{

    int chartH = 400;
    int chartW = 1000;

    //Color[] aryColor = 
    //{
    //     Color.Blue,
    //     Color.DarkOrange,
    //     Color.Purple,
    //     Color.DarkGreen,
    //     Color.DodgerBlue,
    //     Color.Firebrick,
    //     Color.Olive,
    //     Color.Green
    //};
    Color[] aryColor = 
    {
         Color.DodgerBlue,
         Color.Olive,
         Color.DarkOrange,
         Color.Purple,
         Color.DarkGreen,
         Color.Blue,
         Color.Firebrick,
         Color.Green,
         Color.DarkSlateBlue,
         Color.DarkSlateGray,
         Color.Khaki,
         Color.Thistle
    };
    //Color[] aryColor = { Color.FromArgb(179, 29, 64), Color.FromArgb(239, 104, 38), Color.FromArgb(245, 218, 13), Color.FromArgb(134, 196, 63), Color.FromArgb(37, 170, 227), Color.FromArgb(16, 83, 164), Color.FromArgb(88, 55, 146), Color.FromArgb(216, 27, 91), Color.FromArgb(252, 181, 29), Color.FromArgb(247, 238, 47), Color.FromArgb(29, 157, 132), Color.FromArgb(24, 121, 190), Color.FromArgb(13, 84, 166), Color.FromArgb(187, 29, 106), Color.FromArgb(240, 149, 32), Color.FromArgb(204, 219, 40), Color.FromArgb(32, 127, 145), Color.FromArgb(25, 86, 166), Color.FromArgb(45, 57, 141), Color.FromArgb(150, 36, 147) };
    public struct YieldlossInfo
    {
        public string BumpingType;
        public string Part_ID;
        public int TimePeriod;
        public string TimeRange;

        public double TotalOriginal;
        public string FailMode;
        public int nTop;
        public bool xoutscrape;
        public bool lotMerge;
        public bool sf;
        public bool cr;
        public bool fai;
        public string BinCode_Id;
        public string product;
        public string product_part;

    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!this.IsPostBack)
        {
            try
            {
                if (Request["TopN"] != null)
                {
                    if (Request["TopN"].ToString() != "undefined")
                    {
                        if (Request["IsXoutScrap"] != null)
                        {
                            pageInit(((String)Request["P"]), (Request["F"].ToString()), (Request["W"].ToString()), (Request["WI"].ToString()), (Request["Product"].ToString()), (Request["Plant"].ToString()), (Request["Customer"].ToString()), (Request["TYPE"].ToString()), (Request["LotList"].ToString()), (Request["TopN"].ToString()), (Request["IsXoutScrap"].ToString()), (Request["BumpingType"].ToString()), (Request["LotMerge"].ToString()), (Request["TimePeriod"].ToString()));
                        }
                        else
                        {
                            pageInit(((String)Request["P"]), (Request["F"].ToString()), (Request["W"].ToString()), (Request["WI"].ToString()), (Request["Product"].ToString()), (Request["Plant"].ToString()), (Request["Customer"].ToString()), (Request["TYPE"].ToString()), (Request["LotList"].ToString()), (Request["TopN"].ToString()), "", (Request["BumpingType"].ToString()), (Request["LotMerge"].ToString()), (Request["TimePeriod"].ToString()));
                        }
                    }
                    else
                    {
                        if (Request["IsXoutScrap"] != null)
                        {
                            pageInit(((String)Request["P"]), (Request["F"].ToString()), (Request["W"].ToString()), (Request["WI"].ToString()), (Request["Product"].ToString()), (Request["Plant"].ToString()), (Request["Customer"].ToString()), (Request["TYPE"].ToString()), "", (Request["LotList"].ToString()), (Request["IsXoutScrap"].ToString()), (Request["BumpingType"].ToString()), (Request["LotMerge"].ToString()), (Request["TimePeriod"].ToString()));
                        }
                        else
                        {
                            pageInit(((String)Request["P"]), (Request["F"].ToString()), (Request["W"].ToString()), (Request["WI"].ToString()), (Request["Product"].ToString()), (Request["Plant"].ToString()), (Request["Customer"].ToString()), (Request["TYPE"].ToString()), "", (Request["LotList"].ToString()), "", (Request["BumpingType"].ToString()), (Request["LotMerge"].ToString()), (Request["TimePeriod"].ToString()));
                        }
                    }
                }
                //pageInit("IVB L21", "Bump fail", "201327", "201325,201326,201327", "CPU", "All", "INTEL", "PRODUCT", "'N34AHC510111','N34AHC510281'");
                //pageInit("FCS089A", "Bump fail", "201401", "201351,201352,201401", "CPU", "All", "INTEL", "PART", "undefined");
            }
            catch (Exception ex)
            {
            }
        }
    }
    public string ConvertStr2AddMark(string temp)
    {
        string sValue = "";

        if (temp == null)
        {
            return "";
        }
        string[] sSetting = temp.Split(',');
        for (int i = 0; i <= sSetting.Length - 1; i++)
        {
            if (string.IsNullOrEmpty(sValue))
            {
                sValue = "'" + sSetting[i].Trim() + "'";
            }
            else
            {
                sValue += ",'" + sSetting[i].Trim() + "'";
            }
        }
        return sValue;
    }
    public string getMainDT(YieldlossInfo yl)
    {
        string tempReplace = "";
        string tempSQL = "";


        if (yl.FailMode == "Inline異常報廢")
        {

            if (yl.lotMerge == true)
            {
                tempSQL = "select a.DefectCode, a.Fail_Mode, a.BinCode_Id, a.MF_Stage, a.BinCode_Id from dbo.VW_BinCode_Daily_Lot a where 1=1 ";
            }
            else
            {
                tempSQL = "select a.DefectCode, a.Fail_Mode, a.BinCode_Id, a.MF_Stage, a.BinCode_Id from dbo.WB_BinCode_Daily_Lot_NotMerge a where 1=1 ";
            }

            //tempSQL = "select a.DefectCode, a.Fail_Mode, a.BinCode_Id, a.MF_Stage, a.BinCode_Id from dbo.VW_BinCode_Daily_Lot a where 1=1 ";
            tempSQL += "and a.Fail_Mode='Inline異常報廢' ";
            if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
            {
                tempSQL += "and  BumpingType in (" + yl.BumpingType + ")  ";
            }
            else if (!string.IsNullOrEmpty(yl.Part_ID))
            {
                tempSQL += "and Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  ";
            }

            if (yl.fai == false)
            {
                tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 ";
            }

            if (yl.cr == false)
            {

                tempSQL += "and substring(lot_id,9,1)<>'Y' ";
                tempSQL += "and substring(lot_id,9,1)<>'Z' ";
                tempSQL += "and substring(Part_id,7,1)<>'V' ";
            }

            //tempSQL += "and a.category='WB' ";
            tempSQL += "group by a.DefectCode, a.Fail_Mode, a.BinCode_Id, a.MF_Stage, a.BinCode_Id order by a.BinCode_Id ";

        }
        else
        {

            if (yl.lotMerge == true)
            {
                tempSQL += "select top 1 * from dbo.VW_BinCode_Daily_Lot wiht (nolock) where 1=1  ";
            }
            else
            {
                tempSQL += "select top 1 * from dbo.WB_BinCode_Daily_Lot_NotMerge wiht (nolock) where 1=1 ";
            }

            // tempSQL = "select top 1 * from dbo.WB_BinCode_Daily_Lot_NotMerge wiht (nolock) where 1=1 ";


            //tempSQL += "and category='WB' ";
            //tempSQL += "and category='" + yl.product + "' ";


            if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
            {
                tempSQL += "and  BumpingType in (" + yl.BumpingType + ")  ";
            }
            else if (!string.IsNullOrEmpty(yl.Part_ID))
            {
                if (yl.product_part.ToLower() == "part")
                {
                    tempSQL += "and Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  ";
                }
                else
                {
                    tempSQL += "and production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  ";
                }

            }


            tempSQL += "and Fail_Mode='" + yl.FailMode + "'";
            if (yl.cr == false)
            {

                tempSQL += "and substring(lot_id,9,1)<>'Y' ";
                tempSQL += "and substring(lot_id,9,1)<>'Z' ";
                tempSQL += "and substring(Part_id,7,1)<>'V' ";
            }
        }



        if (yl.sf == true & yl.lotMerge == false)
        {
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF");
        }


        return tempSQL;
    }

    public string getLotSql(YieldlossInfo yl)
    {
        //yl.FailMode = "IPP items OOS sorting scrap";

        string tempReplace = "";
        string tempSQL = "";
        if (yl.product != "PPS" && yl.product != "PCB")
        {
            tempSQL = "select distinct C.Customer_Id as Customer, C.Category as 'CPU/CS',A.Lot_Id, b.Production_Type AS Product_ID, a.Part_ID, ";
            if (yl.TimePeriod == 0)
            {
                tempSQL += "SUBSTRING(CONVERT(VARCHAR, a.Datatime, 112), 1, 8) as Datatime ";
            }
            else if (yl.TimePeriod == 1)
            {
                tempSQL += "a.WW as DataTime ";
            }
            else
            {
                tempSQL += "SUBSTRING(CONVERT(VARCHAR, a.Datatime, 112), 1, 6) as Datatime ";
            }
            tempSQL += ",a.WW as WD,Convert(char(19), a.datatime, 120) as Time,DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage,  Fail_Count as QTY, ";
            tempSQL += " a.Original_Input_Qty as Input_QTY, ROUND(Fail_ratio, 3) as Ratio  from vw_BinCode_Daily_RawData A left join ";
            tempSQL += "(select * from VW_BinCode_Daily_Lot where fail_mode = '" + yl.FailMode + "' )  as B ";
            tempSQL += "on A.Lot_Id=B.Lot_Id left join dbo.Customer_Prodction_Mapping_BU_Rename as C on C.Part_Id=B.Part_Id ";
            tempSQL += "where 1 = 1  and ";
            if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
            {
                tempSQL += "c.Bumping_Type in (" + yl.BumpingType + ")  ";
            }
            else if (!string.IsNullOrEmpty(yl.Part_ID))
            {
                if (yl.product_part.ToLower() == "part")
                {
                    tempSQL += "a.Part_Id IN(" + ConvertStr2AddMark(yl.Part_ID) + ") ";
                }
                else
                {
                    tempSQL += "b.production_type IN(" + ConvertStr2AddMark(yl.Part_ID) + ") ";
                }
            }

            if (yl.TimePeriod == 0)
            {
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, a.Datatime, 112), 1, 8)='" + yl.TimeRange + "' ";
            }
            else if (yl.TimePeriod == 1)
            {
                tempSQL += "and a.WW='" + yl.TimeRange + "' ";
            }
            else
            {
                DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
                DateTime eDate = sDate.AddMonths(1);

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, a.Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, a.Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + "  ";
            }
            tempSQL += "order by Time, Lot_ID, BinCode ";
        }

        else
        {


            tempSQL = "select distinct a.Customer_Id as Customer, a.Category as 'CPU/CS', b.Production_Type AS Product_ID, b.Part_ID, ";

            if (yl.TimePeriod == 0)
            {
                tempSQL += "SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime ";
            }
            else if (yl.TimePeriod == 1)
            {
                tempSQL += "WW as DataTime ";
            }
            else
            {
                tempSQL += "SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime ";
            }

            tempSQL += ",WW as WD,Convert(char(19), datatime, 120) as Time, Lot_ID, DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage,  ";


            //tempSQL += ",Convert(char(10), datatime, 120) as WD,Convert(char(19), datatime, 120) as Time, Lot_ID, DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage,  ";
            if (yl.xoutscrape == true)
            {
                tempSQL += "Fail_Count_ByXoutScrap as QTY, Original_Input_Qty as Input_QTY, ROUND(Fail_ratio_ByXoutScrap, 3) as Ratio ";
            }
            else
            {
                tempSQL += "Fail_Count as QTY, Original_Input_Qty as Input_QTY, ROUND(Fail_ratio, 3) as Ratio ";
            }


            if (yl.lotMerge == true)
            {
                tempSQL += "from dbo.Customer_Prodction_Mapping_BU_Rename a, dbo.VW_BinCode_Daily_Lot b ";
            }
            else
            {
                tempSQL += "from dbo.Customer_Prodction_Mapping_BU_Rename a, dbo.WB_BinCode_Daily_Lot_NotMerge b ";
            }


            tempSQL += "where 1 = 1 and a.Part_Id = b.Part_Id and ";

            if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
            {
                tempSQL += "b.BumpingType in (" + yl.BumpingType + ")  ";
            }
            else if (!string.IsNullOrEmpty(yl.Part_ID))
            {
                if (yl.product_part.ToLower() == "part")
                {
                    tempSQL += "b.Part_Id IN(" + ConvertStr2AddMark(yl.Part_ID) + ") ";
                }
                else
                {
                    //tempSQL += "and production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ") and ";
                    tempSQL += "b.production_type IN(" + ConvertStr2AddMark(yl.Part_ID) + ") ";


                }


            }



            tempSQL += "and b.fail_mode = '" + yl.FailMode + "' ";

            if (yl.TimePeriod == 0)
            {
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8)='" + yl.TimeRange + "' ";
            }
            else if (yl.TimePeriod == 1)
            {
                tempSQL += "and WW='" + yl.TimeRange + "' ";
            }
            else
            {
                DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
                DateTime eDate = sDate.AddMonths(1);

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + "  ";


            }


            if (yl.fai == false)
            {
                tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 ";
            }

            if (yl.cr == false)
            {

                tempSQL += "and substring(lot_id,9,1)<>'Y' ";
                tempSQL += "and substring(lot_id,9,1)<>'Z' ";
                tempSQL += "and substring(a.Part_id,7,1)<>'V' ";
            }
            tempSQL += "and Fail_Count>0 ";

            tempSQL += "order by Time, Lot_ID, BinCode ";

            if (yl.sf == true & yl.lotMerge == false)
            {
                tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF");
            }
        }
        return tempSQL;
    }

    public string getLotSql_inline(YieldlossInfo yl, string DB)
    {

        string tempReplace = "";
        string tempSQL = "";

        if (DB == "KSEDA")
        {
            tempSQL = "select Lot_Id,Part_Id ,Station_Id ,Defect_Code ,Fail_Count,DataTime  from KSYield.dbo.MISDefect where Lot_Id in (select distinct Lot_ID ";

        }
        else
        {
            tempSQL = "select Lot_Id,Part_Id ,Station_Id ,Defect_Code ,Fail_Count,DataTime  from Yield.dbo.MISDefect where Lot_Id in (select distinct Lot_ID ";

        }


        if (yl.lotMerge == true)
        {
            tempSQL += "from dbo.Customer_Prodction_Mapping_BU_Rename a, dbo.VW_BinCode_Daily_Lot b ";
        }
        else
        {
            tempSQL += "from dbo.Customer_Prodction_Mapping_BU_Rename a, dbo.WB_BinCode_Daily_Lot_NotMerge b ";
        }


        tempSQL += "where 1 = 1 and a.Part_Id = b.Part_Id and ";

        if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
        {
            tempSQL += "b.BumpingType in (" + yl.BumpingType + ")  ";
        }
        else if (!string.IsNullOrEmpty(yl.Part_ID))
        {
            tempSQL += "b.Part_Id IN(" + ConvertStr2AddMark(yl.Part_ID) + ") ";
        }



        tempSQL += "and b.fail_mode = '" + yl.FailMode + "' ";

        if (yl.TimePeriod == 0)
        {
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8)='" + yl.TimeRange + "' ";
        }
        else if (yl.TimePeriod == 1)
        {
            tempSQL += "and WW='" + yl.TimeRange + "' ";
        }
        else
        {
            DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
            DateTime eDate = sDate.AddMonths(1);

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + "  ";


        }

        if (yl.fai == false)
        {
            tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 ";
        }

        if (yl.cr == false)
        {

            tempSQL += "and substring(lot_id,9,1)<>'Y' ";
            tempSQL += "and substring(lot_id,9,1)<>'Z' ";
            tempSQL += "and substring(a.Part_id,7,1)<>'V' ";
        }

        //tempSQL += ")and Defect_Code ='9C' ";
        tempSQL += ") order by datatime";


        if (yl.sf == true & yl.lotMerge == false)
        {
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF");
        }

        return tempSQL;
    }


    public string getBinCode(YieldlossInfo yl)
    {
        //select ww as yearWW, round((convert(float, sum(case when BinCode_id='BC267' then Fail_count end))/269568 * 100), 2) as BC267 
        //from dbo.VW_BinCode_Daily_Lot where Part_Id in('SNE135C') and WW=201417 group by ww

        string tempReplace = "";
        string tempSQL = "";


        if (yl.TimePeriod == 0)
        {
            tempSQL = "select SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime ";
        }
        else if (yl.TimePeriod == 1)
        {
            tempSQL = "select ww as Datatime ";
        }
        else
        {
            tempSQL = "select SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime ";
        }

        tempSQL += ",MF_Stage,BinCode_Id,BinCode ,DefectCode,";




        if (yl.xoutscrape == true)
        {
            tempSQL += "SUM(Fail_Count_ByXoutScrap) as FailCount ,round((convert(float, SUM(Fail_Count_ByXoutScrap))/";
        }
        else
        {
            tempSQL += "SUM(Fail_Count) as FailCount ,round((convert(float, SUM(Fail_Count))/";
        }

        tempSQL += yl.TotalOriginal.ToString() + "), 6) * 100 as Fail_Ratio from ";


        if (yl.lotMerge == true)
        {
            tempSQL += "dbo.VW_BinCode_Daily_Lot ";
        }
        else
        {
            tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge ";
        }

        tempSQL += "where 1=1  ";

        if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
        {
            tempSQL += "and BumpingType in (" + yl.BumpingType + ")   and ";
        }
        else if (!string.IsNullOrEmpty(yl.Part_ID))
        {

            if (yl.product_part.ToLower() == "part")
            {
                tempSQL += "and Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ") and ";
            }
            else
            {
                tempSQL += "and production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ") and ";

            }


        }

        if (yl.FailMode == "Inline異常報廢")
        {
            tempSQL += "MF_Stage='INLINE'  ";
        }
        else
        {
            tempSQL += "Fail_Mode='" + yl.FailMode + "' ";
        }


        if (yl.fai == false)
        {
            tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 ";
        }

        if (yl.cr == false)
        {

            tempSQL += "and substring(lot_id,9,1)<>'Y' ";
            tempSQL += "and substring(lot_id,9,1)<>'Z' ";
            tempSQL += "and substring(Part_id,7,1)<>'V' ";
        }
        tempSQL += "and Fail_Count>0 "; 

        if (yl.TimePeriod == 0)
        {
            tempSQL += " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ='" + yl.TimeRange + "' group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8)  ";
        }
        else if (yl.TimePeriod == 1)
        {
            tempSQL += " and WW='" + yl.TimeRange + "'  group by ww";
        }
        else
        {
            DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
            DateTime eDate = sDate.AddMonths(1);

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ";

        }




        tempSQL += ",MF_Stage,BinCode_Id,BinCode ,DefectCode ";

        tempSQL += "order by round((convert(float, SUM(Fail_Count))/";

        tempSQL += yl.TotalOriginal.ToString() + "), 6)  desc ";



        if (yl.sf == true & yl.lotMerge == false)
        {
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF");
        }












        //tempSQL+=", round((convert(float, sum(case when BinCode_id='" + yl.BinCode_Id + "' then ";
        //if (yl.xoutscrape == true)
        //{
        //    tempSQL += "Fail_Count_ByXoutScrap ";
        //}
        //else
        //{
        //    tempSQL += "Fail_count ";
        //}

        ////tempSQL +="end))/" + yl.TotalOriginal + " * 100), 2) as " + yl.BinCode_Id + " from ";
        //tempSQL += "end))/" + yl.TotalOriginal + " * 100), 2) as Fail_Ratio from ";

        //if (yl.lotMerge == true)
        //{ 
        //  tempSQL += "dbo.VW_BinCode_Daily_Lot " ;
        //}
        //else 
        //{
        //    tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge ";
        //}

        //if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
        //{
        //    tempSQL += "where BumpingType = '" + yl.BumpingType + "'  and ";
        //}
        //else if (!string.IsNullOrEmpty(yl.Part_ID))
        //{
        //    tempSQL += "where Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ") and ";
        //}

        //if (yl.TimePeriod == 0)
        //{
        //    tempSQL += "SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ='" + yl.TimeRange + "' group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8)  ";
        //}
        //else if (yl.TimePeriod == 1)
        //{
        //    tempSQL += "WW='" + yl.TimeRange + "' group by ww ";
        //}
        //else
        //{
        //    DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
        //    DateTime eDate = sDate.AddMonths(1);

        //    tempSQL += "SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ";

        //}


        return tempSQL;
    }

    public string getTotalOriginal_SQL(YieldlossInfo yl)
    {

        string tempReplace = "";
        string tempSQL = "";
        tempSQL = "select SUM(Original_Input_QTY) from " + "( " + "select distinct Lot_Id ,Original_Input_QTY  " + "from ";

        if (yl.lotMerge == true)
        {
            tempSQL += "dbo.VW_BinCode_Daily_Lot ";
        }
        else
        {
            tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge ";
        }

        tempSQL += "where 1=1  ";

        if (!string.IsNullOrEmpty(yl.BumpingType) & string.IsNullOrEmpty(yl.Part_ID))
        {
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")   ";
        }
        else if (!string.IsNullOrEmpty(yl.Part_ID))
        {
            if (yl.product_part.ToLower() == "part")
            {
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  ";
            }
            else
            {
                tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  ";
            }

        }

        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " + "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " + "and Fail_Mode NOT LIKE '8K%'  ";

        if (yl.TimePeriod == 0)
        {
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " ";
        }
        else if (yl.TimePeriod == 1)
        {
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") ";

        }
        else
        {
            DateTime sDate = System.DateTime.Parse(yl.TimeRange.Substring(0, 4) + "/" + yl.TimeRange.Substring(yl.TimeRange.Length - 2, 2) + "/01");
            DateTime eDate = sDate.AddMonths(1);

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " ";
        }

        if (yl.fai == false)
        {
            tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 ";
        }

        if (yl.cr == false)
        {

            tempSQL += "and substring(lot_id,9,1)<>'Y' ";
            tempSQL += "and substring(lot_id,9,1)<>'Z' ";
            tempSQL += "and substring(Part_id,7,1)<>'V' ";
        }

        tempSQL += ")a ";


        if (yl.sf == true & yl.lotMerge == false)
        {
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF");
        }
        return tempSQL;
    }

    private void pageInit(string part_id, string item, string week, string weekIn, string product, string plant, string customer_id, string product_part, string lot_list, string TopN, string IsXoutScrap, string BumpingType, string LotMerge, string timeperiod)
    {
        YieldlossInfo yl = new YieldlossInfo();
        yl.Part_ID = part_id;
        item = item.Replace("||", "''");
        item = item.Replace("000D000", "''D''");
        yl.FailMode = item;
        Label2.Text = yl.Part_ID + " : " + yl.FailMode;

        yl.TimeRange = week;
        yl.TimePeriod = 1;
        yl.product = product;
        yl.product_part = product_part;

        if (timeperiod == "0")
        {
            yl.TimePeriod = 0;
        }
        else if (timeperiod == "1")
        {
            yl.TimePeriod = 1;
        }
        else
        {
            yl.TimePeriod = 2;
        }


        yl.BumpingType = BumpingType;

        string sValue = "";
        string[] sSetting = BumpingType.Split(',');
        for (int i = 0; (i <= (sSetting.Length - 1)); i++)
        {
            if ((sValue == ""))
            {
                sValue = ("\'"
                            + (sSetting[i].Trim() + "\'"));
            }
            else
            {
                sValue = (sValue + (",\'"
                            + (sSetting[i].Trim() + "\'")));
            }
        }
        yl.BumpingType = sValue;


        yl.lotMerge = true;

        yl.xoutscrape = false;

        if (IsXoutScrap == "True")
        {
            yl.xoutscrape = true;
        }

        if (LotMerge.IndexOf("False") >= 0)
        {
            yl.lotMerge = false;
        }

        yl.cr = true;
        if (LotMerge.IndexOf("CR") > 0)
        {
            yl.cr = false;
        }

        if (LotMerge.IndexOf("SF") > 0)
        {
            yl.lotMerge = false;
            yl.sf = true;
        }

        yl.fai = false;
        if (LotMerge.IndexOf("FAI") > 0)
        {
            yl.fai = true ;
            
        }


        if (yl.product != "PPS" && yl.product != "PCB")
        {
            yl.lotMerge = true;
        }


        //if (LotMerge == "False")
        //{
        //    yl.lotMerge = false;
        //}

        //if (LotMerge == "False_SF")
        //{
        //    yl.lotMerge = false;
        //    yl.sf = true;
        //}

        if (lot_list != "undefined" && lot_list != "")
        {
            String[] strAry = lot_list.Split(new Char[] { ',' });
            lot_list = "";
            for (int i = 0; i < strAry.Length; i++)
            {
                lot_list += "'" + strAry[i] + "',";
            }
            lot_list = lot_list.Substring(0, (lot_list.Length - 1));
        }
        else
        {
            lot_list = "";
        }

        if (BumpingType != "undefined" && BumpingType != "")
        {
            String[] strAry = BumpingType.Split(new Char[] { ',' });
            BumpingType = "";
            for (int i = 0; i < strAry.Length; i++)
            {
                BumpingType += "'" + strAry[i] + "',";
            }
            BumpingType = BumpingType.Substring(0, (BumpingType.Length - 1));
        }
        else
        {
            BumpingType = "";
        }

        if (part_id != "undefined" && part_id != "")
        {
            String[] strAry = part_id.Split(new Char[] { ',' });
            part_id = "";
            for (int i = 0; i < strAry.Length; i++)
            {
                part_id += "'" + strAry[i] + "',";
            }
            part_id = part_id.Substring(0, (part_id.Length - 1));
        }
        else
        {
            part_id = "";
        }

        item = item.Replace("||", "''"); // 因為有 ' 字元的問題, 所以需要跳脫, 在前一頁已經用 00 代替 ' ,不然 javascript 傳不過來
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        string sqlStr = "";
        string conditionStr = "";
        string conditionStr1 = "";
        string conditionStr2a = "";
        string conditionStr2b = "";
        string conditionStr2c = "";
        string conditionStr3a = "";
        string conditionStr3b = "";
        string conditionStr3 = "";
        string customStr = " ";
        string partStr = " ";
        string weekStr = " ";
        string itemStr = " ";
        string topStr = " ";
        string tableName = "";
        string plantStr = "";

        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        DataTable MainDT_Plant = null;
        DataTable MainDT_Total_Plant = null;
        DataTable BinCodeDT = null;
        DataTable workTable = null;
        DataTable LotDT = null;
        DataTable LotDT_Plant = null;

        try
        {
            week = week.Replace("W", "");


            // 建立 DataTable 
            workTable = new DataTable();
            workTable.Columns.Add("DefectCode", Type.GetType("System.String"));
            workTable.Columns.Add("FailMode", Type.GetType("System.String"));
            workTable.Columns.Add("BinCode", Type.GetType("System.String"));
            workTable.Columns.Add("Category", Type.GetType("System.String"));

            conn.Open();
            sqlStr = getMainDT(yl);
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            yl.BinCode_Id = "";

            for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            {

                if (i == 0)
                {
                    sqlStr = getTotalOriginal_SQL(yl);
                    myAdapter = new SqlDataAdapter(sqlStr, conn);
                    DataTable TotalDT = new DataTable();
                    myAdapter.SelectCommand.CommandTimeout = 3600;
                    myAdapter.Fill(TotalDT);

                    conditionStr3b = TotalDT.Rows[0][0].ToString();

                    yl.TotalOriginal = Convert.ToDouble(conditionStr3b);
                }

            }

            sqlStr = getBinCode(yl);
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            BinCodeDT = new DataTable();
            myAdapter.SelectCommand.CommandTimeout = 3600;
            myAdapter.Fill(BinCodeDT);

            // === Thrend Chart ===
            string strLotList = "";
            string strDefectCode = "";
            try
            {
                sqlStr = getLotSql(yl);
                LotDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(LotDT);

                for (int i = 0; i <= (LotDT.Rows.Count - 1); i++)
                {
                    if (i == 0)
                    {
                        strLotList = LotDT.Rows[i]["Lot_Id"].ToString(); 
                    }
                    else 
                    {
                        strLotList +=","+ LotDT.Rows[i]["Lot_Id"].ToString(); 
                    }
                }

                if (MainDT.Rows .Count >0)
                {
                strLotList = ConvertStr2AddMark(strLotList);
                strDefectCode = MainDT.Rows[0]["DefectCode"].ToString() ;
                sqlStr = "select Plant_Id,Defect_code,sum(fail_count*Unit_Length*Unit_Width/ 144 ) as SF,sum(fail_count) as Count   from Yield.dbo. MISDefect a, MES.dbo.ProductInfo b ";
                sqlStr += " where  a.Part_Id=b.Part_No and Lot_Id in (" + strLotList + ") and Defect_Code='" + strDefectCode + "'";
                sqlStr += " group by Defect_code,plant_id";

                MainDT_Plant = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(MainDT_Plant);

                sqlStr = "select Plant_Id,lot_id,Defect_code,sum(fail_count*Unit_Length*Unit_Width/ 144 ) as SF,sum(fail_count) as Count   from Yield.dbo. MISDefect a, MES.dbo.ProductInfo b ";
                sqlStr += " where   a.Part_Id=b.Part_No and Lot_Id in (" + strLotList + ") and Defect_Code='" + strDefectCode + "'";
                sqlStr += " group by Defect_code,plant_id,lot_id";

                LotDT_Plant = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(LotDT_Plant);

                sqlStr = "select Plant_Id,sum(fail_count*Unit_Length*Unit_Width/ 144 ) as SF,sum(fail_count) as Count   from Yield.dbo. MISDefect a, MES.dbo.ProductInfo b ";
                sqlStr += " where  a.Part_Id=b.Part_No and Lot_Id in (" + strLotList + ") ";
                sqlStr += " group by plant_id";

                MainDT_Total_Plant = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(MainDT_Total_Plant);

                }
                
            }
            catch (Exception ex) {
                string strmessage = "";
                strmessage = ex.Message.ToString(); 

            }

            Bump_Detail(ref conn, part_id, item, week, weekIn, product_part, IsXoutScrap, BumpingType);

            if (TopN == "")
            {
                TopN = "1";
            }

            try
            {
                area_Pie(ref MainDT, ref BinCodeDT, ref workTable, week, item, Convert.ToInt32(TopN));
                area_Pie2(ref MainDT_Plant, ref LotDT_Plant, ref workTable, week, item, Convert.ToInt32(TopN));
                area_Pie3(ref MainDT_Total_Plant, ref LotDT_Plant, ref workTable, week, item, Convert.ToInt32(TopN));
            }
            catch (Exception ex) { }

            // weekIn
            area_Thred(ref LotDT, item, week, item, product, plant, yl);

            // --- 加入 RowData ---

            try
            {
                sqlStr = getLotSql(yl);
                LotDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(LotDT);
                if (LotDT.Rows.Count > 0)
                {
                    GV_LotRowData.DataSource = LotDT;
                    GV_LotRowData.DataBind();
                    lab_lotRowData.Text = (item + " RowData");

                    //Fail Detail(Pareto)
                    GV_NewLotRowData.DataSource = LotDT;
                    GV_NewLotRowData.DataBind();
                    lab_NewLotRowData.Text = (item + " RowData");
                }
            }
            catch (Exception ex) { }


            NewPareto_Chart(ref BinCodeDT, Convert.ToInt32(TopN));
            sqlStr = getLotSql_inline(yl, conn.Database);
            LotDT = new DataTable();
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myAdapter.Fill(LotDT);
            inline_rawdata(LotDT, yl);

            conn.Close();
        }
        catch (Exception ex)
        {
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

    }

    private void inline_rawdata(DataTable LotDT, YieldlossInfo yl)
    {
        if (yl.FailMode == "Inline異常報廢")
        {
            TabStrip3.Items[4].Hidden = false;

            GridView1.DataSource = LotDT;
            GridView1.DataBind();
            string lotid = "";





        }
        else
        {
            TabStrip3.Items[4].Hidden = true;
        }


    }


    private void area_Thred(ref DataTable LotDT, string FailMode, string cweek, string item, string product, string plant, YieldlossInfo yl)
    {
        if ((product == "CPU") | (plant.ToUpper() == "ALL"))
        {
            string sTimePeriod = "";
            if (yl.TimePeriod == 0)
            {
                sTimePeriod = "Day";
            }
            else if (yl.TimePeriod == 1)
            {
                sTimePeriod = "Week";
            }
            else
            {
                sTimePeriod = "Month";
            }
            titlePanel.Controls.Add(new LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:middle;font-weight:bold;width:750px;Height:18px'>" + sTimePeriod + " : " + cweek + " [" + item + "]</td></tr>"));
        }
        else
        {
            string sTimePeriod = "";
            if (yl.TimePeriod == 0)
            {
                sTimePeriod = "Day";
            }
            else if (yl.TimePeriod == 1)
            {
                sTimePeriod = "Week";
            }
            else
            {
                sTimePeriod = "Month";
            }
            titlePanel.Controls.Add(new LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:middle;font-weight:bold;width:750px;Height:18px'>" + sTimePeriod + " : " + cweek + " [" + item + "]   Plant : " + plant + "</td></tr>"));
        }

        Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
        Chart.ImageUrl = "temp/yieldT_#SEQ(1000,1)";
        Chart.ImageType = ChartImageType.Png;
        Chart.Palette = ChartColorPalette.Dundas;
        Chart.Height = chartH;
        Chart.Width = chartW;

        Chart.Palette = ChartColorPalette.Dundas;
        Chart.BackColor = Color.White;
        Chart.BackGradientEndColor = Color.Peru;
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
        Chart.BorderStyle = ChartDashStyle.Solid;
        Chart.BorderWidth = 3;
        Chart.BorderColor = Color.DarkBlue;

        Chart.ChartAreas.Add("Default");
        Chart.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
        Chart.ChartAreas["Default"].AxisX.Title = "【" + FailMode + "】";
        Chart.ChartAreas["Default"].AxisX.LabelStyle.Interval = 1;
        Chart.ChartAreas["Default"].AxisX.LabelStyle.FontAngle = -45;
        //文字對齊
        Chart.ChartAreas["Default"].BorderStyle = ChartDashStyle.NotSet;
        Chart.UI.Toolbar.Enabled = false;
        Chart.UI.ContextMenu.Enabled = true;

        Series series = default(Series);
        series = Chart.Series.Add(FailMode);
        series.ChartArea = "Default";
        series.Type = SeriesChartType.Line;
        series.Color = Color.Blue;
        series.MarkerStyle = MarkerStyle.Circle;
        series.MarkerSize = 8;
        series.MarkerColor = Color.DarkBlue;
        series.BorderColor = Color.White;
        series.BorderWidth = 1;
        series.ShowInLegend = false;

        string lot_id = null;
        string wdStr = null;
        string trtmStr = null;
        double value = 0;

        for (int rowIndex = 0; rowIndex <= (LotDT.Rows.Count - 1); rowIndex++)
        {
            //if ((LotDT.Rows[rowIndex]["Total"]) != null)
            //{
            //    lot_id = LotDT.Rows[rowIndex]["Lot_id"].ToString();
            //    wdStr = LotDT.Rows[rowIndex]["WD"].ToString();
            //    trtmStr = LotDT.Rows[rowIndex]["TRTM"].ToString();
            //    value = Convert.ToDouble(LotDT.Rows[rowIndex]["Total"]);
            //    Chart.Series[FailMode].Points.AddXY(lot_id, value);
            //    Chart.Series[FailMode].Points[rowIndex].ToolTip = "[W" + wdStr + "_" + lot_id + "] " + value.ToString() + "%";
            //    Chart.Series[FailMode].Points[rowIndex].Href = "javascript:LinkPoint('" + (lot_id + trtmStr) + "');";
            //}

            if ((LotDT.Rows[rowIndex]["Ratio"]) != null && (LotDT.Rows[rowIndex]["Ratio"]).ToString() !="")
            {
                lot_id = LotDT.Rows[rowIndex]["Lot_id"].ToString();
                wdStr = LotDT.Rows[rowIndex]["WD"].ToString();
                trtmStr = LotDT.Rows[rowIndex]["Time"].ToString();
                value = Convert.ToDouble(LotDT.Rows[rowIndex]["Ratio"]);
                Chart.Series[FailMode].Points.AddXY(lot_id, value);
                Chart.Series[FailMode].Points[rowIndex].ToolTip = "[W" + wdStr + "_" + lot_id + "] " + value.ToString() + "%";
                Chart.Series[FailMode].Points[rowIndex].Href = "javascript:LinkPoint('" + (lot_id + trtmStr) + "');";
            }
            else 
            {
                lot_id = LotDT.Rows[rowIndex]["Lot_id"].ToString();
                wdStr = LotDT.Rows[rowIndex]["WD"].ToString();
                trtmStr = LotDT.Rows[rowIndex]["Time"].ToString();
                value = 0;
                Chart.Series[FailMode].Points.AddXY(lot_id, value);
                Chart.Series[FailMode].Points[rowIndex].ToolTip = "[W" + wdStr + "_" + lot_id + "] " + value.ToString() + "%";
            }
        }

        ThendPanel.Controls.Add(new LiteralControl("<tr><td>"));
        ThendPanel.Controls.Add(Chart);
        ThendPanel.Controls.Add(new LiteralControl("</td></tr>"));
    }

    private void area_Pie(ref DataTable MainDT, ref DataTable BinCodeDT, ref DataTable workDT, string cweek, string item, int TopN)
    {
        //cweek = cweek.Substring(4, 2);
        double bvalue = 0;
        double fvalue = 0;
        string codeID = "";
        int rowIndex = 0;
        DataRow workDR = null;
        // DefectCode, FailMode, BinCode, Category, W ~, Delta

        for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
        {
            workDR = workDT.NewRow();
            workDR[0] = MainDT.Rows[i]["DefectCode"].ToString();
            workDR[1] = MainDT.Rows[i]["Fail_Mode"].ToString();
            workDR[2] = MainDT.Rows[i]["BinCode_Id"].ToString();
            workDR[3] = MainDT.Rows[i]["MF_Stage"].ToString();
            codeID = MainDT.Rows[i]["BinCode_Id"].ToString();
            rowIndex = 4;

            //for (int j = 0; j <= (BinCodeDT.Rows.Count - 1); j++)
            //{
            //    workDR[rowIndex] = BinCodeDT.Rows[j][codeID].ToString();
            //    rowIndex += 1;

            //    if (j == (BinCodeDT.Rows.Count - 2))
            //    {
            //        bvalue = Convert.ToDouble(BinCodeDT.Rows[j][codeID]);
            //    }

            //    if (j == (BinCodeDT.Rows.Count - 1))
            //    {
            //        fvalue = Convert.ToDouble(BinCodeDT.Rows[j][codeID]);
            //    }

        }
        //    for (int j = 0; j <= (BinCodeDT.Rows.Count - 1); j++)
        //    {
        //        workDR[rowIndex] = BinCodeDT.Rows[j]["Fail_Ratio"].ToString();
        //        rowIndex += 1;

        //        if (j == (BinCodeDT.Rows.Count - 2))
        //        {
        //            bvalue = Convert.ToDouble(BinCodeDT.Rows[j]["Fail_Ratio"]);
        //        }

        //        if (j == (BinCodeDT.Rows.Count - 1))
        //        {
        //            fvalue = Convert.ToDouble(BinCodeDT.Rows[j]["Fail_Ratio"]);
        //        }

        //    }

        //    if (bvalue > 0 | fvalue > 0)
        //    {
        //        workDR[rowIndex] = (Math.Round((bvalue - fvalue), 2)).ToString();
        //    }
        //    else
        //    {
        //        workDR[rowIndex] = "0";
        //    }
        //    workDT.Rows.Add(workDR);

        //}

        //workDT.DefaultView.Sort = workDT.Columns[workDT.Columns.Count - 2].Caption + " desc";
        ////DataTable weekGroupDT = workDT.DefaultView.ToTable();

        ////gv_pie.DataSource = workDT;
        //gv_pie.DataSource = workDT.DefaultView.ToTable();
        gv_pie.DataSource = BinCodeDT.DefaultView.ToTable();
        gv_pie.DataBind();
        UtilObj.Set_DataGridRow_OnMouseOver_Color(ref gv_pie, "#FFF68F", gv_pie.AlternatingRowStyle.BackColor);

        //Fail Detail(Pareto)
        //GV_NewLotRowData.DataSource = workDT.DefaultView.ToTable();
        //GV_NewLotRowData.DataBind();
        //lab_NewLotRowData.Text = (item + " RowData");

        // 畫 Pie Chart
        if (MainDT.Rows.Count > 0)
        {
            Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
            ChartArea chartArea1 = new ChartArea();

            Chart.Palette = ChartColorPalette.Dundas;
            Chart.BackColor = Color.White;
            Chart.BackGradientEndColor = Color.Peru;
            Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            Chart.BorderStyle = ChartDashStyle.Solid;
            Chart.BorderWidth = 3;
            Chart.BorderColor = Color.DarkBlue;

            Chart.ImageUrl = "temp/yieldP_#SEQ(1000,1)";
            Chart.ImageType = ChartImageType.Png;
            Chart.Palette = ChartColorPalette.Dundas;
            Chart.ChartAreas.Add(chartArea1);
            Chart.Height = chartH;
            Chart.Width = chartW;

            Series series1 = default(Series);
            series1 = Chart.Series.Add("MQCS");
            series1.BackGradientEndColor = Color.White;
            series1.Type = SeriesChartType.Pie;
            series1.ShowInLegend = true;
            series1.Font = new Font("Verdana", 10);
            series1.FontColor = Color.Red;
            series1.YValueType = ChartValueTypes.Double;
            series1.XValueType = ChartValueTypes.String;

            series1["PieLabelStyle"] = "Outside";
            series1.BorderWidth = 2;
            series1.BorderColor = System.Drawing.Color.FromArgb(26, 59, 105);

            Chart.Legends.Add("Legend1");
            Chart.Legends[0].Enabled = true;
            //Chart.Legends[0].Docking = Docking.Bottom;
            //Chart.Legends(0).Alignment = System.Drawing.StringAlignment.Center
            series1.LegendText = "#VALX [#PERCENT]";

            DataRow[] foundRows = null;
            foundRows = BinCodeDT.Select("Datatime='" + cweek + "'");
            double value = 0;
            string binCodeStr = "";
            string AlisStr = "";

            //if (foundRows.Length > 0)
            //{
            //    for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            //    {
            //        binCodeStr = (String)MainDT.Rows[i]["BinCode_Id"];
            //        AlisStr = (String)MainDT.Rows[i]["BinCode_Id"];
            //        value = Convert.ToDouble(foundRows[0][binCodeStr]);
            //        value = Math.Round(value, 2);
            //        series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value);
            //        series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + "%";
            //    }

            //}

            for (int i = 0; i <= (BinCodeDT.Rows.Count - 1); i++)
            {
                binCodeStr = (String)BinCodeDT.Rows[i]["BinCode_Id"];
                AlisStr = (String)BinCodeDT.Rows[i]["BinCode_Id"];
                value = Convert.ToDouble(BinCodeDT.Rows[i]["Fail_Ratio"]);
                value = Math.Round(value, 3);
                series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value);
                series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + "%";
            }


            PiePanel.Controls.Add(new LiteralControl("<tr><td>"));
            PiePanel.Controls.Add(Chart);
            PiePanel.Controls.Add(new LiteralControl("</td></tr>"));

        }

    }

    private void area_Pie2(ref DataTable MainDT, ref DataTable BinCodeDT, ref DataTable workDT, string cweek, string item, int TopN)
    {
        
        
        
        
        
        
        //cweek = cweek.Substring(4, 2);
        double bvalue = 0;
        double fvalue = 0;
        string codeID = "";
        int rowIndex = 0;
        DataRow workDR = null;
        // DefectCode, FailMode, BinCode, Category, W ~, Delta

        //for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
        //{
        //    workDR = workDT.NewRow();
        //    workDR[0] = MainDT.Rows[i]["DefectCode"].ToString();
        //    workDR[1] = MainDT.Rows[i]["Fail_Mode"].ToString();
        //    workDR[2] = MainDT.Rows[i]["BinCode_Id"].ToString();
        //    workDR[3] = MainDT.Rows[i]["MF_Stage"].ToString();
        //    codeID = MainDT.Rows[i]["BinCode_Id"].ToString();
        //    rowIndex = 4;
        //}

        GridView2.DataSource = BinCodeDT.DefaultView.ToTable();
        GridView2.DataBind();
        UtilObj.Set_DataGridRow_OnMouseOver_Color(ref GridView2, "#FFF68F", gv_pie.AlternatingRowStyle.BackColor);

       
        // 畫 Pie Chart
        if (MainDT.Rows.Count > 0)
        {
            Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
            ChartArea chartArea1 = new ChartArea();

            Chart.Palette = ChartColorPalette.Dundas;
            Chart.BackColor = Color.White;
            Chart.BackGradientEndColor = Color.Peru;
            Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            Chart.BorderStyle = ChartDashStyle.Solid;
            Chart.BorderWidth = 3;
            Chart.BorderColor = Color.DarkBlue;

            Chart.ImageUrl = "temp/yieldP_#SEQ(1000,1)";
            Chart.ImageType = ChartImageType.Png;
            Chart.Palette = ChartColorPalette.Dundas;
            Chart.ChartAreas.Add(chartArea1);
            Chart.Height = chartH;
            Chart.Width = chartW;

            Series series1 = default(Series);
            series1 = Chart.Series.Add("MQCS");
            series1.BackGradientEndColor = Color.White;
            series1.Type = SeriesChartType.Pie;
            series1.ShowInLegend = true;
            series1.Font = new Font("Verdana", 10);
            series1.FontColor = Color.Red;
            series1.YValueType = ChartValueTypes.Double;
            series1.XValueType = ChartValueTypes.String;

            series1["PieLabelStyle"] = "Outside";
            series1.BorderWidth = 2;
            series1.BorderColor = System.Drawing.Color.FromArgb(26, 59, 105);

            Chart.Legends.Add("Legend1");
            Chart.Legends[0].Enabled = true;
            //Chart.Legends[0].Docking = Docking.Bottom;
            //Chart.Legends(0).Alignment = System.Drawing.StringAlignment.Center
            series1.LegendText = "#VALX [#PERCENT]";

            DataRow[] foundRows = null;
            //foundRows = BinCodeDT.Select("Datatime='" + cweek + "'");
            double value = 0;
            string binCodeStr = "";
            string AlisStr = "";

            for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            {
                binCodeStr = (String)MainDT.Rows[i]["Defect_code"];
                AlisStr = (String)MainDT.Rows[i]["Plant_Id"];
                value = Convert.ToDouble(MainDT.Rows[i]["SF"]);
                value = Math.Round(value, 3);
                series1.Points.AddXY(AlisStr + " " + (value.ToString()) + " SF", value);
                series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + " SF";
            }
           
            //for (int i = 0; i <= (BinCodeDT.Rows.Count - 1); i++)
            //{
            //    binCodeStr = (String)BinCodeDT.Rows[i]["BinCode_Id"];
            //    AlisStr = (String)BinCodeDT.Rows[i]["BinCode_Id"];
            //    value = Convert.ToDouble(BinCodeDT.Rows[i]["Fail_Ratio"]);
            //    value = Math.Round(value, 3);
            //    series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value);
            //    series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + "%";
            //}


            PiePanel2.Controls.Add(new LiteralControl("<tr><td>"));
            PiePanel2.Controls.Add(Chart);
            PiePanel2.Controls.Add(new LiteralControl("</td></tr>"));

        }

    }

    private void area_Pie3(ref DataTable MainDT, ref DataTable BinCodeDT, ref DataTable workDT, string cweek, string item, int TopN)
    {

        double bvalue = 0;
        double fvalue = 0;
        string codeID = "";
        int rowIndex = 0;
        DataRow workDR = null;
       

        //GridView2.DataSource = BinCodeDT.DefaultView.ToTable();
        //GridView2.DataBind();
        //UtilObj.Set_DataGridRow_OnMouseOver_Color(ref GridView2, "#FFF68F", gv_pie.AlternatingRowStyle.BackColor);


        // 畫 Pie Chart
        if (MainDT.Rows.Count > 0)
        {
            Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
            ChartArea chartArea1 = new ChartArea();

            Chart.Palette = ChartColorPalette.Dundas;
            Chart.BackColor = Color.White;
            Chart.BackGradientEndColor = Color.Peru;
            Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            Chart.BorderStyle = ChartDashStyle.Solid;
            Chart.BorderWidth = 3;
            Chart.BorderColor = Color.DarkBlue;

            Chart.ImageUrl = "temp/yieldP_#SEQ(1000,1)";
            Chart.ImageType = ChartImageType.Png;
            Chart.Palette = ChartColorPalette.Dundas;
            Chart.ChartAreas.Add(chartArea1);
            Chart.Height = chartH;
            Chart.Width = chartW;

            Series series1 = default(Series);
            series1 = Chart.Series.Add("MQCS");
            series1.BackGradientEndColor = Color.White;
            series1.Type = SeriesChartType.Pie;
            series1.ShowInLegend = true;
            series1.Font = new Font("Verdana", 10);
            series1.FontColor = Color.Red;
            series1.YValueType = ChartValueTypes.Double;
            series1.XValueType = ChartValueTypes.String;

            series1["PieLabelStyle"] = "Outside";
            series1.BorderWidth = 2;
            series1.BorderColor = System.Drawing.Color.FromArgb(26, 59, 105);

            Chart.Legends.Add("Legend1");
            Chart.Legends[0].Enabled = true;
            //Chart.Legends[0].Docking = Docking.Bottom;
            //Chart.Legends(0).Alignment = System.Drawing.StringAlignment.Center
            series1.LegendText = "#VALX [#PERCENT]";

            DataRow[] foundRows = null;
            //foundRows = BinCodeDT.Select("Datatime='" + cweek + "'");
            double value = 0;
            string binCodeStr = "";
            string AlisStr = "";

            for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            {
               // binCodeStr = (String)MainDT.Rows[i]["Defect_code"];
                AlisStr = (String)MainDT.Rows[i]["Plant_Id"];
                value = Convert.ToDouble(MainDT.Rows[i]["SF"]);
                value = Math.Round(value, 3);
                series1.Points.AddXY(AlisStr + " " + (value.ToString()) + " SF", value);
                series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + " SF";
            }

            

            PiePanel3.Controls.Add(new LiteralControl("<tr><td>"));
            PiePanel3.Controls.Add(Chart);
            PiePanel3.Controls.Add(new LiteralControl("</td></tr>"));

        }

    }

    private void Bump_Detail(ref SqlConnection conn, string part_id, string item, string week, string weekIn, string product_part, string IsXoutScrap, string BumpingType)
    {
        TabStrip3.Items[2].Hidden = true;
        DataTable ItemDT = null;
        DataTable yieldDT = null;
        DataTable LotDT = null;
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        string sqlStr = null;
        // IPQC
        string failType = "Bump";

        //If (item.ToUpper).IndexOf("BUMP") >= 0 Then
        if (item.ToUpper().IndexOf("BUMP") >= 0)
        {
            failType = "Bump";
            lab_DetailTitle.Text = "Bump Failure (AOI) Detail Info By Week " + week;
            try
            {
                // 取得最新一週的 Yield 順序的 Items
                if (IsXoutScrap == "True") //匹配報廢回歸(XoutScrap)
                {
                    //Fail_Count_byXoutScrap
                    sqlStr = "select Fail_Mode, ROUND(convert(float,SUM(Fail_Count_byXoutScrap))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE " +
                             "From dbo.VW_BinCode_Detail_Daily_Lot " +
                             "Where 1=1 " +
                             "And category = '" + failType + "' ";
                }
                else
                {
                    sqlStr = "select Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE " +
                             "From dbo.VW_BinCode_Detail_Daily_Lot " +
                             "Where 1=1 " +
                             "And category = '" + failType + "' ";
                }
                if (product_part == "PART")
                {
                    //sqlStr += "And Part_ID='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and Part_ID IN(" + part_id + ") ";
                    }
                }
                else
                {
                    //sqlStr += "And production_type='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and production_type IN(" + part_id + ") ";
                    }
                }
                sqlStr += "And WW=" + week + " " +
                "And Fail_Mode <> '報廢量' " +
                "Group by Fail_Mode " +
                "Order by YIELD_VALUE DESC ";
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                ItemDT = new DataTable();
                myAdapter.Fill(ItemDT);

                if (ItemDT.Rows.Count == 0)
                {
                    return;
                }

                // 取得 Pareto Chart Info
                if (IsXoutScrap == "True") //匹配報廢回歸(XoutScrap)
                {
                    //Fail_Count_byXoutScrap
                    sqlStr = "Select b.yearWW, Fail_Mode, ROUND(convert(float,SUM(Fail_Count_byXoutScrap))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE ";
                }
                else
                {
                    sqlStr = "Select b.yearWW, Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE ";
                }
                sqlStr += "From VW_BinCode_Detail_Daily_Lot a, SystemDateMapping b ";
                sqlStr += "Where 1=1 And a.WW = b.yearWW And a.category='" + failType + "' ";

                if (product_part == "PART")
                {
                    //sqlStr += "And a.Part_ID='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and a.Part_ID IN(" + part_id + ") ";
                    }
                }
                else
                {
                    //sqlStr += "And a.production_type='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and a.production_type IN(" + part_id + ") ";
                    }
                }

                sqlStr += "And a.Fail_Mode <> '報廢量' ";
                sqlStr += "And b.yearWW IN (" + weekIn + ") ";
                sqlStr += "Group by b.yearWW, a.Fail_Mode ";
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                yieldDT = new DataTable();
                myAdapter.Fill(yieldDT);

                // 取得 Lot 的 RowData 
                string[] weekAry = weekIn.Split(new Char[] { ',' });
                sqlStr = "SELECT A.Fail_Mode, ";
                int i = 0;
                for (i = 0; i <= (weekAry.Length - 1); i++)
                {
                    if (i != (weekAry.Length - 1))
                    {
                        sqlStr += "MAX(CASE WHEN A.yearWW=" + weekAry[i] + " THEN A.VALUE END) AS '" + weekAry[i] + "', ";
                    }
                    else
                    {
                        sqlStr += "MAX(CASE WHEN A.yearWW=" + weekAry[i] + " THEN A.VALUE END) AS '" + weekAry[i] + "' ";
                    }
                }
                sqlStr += "FROM ";
                sqlStr += "( ";
                if (IsXoutScrap == "True") //匹配報廢回歸(XoutScrap)
                {
                    //Fail_Count_byXoutScrap
                    sqlStr += "SELECT a.yearWW, b.Fail_Mode, ROUND(convert(float,SUM(b.Fail_Count_byXoutScrap))/SUM(b.Original_Input_QTY) * 100, 3) AS VALUE ";
                }
                else
                {
                    sqlStr += "SELECT a.yearWW, b.Fail_Mode, ROUND(convert(float,SUM(b.Fail_Count))/SUM(b.Original_Input_QTY) * 100, 3) AS VALUE ";
                }
                sqlStr += "From SystemDateMapping a, dbo.VW_BinCode_Detail_Daily_Lot b ";
                sqlStr += "Where 1=1 ";
                sqlStr += "And a.yearWW = b.WW ";
                sqlStr += "And category='" + failType + "' ";
                if (product_part == "PART")
                {
                    //sqlStr += "And Part_ID='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and Part_ID IN(" + part_id + ") ";
                    }
                }
                else
                {
                    //sqlStr += "And production_type='" + part_id + "' ";
                    if (part_id != "")
                    {
                        sqlStr += "and production_type IN(" + part_id + ") ";
                    }
                }
                sqlStr += "And b.Fail_Mode <> '報廢量' ";
                sqlStr += "And a.yearWW IN (" + weekIn + ") ";
                sqlStr += "GROUP BY a.yearWW, Fail_Mode";
                sqlStr += ") A ";
                sqlStr += "GROUP BY A.Fail_Mode ";
                sqlStr += "ORDER BY '" + weekAry[i - 1] + "' DESC";

                myAdapter = new SqlDataAdapter(sqlStr, conn);
                LotDT = new DataTable();
                myAdapter.Fill(LotDT);
                gr_lotview.DataSource = LotDT;
                gr_lotview.DataBind();
                UtilObj.Set_DataGridRow_OnMouseOver_Color(ref gr_lotview, "#FFF68F", gr_lotview.AlternatingRowStyle.BackColor);

                if (yieldDT.Rows.Count > 0 & ItemDT.Rows.Count > 0)
                {
                    Bump_Chart(ref yieldDT, ref ItemDT);
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            TabStrip3.Items[2].Hidden = false;
        }
    }

    private void Bump_Chart(ref DataTable DtSet, ref DataTable setupDT)
    {
        Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
        Chart.ImageUrl = "temp/BumpIPQC_#SEQ(1000,1)";
        Chart.ImageType = ChartImageType.Png;
        Chart.Palette = ChartColorPalette.Dundas;
        Chart.Height = chartH;
        Chart.Width = chartW;

        Chart.Palette = ChartColorPalette.Dundas;
        Chart.BackColor = Color.White;
        Chart.BackGradientEndColor = Color.Peru;
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
        Chart.BorderStyle = ChartDashStyle.Solid;
        Chart.BorderWidth = 3;
        Chart.BorderColor = Color.DarkBlue;

        Chart.ChartAreas.Add("Default");
        Chart.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
        Chart.ChartAreas["Default"].AxisX.LabelStyle.Interval = 1;
        Chart.ChartAreas["Default"].AxisX.LabelStyle.FontAngle = -45;
        //文字對齊
        Chart.ChartAreas["Default"].BorderStyle = ChartDashStyle.NotSet;
        Chart.ChartAreas["Default"].AxisY.LabelStyle.Font = new Font("Arial", 14, GraphicsUnit.Pixel);

        Chart.UI.Toolbar.Enabled = false;
        Chart.UI.ContextMenu.Enabled = true;

        // 找出 Source 所有分類 --> Week 
        DataTable weekGroupDT = UtilObj.fun_DataTable_SelectDistinct(DtSet, "yearWW");
        weekGroupDT.DefaultView.Sort = "yearWW asc";
        weekGroupDT = weekGroupDT.DefaultView.ToTable();

        Series series = default(Series);
        DataRow[] insideRows = null;
        string failMode = null;
        double failValue = 0;
        string weekStr = null;
        int colorInx = 0;
        string scriptStr = "";

        colorInx = (weekGroupDT.Rows.Count - 1);

        for (int toolIndex = 0; toolIndex <= (weekGroupDT.Rows.Count - 1); toolIndex++)
        {
            weekStr = (weekGroupDT.Rows[toolIndex]["yearWW"]).ToString();
            series = Chart.Series.Add(weekStr);
            series.ChartArea = "Default";
            series.Type = SeriesChartType.Column;
            series.Color = aryColor[colorInx];
            series.BorderColor = Color.White;
            series.BorderWidth = 1;


            for (int i = 0; i <= (setupDT.Rows.Count - 1); i++)
            {
                failMode = (setupDT.Rows[i]["Fail_Mode"].ToString().Trim()).Replace("'", "''");
                insideRows = DtSet.Select("yearWW='" + weekStr + "' and Fail_Mode='" + failMode + "'");

                failValue = 0;
                if (insideRows.Length > 0)
                {
                    if (insideRows[0]["YIELD_VALUE"] != null)
                    {
                        failValue = Convert.ToDouble(insideRows[0]["YIELD_VALUE"]);
                    }
                }

                Chart.Series[(weekStr)].Points.AddXY(failMode, failValue);
                Chart.Series[(weekStr)].Points[i].ToolTip = "Week" + weekStr + "\n" + "FailMode=" + failMode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();

            }
            colorInx = (colorInx - 1);

        }
        DetailParetoPanel.Controls.Add(Chart);

    }

    protected void gv_pie_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
    {

        if (e.Row.RowType == DataControlRowType.Header)
        {
            e.Row.Cells[0].Width = Unit.Pixel(80);
            e.Row.Cells[1].Width = Unit.Pixel(80);
            e.Row.Cells[2].Width = Unit.Pixel(80);
            e.Row.Cells[3].Width = Unit.Pixel(80);

        }

        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(50);
            for (int i = 4; i <= (e.Row.Cells.Count - 1); i++)
            {
                e.Row.Cells[i].Width = Unit.Pixel(50);
                e.Row.Cells[i].Text = e.Row.Cells[i].Text + "%";
            }

        }

    }

    protected void gr_lotview_RowDataBound(object sender, System.Web.UI.WebControls.GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(50);
            for (int i = 1; i <= (e.Row.Cells.Count - 1); i++)
            {
                e.Row.Cells[i].Width = Unit.Pixel(50);
                if (e.Row.Cells[i].Text.Length <= 0)
                {
                    e.Row.Cells[i].Text = "0%";
                }
                else
                {
                    e.Row.Cells[i].Text = e.Row.Cells[i].Text + "%";
                }
            }
        }

    }

    protected void gv_pie_RowDataBound1(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                System.Web.UI.WebControls.Label lab = new System.Web.UI.WebControls.Label();
                lab.Text = e.Row.Cells[i].Text;
                e.Row.Cells[i].Controls.Clear();
                ImageButton img = new ImageButton();
                img.ImageUrl = "~/images/s.gif";
                e.Row.Cells[i].Controls.Add(img);
                e.Row.Cells[i].Controls.Add(lab);
                e.Row.Cells[i].Width = Unit.Pixel(150);
            }

        }
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(80);
        }
    }

    protected void gr_lotview_RowDataBound1(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.Header)
        {
            for (int i = 0; i < e.Row.Cells.Count; i++)
            {
                System.Web.UI.WebControls.Label lab = new System.Web.UI.WebControls.Label();
                lab.Text = e.Row.Cells[i].Text;
                e.Row.Cells[i].Controls.Clear();
                ImageButton img = new ImageButton();
                img.ImageUrl = "~/images/s.gif";
                e.Row.Cells[i].Controls.Add(img);
                e.Row.Cells[i].Controls.Add(lab);
                e.Row.Cells[i].Width = Unit.Pixel(100);
            }

        }
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(30);
        }
    }

    protected void GV_LotRowData_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(30);
            String lot = (e.Row.Cells[7].Text);
            String trtm = (e.Row.Cells[6].Text);
            e.Row.ID = (lot + trtm);
            e.Row.Cells[7].Text = "<a name=\"" + (lot + trtm) + "\">" + lot + "</a>";
        }
    }

    //Fail Detail(Pareto)
    protected void GV_NewLotRowData_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Height = Unit.Pixel(30);
            String lot = (e.Row.Cells[7].Text);
            String trtm = (e.Row.Cells[6].Text);
            e.Row.ID = (lot + trtm);
            e.Row.Cells[7].Text = "<a name=\"" + (lot + trtm) + "\">" + lot + "</a>";
        }
    }

    //Fail Detail(Pareto)
    private void area_Pareto(ref DataTable MainDT, ref DataTable BinCodeDT, ref DataTable workDT, string cweek, string item)
    {
        //cweek = cweek.Substring(4, 2);
        double bvalue = 0;
        double fvalue = 0;
        string codeID = "";
        int rowIndex = 0;
        DataRow workDR = null;
        // DefectCode, FailMode, BinCode, Category, W ~, Delta

        for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
        {
            workDR = workDT.NewRow();
            workDR[0] = MainDT.Rows[i]["DefectCode"].ToString();
            workDR[1] = MainDT.Rows[i]["Fail_Mode"].ToString();
            workDR[2] = MainDT.Rows[i]["BinCode_Id"].ToString();
            workDR[3] = MainDT.Rows[i]["MF_Stage"].ToString();
            codeID = MainDT.Rows[i]["BinCode_Id"].ToString();
            rowIndex = 4;

            for (int j = 0; j <= (BinCodeDT.Rows.Count - 1); j++)
            {
                workDR[rowIndex] = BinCodeDT.Rows[j][codeID].ToString();
                rowIndex += 1;

                if (j == (BinCodeDT.Rows.Count - 2))
                {
                    bvalue = Convert.ToDouble(BinCodeDT.Rows[j][codeID]);
                }

                if (j == (BinCodeDT.Rows.Count - 1))
                {
                    fvalue = Convert.ToDouble(BinCodeDT.Rows[j][codeID]);
                }

            }

            if (bvalue > 0 | fvalue > 0)
            {
                workDR[rowIndex] = (Math.Round((bvalue - fvalue), 2)).ToString();
            }
            else
            {
                workDR[rowIndex] = "0";
            }
            workDT.Rows.Add(workDR);
        }
        gv_pie.DataSource = workDT;
        gv_pie.DataBind();
        UtilObj.Set_DataGridRow_OnMouseOver_Color(ref gv_pie, "#FFF68F", gv_pie.AlternatingRowStyle.BackColor);

        // 畫 Pie Chart
        if (MainDT.Rows.Count > 0)
        {
            Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
            ChartArea chartArea1 = new ChartArea();

            Chart.Palette = ChartColorPalette.Dundas;
            Chart.BackColor = Color.White;
            Chart.BackGradientEndColor = Color.Peru;
            Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
            Chart.BorderStyle = ChartDashStyle.Solid;
            Chart.BorderWidth = 3;
            Chart.BorderColor = Color.DarkBlue;

            Chart.ImageUrl = "temp/yieldP_#SEQ(1000,1)";
            Chart.ImageType = ChartImageType.Png;
            Chart.Palette = ChartColorPalette.Dundas;
            Chart.ChartAreas.Add(chartArea1);
            Chart.Height = chartH;
            Chart.Width = chartW;

            Series series1 = default(Series);
            series1 = Chart.Series.Add("MQCS");
            series1.BackGradientEndColor = Color.White;
            series1.Type = SeriesChartType.Pie;
            series1.ShowInLegend = true;
            series1.Font = new Font("Verdana", 10);
            series1.FontColor = Color.Red;
            series1.YValueType = ChartValueTypes.Double;
            series1.XValueType = ChartValueTypes.String;

            series1["PieLabelStyle"] = "Outside";
            series1.BorderWidth = 2;
            series1.BorderColor = System.Drawing.Color.FromArgb(26, 59, 105);

            Chart.Legends.Add("Legend1");
            Chart.Legends[0].Enabled = true;
            //Chart.Legends[0].Docking = Docking.Bottom;
            //Chart.Legends(0).Alignment = System.Drawing.StringAlignment.Center
            series1.LegendText = "#VALX [#PERCENT]";

            DataRow[] foundRows = null;
            foundRows = BinCodeDT.Select("yearWW='" + cweek + "'");
            double value = 0;
            string binCodeStr = "";
            string AlisStr = "";

            if (foundRows.Length > 0)
            {
                for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
                {
                    binCodeStr = (String)MainDT.Rows[i]["BinCode_Id"];
                    AlisStr = (String)MainDT.Rows[i]["BinCode_Id"];
                    value = Convert.ToDouble(foundRows[0][binCodeStr]);
                    value = Math.Round(value, 2);
                    series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value);
                    series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + "%";
                }
            }

            NewParetoPanel.Controls.Add(new LiteralControl("<tr><td>"));
            NewParetoPanel.Controls.Add(Chart);
            NewParetoPanel.Controls.Add(new LiteralControl("</td></tr>"));
        }
    }

    private void NewPareto_Chart(ref DataTable workDT, int TopN)
    {
        Dundas.Charting.WebControl.Chart Chart = new Dundas.Charting.WebControl.Chart();
        Chart.ImageUrl = "temp/BumpIPQC_#SEQ(1000,1)";
        Chart.ImageType = ChartImageType.Png;
        Chart.Palette = ChartColorPalette.Dundas;
        Chart.Height = chartH;
        Chart.Width = chartW;

        Chart.Palette = ChartColorPalette.Dundas;
        Chart.BackColor = Color.White;
        Chart.BackGradientEndColor = Color.Peru;
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;
        Chart.BorderStyle = ChartDashStyle.Solid;
        Chart.BorderWidth = 3;
        Chart.BorderColor = Color.DarkBlue;

        Chart.ChartAreas.Add("Default");
        Chart.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
        Chart.ChartAreas["Default"].AxisX.LabelStyle.Interval = 1;
        Chart.ChartAreas["Default"].AxisX.LabelStyle.FontAngle = -45;
        //文字對齊
        Chart.ChartAreas["Default"].BorderStyle = ChartDashStyle.NotSet;
        Chart.ChartAreas["Default"].AxisY.LabelStyle.Font = new Font("Arial", 14, GraphicsUnit.Pixel);

        if (TopN < 11)
        {
            // Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 18, FontStyle.Bold)
            Chart.ChartAreas["Default"].AxisX.LabelStyle.Font = new Font(FontFamily.GenericSansSerif, 18, FontStyle.Bold);
        }
        else
        {
            //Chart.ChartAreas["Default"].AxisX.LabelStyle.Font = new Font("Arial", 15, GraphicsUnit.Pixel);
            Chart.ChartAreas["Default"].AxisX.LabelStyle.Font = new Font(FontFamily.GenericSansSerif, 15, FontStyle.Bold);
        }



        Chart.UI.Toolbar.Enabled = false;
        Chart.UI.ContextMenu.Enabled = true;

        // 找出 Source 所有分類 --> Week 
        //DataTable weekGroupDT = UtilObj.fun_DataTable_SelectDistinct(DtSet, "yearWW");
        //weekGroupDT.DefaultView.Sort = "yearWW asc";
        //weekGroupDT = weekGroupDT.DefaultView.ToTable();

        workDT.DefaultView.Sort = workDT.Columns[workDT.Columns.Count - 2].Caption + " desc";
        DataTable weekGroupDT = workDT.DefaultView.ToTable();

        Series series = default(Series);
        //DataRow[] insideRows = null;
        string binCode = null;
        double failValue = 0;
        string weekStr = null;
        //int colorInx = 0;
        string scriptStr = "";

        //colorInx = (weekGroupDT.Rows.Count - 1);

        if (workDT.Rows.Count > 0)
        {
            weekStr = workDT.Rows[0][0].ToString();
            series = Chart.Series.Add(weekStr);
            series.ChartArea = "Default";
            series.Type = SeriesChartType.Column;
            series.Color = aryColor[0];
            series.BorderColor = Color.White;
            series.BorderWidth = 1;

            //if (TopN > 0)
            if (TopN > 0)
            {
                if (workDT.Rows.Count > TopN)
                {
                    for (int m = 0; m < TopN; m++)
                    {
                        failValue = 0;
                        binCode = (workDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
                        failValue = Convert.ToDouble(workDT.Rows[m]["Fail_Ratio"]);
                        failValue = Math.Round(failValue, 3);

                        Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
                        Chart.Series[(weekStr)].Points[m].ToolTip = weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
                    }
                }
                else
                {
                    for (int m = 0; m <= (workDT.Rows.Count - 1); m++)
                    {
                        failValue = 0;
                        binCode = (workDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
                        failValue = Convert.ToDouble(workDT.Rows[m]["Fail_Ratio"]);
                        failValue = Math.Round(failValue, 3);

                        Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
                        Chart.Series[(weekStr)].Points[m].ToolTip = weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
                    }
                }
            }
            else
            {
                for (int m = 0; m <= (workDT.Rows.Count - 1); m++)
                {
                    failValue = 0;
                    binCode = (workDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
                    failValue = Convert.ToDouble(workDT.Rows[m]["Fail_Ratio"]);
                    failValue = Math.Round(failValue, 3);

                    Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
                    Chart.Series[(weekStr)].Points[m].ToolTip = weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
                }

            }

        }




        //for (int toolIndex = 0; toolIndex <= (weekGroupDT.Rows.Count - 1); toolIndex++)
        //for (int toolIndex = 4; toolIndex <= (weekGroupDT.Columns.Count - 2); toolIndex++)
        //{
        //    //weekStr = (weekGroupDT.Rows[toolIndex]["yearWW"]).ToString();
        //    weekStr = weekGroupDT.Columns[toolIndex].Caption;
        //    series = Chart.Series.Add(weekStr);
        //    series.ChartArea = "Default";
        //    series.Type = SeriesChartType.Column;
        //    series.Color = aryColor[((weekGroupDT.Columns.Count - 2) - 4) - (toolIndex - 4)];
        //    series.BorderColor = Color.White;
        //    series.BorderWidth = 1;

        //    if (TopN > 0)
        //    {
        //        if (weekGroupDT.Rows.Count > TopN)
        //        {
        //            for (int m = 0; m < TopN; m++)
        //            {
        //                failValue = 0;
        //                binCode = (weekGroupDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
        //                failValue = Convert.ToDouble(weekGroupDT.Rows[m][weekStr]);
        //                failValue = Math.Round(failValue, 2);

        //                Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
        //                Chart.Series[(weekStr)].Points[m].ToolTip = "Week" + weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
        //            }
        //        }
        //        else
        //        {
        //            for (int m = 0; m <= (weekGroupDT.Rows.Count - 1); m++)
        //            {
        //                failValue = 0;
        //                binCode = (weekGroupDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
        //                failValue = Convert.ToDouble(weekGroupDT.Rows[m][weekStr]);
        //                failValue = Math.Round(failValue, 2);

        //                Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
        //                Chart.Series[(weekStr)].Points[m].ToolTip = "Week" + weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
        //            }
        //        }
        //    }
        //    else
        //    {
        //        for (int m = 0; m <= (weekGroupDT.Rows.Count - 1); m++)
        //        {
        //            failValue = 0;
        //            binCode = (weekGroupDT.Rows[m]["BinCode"].ToString().Trim()).Replace("'", "''");
        //            failValue = Convert.ToDouble(weekGroupDT.Rows[m][weekStr]);
        //            failValue = Math.Round(failValue, 2);

        //            Chart.Series[(weekStr)].Points.AddXY(binCode, failValue);
        //            Chart.Series[(weekStr)].Points[m].ToolTip = "Week" + weekStr + "\n" + "BinCode=" + binCode + "\n" + "Value=" + Math.Round(failValue, 5).ToString();
        //        }
        //    }
        //}

        NewParetoPanel.Controls.Add(new LiteralControl("<tr><td>"));
        NewParetoPanel.Controls.Add(Chart);
        NewParetoPanel.Controls.Add(new LiteralControl("</td></tr>"));
    }
}
