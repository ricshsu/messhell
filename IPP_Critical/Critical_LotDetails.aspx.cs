using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using Ext.Net;
using Microsoft.VisualBasic;
using System.Drawing;

public partial class Critical_LotDetails : System.Web.UI.Page
{

    struct Obj 
    {
        public string Category;
        public string DFrom;
        public string DTo;
        public string PartID;
        public string Critical_Item;
        public string valueType;
        public string HL;
        public string HLLOT;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        Obj obj;
        if (!X.IsAjaxRequest)
        {
          this.Window1.Hide();
          if ((String)Session["CriticalLogin"] == "Y") 
          {
              this.Button3.Hide();
          }
        }

        if (!this.IsPostBack) 
        {
            obj = new Obj();
            // Test
            //obj.Category = "WB";
            //obj.DFrom = "2013-02-01";
            //obj.DTo = "2013-04-01";
            //obj.PartID = "ITL134";
            //obj.Critical_Item = "22";
            //obj.valueType = "meanval";
            //obj.HL = "N";
            //obj.HLLOT = "N31BBF210061";
            // Production
            obj.Category = Request["CTYPE"];
            obj.DFrom = Request["DF"];
            obj.DTo = Request["DT"];
            obj.PartID = Request["PART"];
            obj.Critical_Item = Request["CI"];
            obj.valueType = Request["VT"];
            obj.HL = Request["HL"];
            obj.HLLOT = Request["HLLOT"];
            Session["Obj"] = obj;
            pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.Critical_Item, obj.valueType, false, obj.HL, obj.HLLOT);
        }

    }

    public void pageInit(string category, string DFrom, string DTo, string partID, string CItem, string valueType, bool isQuery, string isHL, string HLLOT)
    {
        string sqlStr = "";
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        Dundas.Charting.WebControl.Chart chartObj = default(Dundas.Charting.WebControl.Chart);
        DataTable correlationDt = new DataTable();
        SqlDataAdapter myAdpt = default(SqlDataAdapter);
        DataTable myDt = new DataTable();
        DataTable toolDt = new DataTable();

        try
        {
            if (category == "CPU")
            {
                category = "critical_lot";
            }
            /*else {
                category = "PPS";
            }*/
            conn.Open();
            // 取得 ID 對應 Parameter
            sqlStr = "select Param_Name from Critical_LOT_Params_BU_Rename ";
            sqlStr += "where FType='" + category + "' ";
            sqlStr += "and sub_id='" + CItem + "' ";
            myAdpt = new SqlDataAdapter(sqlStr, conn);
            myAdpt.Fill(myDt);
            CItem = (myDt.Rows[0]["Param_Name"]).ToString().Trim();
            
            // --- 加 Comment SQL---
            sqlStr = "select a.Lot, a.MeasureCount, a.PanelNo, a.part, a.Parametric_Measurement, a.MIS_OP, a.mchno, a.Plant, ";
            sqlStr += "a.meanval, a.maxval, a.minval, a.std, a.cpk, a.cp, a.usl, a.lsl, a.xucl, a.xlcl, a.SUCL, a.SLCL, ";
            sqlStr += "a.WECO_Rule1, a.WECO_Rule2, a.WECO_Rule3, a.WECO_Rule4, a.WECO_Rule5, a.WECO_Rule6, a.WECO_Rule7, a.WECO_Rule8, a.WECO_Rule9, ";
            sqlStr += "Convert(char(19), a.trtm, 120) as trtm, b.Comment, Convert(char(19), b.updateTime, 120) as updateTime ";
            sqlStr += "from Critical_LOT_Data a left join ( ";
            sqlStr += "Select part, lot, meas_item, Comment,updateTime, ";
            sqlStr += "ROW_NUMBER() Over (Partition By part, lot, meas_item Order By updateTime Desc) As Sort ";
            sqlStr += "From dbo.Critical_Lot_Control ";
            sqlStr += ") b on 1=1 ";
            sqlStr += "and b.Sort = 1 ";
            sqlStr += "and a.part = b.Part ";
            sqlStr += "and a.Lot = b.Lot  ";
            sqlStr += "and a.Parametric_Measurement = b.Meas_Item  ";
            sqlStr += "where 1=1  ";
            sqlStr += "and a.MeasureCount = 1  ";
            sqlStr += "and a.part='" + partID + "' ";
            sqlStr += "and a.Parametric_Measurement='" + CItem + "' ";
            sqlStr += "and a.trtm >= '" + DFrom + " 00:00:00' ";
            sqlStr += "and a.trtm <= '" + DTo + " 23:59:59' ";
            sqlStr += "GROUP BY a.Lot, a.MeasureCount, a.PanelNo, a.part, a.Parametric_Measurement, a.MIS_OP, a.mchno, a.Plant, ";
            sqlStr += "a.meanval, a.maxval, a.minval, a.std, a.cpk, a.cp, a.usl, a.lsl, a.xucl, a.xlcl, a.SUCL, a.SLCL, ";
            sqlStr += "a.WECO_Rule1, a.WECO_Rule2, a.WECO_Rule3, a.WECO_Rule4, a.WECO_Rule5, a.WECO_Rule6, a.WECO_Rule7, a.WECO_Rule8, a.WECO_Rule9, ";
            sqlStr += "a.trtm, b.Comment, b.updateTime ";
            sqlStr += "order by a.trtm";

            // --- 未加 Comment ---
            string toolStr = "";
            toolStr = "SELECT Lot, MeasureCount, PanelNo, part, Parametric_Measurement, MIS_OP as MPID, mchno as EQPID, Plant, meanval, maxval,minval,std,samplesize,cpk,cp,usl,lsl,xucl,xlcl,SUCL,SLCL,WECO_Rule1,WECO_Rule2,WECO_Rule3,WECO_Rule4,WECO_Rule5,WECO_Rule6,WECO_Rule7,WECO_Rule8,WECO_Rule9, Convert(char(19), trtm, 120) as trtm ";
            toolStr += "from Critical_LOT_Data ";
            toolStr += "where MeasureCount = 1 ";
            toolStr += "and part='" + partID + "' ";
            toolStr += "and Parametric_Measurement='" + CItem + "' ";
            toolStr += "and trtm >= '" + DFrom + " 00:00:00' ";
            toolStr += "and trtm <= '" + DTo + " 23:59:59' ";
            toolStr += "GROUP BY Lot,MeasureCount,PanelNo,part,Parametric_Measurement,MIS_OP,mchno,Plant,trtm,meanval,maxval,minval,std,samplesize,cpk,cp,usl,lsl,xucl,xlcl,SUCL,SLCL,WECO_Rule1,WECO_Rule2,WECO_Rule3,WECO_Rule4,WECO_Rule5,WECO_Rule6,WECO_Rule7,WECO_Rule8,WECO_Rule9, Convert(char(19), trtm, 120) ";
            toolStr += "ORDER BY MIS_OP, mchno, TRTM";

            myDt = new DataTable();
            myAdpt = new SqlDataAdapter(sqlStr, conn);
            myAdpt.Fill(myDt);

            myAdpt = new SqlDataAdapter(toolStr, conn);
            myAdpt.Fill(toolDt);
            conn.Close();

            if (myDt.Rows.Count > 0)
            {
                Lot_GridView.DataSource = myDt;
                Lot_GridView.DataBind();

                // --- Mean ---
                WecoTrendObj wObj = new WecoTrendObj();
                wObj.chartH = 400;
                wObj.chartW = 1100;
                wObj.valueType = "meanval";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.highlightLot = HLLOT;
                if (isHL == "Y")
                {
                    wObj.isHighlight = true;
                    wObj.HL_Day = DTo;
                }

                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref myDt,ref chartObj,false)))
                {
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[MEAN] " + partID.Replace("'", "''") + "  " + CItem + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    ChartPanel.Controls.Add(chartObj);
                    ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                }

                // --- STD ---
                wObj = new WecoTrendObj();
                wObj.chartH = 400;
                wObj.chartW = 1100;
                wObj.valueType = "std";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.highlightLot = HLLOT;
                if (isHL == "Y")
                {
                    wObj.isHighlight = true;
                    wObj.HL_Day = (txtDateTo.Text.Trim());
                }

                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref myDt, ref chartObj,false)))
                {
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[STD] " + partID.Replace("'", "''") + "  " + CItem + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    ChartPanel.Controls.Add(chartObj);
                    ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                }

                // --- Tool Matching ---
                wObj = new WecoTrendObj();
                wObj.chartH = 400;
                wObj.chartW = 1100;
                wObj.valueType = "meanval";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.highlightLot = HLLOT;
                wObj.ChartByTool = true;
                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref toolDt, ref chartObj,false)))
                {
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[Process Time By Tool] " + partID.Replace("'", "''") + "  " + CItem + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    ChartPanel.Controls.Add(chartObj);
                    ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                }
            }

        }
        catch (Exception ex)
        {}
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

    }

    protected void Lot_GridView_RowDataBound1(object sender, GridViewRowEventArgs e)
    { 
        string lot = "";
        string trtm = "";
        string param = "";
        string part = "";
        string userID = "";
        string comment = "";

        if (e.Row.RowType == DataControlRowType.Header)
        {
            e.Row.Cells[17].Visible = false;
        }

        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            // 為了要由圖連到 RowData
            part = (e.Row.Cells[1].Text);
            lot = (e.Row.Cells[2].Text);
            param = (e.Row.Cells[3].Text);
            trtm = (e.Row.Cells[7].Text);
            comment = (e.Row.Cells[17].Text.Trim());

            if (comment != "&nbsp;") 
            {
                e.Row.ToolTip = comment;
                //"特殊符號 Alt + 41400 ~ 41470";
                e.Row.Cells[0].Text = "★";
            }
       
            e.Row.ID = (lot + trtm);
            e.Row.Cells[1].Text = "<a name=\"" + (lot + trtm) + "\">" + part + "</a>";
            System.Web.UI.WebControls.Button butObj = (System.Web.UI.WebControls.Button)(e.Row.Cells[16].Controls[1]);
            e.Row.Cells[17].Visible = false;

            if (((String)Session["CriticalLogin"]) == "Y") {
                userID = (String)Session["CriticalUser"];
                butObj.Enabled = true;
                butObj.Attributes.Add("onclick", "javascript:openCommand('Critical_Command.aspx','Command','" + part + "','" + lot + "','" + param + "','" + userID + "',event); return false;");
            } else {
                butObj.Enabled = false;
            }

        }
    }

    // Refresh
    protected void Button1_Click(object sender, EventArgs e)
    {
        Obj obj = (Obj)Session["Obj"];
        pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.Critical_Item, obj.valueType, false, obj.HL, obj.HLLOT);
    }

    protected void btnLogin_Click(object sender, DirectEventArgs e)
    {
        this.Window1.Hide();
        string username = this.txtUsername.Text;
        string password = this.txtPassword.Text;
        Session["CriticalLogin"] = "Y";
        Session["CriticalUser"] = username;
        X.Call("pageReFresh");
    }

}