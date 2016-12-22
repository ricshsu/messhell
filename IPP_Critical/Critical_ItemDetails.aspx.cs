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

public partial class Critical_ItemDetails : System.Web.UI.Page
{

    struct Obj 
    {
        public String Category;
        public String DFrom;
        public String DTo;
        public String PartID;
        public String MAIN_ID;
        public String SUB_ID;
        public String YImpact;
        public String KeyModule;
        public String CriticalItem;
        public String EDAItem;
        public String HL;
        public String showProcess;
        public String HLLot;
    }

    protected void Page_Load(object sender, System.EventArgs e)
    {
        Obj obj;
        if (!IsPostBack)
        {
            obj = new Obj();

            // Test
            //obj.Category = "CPU";
            //obj.DFrom = "2012-11-23";
            //obj.DTo = "2012-12-04";
            //obj.PartID = "FCS071A";
            //obj.MAIN_ID = "3";
            //obj.SUB_ID = "36";
            //obj.HL = "Y";

            // Production
            obj.Category = Request["CTYPE"];
            obj.DFrom = Request["DF"];
            obj.DTo = Request["DT"];
            obj.PartID = Request["DP"];
            obj.MAIN_ID = Request["MAIN_ID"];
            obj.SUB_ID = Request["SUB_ID"];
            obj.HL = Request["HL"];
            obj.showProcess = Request["SP"];
            obj.HLLot = Request["HLLOT"];
            Session["CriticalItemObj"] = obj;
            pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.MAIN_ID, obj.SUB_ID, false, obj.HL, "Y", obj.HLLot);
        }

        if (!X.IsAjaxRequest)
        {
            this.Window1.Hide();
            if ((String)Session["CriticalItemLogin"] == "Y")
            {
                this.Button3.Hide();
            }
        }
    }

    private void pageInit(string PCategory, string DFrom, string DTo, string partID, string MAIN_ID, string SUB_ID, bool isQuery, string isHL, string showProcess, string HLLOT)
    {
        Dundas.Charting.WebControl.Chart chartObj;
        WecoTrendObj wObj;
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdpt;
        DataTable myDt;
        string sqlStr = "";
        string measureStr = "";
        string processStr = "";
        string processNormalStr = "";
        string processToolStr = "";
        string processOP = "";
        string ProcessStation = "";
        string MeasureStation = "";
        string yImpact = "";
        string KModule = "";
        string CItem = "";
        string EDAItem = "";
        try
        {
            // --- Source ---
            sqlStr = "select Yield_Impact_Item, Key_Module, Critical_Item, EDA_Item from Daily_CriticalItem_OOC_Monitor_Main_BU_Rename ";
            //sqlStr += "where Customer_id='INTEL' ";
            sqlStr += "where 1=1 ";
            sqlStr += "and Category='" + PCategory + "' ";
            sqlStr += "and MAIN_ID='" + MAIN_ID + "' ";
            sqlStr += "and ID='" + SUB_ID + "' ";
            sqlStr += "Group by Yield_Impact_Item, Key_Module, Critical_Item, EDA_Item";
            conn.Open();
            myAdpt = new SqlDataAdapter(sqlStr, conn);
            myDt = new DataTable();
            myAdpt.Fill(myDt);

            yImpact = (string)myDt.Rows[0]["Yield_Impact_Item"];
            KModule = (string)myDt.Rows[0]["Key_Module"];
            CItem = (string)myDt.Rows[0]["Critical_Item"];
            EDAItem = (string)myDt.Rows[0]["EDA_Item"];
            conn.Close();
        }
        catch (Exception ex){}
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

        // ----- MeasureTime -----
        measureStr += "select c.lot, c.Part_Id as part, c.parametric_measurement, c.MIS_OP, c.MCHNO, ";
        measureStr += "Convert(char(19), c.trtm, 120) as trtm, c.meanval, c.std, c.maxval, c.minval, c.xlcl, c.xucl, c.slcl, c.sucl, c.layer, c.cp, c.cpk, ";
        measureStr += "c.WECO_Rule1, c.WECO_Rule2, c.WECO_Rule3, c.WECO_Rule4, c.WECO_Rule5, c.WECO_Rule6, c.WECO_Rule7, c.WECO_Rule8, c.WECO_Rule9, b.comment, b.updateTime ";
        if (PCategory == "CPU" )
        {
            measureStr += "from view_IPP_CriticalItem_Monitor_BU_Rename c ";
        }
        else if (PCategory == "CS")
        {
            measureStr += "from view_IPP_CriticalItem_Monitor_BU_Rename_CS c ";           
        }
        else
        {
            measureStr += "from view_IPP_CriticalItem_Monitor_wb c ";
        }
        measureStr += "left join ( Select part, lot, meas_item, Comment,updateTime, ROW_NUMBER() Over (Partition By part, lot, meas_item Order By updateTime Desc) As Sort From dbo.Critical_Lot_Control ) b ";
        measureStr += "on 1=1 ";
        measureStr += "and c.part_id = b.Part ";
        measureStr += "and c.Lot = b.Lot ";
        measureStr += "and c.Parametric_Measurement = b.Meas_Item ";
        measureStr += "and b.Sort = 1 ";
        measureStr += "where 1=1 ";

        processStr += "select a.lot, a.part_id as part, a.parametric_measurement, a.MPID, a.EQPID, a.MachineName, c.MIS_OP, c.MCHNO, ";
        //processStr += "Convert(char(19), a.trtm, 120) as processTime, ";
        //processStr += "Convert(char(19), a.MeasureTime, 120) as trtm, ";
        processStr += "Convert(char(19), a.trtm, 120) as trtm, ";
        processStr += "Convert(char(19), a.MeasureTime, 120) as MeasureTime, ";
        processStr += "Convert(char(19), a.InStationDate, 120) as InStationDate, ";
        processStr += "Convert(char(19), a.InMachineDate, 120) as InMachineDate, ";
        processStr += "Convert(char(19), a.OutMachineDate, 120) as OutMachineDate, ";
        processStr += "Convert(char(19), a.OutStationDate, 120) as OutStationDate, ";
        processStr += "a.meanval, a.std, a.maxval, a.minval, a.xlcl, a.xucl, a.slcl, a.sucl, a.layer, a.cp, a.cpk, ";
        processStr += "c.WECO_Rule1, c.WECO_Rule2, c.WECO_Rule3, c.WECO_Rule4, c.WECO_Rule5, c.WECO_Rule6, c.WECO_Rule7, c.WECO_Rule8, c.WECO_Rule9, ";
        processStr += "b.comment, b.updateTime ";
        if (PCategory == "CPU")
        {
            processStr += "from view_IPP_Process_CriticalItem_Monitor a, view_IPP_CriticalItem_Monitor_BU_Rename c ";
        }
        else if (PCategory == "CS")
        {
            processStr += "from view_IPP_Process_CriticalItem_Monitor_cs a, view_IPP_CriticalItem_Monitor_BU_Rename_CS c ";
        }
        else
        {
            processStr += "from view_IPP_Process_CriticalItem_Monitor_wb a, view_IPP_CriticalItem_Monitor_wb c ";
        }
        processStr += "left join (Select part, lot, meas_item, Comment,updateTime, ";
        processStr += "ROW_NUMBER() Over (Partition By part, lot, meas_item Order By updateTime Desc) As Sort  ";
        processStr += "From dbo.Critical_Lot_Control ";
        processStr += ") b on 1=1 ";
        processStr += "and c.part_id = b.Part ";
        processStr += "and c.Lot = b.Lot ";
        processStr += "and c.Parametric_Measurement = b.Meas_Item ";
        processStr += "and b.Sort = 1 ";
        processStr += "where 1=1 ";
        processStr += "and a.MeasureTime = c.trtm ";
        processStr += "and a.Data_Source = c.Data_Source ";
        processStr += "and a.lot = c.lot ";
        processStr += "and a.Part_Id = c.Part_Id ";
        processStr += "and a.Yield_Impact_Item = c.Yield_Impact_Item ";
        processStr += "and a.Key_Module = c.Key_Module ";
        processStr += "and a.Critical_Item = c.Critical_Item ";
        processStr += "and a.EDA_Item = c.EDA_Item ";

        // --- Source SQL ---
        measureStr += "and c.trtm >= '" + DFrom + " 00:00:00' ";
        measureStr += "and c.trtm <= '" + DTo + " 23:59:59' ";
        processStr += "and a.MeasureTime >= '" + DFrom + " 00:00:00' ";
        processStr += "and a.MeasureTime <= '" + DTo + " 23:59:59' ";      
        
        if (partID != "All")
        {
            measureStr += "and c.Part_Id = '" + partID + "' ";
            processStr += "and a.Part_Id = '" + partID + "' ";
        }

        if (yImpact != "All")
        {
            measureStr += "and c.Yield_Impact_Item = '" + yImpact + "' ";
            processStr += "and a.Yield_Impact_Item = '" + yImpact + "' ";
        }

        if (KModule != "All")
        {
            measureStr += "and c.Key_Module = '" + KModule + "' ";
            processStr += "and a.Key_Module = '" + KModule + "' ";
        }

        if (CItem != "All")
        {
            measureStr += "and c.Critical_Item = '" + CItem + "' ";
            processStr += "and a.Critical_Item = '" + CItem + "' ";
        }

        if (EDAItem != "All")
        {
            measureStr += "and c.EDA_Item = '" + EDAItem + "' ";
            processStr += "and a.EDA_Item = '" + EDAItem + "' ";
        }

        measureStr += "group by c.lot, c.Part_Id, c.parametric_measurement, c.MIS_OP, c.MCHNO, c.trtm, ";
        measureStr += "c.meanval, c.std, c.maxval, c.minval, c.xlcl, c.xucl, c.slcl, c.sucl, c.layer, c.cp, c.cpk, ";
        measureStr += "c.WECO_Rule1, c.WECO_Rule2, c.WECO_Rule3, c.WECO_Rule4, c.WECO_Rule5, c.WECO_Rule6, c.WECO_Rule7, c.WECO_Rule8, c.WECO_Rule9, b.comment, b.updateTime ";
        measureStr += "order by c.trtm, c.lot asc ";

        processStr += "group by a.lot, a.part_id , a.parametric_measurement, a.MPID, a.EQPID, a.MachineName,c.MIS_OP, c.MCHNO, ";
        processStr += "a.trtm, a.MeasureTime, ";
        processStr += "a.InStationDate, a.InMachineDate, ";
        processStr += "a.OutMachineDate, a.OutStationDate, ";
        processStr += "a.meanval, a.std, a.maxval, a.minval, a.xlcl, a.xucl, a.slcl, a.sucl, a.layer, a.cp, a.cpk, ";
        processStr += "c.WECO_Rule1, c.WECO_Rule2, c.WECO_Rule3, c.WECO_Rule4, c.WECO_Rule5, c.WECO_Rule6, c.WECO_Rule7, c.WECO_Rule8, c.WECO_Rule9, ";
        processStr += "b.comment, b.updateTime ";

        processNormalStr = processStr + "Order by Convert(char(19), a.trtm, 120), a.lot asc";
        processToolStr = processStr + "Order by a.MPID, a.EqpID, Convert(char(19), a.trtm, 120), a.lot asc";

        try
        {
            // --- Source ---
            conn.Open();
            myAdpt = new SqlDataAdapter(measureStr, conn);
            myDt = new DataTable();
            myAdpt.Fill(myDt);

            if (myDt.Rows.Count > 0)
            {
                // --- By Measure Time ---
                MeasureStation = (myDt.Rows[0]["MIS_OP"] == System.DBNull.Value ? "" : (string)myDt.Rows[0]["MIS_OP"]) + "  " + (myDt.Rows[0]["MCHNO"] == System.DBNull.Value ? "" : (string)myDt.Rows[0]["MCHNO"]);
                wObj = new WecoTrendObj();
                wObj.FunctionType = "Critical_KPP";
                wObj.KPP_Part = partID.Replace("'", "''");
                wObj.KPP_YieldImpact = yImpact.Replace("'", "''");
                wObj.KPP_KeyModule = KModule.Replace("'", "''");
                wObj.KPP_CriticalItem = CItem.Replace("'", "''");
                wObj.KPP_IPP = EDAItem.Replace("'", "''");
                wObj.chartH = 500;
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
                if ((wObj.Call_DrawChart(ref myDt, ref chartObj,false)))
                {
                    Panel1.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[" + MeasureStation + "] Measure Time : " + partID + ":" + EDAItem + ":" + yImpact + ":" + KModule + ":" + CItem + "</td><td style='width:300px'></td></tr>"));
                    Panel1.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:middle;font-weight: bold'>"));
                    Panel1.Controls.Add(chartObj);
                    Panel1.Controls.Add(new LiteralControl("</td></tr>"));
                }

                // ----- By Process Time -----
                myAdpt = new SqlDataAdapter(processNormalStr, conn);
                myDt = new DataTable();
                myAdpt.Fill(myDt);
                processOP = (myDt.Rows[0]["MPID"] == System.DBNull.Value ? "" : (string)myDt.Rows[0]["MPID"]);
                ProcessStation = (myDt.Rows[0]["MPID"] == System.DBNull.Value ? "" : (string)myDt.Rows[0]["MPID"]) + "  " + (myDt.Rows[0]["EQPID"] == System.DBNull.Value ? "" : (string)myDt.Rows[0]["EQPID"]);
                wObj = new WecoTrendObj();
                wObj.FunctionType = "Critical_KPP";
                wObj.KPP_Part = partID.Replace("'", "''");
                wObj.KPP_YieldImpact = yImpact.Replace("'", "''");
                wObj.KPP_KeyModule = KModule.Replace("'", "''");
                wObj.KPP_CriticalItem = CItem.Replace("'", "''");
                wObj.KPP_IPP = EDAItem.Replace("'", "''");
                wObj.chartH = 500;
                wObj.chartW = 1100;
                wObj.valueType = "meanval";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.highlightLot = HLLOT;
                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref myDt, ref chartObj,false)))
                {
                    Panel1.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[" + ProcessStation + "] Process Time : " + partID + ":" + EDAItem + ":" + yImpact + ":" + KModule + ":" + CItem + "</td><td style='width:300px'></td></tr>"));
                    Panel1.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:middle;font-weight: bold'>"));
                    Panel1.Controls.Add(chartObj);
                    Panel1.Controls.Add(new LiteralControl("</td></tr>"));
                }

                GridView1.DataSource = myDt;
                GridView1.DataBind();

                // ----- By Process Tool -----
                myAdpt = new SqlDataAdapter(processToolStr, conn);
                myDt = new DataTable();
                myAdpt.Fill(myDt);
                wObj = new WecoTrendObj();
                wObj.FunctionType = "Critical_Item";
                wObj.KPP_Part = partID.Replace("'", "''");
                wObj.KPP_YieldImpact = yImpact.Replace("'", "''");
                wObj.KPP_KeyModule = KModule.Replace("'", "''");
                wObj.KPP_CriticalItem = CItem.Replace("'", "''");
                wObj.KPP_IPP = EDAItem.Replace("'", "''");
                wObj.chartH = 500;
                wObj.chartW = 1100;
                wObj.valueType = "meanval";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.highlightLot = HLLOT;
                wObj.ChartByTool = true;
                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref myDt, ref chartObj,false)))
                {
                    Panel1.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>[" + processOP + "] Process Time By Tool : " + partID + ":" + EDAItem + ":" + yImpact + ":" + KModule + ":" + CItem + "</td><td style='width:300px'></td></tr>"));
                    Panel1.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:middle;font-weight: bold'>"));
                    Panel1.Controls.Add(chartObj);
                    Panel1.Controls.Add(new LiteralControl("</td></tr>"));
                }

            }

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

    // Refresh
    protected void Button1_Click(object sender, EventArgs e)
    {
        Obj obj = (Obj)Session["CriticalItemObj"];
        pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.MAIN_ID, obj.SUB_ID, false, obj.HL, obj.showProcess, obj.HLLot);
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

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        string lot = "";
        string trtm = "";
        string param = "";
        string part = "";
        string userID = "";
        string comment = "";

        if (e.Row.RowType == DataControlRowType.Header)
        {
            e.Row.Cells[16].Visible = false;
            e.Row.Cells[17].Visible = false;
        }

        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            // 為了要由圖連到 RowData
            part = (e.Row.Cells[1].Text);
            lot = (e.Row.Cells[2].Text);
            param = (e.Row.Cells[3].Text);
            trtm = (e.Row.Cells[8].Text);
            comment = (e.Row.Cells[17].Text);

            e.Row.ID = (lot + trtm);
            e.Row.Cells[2].Text = "<a name=\"" + (lot + trtm) + "\">" + (e.Row.Cells[2].Text) + "</a>";

            if (comment != "&nbsp;")
            {
                e.Row.ToolTip = comment;
                //"特殊符號 Alt + 41400 ~ 41470";
                e.Row.Cells[0].Text = "★";
            }
            // 下註解
            System.Web.UI.WebControls.Button butObj = (System.Web.UI.WebControls.Button)(e.Row.Cells[15].Controls[1]);
            if (((String)Session["CriticalLogin"]) == "Y")
            {
                userID = (String)Session["CriticalUser"];
                butObj.Enabled = true;
                butObj.Attributes.Add("onclick", "javascript:openCommand('Critical_Command.aspx','Command','" + part + "','" + lot + "','" + param + "','" + userID + "',event); return false;");
            }
            else
            {
                butObj.Enabled = false;
            }
            e.Row.Cells[16].Visible = false;
            e.Row.Cells[17].Visible = false;
        }
    }
    
    // Event Correlation
    protected void chb_correlation_CheckedChanged(object sender, System.EventArgs e)
    {
        if (chb_correlation.Checked)
        {
            tr_yieldImpact.Visible = true;
            tr_critical.Visible = true;
            tr_inquery.Visible = true;
        }
        else
        {
            tr_yieldImpact.Visible = false;
            tr_critical.Visible = false;
            tr_inquery.Visible = false;
        }

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        string sqlStr = "";
        DataTable myDT = null;
        SqlDataAdapter myAdapter = default(SqlDataAdapter);

        try
        {
            // -- Data Source --
            conn.Open();

            // -- Yield Impact --
            sqlStr = "select yield_impact_item from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by yield_impact_item";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myDT = new DataTable();
            myAdapter.Fill(myDT);
            UtilObj.FillController(myDT, ref ddlYImpact, 0);

            // -- Key Module --
            sqlStr = "select key_module from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by key_module";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myDT = new DataTable();
            myAdapter.Fill(myDT);
            UtilObj.FillController(myDT, ref ddlKModule, 0);

            // -- Critical Item --
            sqlStr = "select Critical_item from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by Critical_item";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myDT = new DataTable();
            myAdapter.Fill(myDT);
            UtilObj.FillController(myDT, ref ddlCItem, 0);

            // -- Layer --
            sqlStr = "select EDA_ITEM from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by EDA_ITEM";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myDT = new DataTable();
            myAdapter.Fill(myDT);
            UtilObj.FillController(myDT, ref ddl_layer, 0);

            // -- By Date --
            string sTime = System.DateTime.Now.AddDays(-14).ToString("yyyy-MM-dd");
            string eTime = System.DateTime.Now.ToString("yyyy-MM-dd");
            txtDateFrom.Text = sTime;
            txtDateTo.Text = eTime;
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

        Obj obj = (Obj)Session["CriticalItemObj"];
        pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.MAIN_ID, obj.SUB_ID, false, obj.HL, obj.showProcess, obj.HLLot);
    }

    protected void but_Execute_Click(object sender, System.EventArgs e)
    {
        Obj obj = (Obj)Session["CriticalItemObj"];
        pageInit(obj.Category, obj.DFrom, obj.DTo, obj.PartID, obj.MAIN_ID, obj.SUB_ID, false, obj.HL, obj.showProcess, obj.HLLot);
    }

}