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

public partial class Critical_KPPDetails : System.Web.UI.Page
{

    struct Obj 
    {
        public String DFrom;
        public String DTo;
        public String Dsource;
        public String PartID;
        public String YImpact;
        public String KeyModule;
        public String CriticalItem;
        public String EDAItem;
        public String HL;
    }

    protected void Page_Load(object sender, System.EventArgs e)
    {
        Obj obj;
        if (!IsPostBack)
        {
            obj = new Obj();

            // Test
            //obj.DFrom = "2012-11-23";
            //obj.DTo = "2012-12-04";
            //obj.Dsource = "All";
            //obj.PartID = "FCS071A";
            //obj.YImpact = "BUMP";
            //obj.KeyModule = "SR";
            //obj.CriticalItem = "SRO FB Nest";
            //obj.EDAItem = "IPP-SR5";
            //obj.HL = "Y";

            // Production
            obj.DFrom = Request["DF"];
            obj.DTo = Request["DT"];
            obj.Dsource = Request["DS"];
            obj.PartID = Request["DP"];
            obj.YImpact = Request["YI"];
            obj.KeyModule = Request["KM"];
            obj.CriticalItem = Request["CI"];
            obj.EDAItem = Request["EI"];
            obj.HL = Request["HL"];
            Session["CriticalItemObj"] = obj;
            pageInit(obj.DFrom, obj.DTo, obj.Dsource, obj.PartID, obj.YImpact, obj.KeyModule, obj.CriticalItem, obj.EDAItem, false, obj.HL);
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

    private void pageInit(string DFrom, string DTo, string DataSource, string partID, string yImpact, string KModule, string CItem, string EDAItem, bool isQuery, string isHL)
    {
        string corrStr = "";
        string sqlStr = "";
        sqlStr += "select a.lot, a.part_id as part, a.parametric_measurement, a.MPID, a.EQPID, a.MachineName, Convert(char(19), a.trtm, 120) as trtm, Convert(char(19), a.MeasureTime, 120) as MeasureTime, ";
        sqlStr += "Convert(char(19), a.InStationDate, 120) as InStationDate, ";
        sqlStr += "Convert(char(19), a.InMachineDate, 120) as InMachineDate, ";
        sqlStr += "Convert(char(19), a.OutMachineDate, 120) as OutMachineDate, ";
        sqlStr += "Convert(char(19), a.OutStationDate, 120) as OutStationDate, ";
        sqlStr += "a.meanval, a.std, a.maxval, a.minval, a.xlcl, a.xucl, a.slcl, a.sucl, a.layer, a.cp, a.cpk, ";
        sqlStr += "a.WECO_Rule1, a.WECO_Rule2, a.WECO_Rule3, a.WECO_Rule4, a.WECO_Rule5, a.WECO_Rule6, a.WECO_Rule7, a.WECO_Rule8, a.WECO_Rule9, ";
        sqlStr += "b.comment, b.updateTime, a.NCW ";
        sqlStr += "from view_IPP_Process_CriticalItem_Monitor a left join ( ";
        sqlStr += "Select part, lot, meas_item, Comment,updateTime, ";
        sqlStr += "ROW_NUMBER() Over (Partition By part, lot, meas_item Order By updateTime Desc) As Sort  ";
        sqlStr += "From dbo.Critical_Lot_Control ";
        sqlStr += ") b on 1=1 ";
        sqlStr += "and a.part_id = b.Part ";
        sqlStr += "and a.Lot = b.Lot ";
        sqlStr += "and a.Parametric_Measurement = b.Meas_Item ";
        sqlStr += "and b.Sort = 1 ";
        sqlStr += "where 1=1 ";
        corrStr = sqlStr;

        // --- Source SQL ---
        sqlStr += "and a.trtm >= '" + DFrom + " 00:00:00' ";
        sqlStr += "and a.trtm <= '" + DTo + " 23:59:59' ";

        if (DataSource != "All")
        {
            sqlStr += "and a.Data_Source = '" + DataSource + "' ";
            corrStr += "and a.Data_Source = '" + DataSource + "' ";
        }

        if (partID != "All")
        {
            sqlStr += "and a.Part_Id = '" + partID + "' ";
            corrStr += "and a.Part_Id = '" + partID + "' ";
        }

        if (yImpact != "All")
        {
            sqlStr += "and a.Yield_Impact_Item = '" + yImpact + "' ";
        }


        if (KModule != "All")
        {
            sqlStr += "and a.Key_Module = '" + KModule + "' ";
        }

        if (CItem != "All")
        {
            sqlStr += "and a.Critical_Item = '" + CItem + "' ";
        }

        if (EDAItem != "All")
        {
            sqlStr += "and a.EDA_Item = '" + EDAItem + "' ";
        }
        sqlStr += "group by a.lot, a.part_id , a.parametric_measurement, a.MPID, a.EQPID, a.MachineName, a.trtm, a.MeasureTime, ";
        sqlStr += "a.InStationDate, a.InMachineDate, a.OutMachineDate, a.OutStationDate, ";
        sqlStr += "a.meanval, a.std, a.maxval, a.minval, a.xlcl, a.xucl, a.slcl, a.sucl, a.layer, a.cp, a.cpk, ";
        sqlStr += "a.WECO_Rule1, a.WECO_Rule2, a.WECO_Rule3, a.WECO_Rule4, a.WECO_Rule5, a.WECO_Rule6, a.WECO_Rule7, a.WECO_Rule8, a.WECO_Rule9, ";
        sqlStr += "b.comment, b.updateTime, a.NCW ";
        sqlStr += "order by a.trtm, a.lot asc ";

        // --- DataTime ---
        corrStr += "and a.trtm >= '" + (txtDateFrom.Text) + " 00:00:00' ";
        corrStr += "and a.trtm <= '" + (txtDateTo.Text) + " 23:59:59' ";
        // --- Correlation Yield Impact ---
        corrStr += "and a.Yield_Impact_Item = '" + (ddlYImpact.SelectedValue).Trim() + "' ";
        // -- Correlation Key Module --
        corrStr += "and a.Key_Module = '" + (ddlKModule.SelectedValue).Trim() + "' ";
        // -- Correlation Criteria Item --
        corrStr += "and a.Critical_Item = '" + (ddlCItem.SelectedValue).Trim() + "' ";
        // -- Correlation EDA_Item --
        corrStr += "and a.EDA_Item='" + (ddl_layer.SelectedValue).Trim() + "' ";
        corrStr += "order by a.trtm, a.lot desc";

        DataTable myDt = new DataTable();
        DataTable correlationDt = new DataTable();
        SqlDataAdapter myAdpt = default(SqlDataAdapter);
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());

        try
        {
            // --- Source ---
            conn.Open();
            myAdpt = new SqlDataAdapter(sqlStr, conn);
            myAdpt.Fill(myDt);
            conn.Close();

            if (myDt.Rows.Count > 0)
            {
                Dundas.Charting.WebControl.Chart chartObj = default(Dundas.Charting.WebControl.Chart);
                // --- Mean ---
                WecoTrendObj wObj = new WecoTrendObj();
                wObj.FunctionType = "Critical_KPP";
                wObj.KPP_Part = partID.Replace("'", "''");
                wObj.KPP_IPP = EDAItem.Replace("'", "''");
                wObj.KPP_YieldImpact = yImpact.Replace("'", "''");
                wObj.KPP_KeyModule = KModule.Replace("'", "''");
                wObj.KPP_CriticalItem = CItem.Replace("'", "''");
                wObj.chartH = 600;
                wObj.chartW = 1250;
                wObj.valueType = "meanval";
                wObj.txtDateFrom = DFrom;
                wObj.txtDateTo = DTo;
                wObj.notDetail = false;
                wObj.linkToPoint = true;
                wObj.specialOldData = true;
                if (isHL == "Y")
                {
                    wObj.isHighlight = true;
                    wObj.HL_Day = DTo;
                }

                chartObj = new Dundas.Charting.WebControl.Chart();
                if ((wObj.Call_DrawChart(ref myDt, ref chartObj,false)))
                {
                    //Panel1.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>" + partID + ":" + EDAItem + ":" + yImpact + ":" + KModule + ":" + CItem + "</td><td style='width:300px'></td></tr>"));
                    Panel1.Controls.Add(new LiteralControl("<tr><td valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    Panel1.Controls.Add(chartObj);
                    Panel1.Controls.Add(new LiteralControl("</td></tr>"));
                }

                // --- GridView ---
                GridView1.DataSource = myDt;
                GridView1.DataBind();
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
    }

    // Refresh
    protected void Button1_Click(object sender, EventArgs e)
    {
        Obj obj = (Obj)Session["CriticalItemObj"];
        pageInit(obj.DFrom, obj.DTo, obj.Dsource, obj.PartID, obj.YImpact, obj.KeyModule, obj.CriticalItem, obj.EDAItem, false, obj.HL);
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
            e.Row.Cells[21].Visible = false;
        }

        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            // 為了要由圖連到 RowData
            part = (e.Row.Cells[1].Text);
            lot = (e.Row.Cells[2].Text);
            param = (e.Row.Cells[3].Text);
            trtm = (e.Row.Cells[8].Text);
            comment = (e.Row.Cells[21].Text);

            e.Row.ID = (lot + trtm);
            e.Row.Cells[2].Text = "<a name=\"" + (lot + trtm) + "\">" + (e.Row.Cells[2].Text) + "</a>";

            if (comment != "&nbsp;")
            {
                e.Row.ToolTip = comment;
                e.Row.Cells[0].Text = "★"; //"特殊符號 Alt + 41400 ~ 41470";
            }
            // 下註解
            System.Web.UI.WebControls.Button butObj = (System.Web.UI.WebControls.Button)(e.Row.Cells[20].Controls[1]);
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
            e.Row.Cells[21].Visible = false;
            for (int i = 0; i < e.Row.Cells.Count; i++ )
            {
                e.Row.Cells[i].Font.Size = FontUnit.XXSmall;
            }
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
        pageInit(obj.DFrom, obj.DTo, obj.Dsource, obj.PartID, obj.YImpact, obj.KeyModule, obj.CriticalItem, obj.EDAItem, false, obj.HL);

    }

    protected void but_Execute_Click(object sender, System.EventArgs e)
    {
        Obj obj = (Obj)Session["CriticalItemObj"];
        pageInit(obj.DFrom, obj.DTo, obj.Dsource, obj.PartID, obj.YImpact, obj.KeyModule, obj.CriticalItem, obj.EDAItem, false, obj.HL);
    }

}