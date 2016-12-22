using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Ext.Net;
using System.Xml;
using System.Xml.Xsl;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;

public partial class CTF_DataDetails : System.Web.UI.Page
{

    struct Obj 
    {
        public String part_id;
        public String lot_id;
        public String meas_time;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!X.IsAjaxRequest)
        {
            Obj obj = new Obj();
            //obj.part_id = "ATK002";
            //obj.lot_id = "N28C57910011";
            //obj.meas_time = "2012-10-28 23:20:06";
            obj.part_id = Request["PART"];
            obj.lot_id = Request["LOT"];
            obj.meas_time = Request["ET"];
            Session["SaveObj"] = obj;
            this.Store1.DataSource = this.GetDataTable();
            this.Store1.DataBind();
        }
    }

    private DataTable GetDataTable()
    {
        
        Obj obj = (Obj)Session["SaveObj"];
        DataTable table = new DataTable() ;
        SqlConnection conn = null;
        SqlDataAdapter sAdpt = null;
        String sqlStr = "";
        sqlStr += "select Part_Id, lot_id, Machine_id, Meas_item, ";
        sqlStr += "round(Max_Value, 5) as Max_Value, ";
        sqlStr += "round(Min_Value, 5) as Min_Value, ";
        sqlStr += "round(Mean_value, 5) as Mean_value, ";
        sqlStr += "round(Std_value, 5) as Std_value, ";
        sqlStr += "round(Cp, 5) as CP, ";
        sqlStr += "round(Cpk, 5) as CPK, ";
        sqlStr += "RowData ";
        sqlStr += "from CTF_Monitor_Performance_Lot_Summary ";
        sqlStr += "where 1=1 ";
        sqlStr += "and part_id='" + (obj.part_id) + "' ";
        sqlStr += "and lot_id='" + (obj.lot_id) + "' ";
        sqlStr += "and Lot_Meas_End_DataTime='" + (obj.meas_time) + "' ";
        sqlStr += "order by Meas_item ";

        try
        {
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
            conn.Open();
            sAdpt = new SqlDataAdapter(sqlStr, conn);
            sAdpt.Fill(table);
            conn.Close();
        }
        catch (Exception error) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }

        return table;
    }

    protected void Store1_RefreshData(object sender, StoreReadDataEventArgs e)
    {
        this.Store1.DataSource = this.GetDataTable();
        this.Store1.DataBind();
    }

    protected void Store1_Submit(object sender, StoreSubmitDataEventArgs e)
    { 
        string format = this.FormatType.Value.ToString();
        XmlNode xml = e.Xml;
        this.Response.Clear();

        switch (format)
        {
            case "xml":
                string strXml = xml.OuterXml;
                this.Response.AddHeader("Content-Disposition", "attachment; filename=submittedData.xml");
                this.Response.AddHeader("Content-Length", strXml.Length.ToString());
                this.Response.ContentType = "application/xml";
                this.Response.Write(strXml);
                break;
            case "xls":
                this.Response.ContentType = "application/vnd.ms-excel";
                this.Response.AddHeader("Content-Disposition", "attachment; filename=submittedData.xls");
                XslCompiledTransform xtExcel = new XslCompiledTransform();
                xtExcel.Load(Server.MapPath("Excel.xsl"));
                xtExcel.Transform(xml, null, Response.OutputStream);
                break;
            case "csv":
                this.Response.ContentType = "application/octet-stream";
                this.Response.AddHeader("Content-Disposition", "attachment; filename=rowdata.csv");
                XslCompiledTransform xtCsv = new XslCompiledTransform();
                xtCsv.Load(Server.MapPath("Excel.xsl"));
                xtCsv.Transform(xml, null, Response.OutputStream);
                break;
        }
        this.Response.End();

    }

}