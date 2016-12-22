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


public partial class FailDetail_Daily : System.Web.UI.Page
{

    int chartH = 400;
    int chartW = 1000;

    Color[] aryColor = 
    {
		 Color.Blue,
		 Color.DarkOrange,
		 Color.Purple,
		 Color.DarkGreen,
		 Color.DodgerBlue,
		 Color.Firebrick,
		 Color.Olive,
		 Color.Green
	};

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!this.IsPostBack)
        {
            try
            {
                pageInit(((String)Request["C"]), (Request["CA"].ToString()), (Request["P"].ToString()), (Request["F"].ToString()), (Request["D"].ToString()), (Request["PLANT"].ToString()));
                //pageInit("INTEL", "CPU", "SNB P22", "Bump fail", "2013-01-02", "All");
            }
            catch (Exception ex)
            {
            }
        }
    }

    private void pageInit(string customer_id, string category, string production, string failMode, string dateStr, string plant)
    {
        failMode = failMode.Replace("000", "''"); // 因為有 ' 字元的問題, 所以需要跳脫, 在前一頁已經用 000 代替 ' ,不然 javascript 傳不過來
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        string sqlStr = "";
        DataSet ds = null;
        DataTable dt = null;
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        
        try
        {
            // --- Row Data SQL --- 
            sqlStr = "select distinct  ";
            sqlStr += "Convert(char(10), DataTime, 120) as DataTime, ";
            sqlStr += "Part_Id, production_type, Lot_Id, Customer_Id, Fail_Mode, MF_Stage, DefectCode, BinCode, Original_Input_QTY, Fail_Count, round(Fail_ratio, 5) as Fail_ratio, FE_Plant, BE_Plant ";
            sqlStr += "from dbo.BinCode_Daily_Lot ";
            sqlStr += "where 1=1 ";
            sqlStr += "and Customer_ID='{0}' ";
            sqlStr += "and Category='{1}' ";
            sqlStr += "and production_type='{2}' ";
            sqlStr += "and fail_mode='{3}' ";
            sqlStr += "and trtm='{4}' ";
            sqlStr += "order by lot_id, MF_Stage, Fail_Mode";
            sqlStr = string.Format(sqlStr, customer_id, category, production, failMode, dateStr);

            conn.Open();
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            dt = new DataTable();
            myAdapter.Fill(dt);
            conn.Close();

            Lot_GridView.DataSource = dt;
            Lot_GridView.DataBind();
            UtilObj.Set_DataGridRow_OnMouseOver_Color(ref Lot_GridView, "#FFF68F", Lot_GridView.AlternatingRowStyle.BackColor);
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

}