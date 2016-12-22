using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.IO;
using System.Collections;
using System.Web.UI.DataVisualization.Charting;
using System.Drawing;
using System.Xml;
using System.Xml.Linq;

public partial class FailModeQuery : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        but_Execute.Attributes.Add("onclick", "javascript:document.getElementById(\"lab_wait\").innerText='Please wait ......';" +
                                              "javascript:document.getElementById(\"but_Execute\").disabled=true;" +
                                              Page.GetPostBackEventReference(but_Execute));
        if(!this.IsPostBack) 
        {
            pageInit();
        }
    }

    private void pageInit()
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string partStr = " ";

        try
        {
            conn.Open();
            // === Customer ID ===
            sqlStr = "select customer_id from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "group by customer_id order by customer_id";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlCustomer, 0);
            // === CPU / ChipSet ===
            sqlStr = "select category from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "group by category order by category";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlCategory, 0);
            // === Product / Part ===
            if (rbl_BySource.SelectedIndex == 0)
            {
                partStr = "production_id";
            } else {
                partStr = "part_id";
            }
            sqlStr = "select " + partStr + " ";
            sqlStr += "from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "and category='" + (ddlCategory.SelectedValue) + "' ";
            sqlStr += "group by " + partStr + " order by " + partStr;
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlPart, 0);

            // 填入日期
            txtDateFrom.Text = DateTime.Now.AddDays(-14).ToString("yyyy-MM-dd");
            txtDateTo.Text = DateTime.Now.ToString("yyyy-MM-dd");
            // 填入週數
            sqlStr = "select yearWW ";
            sqlStr += "from SystemDateMapping where 1=1 ";
            sqlStr += "and Datetime >= '" + DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd") + "' ";
            sqlStr += "and Datetime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' ";
            sqlStr += "group by yearWW order by yearWW desc";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlWeekStart, 0);
            UtilObj.FillController(MainDT, ref ddlWeekEnd, 0);

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
    }

    // 一般查詢 / 上傳檔案 轉換
    protected void RadioButtonList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        tr_customer.Visible = false;
        tr_cpucs.Visible = false;
        tr_partSelect.Visible = false;
        tr_partData.Visible = false;
        tr_upload.Visible = false;
        tr_Stage.Visible = false;
        tr_time.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_lotlist.Visible = false;
        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_execute.Visible = false;
        tr_result.Visible = false;
        if (rb_mode_selected.SelectedIndex > -1)
        {
            rb_mode_selected.SelectedItem.Selected = false;
        }
 
        if (RadioButtonList1.SelectedIndex == 0)
        {
            tr_customer.Visible = true;
            tr_cpucs.Visible = true;
            tr_partSelect.Visible = true;
            tr_partData.Visible = true;
            tr_time.Visible = true;
        }
        else 
        {
            tr_upload.Visible = true;
            tr_lotlist.Visible = true;
        }
    }

    // Customer 置換
    protected void ddlCustomer_SelectedIndexChanged(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";

        try
        {
            conn.Open();
            // === CPU / ChipSet ===
            sqlStr = "select category from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "group by category order by category";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // CPU / ChipSet 置換
    protected void ddlCategory_SelectedIndexChanged(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string partStr = " ";
        string tableStr = "";

        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_execute.Visible = false;
        tr_result.Visible = false;
        
        if (rb_mode_selected.SelectedIndex > -1)
        {
            rb_mode_selected.SelectedItem.Selected = false;
        }

        try
        {
            conn.Open();
            // === Product / Part ===
            if (rbl_BySource.SelectedIndex == 0)
            {
                partStr = "production_id";
            }
            else
            {
                partStr = "part_id";
            }
            sqlStr = "select " + partStr + " ";
            sqlStr += "from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "and category='" + (ddlCategory.SelectedValue) + "' ";
            sqlStr += "group by " + partStr + " order by " + partStr;
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlPart, 0);
            if (ddlCategory.SelectedValue == "CPU")
            {
                tableStr = " dbo.BinCode_FailMode_Customer_Mapping a, dbo.BinCode b ";
            }
            else
            {
                tableStr = " dbo.CS_BinCode_FailMode_Customer_Mapping a, dbo.CS_BinCode b ";
            }
            sqlStr = "select a.MF_Stage ";
            sqlStr += "from " + tableStr + " ";
            sqlStr += "where 1=1 ";
            sqlStr += "and a.Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
            sqlStr += "group by a.MF_Stage";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillLitsBoxController(MainDT, ref lb_StageSource, 0);
            tr_failMode.Visible = false;
            tr_defectCode.Visible = false;
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // Production / Part 轉換
    protected void rbl_BySource_SelectedIndexChanged(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string partStr = " ";

        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_execute.Visible = false;
        tr_result.Visible = false;

        if (rb_mode_selected.SelectedIndex > -1) 
        {
            rb_mode_selected.SelectedItem.Selected = false;
        }

        try
        {
            conn.Open();
            // === Product / Part ===
            if (rbl_BySource.SelectedIndex == 0)
            {
                partStr = "production_id";
                lab_ProductType_PartID.Text = "Product Type";
            }
            else
            {
                partStr = "part_id";
                lab_ProductType_PartID.Text = "Part ID";
            }
            sqlStr = "select " + partStr + " ";
            sqlStr += "from Customer_Prodction_Mapping where 1=1 ";
            sqlStr += "and fail_function=1 ";
            //sqlStr += "and enable=1 ";
            sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "and category='" + (ddlCategory.SelectedValue) + "' ";
            sqlStr += "group by " + partStr + " order by " + partStr;
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddlPart, 0);
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // By Date
    protected void rb_byDate_CheckedChanged(object sender, EventArgs e)
    {
        rb_byWeek.Checked = false;
        // Date 
        txtDateFrom.Enabled = false;
        txtDateTo.Enabled = false;  
        // Week
        ddlWeekStart.Enabled = false;
        ddlWeekEnd.Enabled = false;
        if (rb_byDate.Checked)
        {
            txtDateFrom.Enabled = true ;
            txtDateTo.Enabled = true;
        } else {
            rb_byWeek.Checked = true;
        }
    }

    // By Week
    protected void rb_byWeek_CheckedChanged(object sender, EventArgs e)
    {
        rb_byDate.Checked = false;
        // Date 
        txtDateFrom.Enabled = false;
        txtDateTo.Enabled = false;
        // Week
        ddlWeekStart.Enabled = false;
        ddlWeekEnd.Enabled = false;
        if (rb_byWeek.Checked)
        {
            ddlWeekStart.Enabled = true;
            ddlWeekEnd.Enabled = true;
        }
        else
        {
            rb_byDate.Checked = true;
        }
    }
    
    // InQuery
    protected void but_Execute_Click(object sender, EventArgs e)
    {
        lab_wait.Text = "";
        if (RadioButtonList1.SelectedIndex == 1 && lb_lotList.Items.Count <= 0) 
        {
            showMessage("請上傳 Lot List !");
            return;
        }

        if (RadioButtonList1.SelectedIndex == 0)
        {
            if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
            {
                failModeMain();
            }
            else 
            {
                ptAOIQuery();
            }
        }
        else 
        {
            lotQuery();
        }
    }
    
    // Fail Mode Main 
    private void failModeMain() 
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        string sqlStr = "";
        string conditionStr = "";
        string lotCondition = "";
        ArrayList partAry = new ArrayList();
        DataTable LotDT = null;
        DataTable RowDT = null;
        String partStr = "";

        // Part Condition
        if (ddlPart.SelectedValue != "All")
        {
            if (rbl_BySource.SelectedIndex == 0)
            {
                conditionStr += "and b.production_type='" + ddlPart.SelectedValue + "' ";
                lotCondition += "and production_type='" + ddlPart.SelectedValue + "' ";
                partStr = (ddlPart.SelectedValue);
            }
            else
            {
                conditionStr += "and b.part_id='" + ddlPart.SelectedValue + "' ";
                lotCondition += "and part_id='" + ddlPart.SelectedValue + "' ";
                partStr = (ddlPart.SelectedValue);
            }
        }
        // DataTime
        if (rb_byDate.Checked) 
        {
            conditionStr += "and b.datatime >= '" + txtDateFrom.Text + " 00:00:00' ";
            conditionStr += "and b.datatime <= '" + txtDateTo.Text + " 23:59:59' ";
            lotCondition += "and datatime >= '" + txtDateFrom.Text + " 00:00:00' ";
            lotCondition += "and datatime <= '" + txtDateTo.Text + " 23:59:59' ";
        }

        if (rb_mode_selected.SelectedIndex == 0) 
        {
            // Fail Mode 
            if (lb_failModeSource.Items.Count != 0) //代表沒有全選
            {
                string failModeStr = "";
                for (int i = 0; i < lb_failModeShow.Items.Count; i++)
                {
                    failModeStr += "'" + (lb_failModeShow.Items[i].Value).Replace("'", "''") + "',";
                }
                failModeStr = failModeStr.Substring(0, (failModeStr.Length - 1));
                conditionStr += "and b.fail_mode in (" + failModeStr + ") ";
            }
        }
        else if (rb_mode_selected.SelectedIndex == 1) 
        {
            // Bin Code 
            if (lb_dcodeSource.Items.Count != 0) //代表沒有全選
            {
                string binCodeStr = "";
                for (int i = 0; i < lb_dcodeShow.Items.Count; i++)
                {
                    binCodeStr += "'" + (lb_dcodeShow.Items[i].Value).Replace("'", "''") + "',";
                }
                binCodeStr = binCodeStr.Substring(0, (binCodeStr.Length - 1));
                conditionStr += "and b.DefectCode in (" + binCodeStr + ") ";
            }
        }

        try
        {
            conn.Open();
            // 依Stage
            String stageStr = "";
            for (int j = 0; j < lb_StageShow.Items.Count; j++)
            {
                stageStr = (lb_StageShow.Items[j].Text);
                if (rb_mode_selected.SelectedIndex == 0)
                {
                    sqlStr = "select b.MF_Stage, b.Fail_Mode, b.Lot_Id, Convert(char(20), Max(b.datatime), 120) as Trtm, Round((convert(float, SUM(Fail_Count))/Max(Original_Input_Qty)), 6) as pvalue ";
                    sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
                    sqlStr += "where 1 = 1 ";
                    sqlStr += "and a.customer_id=b.customer_id ";
                    sqlStr += "and a.Production_Id=b.Production_Type ";
                    sqlStr += "and a.Part_Id=b.Part_Id ";
                    sqlStr += "and b.customer_id='" + ddlCustomer.SelectedValue + "' ";
                    sqlStr += "and b.category='" + ddlCategory.SelectedValue + "' ";
                    sqlStr += "and b.MF_Stage='" + stageStr + "' ";
                    sqlStr += conditionStr;
                    sqlStr += "group by b.MF_Stage, b.Fail_Mode, Lot_Id ";
                    sqlStr += "order by Convert(char(20), Max(b.datatime), 120)";
                    RowDT = new DataTable();
                    myAdapter = new SqlDataAdapter(sqlStr, conn);
                    myAdapter.Fill(RowDT);
                    if (RowDT.Rows.Count > 0)
                    {
                        // 依 FailMode
                        failModeQuery(partStr, stageStr, ref RowDT);
                    }
                    else
                    {
                        ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + partStr + "</td><td style='width:300px'></td></tr>"));
                        ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;'>"));
                        ChartPanel.Controls.Add(new LiteralControl("無資料 !</td></tr>"));
                    }
                }
                else if (rb_mode_selected.SelectedIndex == 1)
                {
                    sqlStr = "select b.MF_Stage, b.DefectCode, b.Lot_Id, Convert(char(20), MAX(b.datatime), 120) as trtm, Round((convert(float, SUM(Fail_Count))/Max(Original_Input_Qty)), 6) as pvalue ";
                    sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
                    sqlStr += "where 1 = 1 ";
                    sqlStr += "and a.customer_id=b.customer_id ";
                    sqlStr += "and a.Production_Id=b.Production_Type ";
                    sqlStr += "and a.Part_Id=b.Part_Id ";
                    sqlStr += "and b.customer_id='" + ddlCustomer.SelectedValue + "' ";
                    sqlStr += "and b.category='" + ddlCategory.SelectedValue + "' ";
                    sqlStr += "and b.MF_Stage='" + stageStr + "' ";
                    sqlStr += conditionStr;
                    sqlStr += "group by b.MF_Stage, b.DefectCode, Lot_Id ";
                    sqlStr += "order by Convert(char(20), Max(b.datatime), 120)";
                    RowDT = new DataTable();
                    myAdapter = new SqlDataAdapter(sqlStr, conn);
                    myAdapter.Fill(RowDT);
                    if (RowDT.Rows.Count > 0)
                    {
                        // 依 Defect Code
                        defectCodeQuery(partStr, stageStr, ref RowDT);
                    }
                    else
                    {
                        ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + partStr + "</td><td style='width:300px'></td></tr>"));
                        ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;'>"));
                        ChartPanel.Controls.Add(new LiteralControl("無資料 !</td></tr>"));
                    }
                }
            }
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // Fail Mode Query
    private void failModeQuery(String part_id, String stage_id, ref DataTable rowDT) 
    {
        String modeStr = "";
        DataRow[] dr;
        Chart chartObj;
        try
        {
            // 取得選擇的 ModeID
            for (int k = 0; k < lb_failModeShow.Items.Count; k++)
            {
                modeStr = lb_failModeShow.Items[k].Text;
                dr = rowDT.Select("Fail_Mode='" + modeStr + "'");
                if (dr.Length > 0) 
                {
                    chartObj = new Chart();
                    chartObj.ChartAreas.Add("Default");
                    chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
                    chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
                    chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
                    chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
                    chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
                    chartObj.ChartAreas["Default"].AxisX.TitleAlignment = StringAlignment.Near;
                    chartObj.Titles.Add("【" + part_id + " Stage:" + stage_id + " Mode:" + modeStr + "】");
                    chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
                    chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);

                    chartObj.ImageLocation = "temp/failModeDetail_#SEQ(1000,1)";
                    chartObj.ImageType = ChartImageType.Png;
                    chartObj.Palette = ChartColorPalette.None;
                    chartObj.Height = Unit.Pixel(500);
                    chartObj.Width = Unit.Pixel(1100);

                    chartObj.BackColor = Color.White;
                    chartObj.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;
                    chartObj.BorderStyle = BorderStyle.None;
                    chartObj.BorderWidth = 0;
                    chartObj.BorderColor = Color.DarkBlue;

                    String urlStr = "javascript:openWindowWithPost('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')";
                    urlStr = String.Format(urlStr, ddlCustomer.SelectedItem.Text, ddlCategory.SelectedItem.Text, (rbl_BySource.SelectedIndex), (ddlPart.SelectedItem.Text), (txtDateFrom.Text.Trim()), (txtDateTo.Text.Trim()), (rb_mode_selected.SelectedIndex), stage_id, modeStr, "ModeID");
                    chartObj.Titles[0].Url = urlStr;

                    createChart(ref chartObj, ref rowDT, modeStr, "Fail_Mode");
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + part_id + " Stage: " + stage_id + " Mode : " + modeStr + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    ChartPanel.Controls.Add(chartObj);
                    ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                }
            }
            tr_chartPanel.Visible = true;
        }
        catch (Exception ex) { showMessage("Error : " + ex.Message); }
    }

    // Defect Code Query
    private void defectCodeQuery(String part_id, String stage_id, ref DataTable rowDT) 
    {
        String modeStr = "";
        DataRow[] dr;
        Chart chartObj;

        try
        {
            // 取得選擇的 ModeID
            for (int k = 0; k < lb_dcodeShow.Items.Count; k++)
            {
                modeStr = lb_dcodeShow.Items[k].Text;
                dr = rowDT.Select("DefectCode='" + modeStr + "'");
                if (dr.Length > 0)
                {
                    chartObj = new Chart();
                    chartObj.ChartAreas.Add("Default");
                    chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
                    chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
                    chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
                    chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
                    chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
                    chartObj.ChartAreas["Default"].AxisX.TitleAlignment = StringAlignment.Near;
                    chartObj.Titles.Add("【" + part_id + " Stage:" + stage_id + " Mode:" + modeStr + "】");
                    chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
                    chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);

                    chartObj.ImageLocation = "temp/failModeDetail_#SEQ(1000,1)";
                    chartObj.ImageType = ChartImageType.Png;
                    chartObj.Palette = ChartColorPalette.None;
                    chartObj.Height = Unit.Pixel(500);
                    chartObj.Width = Unit.Pixel(1100);

                    chartObj.BackColor = Color.White;
                    chartObj.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;
                    chartObj.BorderStyle = BorderStyle.None;
                    chartObj.BorderWidth = 0;
                    chartObj.BorderColor = Color.DarkBlue;

                    String urlStr = "javascript:openWindowWithPost('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')";
                    urlStr = String.Format(urlStr, ddlCustomer.SelectedItem.Text, ddlCategory.SelectedItem.Text, (rbl_BySource.SelectedIndex), (ddlPart.SelectedItem.Text), (txtDateFrom.Text.Trim()), (txtDateTo.Text.Trim()), (rb_mode_selected.SelectedIndex), stage_id, modeStr, "ModeID");
                    chartObj.Titles[0].Url = urlStr;

                    createChart(ref chartObj, ref rowDT, modeStr, "DefectCode");
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + part_id + " Stage: " + stage_id + " Mode : " + modeStr + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                    ChartPanel.Controls.Add(chartObj);
                    ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                }
            }
            tr_chartPanel.Visible = true;
        }
        catch (Exception ex) { showMessage("Error : " + ex.Message); }
    }

    // PT AOI Query
    private void ptAOIQuery() 
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable RowDT = null;
        string modeStr = "";
        string sqlStr = "";
        ArrayList partAry = new ArrayList();
        Chart chartObj;

        try
        {
            conn.Open();
            if (rbl_BySource.SelectedIndex == 0)
            {
                sqlStr = "";
                sqlStr += "select Part_Id from Customer_Prodction_Mapping ";
                sqlStr += "where 1=1 ";
                sqlStr += "and customer_id='" + (ddlCustomer.SelectedValue) + "' ";
                sqlStr += "and category='" + (ddlCategory.SelectedValue) + "' ";
                sqlStr += "and production_id='" + (ddlPart.SelectedValue) + "' ";
                sqlStr += "and Fail_Function = 1 ";
                sqlStr += "group by Part_Id ";
                DataTable partDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(partDT);
                for (int i = 0; i < partDT.Rows.Count; i++)
                {
                    partAry.Add(partDT.Rows[i]["Part_Id"]);
                }
            }
            else 
            {
                partAry.Add(ddlPart.SelectedItem.Text);
            }

            // Mode List 
            string modeList = "";
            for (int i = 0; i < listB_ptaoi_display.Items.Count; i++)
            {
                modeList += "'" + (listB_ptaoi_display.Items[i].Value).Replace("'", "''") + "',";
            }
            modeList = modeList.Substring(0, (modeList.Length - 1));
          
            for (int j = 0; j < partAry.Count; j++)
            {
                sqlStr = "";
                sqlStr += "select a.Lot_ID, b.Group_id as ModeID, Convert(char(20), MAX(datatime), 120) as Trtm, Round((convert(float, SUM(count))/MAX(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from PT_AOI a, ParamGroup b ";
                sqlStr += "where 1=1 ";
                sqlStr += "and a.ModeID=b.Mode_ID ";
                sqlStr += "and a.LayerNo=b.Layer_ID ";
                sqlStr += "and b.STEP_ID='PT_AOI' ";
                sqlStr += "and a.Part='" + partAry[j].ToString() + "' ";
                sqlStr += "and a.LayerNO='" + (ddl_ptaoi_layer.SelectedItem.Text) + "' ";
                sqlStr += "and b.GROUP_ID in (" + modeList + ") ";
                sqlStr += "and a.datatime >= '" + (txtDateFrom.Text) + " 00:00:00' ";
                sqlStr += "and a.datatime <= '" + (txtDateTo.Text) + " 23:59:59' ";
                sqlStr += "group by a.Lot_ID , b.Group_id ";
                sqlStr += "order by Convert(char(20), MAX(a.datatime), 120) ";
                RowDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(RowDT);

                if (RowDT.Rows.Count > 0)
                {
                    // 取得選擇的 ModeID
                    for (int k = 0; k < (listB_ptaoi_display.Items.Count); k++)
                    {
                        modeStr = listB_ptaoi_display.Items[k].Text;
                        chartObj = new Chart();
                        chartObj.ChartAreas.Add("Default");
                        chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
                        chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
                        chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
                        chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
                        chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
                        chartObj.ChartAreas["Default"].AxisX.TitleAlignment = StringAlignment.Near;
                        chartObj.Titles.Add("【" + partAry[j].ToString() + " Layer:" + (ddl_ptaoi_layer.SelectedItem.Text) + " Mode:" + modeStr + "】");
                        chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
                        chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);

                        chartObj.Legends.Add("Legends");                                       //圖例集合
                        chartObj.Legends["Legends"].DockedToChartArea = "Default";             //顯示在圖表內
                        chartObj.Legends["Legends"].BackColor = Color.FromArgb(235, 235, 235); //背景色
                        chartObj.Legends["Legends"].BackHatchStyle = ChartHatchStyle.DarkDownwardDiagonal;
                        chartObj.Legends["Legends"].BorderWidth = 1;
                        chartObj.Legends["Legends"].BorderColor = Color.FromArgb(200, 200, 200);

                        chartObj.ImageLocation = "temp/failByLotDetail_#SEQ(1000,1)";
                        chartObj.ImageType = ChartImageType.Png;
                        chartObj.Palette = ChartColorPalette.None;
                        chartObj.Height = Unit.Pixel(500);
                        chartObj.Width = Unit.Pixel(1100);

                        chartObj.BackColor = Color.White;
                        chartObj.BorderSkin.SkinStyle = BorderSkinStyle.FrameThin2;
                        chartObj.BorderStyle = BorderStyle.None;
                        chartObj.BorderWidth = 0;
                        chartObj.BorderColor = Color.DarkBlue;

                        String urlStr = "javascript:openWindowWithPost('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')";
                        urlStr = String.Format(urlStr, ddlCustomer.SelectedItem.Text, ddlCategory.SelectedItem.Text, (rbl_BySource.SelectedIndex), (ddlPart.SelectedItem.Text), (txtDateFrom.Text.Trim()), (txtDateTo.Text.Trim()), (rb_mode_selected.SelectedIndex), (ddl_ptaoi_layer.SelectedItem.Text), modeStr, "ModeID");
                        chartObj.Titles[0].Url = urlStr;
                        if (createChart(ref chartObj, ref RowDT, modeStr, "ModeID")) 
                        {
                            ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + partAry[j].ToString() + " Mode : " + modeStr + "</td><td style='width:300px'></td></tr>"));
                            ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"));
                            ChartPanel.Controls.Add(chartObj);
                            ChartPanel.Controls.Add(new LiteralControl("</td></tr>"));
                        }
                    }
                }
                else 
                {
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>Part : " + partAry[j].ToString() + "</td><td style='width:300px'></td></tr>"));
                    ChartPanel.Controls.Add(new LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;'>"));
                    ChartPanel.Controls.Add(new LiteralControl("此料號無資料 !</td></tr>"));
                }
            }      
            tr_chartPanel.Visible = true;
            conn.Close();
        }
        catch (Exception ex) { showMessage("Error : " + ex.Message); }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // Charting
    private bool createChart(ref Chart chartObj, ref DataTable rowDT, String ModeStr, String ColName) 
    {
        // ModeID
        bool exeSuccess = false;
        DataRow[] RowDR;
        Series series = chartObj.Series.Add(ModeStr);
        series.ChartArea = "Default";
        series.ChartType = SeriesChartType.Line;
        series.Color = Color.DarkBlue;
        series.MarkerStyle = MarkerStyle.Circle;
        series.MarkerSize = 10;
        series.MarkerColor = Color.DarkBlue;
        series.BorderColor = Color.White;
        series.BorderWidth = 1;

        String ndataStr = "";
        String odataStr = "";
        Double pValue = 0;
        RowDR = rowDT.Select(ColName + "='" + ModeStr + "'", "Trtm");

        if (RowDR.Length > 0) 
        {
            exeSuccess = true;
            for (int i = 0; i < RowDR.Length; i++)
            {
                ndataStr = RowDR[i]["Trtm"].ToString();
                ndataStr = ndataStr.Substring(0, 10);
                pValue = 0;
                if (RowDR[i]["pValue"] != System.DBNull.Value)
                {
                    pValue = Convert.ToDouble(RowDR[i]["pValue"].ToString());
                    pValue = Math.Round(pValue, 4, MidpointRounding.AwayFromZero);
                    chartObj.Series[ModeStr].Points.AddXY(i, pValue);
                    chartObj.Series[ModeStr].Points[i].ToolTip = "Lot:" + RowDR[i]["Lot_ID"].ToString() + "\r\nDateTime:" + RowDR[i]["Trtm"].ToString() + "\r\nValue:" + pValue.ToString();
                }
                else
                {
                    chartObj.Series[ModeStr].Points.AddXY(i, pValue);
                    chartObj.Series[ModeStr].Points[i].ToolTip = "Lot:" + RowDR[i]["Lot_ID"].ToString() + "\r\nDateTime:" + RowDR[i]["Trtm"].ToString() + "\r\nValue:0";
                }
                chartObj.Series[ModeStr].Points[i].MarkerSize = 5;

                if (ndataStr == odataStr)
                {
                    chartObj.Series[ModeStr].Points[i].AxisLabel = " ";
                }
                else
                {
                    chartObj.Series[ModeStr].Points[i].AxisLabel = ndataStr;
                    odataStr = ndataStr;
                }
            }
        }
        return exeSuccess;
    }

    // Lot Query
    private void lotQuery() 
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable LotDT = null;
        string sqlStr = "";
        string conditionStr = "";

        string lotStr = "";
        for (int i = 0; i < lb_lotList.Items.Count; i++)
        {
                lotStr += "'" + lb_lotList.Items[i].Value + "',";
        }
        lotStr = lotStr.Substring(0, (lotStr.Length - 1));
        conditionStr += "and b.lot_id in (" + lotStr + ") ";

        try
        {
            conn.Open();
            sqlStr = "";
            sqlStr += "select a.Customer_Id as Customer, a.Category as 'CPU/CS', b.Production_Type AS Product_ID, ";
            sqlStr += "b.Part_ID, WW as Week, Convert(char(19), datatime, 120) as Time, ";
            sqlStr += "Lot_ID, DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage, Fail_Count as QTY, ";
            sqlStr += "ROUND(Fail_ratio, 2) as Ratio ";
            sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
            sqlStr += "where 1 = 1 ";
            sqlStr += "and a.Production_Id = b.Production_Type ";
            sqlStr += "and a.Part_Id = b.Part_Id ";
            sqlStr += conditionStr;
            sqlStr += "order by b.Lot_ID, b.MF_stage, b.BinCode ";

            LotDT = new DataTable();
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myAdapter.Fill(LotDT);
            if (LotDT.Rows.Count > 0)
            {
                GV_LotRowData.DataSource = LotDT;
                GV_LotRowData.DataBind();
                UtilObj.Set_DataGridRow_OnMouseOver_Color(ref GV_LotRowData, "#FFF68F", GV_LotRowData.AlternatingRowStyle.BackColor);
                lab_wait.Text = "Data Count : " + LotDT.Rows.Count;
            }
            else
            {
                lab_wait.Text = "No Data !";
            }
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    
    }

    // 上傳 Lot 
    protected void but_Uupload_Click(object sender, EventArgs e)
    {
        StreamReader file = default(StreamReader);
        string line = null;
        string saveFullName = Page.MapPath(".") + "\\upload\\";

        if ((uf_UfilePath.HasFile))
        {
            string fileName = uf_UfilePath.FileName;
            saveFullName += fileName;
            uf_UfilePath.SaveAs(saveFullName);
            FileInfo fileN = new FileInfo(saveFullName);
            if ((fileN.Extension).ToUpper() == ".TXT" || (fileN.Extension).ToUpper() == ".CSV")
            {
                try
                {
                    ArrayList lotAry = new ArrayList();
                    file = new StreamReader(saveFullName, System.Text.Encoding.Default);
                    line = file.ReadLine();
                    lb_lotList.Items.Clear();
                    while (line != null)
                    {
                        lb_lotList.Items.Add(line);
                        line = file.ReadLine();
                    }
                    file.Close();
                }
                catch (Exception ex){}
            }
            else
            {
                showMessage("請上傳 TXT or CSV !!!");
            }
        }
    }

    private void showMessage(String msgStr) 
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.Append("<script language='javascript'>");
        sb.Append("alert('" + msgStr + "');");
        sb.Append("</script>");
        ClientScriptManager myCSManager = Page.ClientScript;
        myCSManager.RegisterStartupScript(this.GetType(), "SetStatusScript", sb.ToString());
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        showMessage(e.ToString());
    }

    // Stage >>
    protected void but_stageTo_Click(object sender, EventArgs e)
    {
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        moveList(ref lb_StageSource, ref lb_StageShow);
    }
    
    // Stage <<
    protected void but_stageBack_Click(object sender, EventArgs e)
    {
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        moveList(ref lb_StageShow, ref lb_StageSource);
    }
    
    // Stage OK
    protected void but_stageOK_Click(object sender, EventArgs e)
    {
        lb_failModeSource.Items.Clear();
        lb_failModeShow.Items.Clear();
        lb_dcodeSource.Items.Clear();
        lb_dcodeShow.Items.Clear();
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string tableStr = " dbo.BinCode_FailMode_Customer_Mapping a, dbo.BinCode b ";
        string stageStr = "";
        if (lb_StageShow.Items.Count <= 0) 
        {
            showMessage("請選擇 Stage !");
            return;
        }
        try
        {
            for (int i = 0; i < lb_StageShow.Items.Count; i++ )
            {
                stageStr += "'" + (lb_StageShow.Items[i].Value).Replace("'", "''") + "',";
            }
            stageStr = stageStr.Substring(0, (stageStr.Length - 1));
            conn.Open();
            // === Fail Mode ===
            if (ddlCategory.SelectedValue == "CPU")
            {
                tableStr = " dbo.BinCode_FailMode_Customer_Mapping a, dbo.BinCode b ";
            }
            else
            {
                tableStr = " dbo.CS_BinCode_FailMode_Customer_Mapping a, dbo.CS_BinCode b ";
            }

            if (rb_mode_selected.SelectedIndex == 0) 
            {
                sqlStr = "select b.BinCode ";
                sqlStr += "from " + tableStr;
                sqlStr += "where 1=1 ";
                sqlStr += "and a.Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
                sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
                sqlStr += "and a.MF_Stage in (" + stageStr + ") ";
                sqlStr += "group by b.BinCode";
            }
            else if (rb_mode_selected.SelectedIndex == 1)
            {
                sqlStr = "select b.DefectCode_Id ";
                sqlStr += "from " + tableStr;
                sqlStr += "where 1=1 ";
                sqlStr += "and a.Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
                sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
                sqlStr += "and a.MF_Stage in (" + stageStr + ") ";
                sqlStr += "group by b.DefectCode_Id";
            }

            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);

            tr_failMode.Visible = false;
            tr_defectCode.Visible = false;
            if (MainDT.Rows.Count > 0)
            {
                if (rb_mode_selected.SelectedIndex == 0)
                {
                    tr_failMode.Visible = true;
                    UtilObj.FillLitsBoxController(MainDT, ref lb_failModeSource, 0);
                }
                else if (rb_mode_selected.SelectedIndex == 1)
                {
                    tr_defectCode.Visible = true;
                    UtilObj.FillLitsBoxController(MainDT, ref lb_dcodeSource, 0);
                }
            }
            else 
            {
                showMessage("沒有資料!");
            }
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }
    
    // Fail Mode >>
    protected void but_failTo_Click(object sender, EventArgs e)
    {
        tr_defectCode.Visible = false;
        moveList(ref lb_failModeSource, ref lb_failModeShow);
    }
    
    // Fail Mode <<
    protected void but_failBack_Click(object sender, EventArgs e)
    {
        tr_defectCode.Visible = false;
        moveList(ref lb_failModeShow, ref lb_failModeSource);
    }
    
    // Fail Mode OK
    protected void but_failModeOK_Click(object sender, EventArgs e)
    {
        if (lb_failModeShow.Items.Count <= 0)
        {
            moveListAll(ref lb_failModeSource, ref lb_failModeShow);
        }
        tr_execute.Visible = true;
    }
    
    // Defect Code >>
    protected void but_dcodeTo_Click(object sender, EventArgs e)
    {
        moveList(ref lb_dcodeSource, ref lb_dcodeShow);
    }
    
    // Defect Code <<
    protected void but_dcodeBack_Click(object sender, EventArgs e)
    {
        moveList(ref lb_dcodeShow, ref lb_dcodeSource);
    }
    
    // Defect Code OK
    protected void but_DefectCode_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string tableStr = " dbo.BinCode_FailMode_Customer_Mapping a, dbo.BinCode b ";
        string stageStr = "";
        string failModeStr = "";
        if (lb_failModeShow.Items.Count <= 0)
        {
            showMessage("請選擇 Fail Mode !");
            return;
        }

        try
        {
            for (int i = 0; i < lb_StageShow.Items.Count; i++)
            {
                stageStr += "'" + (lb_StageShow.Items[i].Value).Replace("'", "''") + "',";
            }
            stageStr = stageStr.Substring(0, (stageStr.Length - 1));
            for (int i = 0; i < lb_failModeShow.Items.Count; i++)
            {
                failModeStr += "'" + (lb_failModeShow.Items[i].Value).Replace("'", "''") + "',";
            }
            failModeStr = failModeStr.Substring(0, (failModeStr.Length - 1));
            conn.Open();
            // === Fail Mode ===
            if (ddlCategory.SelectedValue == "CPU")
            {
                tableStr = " dbo.BinCode_FailMode_Customer_Mapping a, dbo.BinCode b ";
            }
            else
            {
                tableStr = " dbo.CS_BinCode_FailMode_Customer_Mapping a, dbo.CS_BinCode b ";
            }
            sqlStr = "select b.DefectCode_Id ";
            sqlStr += "from " + tableStr;
            sqlStr += "where 1=1 ";
            sqlStr += "and a.Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
            sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
            sqlStr += "and a.MF_Stage in (" + stageStr + ") ";
            sqlStr += "and b.BinCode in (" + stageStr + ") ";
            sqlStr += "group by b.DefectCode_Id";
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);

            tr_failMode.Visible = false;
            tr_defectCode.Visible = false;
            if (MainDT.Rows.Count > 0)
            {
                tr_failMode.Visible = true;
                UtilObj.FillLitsBoxController(MainDT, ref lb_failModeSource, 0);
            }
            else
            {
                showMessage("沒有資料!");
            }
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    private void moveList(ref ListBox sourceList, ref ListBox destList)
    {
        ArrayList sourceAry = new ArrayList();
        ArrayList DestAry = new ArrayList();

        for (int i = 0; i < sourceList.Items.Count; i++)
        {
            if (sourceList.Items[i].Selected)
            {
                DestAry.Add(sourceList.Items[i].Value);
            }
            else
            {
                sourceAry.Add(sourceList.Items[i].Value);
            }
        }
        sourceList.Items.Clear();

        for (int i = 0; i < sourceAry.Count; i++)
        {
            sourceList.Items.Add(sourceAry[i].ToString());
        }
        for (int i = 0; i < DestAry.Count; i++)
        {
            destList.Items.Add(DestAry[i].ToString());
        }

    }
    private void moveListAll(ref ListBox sourceList, ref ListBox destList)
    {
        for (int i = 0; i < sourceList.Items.Count; i++)
        {
            destList.Items.Add(sourceList.Items[i].Value);
        }
        sourceList.Items.Clear();
    }

    protected void rb_mode_selected_SelectedIndexChanged(object sender, EventArgs e)
    {
        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_execute.Visible = false;
        tr_result.Visible = false;

        listB_ptaoi_source.Items.Clear();
        listB_ptaoi_display.Items.Clear();
        lb_StageSource.Items.Clear();
        lb_StageShow.Items.Clear();
        lb_failModeSource.Items.Clear();
        lb_failModeShow.Items.Clear();
        lb_dcodeSource.Items.Clear();
        lb_dcodeShow.Items.Clear();

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        string sqlStr = "";
        string partStr = " ";

        if ((rb_mode_selected.SelectedIndex == 0) || (rb_mode_selected.SelectedIndex == 1)) 
        {// Fail Mode / Defect Code
            tr_Stage.Visible = true;
            if (ddlCategory.SelectedValue == "CPU")
            {
                sqlStr = "select MF_Stage ";
                sqlStr += "from BinCode_FailMode_Customer_Mapping ";
                sqlStr += "where 1=1 ";
                sqlStr += "and Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
                sqlStr += "group by MF_Stage";
            }
            else
            {
                sqlStr = "select MF_Stage ";
                sqlStr += "from CS_BinCode_FailMode_Customer_Mapping ";
                sqlStr += "where 1=1 ";
                sqlStr += "and Customer_Id = '" + (ddlCustomer.SelectedValue) + "' ";
                sqlStr += "group by MF_Stage";
            }
        }
        else if (rb_mode_selected.SelectedIndex == 2) 
        {// PT AOI
            tr_ptaoi.Visible = true;
            sqlStr = "Select Layer_ID ";
            sqlStr += "FROM ParamGroup ";
            sqlStr += "WHERE STEP_ID='PT_AOI' ";
            sqlStr += "GROUP BY Layer_ID ";
            sqlStr += "ORDER BY Layer_ID";

            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            UtilObj.FillController(MainDT, ref ddl_ptaoi_layer, 0);

            sqlStr = "Select Group_ID ";
            sqlStr += "FROM ParamGroup ";
            sqlStr += "WHERE Layer_ID='" + ddl_ptaoi_layer.SelectedValue + "' ";
            sqlStr += "AND STEP_ID='PT_AOI' ";
            sqlStr += "GROUP BY Group_ID ";
            sqlStr += "ORDER BY Group_ID";
        }

        try
        {
            conn.Open();
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);
            if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
            {// Fail Mode
                tr_Stage.Visible = true;
                UtilObj.FillLitsBoxController(MainDT, ref lb_StageSource, 0);
            }
            else if (rb_mode_selected.SelectedIndex == 2)
            {// PT AOI
                tr_ptaoi.Visible = true;
                UtilObj.FillLitsBoxController(MainDT, ref listB_ptaoi_source, 0);
            }
            conn.Close();
        }
        catch (Exception ex) { }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }
    protected void ddl_ptaoi_layer_SelectedIndexChanged(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter dtAdp;
        DataTable dt;
        String sqlStr = "";

        listB_ptaoi_source.Items.Clear();
        listB_ptaoi_display.Items.Clear();
        
        try
        {
            sqlStr = "Select Group_ID ";
            sqlStr += "FROM ParamGroup ";
            sqlStr += "WHERE Layer_ID='" + ddl_ptaoi_layer.SelectedValue + "' ";
            sqlStr += "AND STEP_ID='PT_AOI' ";
            sqlStr += "GROUP BY Group_ID ";
            sqlStr += "ORDER BY Group_ID";

            conn.Open();
            dtAdp = new SqlDataAdapter(sqlStr, conn);
            dt = new DataTable();
            dtAdp.Fill(dt);
            UtilObj.FillLitsBoxController(dt, ref listB_ptaoi_source, 0);
            conn.Close();
        }
        catch (Exception error)
        {
            throw error;
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    // >> PT AOI 
    protected void but_ptaoi_right_Click(object sender, EventArgs e)
    {
        moveList(ref listB_ptaoi_source, ref listB_ptaoi_display);
    }

    // << PT AOI
    protected void but_ptaoi_left_Click(object sender, EventArgs e)
    {
        moveList(ref listB_ptaoi_display, ref listB_ptaoi_source);
    }

    // PT AOI OK
    protected void but_ptaoi_decision_Click(object sender, EventArgs e)
    {
        if (listB_ptaoi_display.Items.Count <= 0)
        {
            moveListAll(ref listB_ptaoi_source, ref listB_ptaoi_display);
        }
        tr_execute.Visible = true;
    }

    // Defect Code OK
    protected void but_DefectCodeOK_Click(object sender, EventArgs e)
    {
        if (lb_dcodeShow.Items.Count <= 0)
        {
            moveListAll(ref lb_dcodeSource, ref lb_dcodeShow);
        }
        tr_execute.Visible = true;
    }

    protected void ddlPart_SelectedIndexChanged(object sender, EventArgs e)
    {
        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;
        tr_execute.Visible = false;
        tr_result.Visible = false;
        if (rb_mode_selected.SelectedIndex > -1)
        {
            rb_mode_selected.SelectedItem.Selected = false;
        }
    }

    protected void txtDateFrom_TextChanged(object sender, EventArgs e)
    {

    }
}