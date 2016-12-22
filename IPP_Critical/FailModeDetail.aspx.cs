using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Collections;
using System.Web.UI.DataVisualization.Charting;
using System.Drawing;

public partial class FailModeDetail : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        but_Execute.Attributes.Add("onclick", "javascript:document.getElementById(\"lab_wait\").innerText='Please wait ......';" +
                                   "javascript:document.getElementById(\"but_Execute\").disabled=true;" +
                                   Page.GetPostBackEventReference(but_Execute));
        if (!IsPostBack) 
        {
            try {
                pageInit();
            } catch(Exception error){}
        }
    }

    private void pageInit() 
    {
        // Test 
        //this.lab_custom.Text = "INTEL";
        //this.lab_product.Text = "CPU";
        //this.lab_productType.Text = "1";
        //this.lab_partID.Text = "FCS071A";
        //this.lab_dateF.Text = "2013-03-01";
        //this.lab_dateT.Text = "2013-03-19";
        //this.lab_modeType.Text = "2";
        //this.lab_PA.Text = "04";
        //this.lab_PB.Text = "51Other";
        //this.lab_PC.Text = "ModeID";

        this.lab_custom.Text = Request["customerID"].ToString();       // INTEL
        this.lab_product.Text = Request["product"].ToString();         // CPU / CS
        this.lab_productType.Text = Request["productType"].ToString(); // 0. Product ID, 1. Part ID
        this.lab_partID.Text = Request["partID"].ToString();           // FCS071A
        this.lab_dateF.Text = Request["dateF"].ToString();             // yyyy-MM-dd
        this.lab_dateT.Text = Request["dateT"].ToString();             // yyyy-MM-dd
        this.lab_modeType.Text = Request["modeType"].ToString();       // 0. FailMode, 1. DefectCode, 2. PT_AOI
        this.lab_PA.Text = Request["PA"].ToString();                   // Fail Mode -- Stage, PT AOI -- Layer
        this.lab_PB.Text = Request["PB"].ToString();                   // Fail Mode -- Fail / Defect, PT AOI -- Mode
        this.lab_PC.Text = Request["PC"].ToString();                   // Colume Name [Fail -- fail_mode / DefectCode, PT AOI -- ModeID]

        rb_mode_selected.Items[0].Enabled = false;
        rb_mode_selected.Items[1].Enabled = false;
        rb_mode_selected.Items[2].Enabled = false;

        if ((this.lab_modeType.Text == "0")) 
        {
            rb_mode_selected.Items[0].Enabled = true;
            failModeMain(false);
        }
        else if ((this.lab_modeType.Text == "1"))
        {
            rb_mode_selected.Items[1].Enabled = true;
            failModeMain(false);
        }
        else if ((this.lab_modeType.Text == "2")) 
        {
            rb_mode_selected.Items[2].Enabled = true;
            ptAOIQuery(false);
        }
    }

    #region Tool

    private void showMessage(String msgStr)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.Append("<script language='javascript'>");
        sb.Append("alert('" + msgStr + "');");
        sb.Append("</script>");
        ClientScriptManager myCSManager = Page.ClientScript;
        myCSManager.RegisterStartupScript(this.GetType(), "SetStatusScript", sb.ToString());
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
    // >> 
    protected void but_right_Click(object sender, EventArgs e)
    {
        Button butObj = (Button)sender;
        if (butObj.ID == "but_ptaoi_right")
        {
            moveList(ref listB_ptaoi_source, ref listB_ptaoi_display);
        }
        else if (butObj.ID == "but_stageTo")
        {
            moveList(ref lb_StageSource, ref lb_StageShow);
        }
        else if (butObj.ID == "but_failTo")
        {
            moveList(ref lb_failModeSource, ref lb_failModeShow);
        }
        else if (butObj.ID == "but_dcodeTo")
        {
            moveList(ref lb_dcodeSource, ref lb_dcodeShow);
        }
    }
    // <<
    protected void but_left_Click(object sender, EventArgs e)
    {
        Button butObj = (Button)sender;
        if (butObj.ID == "but_ptaoi_left")
        {
            moveList(ref listB_ptaoi_display, ref listB_ptaoi_source);
        }
        else if (butObj.ID == "but_stageBack")
        {
            moveList(ref lb_StageShow, ref lb_StageSource);
        }
        else if (butObj.ID == "but_failBack")
        {
            moveList(ref lb_failModeShow, ref lb_failModeSource);
        }
        else if (butObj.ID == "but_dcodeBack")
        {
            moveList(ref lb_dcodeShow, ref lb_dcodeSource);
        }
    }

    #endregion

    // Mode Select
    protected void rb_mode_selected_SelectedIndexChanged(object sender, EventArgs e)
    {
        modeSelect();
    }
    private void modeSelect() 
    {
        tr_ptaoi.Visible = false;
        tr_Stage.Visible = false;
        tr_failMode.Visible = false;
        tr_defectCode.Visible = false;

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

        if ((rb_mode_selected.SelectedIndex == 0) || (rb_mode_selected.SelectedIndex == 1))
        {// Fail Mode / Defect Code
            tr_Stage.Visible = true;
            if (lab_product.Text == "CPU")
            {
                sqlStr = "select MF_Stage ";
                sqlStr += "from BinCode_FailMode_Customer_Mapping ";
                sqlStr += "where 1=1 ";
                sqlStr += "and Customer_Id = '" + (lab_custom.Text) + "' ";
                sqlStr += "group by MF_Stage";
            }
            else
            {
                sqlStr = "select MF_Stage ";
                sqlStr += "from CS_BinCode_FailMode_Customer_Mapping ";
                sqlStr += "where 1=1 ";
                sqlStr += "and Customer_Id = '" + (lab_custom.Text) + "' ";
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
            for (int i = 0; i < ddl_ptaoi_layer.Items.Count; i++)
            {
                if ((ddl_ptaoi_layer.Items[i].Value).Trim() == (lab_PA.Text)) 
                {
                    ddl_ptaoi_layer.SelectedIndex = i;
                }
            }
            ddl_ptaoi_layer.Enabled = false;

            sqlStr = "Select Group_ID as ModeID ";
            sqlStr += "FROM ParamGroup ";
            sqlStr += "WHERE Layer_ID='" + (ddl_ptaoi_layer.SelectedValue) + "' ";
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

    // Layer Select
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
            sqlStr = "Select GROUP_ID as ModeID ";
            sqlStr += "FROM ParamGroup ";
            sqlStr += "WHERE Layer_ID='" + ddl_ptaoi_layer.SelectedValue + "' ";
            sqlStr += "AND STEP_ID='PT_AOI' ";
            sqlStr += "GROUP BY GROUP_ID ";
            sqlStr += "ORDER BY GROUP_ID";

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
    
    // PT AOI OK
    protected void but_ptaoi_decision_Click(object sender, EventArgs e)
    {
        if (listB_ptaoi_display.Items.Count <= 0)
        {
            moveListAll(ref listB_ptaoi_source, ref listB_ptaoi_display);
        }
        tr_execute.Visible = true;
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
            for (int i = 0; i < lb_StageShow.Items.Count; i++)
            {
                stageStr += "'" + (lb_StageShow.Items[i].Value).Replace("'", "''") + "',";
            }
            stageStr = stageStr.Substring(0, (stageStr.Length - 1));
            conn.Open();
            // === Fail Mode ===
            if (lab_product.Text == "CPU")
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
                sqlStr += "and a.Customer_Id = '" + (lab_custom.Text) + "' ";
                sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
                sqlStr += "and a.MF_Stage IN (" + stageStr + ") ";
                sqlStr += "group by b.BinCode";
            }
            else if (rb_mode_selected.SelectedIndex == 1)
            {
                sqlStr = "select b.DefectCode_Id ";
                sqlStr += "from " + tableStr;
                sqlStr += "where 1=1 ";
                sqlStr += "and a.Customer_Id = '" + (lab_custom.Text) + "' ";
                sqlStr += "and a.BinCode_Id = b.BinCode_Id ";
                sqlStr += "and a.MF_Stage IN (" + stageStr + ") ";
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

    // Fail Mode OK
    protected void but_failModeOK_Click(object sender, EventArgs e)
    {
        if (lb_failModeShow.Items.Count <= 0)
        {
            moveListAll(ref lb_failModeSource, ref lb_failModeShow);
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

    // Query
    protected void but_Execute_Click(object sender, EventArgs e)
    {
        if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
        {
            failModeMain(true);
        }
        else
        {
            ptAOIQuery(true);
        }
    }

    // Fail Mode Main 
    private void failModeMain(Boolean isResend)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        string sqlStr = "";
        string conditionStr = "";
        ArrayList partAry = new ArrayList();
        DataTable RowDT = null;

        if (this.lab_productType.Text == "0")
        {// Product
            conditionStr += "and b.production_type='" + (lab_partID.Text) + "' ";
        }
        else
        {
            conditionStr += "and b.part_id='" + (lab_partID.Text) + "' ";
        }
        // DataTime
        conditionStr += "and b.datatime >= '" + (lab_dateF.Text) + " 00:00:00' ";
        conditionStr += "and b.datatime <= '" + (lab_dateT.Text) + " 23:59:59' ";

        try
        {
            conn.Open();
            sqlStr = "";
            if (lab_modeType.Text == "0")
            {
                sqlStr = "select b.MF_Stage, b.Fail_Mode, b.Lot_Id, Convert(char(20), MAX(b.datatime), 120) as trtm, ";
                sqlStr += "SUM(Fail_Count) as QTY, Max(Original_Input_Qty) as Original_Input_Qty, Round((convert(float, SUM(Fail_Count))/Max(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
                sqlStr += "where 1=1 ";
                sqlStr += "and a.customer_id=b.customer_id ";
                sqlStr += "and a.Production_Id=b.Production_Type ";
                sqlStr += "and a.Part_Id=b.Part_Id ";
                sqlStr += "and b.customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and b.category='" + (lab_product.Text) + "' ";
                sqlStr += "and b.MF_Stage='" + (lab_PA.Text) + "' ";
                sqlStr += "and b.Fail_Mode='" + (lab_PB.Text) + "' ";
                sqlStr += conditionStr;
                sqlStr += "group by b.MF_Stage, b.Fail_Mode, Lot_Id ";
                sqlStr += "order by Convert(char(20), Max(b.datatime), 120)";
                RowDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(RowDT);
                // 依 Fail Mode
                failModeQuery(ref RowDT, isResend);
            }
            else if (lab_modeType.Text == "1")
            {
                sqlStr = "select b.MF_Stage, b.DefectCode, b.Lot_Id, Convert(char(20), MAX(b.datatime), 120) as trtm, ";
                sqlStr += "SUM(Fail_Count) as QTY, Max(Original_Input_Qty) as Original_Input_Qty, Round((convert(float, SUM(Fail_Count))/Max(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
                sqlStr += "where 1 = 1 ";
                sqlStr += "and a.customer_id=b.customer_id ";
                sqlStr += "and a.Production_Id=b.Production_Type ";
                sqlStr += "and a.Part_Id=b.Part_Id ";
                sqlStr += "and b.customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and b.category='" + (lab_product.Text) + "' ";
                sqlStr += "and b.MF_Stage='" + (lab_PA.Text) + "' ";
                sqlStr += "and b.DefectCode='" + (lab_PB.Text) + "' ";
                sqlStr += conditionStr;
                sqlStr += "group by b.MF_Stage, b.DefectCode, Lot_Id ";
                sqlStr += "order by Convert(char(20), Max(b.datatime), 120)";
                RowDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(RowDT);
                // 依 Defect Code
                defectCodeQuery(ref RowDT, isResend);
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
    private void failModeQuery(ref DataTable rowDT, Boolean isResend)
    {
        String modeStr = "";
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable ResendDT = new DataTable();

        ArrayList paramAry = new ArrayList();
        if (rb_mode_selected.SelectedIndex != -1)
        {
            for (int i = 0; i < lb_failModeShow.Items.Count; i++)
            {
                if ((lb_failModeShow.Items[i].Value) != (lab_PB.Text))
                {
                    paramAry.Add(lb_failModeShow.Items[i].Value);
                }
            }
        }

        try
        {
            if (isResend)
            {
                // Fail Mode 
                string failModeStr = "";
                for (int i = 0; i < lb_failModeShow.Items.Count; i++)
                {
                    failModeStr += "'" + (lb_failModeShow.Items[i].Value).Replace("'", "''") + "',";
                }
                failModeStr = failModeStr.Substring(0, (failModeStr.Length - 1));

                String sqlStr = "";
                sqlStr += "select MF_Stage, Fail_Mode, Lot_ID, SUM(Fail_Count) as QTY, Round((convert(float, SUM(Fail_Count))/MAX(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from BinCode_Daily_Lot  ";
                sqlStr += "where 1=1 ";
                sqlStr += "and customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and category='" + (lab_product.Text) + "' ";
                sqlStr += "and Fail_Mode IN (" + failModeStr + ") ";
                sqlStr += "and Lot_ID IN (";
                sqlStr += "select Lot_ID ";
                sqlStr += "from BinCode_Daily_Lot ";
                sqlStr += "where 1=1 ";
                sqlStr += "and customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and category='" + (lab_product.Text) + "' ";
                sqlStr += "and MF_Stage='" + (lab_PA.Text) + "' ";
                sqlStr += "and Fail_Mode='" + (lab_PB.Text) + "' ";
                sqlStr += "and datatime >= '" + (lab_dateF.Text) + " 00:00:00' ";
                sqlStr += "and datatime <= '" + (lab_dateT.Text) + " 23:59:59' ";
                sqlStr += "group by Lot_ID) ";
                sqlStr += "group by MF_Stage, Fail_Mode, Lot_ID ";
                conn.Open();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(ResendDT);
                conn.Close();
            }

            // 取得選擇的 ModeID
            modeStr = (lab_PB.Text);
            chartObj.ChartAreas.Clear();
            chartObj.Legends.Clear();
            chartObj.Titles.Clear();
            chartObj.ChartAreas.Add("Default");
            chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
            chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
            chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
            chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
            chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
            chartObj.ChartAreas["Default"].AxisX.TitleAlignment = System.Drawing.StringAlignment.Near;
            chartObj.Titles.Add("【" + (lab_partID.Text) + " Stage:" + (lab_PA.Text) + " Mode:" + (lab_PB.Text) + "】");
            chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
            chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);
            // --- Legends ---
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

            // --- Chart ---
            lab_chartTitle.Text = "(Part : " + (lab_partID.Text) + " Stage : " + (lab_PA.Text) + " Mode : " + modeStr + ")";
            createChart(ref chartObj, ref rowDT, "MF_Stage", "Fail_Mode", ref paramAry, isResend, ref ResendDT);
            // --- DataGrid View ---
            tr_chartPanel.Visible = true;
            bindGridView(ref rowDT, "MF_Stage", "Fail_Mode", ref paramAry, isResend, ref ResendDT);

        }
        catch (Exception ex) { showMessage("Error : " + ex.Message); }
        finally 
        { 
          if(conn.State == ConnectionState.Open)
          {
              conn.Close();
          }
        }
    }

    // Defect Code Query
    private void defectCodeQuery(ref DataTable rowDT, Boolean isResend)
    {
        String modeStr = "";
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable ResendDT = new DataTable();

        ArrayList paramAry = new ArrayList();
        if (rb_mode_selected.SelectedIndex != -1)
        {
            for (int i = 0; i < lb_dcodeShow.Items.Count; i++)
            {
                if ((lb_dcodeShow.Items[i].Value) != (lab_PB.Text))
                {
                    paramAry.Add(lb_dcodeShow.Items[i].Value);
                }
            }
        }

        try
        {
            if (isResend)
            {
                // Defect Code
                string failModeStr = "";
                for (int i = 0; i < lb_dcodeShow.Items.Count; i++)
                {
                    failModeStr += "'" + (lb_dcodeShow.Items[i].Value).Replace("'", "''") + "',";
                }
                failModeStr = failModeStr.Substring(0, (failModeStr.Length - 1));

                String sqlStr = "";
                sqlStr += "select MF_Stage, DefectCode, Lot_ID, SUM(Fail_Count) as QTY, Round((convert(float, SUM(Fail_Count))/MAX(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from BinCode_Daily_Lot  ";
                sqlStr += "where 1=1 ";
                sqlStr += "and customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and category='" + (lab_product.Text) + "' ";
                sqlStr += "and DefectCode IN (" + failModeStr + ") ";
                sqlStr += "and Lot_ID IN (";
                sqlStr += "select Lot_ID ";
                sqlStr += "from BinCode_Daily_Lot ";
                sqlStr += "where 1=1 ";
                sqlStr += "and customer_id='" + (lab_custom.Text) + "' ";
                sqlStr += "and category='" + (lab_product.Text) + "' ";
                sqlStr += "and MF_Stage='" + (lab_PA.Text) + "' ";
                sqlStr += "and DefectCode='" + (lab_PB.Text) + "' ";
                sqlStr += "and datatime >= '" + (lab_dateF.Text) + " 00:00:00' ";
                sqlStr += "and datatime <= '" + (lab_dateT.Text) + " 23:59:59' ";
                sqlStr += "group by Lot_ID) ";
                sqlStr += "group by MF_Stage, DefectCode, Lot_ID ";
                conn.Open();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(ResendDT);
                conn.Close();
            }

            // 取得選擇的 ModeID
            modeStr = (lab_PB.Text);
            chartObj.ChartAreas.Clear();
            chartObj.Legends.Clear();
            chartObj.Titles.Clear();
            chartObj.ChartAreas.Add("Default");
            chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
            chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
            chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
            chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
            chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
            chartObj.ChartAreas["Default"].AxisX.TitleAlignment = System.Drawing.StringAlignment.Near;
            chartObj.Titles.Add("【" + (lab_partID.Text) + " Stage:" + (lab_PA.Text) + " Mode:" + (lab_PB.Text) + "】");
            chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
            chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);
            // --- Legends ---
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

            // --- Chart ---
            lab_chartTitle.Text = "(Part : " + (lab_partID.Text) + " Stage : " + (lab_PA.Text) + " Mode : " + modeStr + ")";
            createChart(ref chartObj, ref rowDT, "MF_Stage", "DefectCode", ref paramAry, isResend, ref ResendDT);
            // --- DataGrid View ---
            tr_chartPanel.Visible = true;
            bindGridView(ref rowDT, "MF_Stage", "DefectCode", ref paramAry, isResend, ref ResendDT);
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

    // PT AOI Query
    private void ptAOIQuery(Boolean isResend)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable LotDT = null;
        DataTable RowDT = null;
        DataTable ResendDT = new DataTable();
        string modeStr = "";
        string sqlStr = "";
        
        ArrayList paramAry = new ArrayList();
        if (rb_mode_selected.SelectedIndex != -1)
        {
            for (int i = 0; i < listB_ptaoi_display.Items.Count; i++)
            {
                if ((listB_ptaoi_display.Items[i].Value) != lab_PB.Text)
                {
                    paramAry.Add(listB_ptaoi_display.Items[i].Value);
                }
                else
                {
                    if (ddl_ptaoi_layer.SelectedValue.ToString() != lab_PA.Text)
                    {
                        paramAry.Add(listB_ptaoi_display.Items[i].Value);
                    }
                }
            }
        }

        try
        {
            conn.Open();
            sqlStr = "select a.Lot_ID, b.Group_id as ModeID, Convert(char(20), MAX(datatime), 120) as Trtm, ";
            sqlStr += "SUM(count) as QTY, MAX(Original_Input_Qty) AS Original_Input_Qty, Round((convert(float, SUM(count))/MAX(Original_Input_Qty)), 6) as pvalue ";
            sqlStr += "from PT_AOI a, ParamGroup b ";
            sqlStr += "where 1=1 ";
            sqlStr += "and a.ModeID=b.Mode_ID ";
            sqlStr += "and a.LayerNo=b.Layer_ID ";
            sqlStr += "and b.STEP_ID='PT_AOI' ";
            sqlStr += "and a.Part='" + (lab_partID.Text) + "' ";
            sqlStr += "and a.LayerNO='" + (lab_PA.Text) + "' ";
            sqlStr += "and b.Group_ID='" + (lab_PB.Text) + "' ";
            sqlStr += "and a.datatime >= '" + (lab_dateF.Text) + " 00:00:00' ";
            sqlStr += "and a.datatime <= '" + (lab_dateT.Text) + " 23:59:59' ";
            sqlStr += "group by a.Lot_ID , b.Group_id ";
            sqlStr += "order by Convert(char(20), MAX(a.datatime), 120) ";
            RowDT = new DataTable();
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            myAdapter.Fill(RowDT);

            if (isResend) 
            {
                String modeList = "";
                for (int i = 0; i < listB_ptaoi_display.Items.Count; i++)
                {
                    modeList += "'" + ((listB_ptaoi_display.Items[i].Value).Replace("'", "''")) + "',";
                }
                modeList = modeList.Substring(0, (modeList.Length - 1));
                sqlStr = "";
                sqlStr += "select a.LayerNO, b.Group_ID as ModeID, a.Lot_ID, SUM(Count) as QTY, Round((convert(float, SUM(count))/MAX(Original_Input_Qty)), 6) as pvalue ";
                sqlStr += "from PT_AOI a, ParamGroup b  ";
                sqlStr += "where 1=1 ";
                sqlStr += "and a.LayerNo=b.Layer_ID ";
                sqlStr += "and a.ModeID=b.Mode_ID ";
                sqlStr += "and b.STEP_ID='PT_AOI' ";
                sqlStr += "and a.Part='" + ((lab_partID.Text).ToUpper()) + "' ";
                sqlStr += "and a.LayerNO='" + ((ddl_ptaoi_layer.SelectedValue.ToString()).ToUpper()) + "' ";
                sqlStr += "and b.Group_ID IN (" + modeList + ") ";
                sqlStr += "and a.Lot_ID IN (";
                sqlStr += "select Lot_ID ";
                sqlStr += "from PT_AOI  ";
                sqlStr += "where 1=1 ";
                sqlStr += "and Part='" + (lab_partID.Text) + "' ";
                sqlStr += "and LayerNO='" + (lab_PA.Text) + "' ";
                sqlStr += "and datatime >= '" + (lab_dateF.Text) + " 00:00:00' ";
                sqlStr += "and datatime <= '" + (lab_dateT.Text) + " 23:59:59' ";
                sqlStr += "group by Lot_ID) ";
                sqlStr += "group by a.LayerNO, b.Group_ID, a.Lot_ID ";
                ResendDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(ResendDT);
            }
            conn.Close();

            if (RowDT.Rows.Count > 0)
            {
                modeStr = (lab_PB.Text);
                chartObj.ChartAreas.Clear();
                chartObj.Legends.Clear();
                chartObj.Titles.Clear();
                chartObj.ChartAreas.Add("Default");
                chartObj.ChartAreas["Default"].AxisX.LabelStyle.Interval = 0.5;
                chartObj.ChartAreas["Default"].AxisX.LabelStyle.Angle = -90;
                chartObj.ChartAreas["Default"].BorderDashStyle = ChartDashStyle.NotSet;
                chartObj.ChartAreas["Default"].AxisX.MajorGrid.Enabled = false;
                chartObj.ChartAreas["Default"].AxisY.LabelStyle.Format = "P2";
                chartObj.ChartAreas["Default"].AxisX.TitleAlignment = StringAlignment.Near;
                chartObj.Titles.Add("【" + (lab_partID.Text) + " Layer:" + (lab_PA.Text) + " Mode:" + (lab_PB.Text) + "】");
                chartObj.Titles[0].Font = new Font("Arial", 8, FontStyle.Regular);
                chartObj.ChartAreas["Default"].AxisX.TitleFont = new Font("Arial", 6, FontStyle.Regular);
                // --- Legends ---
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
                
                // --- Chart ---
                lab_chartTitle.Text = "(Part : " + (lab_partID.Text) + " Layer : " + (lab_PA.Text) + " Mode : " + modeStr + ")";
                createChart(ref chartObj, ref RowDT, "LayerNO", "ModeID", ref paramAry, isResend, ref ResendDT);
                // --- DataGrid View ---
                tr_chartPanel.Visible = true;
                bindGridView(ref RowDT, "LayerNO", "ModeID", ref paramAry, isResend, ref ResendDT);
            }
            else
            {
                lab_chartTitle.Text = "(此料號無資料 !)";
            }
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
    private void createChart(ref Chart chartObj, ref DataTable rowDT, String Col1Name, String Col2Name, ref ArrayList paramAry, Boolean isResend, ref DataTable resendDT)
    {
        DataRow[] RowDR;
        Series series;
        String ndataStr = "";
        String odataStr = "";
        Double pValue = 0;
        String oriMode = (lab_PB.Text);
        String Stage = (lab_PA.Text);
        String ModeStr = (lab_PB.Text);
        String seriesName = "";
        String LotID, TimeStr;

        // --- Step1. 先加入原本的參數 ---
        seriesName = Stage + "[" + ModeStr + "]";
        chartObj.Series.Clear();
        series = chartObj.Series.Add(seriesName);
        series.ChartArea = "Default";
        series.ChartType = SeriesChartType.Line;
        series.Color = Color.DarkBlue;
        series.MarkerStyle = MarkerStyle.Circle;
        series.MarkerSize = 8;
        series.MarkerColor = Color.DarkBlue;
        series.BorderColor = Color.White;
        series.BorderWidth = 1;
        for (int i = 0; i < rowDT.Rows.Count; i++)
        {
            LotID = rowDT.Rows[i]["Lot_ID"].ToString().Trim();
            TimeStr = rowDT.Rows[i]["Trtm"].ToString().Trim();
            ndataStr = TimeStr;
            ndataStr = ndataStr.Substring(0, 10);
            pValue = 0;
            if (rowDT.Rows[i]["pValue"] != System.DBNull.Value)
            {
                pValue = Convert.ToDouble(rowDT.Rows[i]["pValue"].ToString());
                pValue = Math.Round(pValue, 4, MidpointRounding.AwayFromZero);
                chartObj.Series[seriesName].Points.AddXY(i, pValue);
                chartObj.Series[seriesName].Points[i].ToolTip = "Lot:" + LotID + "\r\nDateTime:" + TimeStr + "\r\nValue:" + pValue.ToString();
            }
            else
            {
                chartObj.Series[seriesName].Points.AddXY(i, pValue);
                chartObj.Series[seriesName].Points[i].ToolTip = "Lot:" + rowDT.Rows[i]["Lot_ID"].ToString() + "\r\nDateTime:" + rowDT.Rows[i]["Trtm"].ToString() + "\r\nValue:0";
            }
            chartObj.Series[seriesName].Points[i].Url = "javascript:LinkPoint('" + ((LotID + TimeStr)) + "');";

            if (ndataStr == odataStr)
            {
                chartObj.Series[seriesName].Points[i].AxisLabel = " ";
            }
            else
            {
                chartObj.Series[seriesName].Points[i].AxisLabel = ndataStr;
                odataStr = ndataStr;
            }
        }
        chartObj.Series[seriesName].IsVisibleInLegend = true;
        chartObj.Series[seriesName].LegendText = (lab_PB.Text);

        if (isResend) 
        {
            String chartTitle = (lab_PB.Text) + ",";
            Stage = "";
            // --- Step2. 加入其他參數 ---
            for (int a = 0; a < paramAry.Count; a++)
            {
                ModeStr = paramAry[a].ToString();
                chartTitle += (ModeStr + ",");
                if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
                {
                    Stage = "";
                    seriesName = "[" + ModeStr + "]";
                }
                else
                {
                    Stage = (ddl_ptaoi_layer.SelectedValue.ToString());
                    seriesName = Stage + "[" + ModeStr + "]";
                }
                
                series = chartObj.Series.Add(seriesName);
                series.ChartArea = "Default";
                series.ChartType = SeriesChartType.Line;
                series.Color = Color.DarkGreen;
                series.MarkerStyle = MarkerStyle.Circle;
                series.MarkerSize = 8;
                series.MarkerColor = Color.DarkGreen;
                series.BorderColor = Color.White;
                series.BorderWidth = 1;
                for (int i = 0; i < rowDT.Rows.Count; i++)
                {
                    LotID = rowDT.Rows[i]["Lot_ID"].ToString().Trim();
                    TimeStr = rowDT.Rows[i]["Trtm"].ToString().Trim();
                    if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
                    {
                        RowDR = resendDT.Select("Lot_ID='" + LotID + "' and " + Col2Name + "='" + ModeStr + "'");
                    }
                    else 
                    {
                        RowDR = resendDT.Select("Lot_ID='" + LotID + "' and " + Col1Name + "='" + Stage + "' and " + Col2Name + "='" + ModeStr + "'");
                    }
                    ndataStr = TimeStr;
                    ndataStr = ndataStr.Substring(0, 10);
                    pValue = 0;
                    if (RowDR.Length > 0)
                    {
                        pValue = Convert.ToDouble(RowDR[0]["pValue"].ToString());
                        pValue = Math.Round(pValue, 4, MidpointRounding.AwayFromZero);
                        chartObj.Series[seriesName].Points.AddXY(i, pValue);
                        chartObj.Series[seriesName].Points[i].ToolTip = "Lot:" + rowDT.Rows[i]["Lot_ID"].ToString() + "\r\nDateTime:" + rowDT.Rows[i]["Trtm"].ToString() + "\r\nValue:" + pValue.ToString();
                    }
                    else
                    {
                        chartObj.Series[seriesName].Points.AddXY(i, pValue);
                        chartObj.Series[seriesName].Points[i].ToolTip = "Lot:" + rowDT.Rows[i]["Lot_ID"].ToString() + "\r\nDateTime:" + rowDT.Rows[i]["Trtm"].ToString() + "\r\nValue:0";
                    }
                    chartObj.Series[seriesName].Points[i].Url = "javascript:LinkPoint('" + ((LotID + TimeStr)) + "');";
                }
                chartObj.Series[seriesName].IsVisibleInLegend = true;
                chartObj.Series[seriesName].LegendText = ModeStr;
            }

            chartTitle = chartTitle.Substring(0, (chartTitle.Length - 1));
            chartObj.Titles[0].Text = chartTitle;

            // --- Step3. Summary ---
            Double SummaryOty = 0;
            series = chartObj.Series.Add("Summary");
            series.ChartArea = "Default";
            series.ChartType = SeriesChartType.Line;
            series.Color = Color.Red;
            series.MarkerStyle = MarkerStyle.Triangle;
            series.MarkerSize = 10;
            series.MarkerColor = Color.Red;
            series.BorderColor = Color.White;
            series.BorderWidth = 1;
            for (int i = 0; i < rowDT.Rows.Count; i++)
            {
                SummaryOty = 0;
                LotID = rowDT.Rows[i]["Lot_ID"].ToString().Trim();
                TimeStr = rowDT.Rows[i]["Trtm"].ToString().Trim();
                ndataStr = TimeStr;
                ndataStr = ndataStr.Substring(0, 10);
                // 原本參數
                Stage = (lab_PA.Text);
                ModeStr = (lab_PB.Text);
                if (rowDT.Rows[i]["pValue"] != System.DBNull.Value)
                {
                    pValue = Convert.ToDouble(rowDT.Rows[i]["pValue"].ToString());
                    pValue = Math.Round(pValue, 4, MidpointRounding.AwayFromZero);
                    SummaryOty += pValue;
                }
                else
                {
                    SummaryOty += 0;
                }
                
                // 其他參數
                for (int j = 0; j < paramAry.Count; j++)
                {
                    Stage = (ddl_ptaoi_layer.SelectedValue.ToString());
                    ModeStr = paramAry[j].ToString();
                    if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
                    {
                        RowDR = resendDT.Select("Lot_ID='" + (rowDT.Rows[i]["Lot_ID"].ToString()) + "' and " + Col2Name + "='" + ModeStr + "'");
                    }
                    else 
                    {
                        RowDR = resendDT.Select("Lot_ID='" + (rowDT.Rows[i]["Lot_ID"].ToString()) + "' and " + Col1Name + "='" + Stage + "' and " + Col2Name + "='" + ModeStr + "'");
                    }
                    if (RowDR.Length > 0)
                    {
                        pValue = Convert.ToDouble(RowDR[0]["pValue"].ToString());
                        pValue = Math.Round(pValue, 4, MidpointRounding.AwayFromZero);
                        SummaryOty += pValue;
                    }
                    else
                    {
                        SummaryOty += 0;
                    }
                }
                chartObj.Series["Summary"].Points.AddXY(i, SummaryOty);
                chartObj.Series["Summary"].Points[i].ToolTip = "Lot:" + LotID + "\r\nDateTime:" + TimeStr + "\r\nValue:" + SummaryOty.ToString();
                chartObj.Series["Summary"].Points[i].Url = "javascript:LinkPoint('" + ((LotID + TimeStr)) + "');";
                chartObj.Series["Summary"].IsVisibleInLegend = true;
                chartObj.Series["Summary"].LegendText = "Summary";
            }
        }

    }

    // GridView
    private void bindGridView(ref DataTable rowDT, String Col1Name, String Col2Name, ref ArrayList modeAry, Boolean isResend, ref DataTable resendDT) 
    {
        String colName = "";
        DataRow workDR;
        DataRow[] DRAry;
        double Num = 0;
        double rate = 0;
        double summaryNum = 0;
        double InQty = 0;
        
        // --- DataTable 定義 ---
        DataTable workDT = new DataTable();
        workDT.Columns.Add("Lot ID", Type.GetType("System.String"));
        workDT.Columns.Add("Trtm", Type.GetType("System.String"));
        if (this.lab_modeType.Text == "0")
        {
            workDT.Columns.Add("Stage", Type.GetType("System.String"));
            colName = "";
        }
        else if (this.lab_modeType.Text == "1")
        {
            workDT.Columns.Add("Stage", Type.GetType("System.String"));
            colName = "Code_";
        }
        else if (this.lab_modeType.Text == "2")
        {
            workDT.Columns.Add("Layer", Type.GetType("System.String"));
            colName = (lab_PA.Text) + "_";
        }
        workDT.Columns.Add("Ori_Input_QTY", Type.GetType("System.String"));
        workDT.Columns.Add((lab_PB.Text), Type.GetType("System.String"));
        workDT.Columns.Add((lab_PB.Text) + "%", Type.GetType("System.String"));

        for (int i = 0; i < modeAry.Count; i++)
        {
            if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
            {
                workDT.Columns.Add((modeAry[i].ToString()), Type.GetType("System.String"));
                workDT.Columns.Add((modeAry[i].ToString()) + "%", Type.GetType("System.String"));
            }
            else if (this.lab_modeType.Text == "2")
            {
                workDT.Columns.Add(((ddl_ptaoi_layer.SelectedValue.ToString()) + "_" + modeAry[i].ToString()), Type.GetType("System.String"));
                workDT.Columns.Add(((ddl_ptaoi_layer.SelectedValue.ToString()) + "_" + modeAry[i].ToString()) + "%", Type.GetType("System.String"));
            }
        }

        if (modeAry.Count > 0)
        {
            workDT.Columns.Add("Summary", Type.GetType("System.String"));
            workDT.Columns.Add("Summary%", Type.GetType("System.String"));
        }

        // --- DataTable 放值 ---
        for (int i = 0; i < rowDT.Rows.Count; i++)
        {
            summaryNum = 0;
            workDR = workDT.NewRow();
            workDR[0] = rowDT.Rows[i]["Lot_ID"].ToString();
            workDR[1] = rowDT.Rows[i]["Trtm"].ToString();
            workDR[2] = (lab_PA.Text);
            workDR[3] = rowDT.Rows[i]["Original_Input_Qty"].ToString();
            InQty = Convert.ToDouble(rowDT.Rows[i]["Original_Input_Qty"].ToString());
            
            // 加入原本的
            if (rowDT.Rows[i]["QTY"] != System.DBNull.Value)
            {
                workDR[4] = (rowDT.Rows[i]["QTY"]).ToString();
                Num = Convert.ToDouble((rowDT.Rows[i]["QTY"]).ToString());
                summaryNum += Num;
            }
            else
            {
                workDR[4] = "0";
                Num = 0;
                summaryNum += 0;
            }
            rate = Math.Round(((Num / InQty) * 100), 4, MidpointRounding.AwayFromZero);
            workDR[5] = rate.ToString() + "%";
            
            // 加入選擇的
            int index = 6;
            if (isResend) 
            {
                String modeStr = "";
                for (int a = 0; a < (modeAry.Count); a++)
                {
                    modeStr = modeAry[a].ToString();
                    if (rb_mode_selected.SelectedIndex == 0 || rb_mode_selected.SelectedIndex == 1)
                    {
                        DRAry = resendDT.Select("Lot_ID='" + (rowDT.Rows[i]["Lot_ID"].ToString()) + "' AND " + Col2Name + "='" + modeStr + "'");
                    }
                    else
                    {
                        String Stage = (ddl_ptaoi_layer.SelectedValue.ToString());
                        DRAry = resendDT.Select("Lot_ID='" + (rowDT.Rows[i]["Lot_ID"].ToString()) + "' AND " + Col1Name + "='" + Stage + "' AND " + Col2Name + "='" + modeStr + "'");
                    }

                    if (DRAry.Length > 0)
                    {
                        workDR[index] = DRAry[0]["QTY"].ToString();
                        Num = Convert.ToDouble(DRAry[0]["QTY"].ToString());
                        summaryNum += Num;
                    }
                    else
                    {
                        workDR[index] = "0";
                        Num = 0;
                        summaryNum += 0;
                    }
                    index += 1;
                    rate = Math.Round(((Num / InQty) * 100), 4, MidpointRounding.AwayFromZero);
                    workDR[index] = rate.ToString() + "%";
                    index += 1;

                }

                if (modeAry.Count > 0)
                {
                    workDR[index] = summaryNum;
                    rate = Math.Round(((summaryNum / InQty) * 100), 4, MidpointRounding.AwayFromZero);
                    index += 1;
                    workDR[index] = rate.ToString() + "%";
                }
            }
            workDT.Rows.Add(workDR);
        }
        
        tr_result.Visible = true;
        GV_LotRowData.DataSource = workDT;
        GV_LotRowData.DataBind();
    }

    protected void GV_LotRowData_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        string lot = "";
        string trtm = "";
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            lot = (e.Row.Cells[0].Text).Trim().ToString();
            trtm = (e.Row.Cells[1].Text).Trim().ToString();
            //System.Web.UI.WebControls.Label lab = (System.Web.UI.WebControls.Label)(e.Row.Cells[0].Controls[1]);
            e.Row.ID = (lot + trtm);
            e.Row.Cells[0].Text = "<a name=\"" + (lot + trtm) + "\">" + (e.Row.Cells[0].Text) + "</a>";
        }
    }
}