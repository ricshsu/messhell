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


public partial class FailDetail : System.Web.UI.Page
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
                pageInit(((String)Request["P"]), (Request["F"].ToString()), (Request["W"].ToString()), (Request["WI"].ToString()), (Request["Product"].ToString()), (Request["Plant"].ToString()));
                //pageInit("SNB P22", "Bump fail", "52", "49,50,51,52", "CPU", "All");
            }
            catch (Exception ex)
            {
            }
        }
    }

    private void pageInit(string part_id, string item, string week, string weekIn, string product, string plant)
    {
        item = item.Replace("000", "''"); // 因為有 ' 字元的問題, 所以需要跳脫, 在前一頁已經用 000 代替 ' ,不然 javascript 傳不過來
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["iSVRConnectionString"].ToString());
        string sqlStr = "";
        string conditionStr = "";
        string customStr = " ";
        string partStr = " ";
        string weekStr = " ";
        string itemStr = " ";
        string topStr = " ";
        string tableName = "";
        string plantStr = "";

        SqlDataAdapter myAdapter = default(SqlDataAdapter);
        DataTable MainDT = null;
        DataTable BinCodeDT = null;
        DataTable workTable = null;
        DataTable LotDT = null;

        //TabStrip3.Items[2].Hidden = true;

        try
        {
            week = week.Replace("W", "");
            if (product == "CPU")
            {
                tableName = " BinCode_FailMode_Customer_Mapping b, BinCode c ";
                plantStr = " ";
            }
            else
            {
                tableName = " CS_BinCode_FailMode_Customer_Mapping b, CS_BinCode c ";
                plantStr = " and a.plant='" + plant + "' ";
            }

            // 建立 DataTable 
            workTable = new DataTable();
            workTable.Columns.Add("DefectCode", Type.GetType("System.String"));
            workTable.Columns.Add("FailMode", Type.GetType("System.String"));
            workTable.Columns.Add("BinCode", Type.GetType("System.String"));
            workTable.Columns.Add("Category", Type.GetType("System.String"));

            conn.Open();

            // === Main Table ===
            sqlStr = "";
            sqlStr += "select c.DefectCode_Id, a.Fail_Mode, c.BinCode, b.MF_Stage, b.BinCode_Id ";
            sqlStr += "from dbo.BinCode_Summary a, " + tableName;
            sqlStr += "where 1=1 ";
            sqlStr += plantStr;
            sqlStr += "and a.Fail_Mode=b.FailMode ";
            sqlStr += "and a.Fail_Mode='{0}' ";
            sqlStr += "and a.Part_Id='{1}' ";
            sqlStr += "and b.BinCode_Id=c.BinCode_Id ";
            sqlStr += "and a.WW IN (select ww from SystemDateMapping where yearWW={2} group by ww) ";
            sqlStr += "group by c.DefectCode_Id, a.Fail_Mode, c.BinCode, b.MF_Stage, b.BinCode_Id ";
            sqlStr += "order by b.BinCode_Id";
            sqlStr = string.Format(sqlStr, item, part_id, week);
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            MainDT = new DataTable();
            myAdapter.Fill(MainDT);

            for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            {
                conditionStr += "round((convert(float, SUM(" + MainDT.Rows[i]["BinCode_Id"] + "))/SUM(Original_Input_QTY) * 100), 2) as " + MainDT.Rows[i]["BinCode_Id"].ToString().Trim() + ",";
            }
            conditionStr = conditionStr.Substring(0, (conditionStr.Length - 1));

            if ((product == "CPU") | (plant.ToUpper() == "ALL"))
            {
                plantStr = " ";
            }
            else
            {
                plantStr = " and a.fe_plant_id='" + plant + "' ";
            }

            // === BinCode Data === ' 取最後一筆畫圓餅圖
            sqlStr = "";
            sqlStr += "select c.yearWW,";
            sqlStr += conditionStr;
            sqlStr += " from BinCode_Daily_RawData a, Customer_Prodction_Mapping b, SystemDateMapping c ";
            sqlStr += " where 1=1";
            sqlStr += plantStr;
            sqlStr += " and a.Part_Id=b.Part_Id";
            sqlStr += " and a.WW = c.WW";
            sqlStr += " and b.production_id='{0}' ";
            sqlStr += " and c.yearWW IN ({1}) ";
            sqlStr += " GROUP BY c.yearWW";
            sqlStr += " ORDER BY c.yearWW";
            sqlStr = string.Format(sqlStr, part_id, weekIn);
            myAdapter = new SqlDataAdapter(sqlStr, conn);
            BinCodeDT = new DataTable();
            myAdapter.Fill(BinCodeDT);

            // === Thrend Chart ===
            conditionStr = "";
            for (int i = 0; i <= (MainDT.Rows.Count - 1); i++)
            {
                conditionStr += MainDT.Rows[i]["BinCode_Id"] + "+";
            }
            conditionStr = conditionStr.Substring(0, (conditionStr.Length - 1));

            sqlStr = "";
            sqlStr += "select WW, Convert(char(10), datatime, 120) as WD, lot_id, ";
            sqlStr += "round((convert(float, ({0}))/(Original_Input_QTY) * 100), 2) as Total ";
            sqlStr += "from BinCode_Daily_RawData a, Customer_Prodction_Mapping b ";
            sqlStr += "where 1=1 ";
            sqlStr += plantStr;
            sqlStr += "and a.Part_Id=b.Part_Id ";
            sqlStr += "and b.production_id='{1}' ";
            sqlStr += "and WW IN (select ww from SystemDateMapping where yearWW={2} group by ww) ";
            sqlStr += "and Original_Input_QTY <> 0 ";
            sqlStr += "ORDER BY WW, WD";
            sqlStr = string.Format(sqlStr, conditionStr, part_id, week);
            try
            {
                LotDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(LotDT);
            }
            catch (Exception ex){}

            Bump_Detail(ref conn, part_id, item, week, weekIn);

            // 加入週數
            for (int i = 0; i <= (BinCodeDT.Rows.Count - 1); i++)
            {
                workTable.Columns.Add(((String)((BinCodeDT.Rows[i][0]).ToString())), Type.GetType("System.String"));
            }
            workTable.Columns.Add("Delta", Type.GetType("System.String"));

            // Area Pie & ViewGrid 
            area_Pie(ref MainDT, ref BinCodeDT, ref workTable, week, item);

            // weekIn
            area_Thred(ref LotDT, item, week, item, product, plant);

            // --- 加入 RowData ---
            sqlStr = "";
            sqlStr += "select a.Customer_Id as Customer, a.Category as 'CPU/CS', b.Production_Type AS Product_ID, ";
            sqlStr += "b.Part_ID, WW as Week, Convert(char(19), datatime, 120) as Time, ";
            sqlStr += "Lot_ID, DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage, Fail_Count as QTY, ";
            sqlStr += "ROUND(Fail_ratio, 2) as Ratio ";
            sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b ";
            sqlStr += "where 1 = 1 ";
            sqlStr += "and a.Production_Id = b.Production_Type ";
            sqlStr += "and a.Part_Id = b.Part_Id ";
            sqlStr += "and b.Production_Type = '{0}' ";
            sqlStr += "and b.fail_mode = '{1}' ";
            sqlStr += "and WW IN (select ww from SystemDateMapping where yearWW={2} group by ww) ";
            sqlStr += "order by Time, Lot_ID, BinCode ";
            sqlStr = string.Format(sqlStr, part_id, item, week);

            try
            {
                LotDT = new DataTable();
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                myAdapter.Fill(LotDT);
                if (LotDT.Rows.Count > 0)
                {
                    GV_LotRowData.DataSource = LotDT;
                    GV_LotRowData.DataBind();
                    UtilObj.Set_DataGridRow_OnMouseOver_Color(ref GV_LotRowData, "#FFF68F", GV_LotRowData.AlternatingRowStyle.BackColor);
                    lab_lotRowData.Text = (item + " RowData");
                }

            }
            catch (Exception ex) { }

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

    private void area_Thred(ref DataTable LotDT, string FailMode, string cweek, string item, string product, string plant)
    {
        if ((product == "CPU") | (plant.ToUpper() == "ALL"))
        {
            titlePanel.Controls.Add(new LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:middle;font-weight:bold;width:750px;Height:18px'>Week : " + cweek + " [" + item + "]</td></tr>"));
        }
        else
        {
            titlePanel.Controls.Add(new LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:middle;font-weight:bold;width:750px;Height:18px'>Week : " + cweek + " [" + item + "]   Plant : " + plant + "</td></tr>"));
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
        //Chart.ChartAreas("Default").AxisY.Interval =
        //Chart.ChartAreas("Default").AxisY.Maximum = 20
        //Chart.ChartAreas("Default").AxisY.Minimum = -20

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

        string wdStr = null;
        string lot_id = null;
        double value = 0;

        for (int rowIndex = 0; rowIndex <= (LotDT.Rows.Count - 1); rowIndex++)
        {
            if ((LotDT.Rows[rowIndex]["Total"]) != null)
            {
                wdStr = LotDT.Rows[rowIndex]["WD"].ToString();
                lot_id = LotDT.Rows[rowIndex]["Lot_id"].ToString();
                value = Convert.ToDouble(LotDT.Rows[rowIndex]["Total"]);
                Chart.Series[FailMode].Points.AddXY(lot_id, value);
                Chart.Series[FailMode].Points[rowIndex].ToolTip = "[W" + wdStr + "_" + lot_id + "] " + value.ToString() + "%";
            }
        }

        ThendPanel.Controls.Add(new LiteralControl("<tr><td>"));
        ThendPanel.Controls.Add(Chart);
        ThendPanel.Controls.Add(new LiteralControl("</td></tr>"));

    }

    private void area_Pie(ref DataTable MainDT, ref DataTable BinCodeDT, ref DataTable workDT, string cweek, string item)
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
            workDR[0] = MainDT.Rows[i]["DefectCode_Id"].ToString();
            workDR[1] = MainDT.Rows[i]["Fail_Mode"].ToString();
            workDR[2] = MainDT.Rows[i]["BinCode"].ToString();
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
        if (MainDT.Rows.Count > 1)
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
                    AlisStr = (String)MainDT.Rows[i]["BinCode"];
                    value = Convert.ToDouble(foundRows[0][binCodeStr]);
                    value = Math.Round(value, 2);
                    series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value);
                    series1.Points[i].ToolTip = AlisStr + " : " + (value.ToString()) + "%";
                }

            }

            PiePanel.Controls.Add(new LiteralControl("<tr><td>"));
            PiePanel.Controls.Add(Chart);
            PiePanel.Controls.Add(new LiteralControl("</td></tr>"));

        }

    }

    private void Bump_Detail(ref SqlConnection conn, string part_id, string item, string week, string weekIn)
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
                sqlStr = "select Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE " + "From dbo.BinCode_Detail_Daily_Lot " + "Where 1=1 " + "And category = '" + failType + "' " + "And production_type='" + part_id + "' " + "And WW=" + (week.Substring(4, 2)) + " " + "Group by Fail_Mode " + "Order by YIELD_VALUE DESC ";
                myAdapter = new SqlDataAdapter(sqlStr, conn);
                ItemDT = new DataTable();
                myAdapter.Fill(ItemDT);

                if (ItemDT.Rows.Count == 0)
                {
                    return;
                }

                // 取得 Pareto Chart Info
                sqlStr = "Select b.yearWW, Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE ";
                sqlStr += "From BinCode_Detail_Daily_Lot a, SystemDateMapping b ";
                sqlStr += "Where 1=1 And a.WW = b.WW And a.category='" + failType + "' And a.production_type='" + part_id + "' ";
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
                sqlStr += "SELECT a.yearWW, b.Fail_Mode, ROUND(convert(float,SUM(b.Fail_Count))/SUM(b.Original_Input_QTY) * 100, 3) AS VALUE ";
                sqlStr += "From SystemDateMapping a, dbo.BinCode_Detail_Daily_Lot b ";
                sqlStr += "Where 1=1 ";
                sqlStr += "And a.WW = b.WW ";
                sqlStr += "And category='" + failType + "' ";
                sqlStr += "And production_type='" + part_id + "' ";
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
        }
    }

}