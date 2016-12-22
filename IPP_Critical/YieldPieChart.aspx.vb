Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing
Imports Dundas.Charting.WebControl

Partial Class YieldPieChart
    Inherits System.Web.UI.Page

    Dim chartH As Integer = 400
    Dim chartW As Integer = 800
    Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.Green, Color.Firebrick, Color.DodgerBlue, Color.Olive, Color.DarkGreen, Color.Red, Color.Gold, Color.Gray, Color.Cyan}

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not Me.IsPostBack Then
            Try
                pageInit((Request.Form("P").ToString), (Request.Form("F").ToString), (Request.Form("W").ToString), (Request.Form("WI").ToString), (Request.Form("Product").ToString), (Request.Form("Plant").ToString))
                'pageInit("Cougar Point DT BGA-9 e2", "Bump fail", "52", "49,50,51,52", "CS", "All")
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub pageInit(ByVal part_id As String, ByVal item As String, ByVal week As String, ByVal weekIn As String, ByVal product As String, ByVal plant As String)

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim conditionStr As String = ""
        Dim customStr As String = " "
        Dim partStr As String = " "
        Dim weekStr As String = " "
        Dim itemStr As String = " "
        Dim topStr As String = " "
        Dim tableName As String = ""
        Dim plantStr As String = ""

        Dim myAdapter As SqlDataAdapter
        Dim MainDT, BinCodeDT, workTable, LotDT As DataTable

        Try

            week = week.Replace("W", "")
            If product = "CPU" Then
                tableName = " BinCode_FailMode_Customer_Mapping b, BinCode c "
                plantStr = " "
            Else
                tableName = " CS_BinCode_FailMode_Customer_Mapping b, CS_BinCode c "
                plantStr = " and a.plant='" + plant + "' "
            End If

            ' 建立 DataTable 
            workTable = New DataTable
            workTable.Columns.Add("DefectCode", Type.GetType("System.String"))
            workTable.Columns.Add("FailMode", Type.GetType("System.String"))
            workTable.Columns.Add("BinCode", Type.GetType("System.String"))
            workTable.Columns.Add("Category", Type.GetType("System.String"))

            conn.Open()

            ' === Main Table ===
            sqlStr = ""
            sqlStr += "select c.DefectCode_Id, a.Fail_Mode, c.BinCode, b.MF_Stage, b.BinCode_Id "
            sqlStr += "from dbo.BinCode_Summary a, " + tableName
            sqlStr += "where 1=1 "
            sqlStr += plantStr
            sqlStr += "and a.Fail_Mode=b.FailMode "
            sqlStr += "and a.Fail_Mode='{0}' "
            sqlStr += "and a.Part_Id='{1}' "
            sqlStr += "and b.BinCode_Id=c.BinCode_Id "
            sqlStr += "and a.WW IN ({2}) "
            sqlStr += "group by c.DefectCode_Id, a.Fail_Mode, c.BinCode, b.MF_Stage, b.BinCode_Id "
            sqlStr += "order by b.BinCode_Id"
            sqlStr = String.Format(sqlStr, item, part_id, week)
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            MainDT = New DataTable
            myAdapter.Fill(MainDT)

            For i As Integer = 0 To (MainDT.Rows.Count - 1)
                conditionStr += "round((convert(float, SUM(" + MainDT.Rows(i)("BinCode_Id") + "))/SUM(Original_Input_QTY) * 100), 2) as " + MainDT.Rows(i)("BinCode_Id").ToString().Trim + ","
            Next
            conditionStr = conditionStr.Substring(0, (conditionStr.Length - 1))

            If (product = "CPU") Or (plant.ToUpper = "ALL") Then
                plantStr = " "
            Else
                plantStr = " and a.fe_plant_id='" + plant + "' "
            End If

            ' === BinCode Data === ' 取最後一筆畫圓餅圖
            sqlStr = ""
            sqlStr += "select WW,"
            sqlStr += conditionStr
            sqlStr += " from BinCode_Daily_RawData a, Customer_Prodction_Mapping b "
            sqlStr += " where 1=1"
            sqlStr += plantStr
            sqlStr += " and a.Part_Id=b.Part_Id"
            sqlStr += " and b.production_id='{0}' "
            sqlStr += " and WW in ({1})"
            sqlStr += " GROUP BY WW"
            sqlStr += " ORDER BY WW"
            sqlStr = String.Format(sqlStr, part_id, weekIn)
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            BinCodeDT = New DataTable
            myAdapter.Fill(BinCodeDT)

            ' === Thrend Chart ===
            conditionStr = ""
            For i As Integer = 0 To (MainDT.Rows.Count - 1)
                conditionStr += MainDT.Rows(i)("BinCode_Id") + "+"
            Next
            conditionStr = conditionStr.Substring(0, (conditionStr.Length - 1))

            sqlStr = ""
            sqlStr += "select WW, WD, lot_id, "
            sqlStr += "round((convert(float, ({0}))/(Original_Input_QTY) * 100), 2) as Total "
            sqlStr += "from BinCode_Daily_RawData a, Customer_Prodction_Mapping b "
            sqlStr += "where 1=1 "
            sqlStr += plantStr
            sqlStr += "and a.Part_Id=b.Part_Id "
            sqlStr += "and b.production_id='{1}' "
            sqlStr += "and WW={2} "
            sqlStr += "and Original_Input_QTY <> 0 "
            sqlStr += "ORDER BY WW, WD"
            sqlStr = String.Format(sqlStr, conditionStr, part_id, week)
            Try
                LotDT = New DataTable
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(LotDT)
            Catch ex As Exception

            End Try

            Bump_Detail(conn, part_id, item, week, weekIn)

            ' 加入週數
            For i As Integer = 0 To (BinCodeDT.Rows.Count - 1)
                workTable.Columns.Add("W" + ((BinCodeDT.Rows(i)(0)).ToString()).PadLeft(2, "0"), Type.GetType("System.String"))
            Next
            workTable.Columns.Add("Delta", Type.GetType("System.String"))

            ' Area Pie & ViewGrid 
            area_Pie(MainDT, BinCodeDT, workTable, week, item)

            ' weekIn
            area_Thred(LotDT, item, week, item, product, plant)

            ' --- 加入 RowData ---
            sqlStr = ""
            sqlStr += "select a.Customer_Id as Customer, a.Category as 'CPU/CS', b.Production_Type AS Product_ID, "
            sqlStr += "b.Part_ID, WW as Week, Convert(char(19), datatime, 120) as Time, "
            sqlStr += "Lot_ID, DefectCode, Fail_Mode as FailMode, BinCode, MF_Stage as Stage, Fail_Count as QTY, "
            sqlStr += "ROUND(Fail_ratio, 2) as Ratio "
            sqlStr += "from dbo.Customer_Prodction_Mapping a, dbo.BinCode_Daily_Lot b "
            sqlStr += "where 1 = 1 "
            sqlStr += "and a.Production_Id = b.Production_Type "
            sqlStr += "and a.Part_Id = b.Part_Id "
            sqlStr += "and b.Production_Type = '{0}' "
            sqlStr += "and b.fail_mode = '{1}' "
            sqlStr += "and WW = {2} "
            sqlStr += "order by DataTime "
            sqlStr = String.Format(sqlStr, part_id, item, week)

            Try
                LotDT = New DataTable
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(LotDT)
                If LotDT.Rows.Count > 0 Then
                    Me.tr_RowData.Visible = True
                    GV_LotRowData.DataSource = LotDT
                    GV_LotRowData.DataBind()
                    UtilObj.Set_DataGridRow_OnMouseOver_Color(GV_LotRowData, "#FFF68F", GV_LotRowData.AlternatingRowStyle.BackColor)
                    lab_lotRowData.Text = item + " RowData"
                Else
                    Me.tr_RowData.Visible = False
                End If

                conn.Close()
            Catch ex As Exception

            End Try

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub area_Thred(ByRef LotDT As DataTable, ByVal FailMode As String, ByVal cweek As String, ByVal item As String, ByVal product As String, ByVal plant As String)

        If (product = "CPU") Or (plant.ToUpper = "ALL") Then
            ThendPanel.Controls.Add(New LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:x-large;font-weight:bold;width:750px'>Week : " & cweek & " [" + item + "]</td></tr>"))
        Else
            ThendPanel.Controls.Add(New LiteralControl("<tr><td class='Table_One_Title' valign=middle align='center' style='font-size:x-large;font-weight:bold;width:750px'>Week : " & cweek & " [" + item + "]   Plant : " & plant & "</td></tr>"))
        End If

        Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.DodgerBlue, Color.Firebrick, Color.Olive, Color.Green}
        Dim Chart As New Dundas.Charting.WebControl.Chart()
        Chart.ImageUrl = "temp/yieldT_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = chartH
        Chart.Width = chartW

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.Title = "【" + FailMode + "】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        'Chart.ChartAreas("Default").AxisY.Interval =
        'Chart.ChartAreas("Default").AxisY.Maximum = 20
        'Chart.ChartAreas("Default").AxisY.Minimum = -20

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        series = Chart.Series.Add(FailMode)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = aryColor(0)
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 8
        series.MarkerColor = Color.DarkBlue
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series.ShowInLegend = False

        Dim wdStr As String
        Dim lot_id As String
        Dim value As Double

        For rowIndex As Integer = 0 To (LotDT.Rows.Count - 1)
            If Not IsDBNull(LotDT.Rows(rowIndex)("Total")) Then
                wdStr = LotDT.Rows(rowIndex)("WD").ToString
                lot_id = LotDT.Rows(rowIndex)("Lot_id").ToString
                value = CType(LotDT.Rows(rowIndex)("Total"), Double)
                Chart.Series(FailMode).Points.AddXY(lot_id, value)
                Chart.Series(FailMode).Points(rowIndex).ToolTip = "[W" + wdStr + "_" + lot_id + "] " + value.ToString + "%"
            End If
        Next

        ThendPanel.Controls.Add(New LiteralControl("<tr><td>"))
        ThendPanel.Controls.Add(Chart)
        ThendPanel.Controls.Add(New LiteralControl("</td></tr>"))

    End Sub

    Private Sub area_Pie(ByRef MainDT As DataTable, ByRef BinCodeDT As DataTable, ByRef workDT As DataTable, ByVal cweek As String, ByVal item As String)

        Dim bvalue As Double = 0
        Dim fvalue As Double = 0
        Dim codeID As String = ""
        Dim rowIndex As Integer = 0
        Dim workDR As DataRow
        ' DefectCode, FailMode, BinCode, Category, W ~, Delta
        For i As Integer = 0 To (MainDT.Rows.Count - 1)

            workDR = workDT.NewRow
            workDR(0) = MainDT.Rows(i)("DefectCode_Id").ToString()
            workDR(1) = MainDT.Rows(i)("Fail_Mode").ToString()
            workDR(2) = MainDT.Rows(i)("BinCode").ToString()
            workDR(3) = MainDT.Rows(i)("MF_Stage").ToString()
            codeID = MainDT.Rows(i)("BinCode_Id").ToString()
            rowIndex = 4

            For j As Integer = 0 To (BinCodeDT.Rows.Count - 1)

                workDR(rowIndex) = BinCodeDT.Rows(j)(codeID).ToString()
                rowIndex += 1

                If j = (BinCodeDT.Rows.Count - 2) Then
                    bvalue = CType(BinCodeDT.Rows(j)(codeID), Double)
                End If

                If j = (BinCodeDT.Rows.Count - 1) Then
                    fvalue = CType(BinCodeDT.Rows(j)(codeID), Double)
                End If

            Next

            If bvalue > 0 Or fvalue > 0 Then
                workDR(rowIndex) = (Math.Round((bvalue - fvalue), 2)).ToString
            Else
                workDR(rowIndex) = "0"
            End If
            workDT.Rows.Add(workDR)

        Next
        gv_pie.DataSource = workDT
        gv_pie.DataBind()
        UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_pie, "#FFF68F", gv_pie.AlternatingRowStyle.BackColor)

        If MainDT.Rows.Count > 1 Then ' 畫 Pie Chart

            Dim Chart As New Dundas.Charting.WebControl.Chart()
            Dim chartArea1 As ChartArea = New ChartArea()

            Chart.Palette = ChartColorPalette.Dundas
            Chart.BackColor = Color.White
            Chart.BackGradientEndColor = Color.Peru
            Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
            Chart.BorderStyle = ChartDashStyle.Solid
            Chart.BorderWidth = 3
            Chart.BorderColor = Color.DarkBlue

            'chartArea1.BackColor = Color.Transparent
            'chartArea1.BackGradientEndColor = Color.Transparent
            'chartArea1.BackGradientType = GradientType.None
            'chartArea1.BorderColor = Color.Black
            'chartArea1.BorderWidth = 1
            'chartArea1.BorderStyle = ChartDashStyle.Solid
            'chartArea1.ShadowColor = Color.Transparent
            'chartArea1.AxisY.LineColor = Color.FromArgb(64)
            'chartArea1.AxisX.Interlaced = True
            'chartArea1.AxisX.InterlacedColor = Color.FromArgb(15)
            'chartArea1.AxisX.LineColor = Color.FromArgb(64)

            Chart.ImageUrl = "temp/yieldP_#SEQ(1000,1)"
            Chart.ImageType = ChartImageType.Png
            Chart.Palette = ChartColorPalette.Dundas
            Chart.ChartAreas.Add(chartArea1)
            Chart.Height = chartH
            Chart.Width = chartW

            Dim series1 As Series
            series1 = Chart.Series.Add("MQCS")
            series1.BackGradientEndColor = Color.White
            series1.Type = SeriesChartType.Pie
            series1.ShowInLegend = True
            series1.Font = New Font("Verdana", 10)
            series1.FontColor = Color.Red
            series1.YValueType = ChartValueTypes.Double
            series1.XValueType = ChartValueTypes.String

            series1("PieLabelStyle") = "Outside"
            series1.BorderWidth = 2
            series1.BorderColor = System.Drawing.Color.FromArgb(26, 59, 105)

            Chart.Legends.Add("Legend1")
            Chart.Legends(0).Enabled = True
            Chart.Legends(0).Docking = Docking.Bottom
            'Chart.Legends(0).Alignment = System.Drawing.StringAlignment.Center
            series1.LegendText = "#VALX [#PERCENT]"

            Dim foundRows() As DataRow
            foundRows = BinCodeDT.Select("WW='" + cweek + "'")
            Dim value As Double = 0
            Dim binCodeStr As String = ""
            Dim AlisStr As String = ""
            If foundRows.Length > 0 Then

                For i As Integer = 0 To (MainDT.Rows.Count - 1)
                    binCodeStr = MainDT.Rows(i)("BinCode_Id")
                    AlisStr = MainDT.Rows(i)("BinCode")
                    value = CType(foundRows(0).Item(binCodeStr), Double)
                    value = Math.Round(value, 2)
                    series1.Points.AddXY(AlisStr + " " + (value.ToString()) + "%", value)
                    series1.Points(i).ToolTip = AlisStr + " : " + (value.ToString()) + "%"
                Next

            End If

            PiePanel.Controls.Add(New LiteralControl("<tr><td>"))
            PiePanel.Controls.Add(Chart)
            PiePanel.Controls.Add(New LiteralControl("</td></tr>"))

        End If

    End Sub

    Private Sub Bump_Detail(ByRef conn As SqlConnection, ByVal part_id As String, ByVal item As String, ByVal week As String, ByVal weekIn As String)

        Dim ItemDT, yieldDT, LotDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sqlStr As String

        Dim failType As String = "Bump" ' IPQC
        If (item.ToUpper).IndexOf("BUMP") >= 0 Then

            If (item.ToUpper).IndexOf("BUMP") >= 0 Then
                failType = "Bump"
                lab_DetailTitle.Text = "Bump Failure (AOI) Detail Info By Week " + week
            End If

            Try
                ' 取得最新一週的 Yield 順序的 Items
                sqlStr = "select Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE " +
                         "From dbo.BinCode_Detail_Daily_Lot " +
                         "Where 1=1 " +
                         "And category = '" + failType + "' " +
                         "And production_type='" + part_id + "' " +
                         "And WW=" + week + " " +
                         "Group by Fail_Mode " +
                         "Order by YIELD_VALUE DESC "
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                ItemDT = New DataTable
                myAdapter.Fill(ItemDT)

                If ItemDT.Rows.Count = 0 Then
                    Exit Sub
                End If

                ' 取得 Pareto Chart Info
                sqlStr = "Select WW, Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) as YIELD_VALUE " +
                         "From dbo.BinCode_Detail_Daily_Lot " +
                         "Where 1=1 " +
                         "And category='" + failType + "' " +
                         "And production_type='" + part_id + "' " +
                         "And WW in (" + weekIn + ") " +
                         "Group by WW, Fail_Mode "
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                yieldDT = New DataTable
                myAdapter.Fill(yieldDT)

                ' 取得 Lot 的 RowData 
                Dim weekAry() As String = weekIn.Split(",")
                sqlStr = "SELECT A.Fail_Mode, "
                Dim i As Integer = 0
                For i = 0 To (weekAry.Length - 1)
                    If i <> (weekAry.Length - 1) Then
                        sqlStr += "MAX(CASE WHEN A.WW=" + weekAry(i) + " THEN A.VALUE END) AS 'W" + weekAry(i) + "', "
                    Else
                        sqlStr += "MAX(CASE WHEN A.WW=" + weekAry(i) + " THEN A.VALUE END) AS 'W" + weekAry(i) + "' "
                    End If
                Next
                sqlStr += "FROM "
                sqlStr += "( "
                sqlStr += "SELECT WW, Fail_Mode, ROUND(convert(float,SUM(Fail_Count))/SUM(Original_Input_QTY) * 100, 3) AS VALUE "
                sqlStr += "From dbo.BinCode_Detail_Daily_Lot "
                sqlStr += "Where 1=1 "
                sqlStr += "And category='" + failType + "' "
                sqlStr += "And production_type='" + part_id + "' "
                sqlStr += "And WW in (" + weekIn + ") "
                sqlStr += "GROUP BY WW, Fail_Mode"
                sqlStr += ") A "
                sqlStr += "GROUP BY A.Fail_Mode "
                sqlStr += "ORDER BY 'W" + weekAry(i - 1) + "' DESC"

                myAdapter = New SqlDataAdapter(sqlStr, conn)
                LotDT = New DataTable
                myAdapter.Fill(LotDT)
                gr_lotview.DataSource = LotDT
                gr_lotview.DataBind()

                If yieldDT.Rows.Count > 0 And ItemDT.Rows.Count > 0 Then
                    Bump_Chart(yieldDT, ItemDT)
                    Me.tr_paretoChart.Visible = True
                End If

            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        End If

    End Sub

    Private Sub Bump_Chart(ByRef DtSet As DataTable, ByRef setupDT As DataTable)

        Dim Chart As New Dundas.Charting.WebControl.Chart()
        Chart.ImageUrl = "temp/BumpIPQC_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = chartH
        Chart.Width = chartW

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        ' 找出 Source 所有分類 --> Week 
        Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, "WW")
        weekGroupDT.DefaultView.Sort = "WW asc"
        weekGroupDT = weekGroupDT.DefaultView.ToTable

        Dim series As Series
        Dim insideRows() As DataRow
        Dim failMode As String
        Dim failValue As Double
        Dim weekStr As String
        Dim colorInx As Integer = 0
        Dim scriptStr As String = ""

        colorInx = (weekGroupDT.Rows.Count - 1)
        For toolIndex As Integer = 0 To (weekGroupDT.Rows.Count - 1)

            weekStr = (weekGroupDT.Rows(toolIndex)("WW")).ToString
            series = Chart.Series.Add(("WW" + weekStr))
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Column
            series.Color = aryColor(colorInx)
            series.BorderColor = Color.White
            series.BorderWidth = 1

            For i As Integer = 0 To (setupDT.Rows.Count - 1)

                failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "''")
                insideRows = DtSet.Select("WW='" + weekStr + "' and Fail_Mode='" + failMode + "'")

                failValue = 0
                If insideRows.Length > 0 Then
                    If Not IsDBNull(insideRows(0).Item("YIELD_VALUE")) Then
                        failValue = CType(insideRows(0).Item("YIELD_VALUE"), Double)
                    End If
                End If

                Chart.Series(("WW" + weekStr)).Points.AddXY(failMode, failValue)
                Chart.Series(("WW" + weekStr)).Points(i).ToolTip = "Week" & weekStr & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString

            Next
            colorInx = (colorInx - 1)

        Next
        DetailParetoPanel.Controls.Add(Chart)

    End Sub

    Protected Sub gv_pie_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_pie.RowDataBound

        If e.Row.RowType = DataControlRowType.Header Then

            e.Row.Cells(0).Width = Unit.Pixel(80)
            e.Row.Cells(1).Width = Unit.Pixel(80)
            e.Row.Cells(2).Width = Unit.Pixel(80)
            e.Row.Cells(3).Width = Unit.Pixel(80)

        End If

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Height = Unit.Pixel(50)
            For i As Integer = 4 To (e.Row.Cells.Count - 1)
                e.Row.Cells(i).Width = Unit.Pixel(50)
                e.Row.Cells(i).Text = e.Row.Cells(i).Text + "%"
            Next

        End If

    End Sub

    Protected Sub gr_lotview_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gr_lotview.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Height = Unit.Pixel(50)
            For i As Integer = 1 To (e.Row.Cells.Count - 1)
                e.Row.Cells(i).Width = Unit.Pixel(50)
                If e.Row.Cells(i).Text.Length <= 0 Then
                    e.Row.Cells(i).Text = "0%"
                Else
                    e.Row.Cells(i).Text = e.Row.Cells(i).Text + "%"
                End If
            Next
        End If

    End Sub

End Class
