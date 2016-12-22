Imports System.Data.SqlClient
Imports System.Data
Imports Dundas.Charting.WebControl
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports System.Diagnostics


Partial Class CTF_Info
    Inherits System.Web.UI.Page

    Private Structure IPPData
        Dim dt_data As DataTable
    End Structure

    Private Structure IPPInfo
        Dim sParts As String
        Dim sSatrtTime As String
        Dim sEndTime As String
        Dim sMIS As String
        Dim sLayer As String
        Dim sPARA As String
        Dim sLot As String
        Dim sSLI As String
    End Structure

    Private Structure IPPChart
        Dim dt As DataTable
        Dim ParametricItem As String
        Dim nType As String
        Dim xLCL As Double
        Dim xUCL As Double
        Dim xCL As Double
        Dim xSigma As Double
        Dim sLCL As Double
        Dim sUCL As Double
        Dim sCL As Double
        Dim sSigma As Double
        Dim sLayer As String
        Dim sMIS As String
        Dim YValue As String
        Dim XLableValue As String
        Dim nMIN As Double
        Dim nMAX As Double
        Dim sLot As String
        Dim sSLI As String
        Dim yImpact As String
        Dim kModule As String
        Dim Critical As String
        Dim partId As String
        Dim edaItem As String
    End Structure

    Dim JSCRIPT As String
    Dim gChartH As Integer = 300
    Dim gChartW As Integer = 500

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait ......';" & _
                                                 "javascript:document.getElementById('but_Execute').disabled=true;" & _
                                                  Me.Page.GetPostBackEventReference(but_Execute))
        If Not Me.IsPostBack Then
            pageInit()
            If Request.QueryString("RE") <> Nothing Then
                Dim myDT As DataTable = New DataTable
                bindGridData(myDT)
            End If
        End If

        Me.but_Execute.Attributes.Add("onclick", "javascript:Inquery();")
        Me.but_parse.Attributes.Add("onclick", "javascript:openTransfer();")

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try

            ' -- Part ID --
            conn.Open()
            sqlStr = "select max(part_id) from dbo.CTF_Monitor_Performance_Lot_Summary where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += "group by part_id order by part_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlPart, 0)

            ' -- Meas_Item --
            sqlStr = "select max(Meas_Item) from dbo.CTF_Monitor_Performance_Lot_Summary where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += "group by Meas_Item order by Meas_Item"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlMeasItem, 1)

            ' -- Machine ID --
            sqlStr = "select max(machine_id) from dbo.CTF_Monitor_Performance_Lot_Summary where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += "group by machine_id order by machine_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlMachineID, 1)

            ' -- Lot ID --
            sqlStr = "select max(lot_id) from dbo.CTF_Monitor_Performance_Lot_Summary where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += "group by lot_id order by lot_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_lotSource, 1)

            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Part Change Step1.
    Protected Sub ddlPart_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlPart.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_rowData.Visible = False

        If Me.ddlPart.SelectedIndex > 0 Then

            Try

                conn.Open()
                ' -- Meas_Item --
                sqlStr = "select max(Meas_Item) from dbo.CTF_Monitor_Performance_Lot_Summary "
                sqlStr += "where part_id='" + (ddlPart.SelectedValue).Trim() + "' "
                sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
                sqlStr += "group by Meas_Item order by Meas_Item"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillController(myDT, ddlMeasItem, 1)

                ' -- Machine ID --
                sqlStr = "select max(machine_id) from dbo.CTF_Monitor_Performance_Lot_Summary "
                sqlStr += "where part_id='" + (ddlPart.SelectedValue).Trim() + "' "
                sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
                sqlStr += "group by machine_id order by machine_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillController(myDT, ddlMachineID, 1)

                ' -- Lot ID --
                sqlStr = "select max(lot_id) from dbo.CTF_Monitor_Performance_Lot_Summary "
                sqlStr += "where part_id='" + (ddlPart.SelectedValue).Trim() + "' "
                sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
                sqlStr += "group by lot_id order by lot_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                lb_lotShow.Items.Clear()
                UtilObj.FillLitsBoxController(myDT, lb_lotSource, 1)

                conn.Close()
            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        Else
            pageInit()
        End If

    End Sub

    ' MeasItem Change Step2.
    Protected Sub ddlMeasItem_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlMeasItem.SelectedIndexChanged

        Chart_Panel.Visible = False
        tr_rowData.Visible = False

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim partsql As String = " "
        Dim measItemSql As String = " "
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        If Me.ddlPart.SelectedIndex >= 0 Then
            partsql = " and part_id='" + (ddlPart.SelectedValue).Trim() + "' "
        End If

        If Me.ddlMeasItem.SelectedIndex > 0 Then
            measItemSql = "and meas_item='" + (ddlMeasItem.SelectedValue).Trim() + "' "
        End If

        Try

            conn.Open()
            ' -- Machine ID --
            sqlStr = "select max(machine_id) from dbo.CTF_Monitor_Performance_Lot_Summary "
            sqlStr += "where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += partsql
            sqlStr += measItemSql
            sqlStr += "group by machine_id order by machine_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlMachineID, 1)

            ' -- Lot ID --
            sqlStr = "select max(lot_id) from dbo.CTF_Monitor_Performance_Lot_Summary "
            sqlStr += "where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += partsql
            sqlStr += measItemSql
            sqlStr += "group by lot_id order by lot_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_lotSource, 1)
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    ' Machine Change Step3.
    Protected Sub ddlMachineID_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlMachineID.SelectedIndexChanged

        Chart_Panel.Visible = False
        tr_rowData.Visible = False

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim partsql As String = " "
        Dim measItemSql As String = " "
        Dim machineSql As String = " "
        Dim lotsql As String = " "
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        If Me.ddlPart.SelectedIndex >= 0 Then
            partsql = " and part_id='" + (ddlPart.SelectedValue).Trim() + "' "
        End If

        If Me.ddlMeasItem.SelectedIndex > 0 Then
            measItemSql = " and meas_item='" + (ddlMeasItem.SelectedValue).Trim() + "' "
        End If

        If Me.ddlMachineID.SelectedIndex > 0 Then
            machineSql = "and machine_id='" + (ddlMachineID.SelectedValue).Trim() + "' "
        End If

        Try

            conn.Open()
            ' -- Lot ID --
            sqlStr = "select max(lot_id) from dbo.CTF_Monitor_Performance_Lot_Summary "
            sqlStr += "where 1=1 "
            sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
            sqlStr += partsql
            sqlStr += measItemSql
            sqlStr += machineSql
            sqlStr += "group by lot_id order by lot_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_lotSource, 1)
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click

        tr_chart.Visible = False
        tr_rowData.Visible = False

        If lb_lotShow.Items.Count = 0 Then
            For i As Integer = 0 To (lb_lotSource.Items.Count - 1)
                lb_lotShow.Items.Add(lb_lotSource.Items(i).Value)
            Next
            lb_lotSource.Items.Clear()
        End If

        Dim myDT As DataTable = New DataTable
        bindGridData(myDT)

        If (cb_showchart.Checked) And (myDT.Rows.Count > 0) Then

            Dim txtString As String = (Me.txb_chart_value.Value)
            Dim chartInfo As String() = txtString.Split(New Char() {"-"})

            ' chartInfo(0) -- Chart Type  [0, Thend] [1, Bar] [2, BoxPlot]
            ' chartInfo(1) -- Chart Value [0, Mean]  [1, Std]    [2, Both]
            ' chartInfo(2) -- Group by    [0 False]  [1 True]
            ' Thend

            TrendChart(myDT, "2", True)

            'If chartInfo(0) = "0" Then
            '    ' Thend
            '    TrendChart(myDT, chartInfo(1), chartInfo(2))
            'ElseIf chartInfo(0) = "1" Then
            '    ' Bar
            '    BarChart(myDT, chartInfo(1), chartInfo(2))
            'End If

            Me.tr_chart.Visible = True

        End If

    End Sub

    Protected Sub gv_data_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_data.RowDataBound

        If e.Row.RowType = DataControlRowType.Header Then
            e.Row.Cells(13).Text = "Spec"
        End If

        If e.Row.RowType = DataControlRowType.DataRow Then

            e.Row.Cells(1).Text = "<a href=""#"" onclick=""Javascript:window.open('CTF_DataDetails.aspx?PART=" + (e.Row.Cells(0).Text) + "&LOT=" + (e.Row.Cells(1).Text) + "&ET=" + (e.Row.Cells(12).Text) + "','revert','width=1150,height=500,resizable=yes, scrollbars=yes')"">" + (e.Row.Cells(1).Text) + "</a>"
            e.Row.Cells(13).Text = "<a href=""#"" onclick=""Javascript:window.open('CTF_SpecSetup.aspx?PART=" + (e.Row.Cells(0).Text) + "','revert','width=520,height=630')"">Setup</a>"

            ' 看 CPK 是否小於 1.33, 有要變紅色
            If e.Row.Cells(10).Text.Length > 0 Then
                Dim cpk As Double = Convert.ToDouble(e.Row.Cells(10).Text)
                If cpk < 1.33 Then
                    e.Row.Cells(10).BackColor = Color.Red
                End If
            End If

            For i As Integer = 0 To (e.Row.Cells.Count - 1)
                e.Row.Cells(i).Font.Size = FontUnit.XXSmall
            Next

        End If

    End Sub

    Private Sub bindGridData(ByRef myDT As DataTable)

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myAdapter As SqlDataAdapter

        sqlStr += "select Part_Id, lot_id, Machine_id, meas_item, "
        sqlStr += "Convert(char(19), lot_meas_start_datatime, 120) as lot_meas_start_datatime, "
        sqlStr += "Convert(char(19), lot_meas_end_datatime, 120) as lot_meas_end_datatime, "
        sqlStr += "round(mean_value, 4) as mean_value, "
        sqlStr += "round(Std_Value, 4) as Std_Value, "
        sqlStr += "round(Max_Value, 4) as Max_Value, "
        sqlStr += "round(Min_Value, 4) as Min_Value, "
        sqlStr += "round(Cp, 4) as Cp, "
        sqlStr += "round(Cpk, 4) as Cpk, "
        sqlStr += "round(CSL, 4) as CSL, "
        sqlStr += "USL, LSL from CTF_Monitor_Performance_Lot_Summary "
        sqlStr += "where 1=1 "
        sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "

        If ddlPart.SelectedIndex >= 0 Then
            sqlStr += "and part_id='" + (ddlPart.SelectedValue) + "' "
        End If

        If ddlMeasItem.SelectedIndex > 0 Then
            sqlStr += "and Meas_Item='" + (ddlMeasItem.SelectedValue) + "' "
        End If

        If ddlMachineID.SelectedIndex > 0 Then
            sqlStr += "and Machine_ID='" + (ddlMachineID.SelectedValue) + "' "
        End If

        If lb_lotShow.Items.Count > 0 Then
            Dim lotSet As String = ""
            For i As Integer = 0 To (lb_lotShow.Items.Count - 1)
                lotSet += "'" + lb_lotShow.Items(i).Value + "',"
            Next
            lotSet = lotSet.Substring(0, lotSet.Length - 1)
            sqlStr += "and lot_id in (" + lotSet + ")"
        End If

        sqlStr += "order by Part_Id, Machine_Id, Lot_Id, Meas_Item"

        Try

            conn.Open()
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.Fill(myDT)

            If (myDT.Rows.Count > 0) Then
                gv_data.DataSource = myDT
                gv_data.DataBind()
                UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_data, "#FFF68F", gv_data.AlternatingRowStyle.BackColor)
                tr_rowData.Visible = True
                lab_wait.Visible = False
            Else
                lab_wait.Visible = True
                lab_wait.Text = "No Data !!!"
            End If
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

#Region "Trend Chart"

    ' Trend Chart
    Private Sub TrendChart(ByRef myDT As DataTable, ByVal chartValue As String, ByVal isGroup As Boolean)

        Dim itemAry As ArrayList = New ArrayList
        If ddlMeasItem.SelectedIndex = 0 Then
            For i As Integer = 1 To (ddlMeasItem.Items.Count - 1)
                itemAry.Add((ddlMeasItem.Items(i).Value))
            Next
        Else
            itemAry.Add((ddlMeasItem.SelectedValue))
        End If

        For i As Integer = 0 To (itemAry.Count - 1)

            Dim drResult() As DataRow = myDT.Select("meas_item='" + (itemAry(i).ToString).Replace("'", "''") + "'")
            Dim dtFilter As DataTable
            Dim dr As DataRow
            dtFilter = myDT.Clone
            For x = 0 To drResult.Length - 1
                dr = drResult(x)
                dtFilter.LoadDataRow(dr.ItemArray, False)
            Next
            dtFilter.CaseSensitive = True

            Chart_Panel.Controls.Add(New LiteralControl("<tr><td class='Table_Two_Title' valign=middle align='center' style='font-size:middle;font-weight: bold'>" & (ddlPart.SelectedValue) & " " & (itemAry(i).ToString).Replace("'", "''") & "</td></tr>"))
            Dim Chart As Dundas.Charting.WebControl.Chart
            If chartValue = 0 Then ' Mean
                gChartH = 600
                gChartW = 1080
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawTrendChart(Chart, dtFilter, (itemAry(i).ToString), "mean_value")

                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            ElseIf chartValue = 1 Then ' Std
                gChartH = 600
                gChartW = 1080
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawTrendChart(Chart, dtFilter, (itemAry(i).ToString), "Std_Value")
                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            Else
                gChartH = 400
                gChartW = 500
                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawTrendChart(Chart, dtFilter, (itemAry(i).ToString), "mean_value")
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td><td>"))
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawTrendChart(Chart, dtFilter, (itemAry(i).ToString), "Std_Value")
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            End If

        Next

    End Sub

    ' Draw Trend Chart
    Private Sub DrawTrendChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByVal meas_Item As String, ByVal valueType As String)

        Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.DodgerBlue, Color.Firebrick, Color.Olive, Color.Green}
        Chart.ImageUrl = "temp/CTF_Bihon_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH)
        Chart.Width = Unit.Pixel(gChartW)

        If valueType = "mean_value" Then
            Chart.Titles.Add("Mean")
        Else
            Chart.Titles.Add("Std")
        End If
        Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(0).Color = Color.DarkBlue
        'Chart.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (cinfo.ParametricItem) + "','" + (valueType) + "','" + isHLStr + "')"

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisX.Title = "【" + meas_Item + "】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -90 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        Dim toolGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, "Machine_id")
        Dim foundRows() As DataRow
        Dim lot_id As String
        Dim mean_Value As Double

        Dim first As Boolean = True
        Dim chartMax As Double = 0
        Dim chartMin As Double = 0
        Dim dataMax As Double = 0
        Dim dataMin As Double = 0
        Dim maxTool As String = "" ' 找機台最多的點, 畫上下界
        Dim maxToolPoint As Integer = 0
        Dim hrefStr, sPartID, sMItem, sTool As String

        For toolIndex As Integer = 0 To (toolGroupDT.Rows.Count - 1)

            foundRows = DtSet.Select("Machine_id='" + ((toolGroupDT.Rows(toolIndex)("Machine_id")).Replace("'", "''")) + "'", "Machine_id, lot_meas_start_datatime")
            series = Chart.Series.Add((toolGroupDT.Rows(toolIndex)("Machine_id")).ToString)
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Line
            series.Color = aryColor(toolIndex)
            series.MarkerStyle = MarkerStyle.Circle
            series.MarkerSize = 8
            series.MarkerColor = Color.DarkBlue
            series.BorderColor = Color.White
            series.BorderWidth = 1

            If maxToolPoint < foundRows.Length Then
                maxToolPoint = foundRows.Length
                maxTool = (toolGroupDT.Rows(toolIndex)("Machine_id")).Replace("'", "''")
            End If

            For rowIndex As Integer = 0 To (foundRows.Length - 1)

                If first Then

                    Try
                        chartMax = CType(foundRows(rowIndex)("USL"), Double)
                        chartMin = CType(foundRows(rowIndex)("LSL"), Double)
                        dataMax = CType(foundRows(rowIndex)(valueType), Double)
                        dataMin = CType(foundRows(rowIndex)(valueType), Double)
                    Catch ex As Exception
                    End Try
                    first = False

                End If

                If Not IsDBNull(foundRows(rowIndex).Item(valueType)) Then

                    lot_id = foundRows(rowIndex).Item("Lot_id").ToString
                    mean_Value = CType(foundRows(rowIndex).Item(valueType), Double)

                    series.Points.AddXY(lot_id, mean_Value)
                    series.Points(rowIndex).ToolTip = "Lot ID=" & lot_id & vbCrLf & "Value=" & Math.Round(mean_Value, 5)

                    sPartID = (foundRows(rowIndex).Item("Part_Id").ToString)
                    sMItem = (foundRows(rowIndex).Item("meas_item").ToString)
                    sTool = (foundRows(rowIndex).Item("Machine_id").ToString)
                    hrefStr = "javascript:openWindowWithPost('{0}','{1}','{2}','{3}')"
                    hrefStr = String.Format(hrefStr, sPartID, sMItem, sTool, valueType)
                    series.Points(rowIndex).Href = hrefStr

                    If dataMax < mean_Value Then
                        dataMax = mean_Value
                    End If

                    If dataMin > mean_Value Then
                        dataMin = mean_Value
                    End If

                End If

            Next

        Next

        If dataMax < chartMax Then
            dataMax = chartMax
        End If

        If dataMin > chartMin Then
            dataMin = chartMin
        End If

        Dim ntemp As Double
        Dim nInterval As Double = Math.Round((dataMax - dataMin) / 5, 4)
        If nInterval <> 0 Then
            ' --- Max ---
            ntemp = (dataMax + nInterval)
            ntemp = Math.Round(ntemp, 2)
            Chart.ChartAreas("Default").AxisY.Maximum = ntemp
            ' --- Min ---
            ntemp = (dataMin - nInterval)
            ntemp = Math.Round(ntemp, 2)
            Chart.ChartAreas("Default").AxisY.Minimum = ntemp
        End If

        If maxToolPoint > 0 Then

            foundRows = DtSet.Select("Machine_id='" + maxTool + "'", "Machine_id, lot_meas_start_datatime")
            Dim dtFilter As DataTable
            Dim dr As DataRow
            dtFilter = DtSet.Clone
            For x = 0 To foundRows.Length - 1
                dr = foundRows(x)
                dtFilter.LoadDataRow(dr.ItemArray, False)
            Next
            dtFilter.CaseSensitive = True

            Try
                Chart_USL(dtFilter, Chart)
                Chart_LSL(dtFilter, Chart)
            Catch ex As Exception

            End Try

        End If

    End Sub
    Public Sub Chart_USL(ByRef dt As DataTable, ByRef Chart As Dundas.Charting.WebControl.Chart)

        Dim x_data As String
        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("USL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Red
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series.ShowInLegend = False
        series("LabelStyle") = "Top"

        If dt.Rows.Count > 0 Then

            For i As Integer = 0 To dt.Rows.Count - 1

                If Not IsDBNull(dt.Rows(i)("USL")) Then
                    x_data = dt.Rows(i)("Lot_id").ToString() + "[" + dt.Rows(i)("lot_meas_end_datatime").ToString() + "]"
                    tmpDouble = CType(dt.Rows(i)("USL"), Double)
                    series.Points.AddXY(x_data, tmpDouble)
                End If

            Next

            series.Points(series.Points.Count - 1).Label = String.Format("USL:{0:##0.###}", dt.Rows(dt.Rows.Count - 1)("USL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow

            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("USL").LegendText = "USL"

    End Sub
    Public Sub Chart_LSL(ByRef dt As DataTable, ByRef Chart As Dundas.Charting.WebControl.Chart)

        Dim x_data As String
        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("LSL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.DarkRed
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series.ShowInLegend = False
        series("LabelStyle") = "Top"

        'Series Data
        If dt.Rows.Count > 0 Then

            For i As Integer = 0 To (dt.Rows.Count - 1)

                If Not IsDBNull(dt.Rows(i)("LSL")) Then
                    x_data = dt.Rows(i)("Lot_id").ToString() + "[" + dt.Rows(i)("lot_meas_end_datatime").ToString() + "]"
                    tmpDouble = CType(dt.Rows(i)("LSL"), Double)
                    series.Points.AddXY(x_data, tmpDouble)
                End If

            Next

            series.Points(series.Points.Count - 1).Label = String.Format("LSL:{0:##0.###}", dt.Rows(dt.Rows.Count - 1)("LSL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("LSL").LegendText = "LSL"

    End Sub
    Public Sub Chart_CSL(ByRef dt As DataTable, ByRef Chart As Dundas.Charting.WebControl.Chart)

        Dim x_data As String
        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("CSL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Coral
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series.ShowInLegend = False
        series("LabelStyle") = "Top"

        If dt.Rows.Count > 0 Then

            For i As Integer = 0 To dt.Rows.Count - 1

                If Not IsDBNull(dt.Rows(i)("CSL")) Then
                    x_data = dt.Rows(i)("Lot_id").ToString() + "[" + dt.Rows(i)("lot_meas_end_datatime").ToString() + "]"
                    tmpDouble = CType(dt.Rows(i)("CSL"), Double)
                    series.Points.AddXY(x_data, tmpDouble)
                End If

            Next

            series.Points(series.Points.Count - 1).Label = String.Format("CSL:{0:##0.###}", dt.Rows(dt.Rows.Count - 1)("CSL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("CSL").LegendText = "CSL"

    End Sub


#End Region

#Region "Bar Chart"
    ' Bar Chart
    Private Sub BarChart(ByRef myDT As DataTable, ByVal chartValue As String, ByVal isGroup As Boolean)

        Dim itemAry As ArrayList = New ArrayList
        If ddlMeasItem.SelectedIndex = 0 Then
            For i As Integer = 1 To (ddlMeasItem.Items.Count - 1)
                itemAry.Add((ddlMeasItem.Items(i).Value))
            Next
        Else
            itemAry.Add((ddlMeasItem.SelectedValue))
        End If

        For i As Integer = 0 To (itemAry.Count - 1)

            Dim drResult() As DataRow = myDT.Select("meas_item='" + (itemAry(i).ToString).Replace("'", "''") + "'")
            Dim dtFilter As DataTable
            Dim dr As DataRow
            dtFilter = myDT.Clone
            For x = 0 To drResult.Length - 1
                dr = drResult(x)
                dtFilter.LoadDataRow(dr.ItemArray, False)
            Next
            dtFilter.CaseSensitive = True

            Chart_Panel.Controls.Add(New LiteralControl("<tr><td class='Table_Two_Title' valign=middle align='center' style='font-size:x-large;font-weight: bold'>" & (itemAry(i).ToString).Replace("'", "''") & "</td></tr>"))
            Dim Chart As Dundas.Charting.WebControl.Chart
            If chartValue = 0 Then ' Mean
                gChartH = 600
                gChartW = 1080
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawBarChart(Chart, dtFilter, (itemAry(i).ToString), "mean_value")
                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            ElseIf chartValue = 1 Then ' Std
                gChartH = 600
                gChartW = 1080
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawBarChart(Chart, dtFilter, (itemAry(i).ToString), "Std_Value")
                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            Else
                gChartH = 400
                gChartW = 510
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawBarChart(Chart, dtFilter, (itemAry(i).ToString), "mean_value")
                Chart_Panel.Controls.Add(New LiteralControl("<tr><td>"))
                Chart_Panel.Controls.Add(Chart)
                Chart = New Dundas.Charting.WebControl.Chart()
                DrawBarChart(Chart, dtFilter, (itemAry(i).ToString), "Std_Value")
                Chart_Panel.Controls.Add(Chart)
                Chart_Panel.Controls.Add(New LiteralControl("</td></tr>"))
            End If

        Next

    End Sub
    ' Draw Bar Chart
    Private Sub DrawBarChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByVal meas_Item As String, ByVal valueType As String)

        Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.DodgerBlue, Color.Firebrick, Color.Olive, Color.Green}
        Chart.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH) 'big
        Chart.Width = Unit.Pixel(gChartW) 'midd

        If valueType = "mean_value" Then
            Chart.Titles.Add("Mean")
        Else
            Chart.Titles.Add("Std")
        End If
        Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(0).Color = Color.DarkBlue

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.Title = "【" + meas_Item + "】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -90 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        Dim toolGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, "Machine_id")
        Dim foundRows() As DataRow
        Dim lot_id As String
        Dim mean_Value As Double

        For toolIndex As Integer = 0 To (toolGroupDT.Rows.Count - 1)

            foundRows = DtSet.Select("Machine_id='" + (toolGroupDT.Rows(toolIndex)("Machine_id")) + "'", "Machine_id, lot_meas_start_datatime")
            series = Chart.Series.Add((toolGroupDT.Rows(toolIndex)("Machine_id")).ToString)
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Column
            series.Color = aryColor(toolIndex)
            series.BorderColor = Color.White
            series.BorderWidth = 1

            For rowIndex As Integer = 0 To (foundRows.Length - 1)
                If Not IsDBNull(foundRows(rowIndex).Item("mean_value")) Then
                    lot_id = foundRows(rowIndex).Item("Lot_id").ToString
                    mean_Value = CType(foundRows(rowIndex).Item("mean_value"), Double)
                    Chart.Series((toolGroupDT.Rows(toolIndex)("Machine_id")).ToString).Points.AddXY(lot_id, mean_Value)
                End If
            Next

        Next

    End Sub
#End Region

#Region "BoxPlot Chart"
    ' BoxPlot Chart
    Private Sub BoxPlotChart(ByRef myDT As DataTable, ByVal chartValue As String, ByVal isGroup As Boolean)



    End Sub
#End Region

    ' To >>
    Protected Sub but_lotTo_Click(sender As Object, e As System.EventArgs) Handles but_lotTo.Click

        txb_lotinput.Text = ""
        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList
        For i As Integer = 0 To (lb_lotSource.Items.Count - 1)
            If lb_lotSource.Items(i).Selected Then
                DestAry.Add(lb_lotSource.Items(i).Value)
            Else
                sourceAry.Add(lb_lotSource.Items(i).Value)
            End If
        Next

        lb_lotSource.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_lotSource.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_lotShow.Items.Add(DestAry(i).ToString())
        Next

    End Sub
    ' Back <<
    Protected Sub but_lotBack_Click(sender As Object, e As System.EventArgs) Handles but_lotBack.Click

        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList

        For i As Integer = 0 To (lb_lotShow.Items.Count - 1)
            If lb_lotShow.Items(i).Selected Then
                DestAry.Add(lb_lotShow.Items(i).Value)
            Else
                sourceAry.Add(lb_lotShow.Items(i).Value)
            End If
        Next

        lb_lotShow.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_lotShow.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_lotSource.Items.Add(DestAry(i).ToString())
        Next
    End Sub

End Class
