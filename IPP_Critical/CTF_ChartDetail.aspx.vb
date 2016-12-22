Imports System.Data
Imports System.Data.SqlClient
Imports Dundas.Charting.WebControl
Imports System.Drawing

Partial Class CTF_ChartDetail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then

            Dim PartID As String = Request("PART")
            Dim MItem As String = Request("ITEM")
            Dim Tool As String = Request("TOOL")
            Dim valueType As String = Request("TYPE")

            'Dim PartID As String = "JYKB33"
            'Dim MItem As String = "Diameter"
            'Dim Tool As String = "TA20"
            'Dim valueType As String = "mean_value"

            pageInit(PartID, MItem, Tool, valueType)
        End If

    End Sub

    Private Sub pageInit(ByVal partID As String, ByVal MItem As String, ByVal Tool As String, ByVal valueType As String)


        Dim corrStr As String = ""
        Dim sqlStr As String = ""

        sqlStr += "select Part_Id, lot_id, Machine_id, meas_item, "
        sqlStr += "Convert(char(19), lot_meas_start_datatime, 120) as lot_meas_start_datatime, "
        sqlStr += "Convert(char(19), lot_meas_end_datatime, 120) as lot_meas_end_datatime, "
        sqlStr += "round(mean_value, 5) as mean_value, "
        sqlStr += "round(Std_Value, 5) as Std_Value, "
        sqlStr += "round(Max_Value, 5) as Max_Value, "
        sqlStr += "round(Min_Value, 5) as Min_Value, "
        sqlStr += "round(Cp, 5) as Cp, "
        sqlStr += "round(Cpk, 5) as Cpk, "
        sqlStr += "round(CSL, 5) as CSL, "
        sqlStr += "USL, LSL from CTF_Monitor_Performance_Lot_Summary "
        sqlStr += "where 1=1 "
        sqlStr += "and meas_item not in ('Diameter_1', 'Diameter_2', 'Diameter_3', 'Diameter_4') "
        sqlStr += "and Part_Id='" + partID + "' "
        sqlStr += "and meas_item='" + MItem + "' "
        sqlStr += "and Machine_id='" + Tool + "' "
        sqlStr += "order by lot_meas_end_datatime "

        Dim myDt As New DataTable
        Dim correlationDt As New DataTable
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)

        Try
            ' --- Source ---
            conn.Open()
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myAdpt.Fill(myDt)
            conn.Close()
            ' --- Mean Or Std ---
           
            Dim chartObj As New Dundas.Charting.WebControl.Chart()

            DrawChart(myDt, chartObj, MItem, valueType)
            Chart_USL(myDt, chartObj)
            Chart_LSL(myDt, chartObj)
            Chart_CSL(myDt, chartObj)
            ChartPanel.Controls.Add(chartObj)

            Lot_GridView.DataSource = myDt
            Lot_GridView.DataBind()
            UtilObj.Set_DataGridRow_OnMouseOver_Color(Lot_GridView, "#FFF68F", Lot_GridView.AlternatingRowStyle.BackColor)
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub DrawChart(ByRef dt As DataTable, ByRef chart As Dundas.Charting.WebControl.Chart, ByVal MItem As String, ByVal valueType As String)

        chart.ImageUrl = "temp/CTF_Detail_#SEQ(1000,1)"
        chart.ImageType = ChartImageType.Png
        chart.Palette = ChartColorPalette.Dundas
        chart.Height = Unit.Pixel(500)
        chart.Width = Unit.Pixel(800)

        If valueType = "mean_value" Then
            chart.Titles.Add("Mean")
        Else
            chart.Titles.Add("Std")
        End If
        chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        chart.Titles(0).Color = Color.DarkBlue

        chart.Palette = ChartColorPalette.Dundas
        chart.BackColor = Color.White
        chart.BackGradientEndColor = Color.Peru
        chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        chart.BorderStyle = ChartDashStyle.Solid
        chart.BorderWidth = 3
        chart.BorderColor = Color.DarkBlue

        chart.ChartAreas.Add("Default")
        chart.ChartAreas("Default").AxisX.Title = "【" + MItem + "】"
        chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

        chart.UI.Toolbar.Enabled = False
        chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        Dim lot_id As String
        Dim mean_Value As Double

        series = chart.Series.Add(MItem)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.DarkBlue
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 8
        series.MarkerColor = Color.DarkBlue
        series.BorderColor = Color.White
        series.BorderWidth = 1

        Dim first As Boolean = True
        Dim chartMax As Double = 0
        Dim chartMin As Double = 0
        Dim dataMax As Double = 0
        Dim dataMin As Double = 0
        For rowIndex As Integer = 0 To (dt.Rows.Count - 1)

            If Not IsDBNull(dt.Rows(rowIndex).Item(valueType)) Then

                If first Then
                    chartMax = CType(dt.Rows(rowIndex)("USL"), Double)
                    chartMin = CType(dt.Rows(rowIndex)("LSL"), Double)
                    dataMax = CType(dt.Rows(rowIndex)(valueType), Double)
                    dataMin = CType(dt.Rows(rowIndex)(valueType), Double)
                    first = False
                End If

                lot_id = dt.Rows(rowIndex).Item("Lot_id").ToString() + "[" + dt.Rows(rowIndex).Item("lot_meas_end_datatime").ToString() + "]"
                mean_Value = CType(dt.Rows(rowIndex).Item(valueType), Double)

                series.Points.AddXY(lot_id, mean_Value)
                series.Points(rowIndex).ToolTip = "Lot ID=" & lot_id & vbCrLf & "Value=" & Math.Round(mean_Value, 5)

                If dataMax < mean_Value Then
                    dataMax = mean_Value
                End If

                If dataMin > mean_Value Then
                    dataMin = mean_Value
                End If

                If valueType = "mean_value" Then
                    ' 只有平均值才將超過上下界的標為紅色
                    If mean_Value > CType(dt.Rows(rowIndex)("USL"), Double) Then
                        series.Points(rowIndex).MarkerColor = Color.Red
                    End If

                    If mean_Value < CType(dt.Rows(rowIndex)("LSL"), Double) Then
                        series.Points(rowIndex).MarkerColor = Color.Red
                    End If
                End If

            End If

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
            chart.ChartAreas("Default").AxisY.Maximum = ntemp
            ' --- Min ---
            ntemp = (dataMin - nInterval)
            ntemp = Math.Round(ntemp, 2)
            chart.ChartAreas("Default").AxisY.Minimum = ntemp  
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

End Class
