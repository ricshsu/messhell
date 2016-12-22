Imports System.Data
Imports System.Data.SqlClient
Imports System.Drawing
Imports Dundas.Charting.WebControl

Partial Class DailyReportDetail_Test
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not Me.IsPostBack Then
            Try
                ' 如果 PartID = CustomerID 代表選擇 "廠商Group"
                pageInit((Request.Form("customerID").ToString), (Request.Form("productType").ToString), (Request.Form("partID").ToString), (Request.Form("yieldType").ToString), (Request.Form("RangeTime").ToString), (Request.Form("RangeType").ToString), (Request.Form("DataType").ToString))
            Catch ex As Exception
                Dim sError As String = ex.ToString()
            End Try
        End If
    End Sub

    Private Sub pageInit(ByVal CustomerID As String, ByVal ProductType As String, ByVal partID As String, ByVal yieldType As String, ByVal RangeTime As String, ByVal RangeType As String, ByVal DataType As String)

        Dim status As Boolean = False
        Dim sqlStr As String = ""
        Dim itemDT, myDT, rowDT As DataTable
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdapter As SqlDataAdapter
        Dim monthNextStr As String
        Dim ConditionStr As String = ""
        Dim BKMSql As String = "and BKM='N' "
        Dim AMD_FAILMODESql As String = "and AMD_FAILMODE='N' "

        If partID = "Summary" Then
            partID = CustomerID
        End If

        If RangeType = "D" Then

            ' --- Day ---
            If CustomerID = partID Then
                ConditionStr += "and Customer_Id='{0}' "
            Else
                ConditionStr += "and Part_Id='{0}' "
            End If
            ConditionStr += "and DataTime >= '{1}' "
            ConditionStr += "and DataTime <= '{2}' "
            If RangeTime.IndexOf("~") >= 0 Then
                Dim stimeTemp As String = ""
                Dim etimeTemp As String = ""
                Try
                    stimeTemp = RangeTime.Split(New Char() {"~"})(0)
                    etimeTemp = RangeTime.Split(New Char() {"~"})(1)
                Catch ex As Exception
                End Try
                ConditionStr = String.Format(ConditionStr, partID,
                                            (stimeTemp.Substring(0, 8) + " " + stimeTemp.Substring(8, 2) + ":00:00"),
                                            (etimeTemp.Substring(0, 8) + " " + etimeTemp.Substring(8, 2) + ":00:00"))
            Else
                If (RangeTime = (DateTime.Now.ToString("yyyyMMdd"))) Then
                    Dim mWD As DateTime = DateTime.ParseExact(RangeTime, "yyyyMMdd", Nothing)
                    ConditionStr = String.Format(ConditionStr, partID, ((mWD.AddDays(-1).ToString("yyyyMMdd")) + " 16:00:00"), (RangeTime + " 16:00:00"))
                Else
                    ConditionStr = String.Format(ConditionStr, partID, (RangeTime + " 00:00:00"), (RangeTime + " 23:59:59"))
                End If
            End If

        ElseIf RangeType = "W" Then

            ' --- Week --
            If CustomerID = partID Then
                ConditionStr += "and Customer_Id='{0}' "
            Else
                ConditionStr += "and Part_Id='{0}' "
            End If
            ConditionStr += "and YearWW={1} "
            ConditionStr = String.Format(ConditionStr, partID, (RangeTime))

        Else

            ' --- Month ---
            'Dim monthNext As Date = (DateTime.ParseExact(RangeTime, "yyyy-MM", Nothing))
            Dim monthNext As Date = (DateTime.ParseExact(RangeTime, "yyyyMM", Nothing))
            monthNextStr = monthNext.AddMonths(1).ToString("yyyyMM")
            If CustomerID = partID Then
                ConditionStr += "and Customer_Id='{0}' "
            Else
                ConditionStr += "and Part_Id='{0}' "
            End If
            ConditionStr += "and substring(CONVERT(varchar, DataTime, 112),1,8) >='{1}' "
            ConditionStr += "and substring(CONVERT(varchar, DataTime, 112),1,8) < '{2}' "
            ConditionStr = String.Format(ConditionStr, partID, (RangeTime + "01"), (monthNextStr + "01"))

        End If

        Try

            conn.Open()

            ' --- 取得 Chart Data ---
            sqlStr = ""
            sqlStr += "select part_id as part, Lot_id as lot, "
            sqlStr += "Convert(char(19), DataTime, 120) as DataTime, "
            sqlStr += "ROUND((yield), 4) as yield "
            If CustomerID.Equals("INTEL") Then
                If DataType.Equals("NORMAL") Then
                    sqlStr += "from dbo.Yield_Daily_RawData "
                Else
                    sqlStr += "from dbo.BKM_Yield_Daily_RawData "
                    BKMSql = "and BKM='Y' "
                End If
            ElseIf CustomerID.Equals("AMD") Then
                If DataType.Equals("NORMAL") Then
                    sqlStr += "from dbo.Yield_Daily_RawData "
                    AMD_FAILMODESql = "and AMD_FAILMODE='Y' "
                Else
                    sqlStr += "from dbo.Yield_Daily_RawData_NonIntel "
                End If
            Else
                If DataType.Equals("WB") Then
                    sqlStr += " from WB_Yield_Daily_RawData "
                Else
                    sqlStr += " from Yield_Daily_RawData_NonIntel "
                End If
            End If

            sqlStr += "where 1=1 "
            sqlStr += "and yield_category='" + yieldType + "' "
            sqlStr += ConditionStr
            sqlStr += "order by DataTime"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' --- 取得 Yield Item ---
            sqlStr = ""
            sqlStr += "select yield_category, SEQ "
            sqlStr += "from dbo.Yield_CATEGORY_Mapping "
            sqlStr += "where customer_id ='" + CustomerID + "' "
            sqlStr += "and yield_type ='" + ProductType + "' "
            sqlStr += "and SEQ != 1 "
            sqlStr += BKMSql
            sqlStr += AMD_FAILMODESql
            sqlStr += "group by yield_category, SEQ order by SEQ"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            itemDT = New DataTable
            myAdapter.Fill(itemDT)

            ' --- 畫 Chart ---
            Try
                area_Thred(partID, myDT, yieldType, RangeType, RangeTime)
            Catch ex As Exception
            End Try

            ' --- 取得 DataTable 資料 --- 

            Dim colStr As String = " Lot_id as lot, Convert(char(19), DataTime, 120) as DataTime, FE_Plant_id as FE, BE_Plant_id as BE, "
            Dim groupStr As String = " group by Lot_id, DataTime, FE_Plant_id, BE_Plant_id "

            If DataType.Equals("WB") Then
                colStr = " Lot_id as lot, Convert(char(19), DataTime, 120) as DataTime, FE_Plant_id as FE, "
                groupStr = " group by Lot_id, DataTime, FE_Plant_id "
            End If

            sqlStr = "select "
            sqlStr += colStr

            For x As Integer = 0 To (itemDT.Rows.Count - 1)
                sqlStr += "max(case when Yield_Category = '" + (itemDT.Rows(x)(0).ToString) + "' then Input_QTY else 0 end) as '" + (itemDT.Rows(x)(0).ToString) + "In',"
                sqlStr += "max(case when Yield_Category = '" + (itemDT.Rows(x)(0).ToString) + "' then Output_QTY else 0 end) as '" + (itemDT.Rows(x)(0).ToString) + "Out',"
                sqlStr += "max(case when Yield_Category = '" + (itemDT.Rows(x)(0).ToString) + "' then ROUND((yield), 4) else 0 end) as '" + (itemDT.Rows(x)(0).ToString) + "',"
            Next
            sqlStr = sqlStr.Substring(0, (sqlStr.Length - 1))

            If CustomerID.Equals("INTEL") Then
                If DataType.Equals("NORMAL") Then
                    sqlStr += " from dbo.Yield_Daily_RawData "
                    sqlStr += " where customer_id='INTEL' "
                Else
                    sqlStr += " from dbo.BKM_Yield_Daily_RawData "
                    sqlStr += " where 1=1 "
                End If
            ElseIf CustomerID.Equals("AMD") Then
                If DataType.Equals("NORMAL") Then
                    sqlStr += " from dbo.Yield_Daily_RawData "
                    sqlStr += " where customer_id='AMD' "
                Else
                    sqlStr += " from dbo.Yield_Daily_RawData_NonIntel "
                    sqlStr += " where 1=1 "
                End If
            Else
                If DataType.Equals("WB") Then
                    sqlStr += " from WB_Yield_Daily_RawData "
                Else
                    sqlStr += " from Yield_Daily_RawData_NonIntel "
                End If
                sqlStr += " where 1=1 "
            End If

            sqlStr += ConditionStr
            sqlStr += groupStr
            sqlStr += " order by DataTime "

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            rowDT = New DataTable
            myAdapter.Fill(rowDT)
            conn.Close()

            ' --- Data Grid ---
            gv_lotYield.DataSource = rowDT
            gv_lotYield.DataBind()

        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub area_Thred(ByVal Part_id As String, ByRef LotDT As DataTable, ByVal FailMode As String, ByVal rangeType As String, ByVal rangeTime As String)

        Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.DodgerBlue, Color.Firebrick, Color.Olive, Color.Green}
        Dim Chart As New Dundas.Charting.WebControl.Chart()
        Chart.ImageUrl = "temp/yieldDetail_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(400) 'big
        Chart.Width = Unit.Pixel(1100) 'midd

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisX.Title = "【" + Part_id + " " + FailMode + " " + rangeType + rangeTime + "】"
        Chart.ChartAreas("Default").AxisX.TitleColor = Color.Red
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -60 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

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

        Dim lot_id As String
        Dim timeStr As String
        Dim value As Double
        Dim minValue As Double = 0
        Dim maxValue As Double = 0
        Dim firstTimeIn As Boolean = True

        For rowIndex As Integer = 0 To (LotDT.Rows.Count - 1)

            If Not IsDBNull(LotDT.Rows(rowIndex)("yield")) Then

                lot_id = LotDT.Rows(rowIndex)("LOT").ToString
                timeStr = LotDT.Rows(rowIndex)("DataTime").ToString
                value = ((CType(LotDT.Rows(rowIndex)("yield"), Double)) * 100)

                If (LotDT.Rows.Count = 1) Then
                    series.Type = SeriesChartType.Point
                End If

                Chart.Series(FailMode).Points.AddXY(rowIndex, value)
                Chart.Series(FailMode).Points(rowIndex).AxisLabel = lot_id
                Chart.Series(FailMode).Points(rowIndex).ToolTip = "[" + lot_id + "  " + timeStr + "] " + value.ToString + "%"
                Chart.Series(FailMode).Points(rowIndex).Href = "javascript:LinkPoint('" + (lot_id + timeStr) + "');"

                If firstTimeIn Then
                    maxValue = value
                    minValue = value
                    firstTimeIn = False
                End If

                If value > maxValue Then
                    maxValue = value
                End If

                If value < minValue Then
                    minValue = value
                End If

            End If

        Next

        Dim Maxntemp As Double
        Dim Minntemp As Double
        Dim nInterval As Double = Math.Round((maxValue - minValue), 2, MidpointRounding.AwayFromZero)
        If nInterval <> 0 Then
            ' --- Max ---
            Maxntemp = (maxValue + nInterval)
            Maxntemp = Math.Round(Maxntemp, 2, MidpointRounding.AwayFromZero)
            ' --- Min ---
            Minntemp = (minValue - nInterval)
            Minntemp = Math.Round(Minntemp, 2, MidpointRounding.AwayFromZero)
        End If

        If maxValue = minValue Then
            Maxntemp = maxValue + 5
            Minntemp = minValue - 5
        End If

        If Maxntemp > 100 Then
            Maxntemp = 100
        End If

        getSPECLine(LotDT, Chart, Part_id, FailMode)
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        ThendPanel.Controls.Add(New LiteralControl("<tr><td>"))
        ThendPanel.Controls.Add(Chart)
        ThendPanel.Controls.Add(New LiteralControl("</td></tr>"))

    End Sub

    Protected Sub gv_lotYield_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_lotYield.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then

            For i As Integer = 0 To (e.Row.Cells.Count - 1)

                Dim lot As String = (e.Row.Cells(1).Text)
                Dim trtm As String = (e.Row.Cells(2).Text)
                e.Row.ID = (lot + trtm)
                Dim txt As String = ""

                If i = 1 Then
                    txt = "<a name=""" + (lot + trtm) + """>" + (e.Row.Cells(1).Text) + "</a>"
                Else
                    txt = e.Row.Cells(i).Text
                End If

                Dim lab As New System.Web.UI.WebControls.Label
                lab.Text = txt
                lab.Height = Unit.Pixel(20)
                lab.Width = Unit.Pixel(100)
                lab.Font.Size = FontUnit.XSmall
                e.Row.Cells(i).Controls.Add(lab)

            Next

        End If


    End Sub

    Private Sub getSPECLine(ByRef dt As DataTable, ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal Part_ID As String, ByVal yieldType As String)

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim ucl As Double = 0
        Dim lcl As Double = 0

        Try
            conn.Open()
            sqlStr = "SELECT "
            sqlStr += "UCL_Sigma, LCL_Sigma, "
            sqlStr += "CASE WHEN (UCL_Sigma = 3) THEN UCL_3S  WHEN (UCL_Sigma = 4) THEN UCL_4S WHEN (UCL_Sigma = 5) THEN UCL_5S ELSE NULL END as UCL,"
            sqlStr += "CASE WHEN (LCL_Sigma = 3) THEN LCL_3S  WHEN (LCL_Sigma = 4) THEN LCL_4S WHEN (LCL_Sigma = 5) THEN LCL_5S ELSE NULL END as LCL "
            sqlStr += "FROM dbo.Yield_SPEC "
            sqlStr += "WHERE 1=1 "
            sqlStr += "AND Part_Id='{0}' "
            sqlStr += "AND UPPER(OOC_Code)='{1}' "
            sqlStr = String.Format(sqlStr, Part_ID, (yieldType.ToUpper()))
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            If myDT.Rows.Count > 0 Then

                If Not IsDBNull(myDT.Rows(0)("UCL")) Then
                    ucl = Math.Round((CType(myDT.Rows(0)("UCL").ToString(), Double) * 100), 2, MidpointRounding.AwayFromZero)
                    Chart_SPEC(dt, Chart, (myDT.Rows(0)("UCL_Sigma").ToString() + "_UCL"), ucl.ToString())
                End If

                If Not IsDBNull(myDT.Rows(0)("LCL")) Then
                    lcl = Math.Round((CType(myDT.Rows(0)("LCL").ToString(), Double) * 100), 2, MidpointRounding.AwayFromZero)
                    Chart_SPEC(dt, Chart, (myDT.Rows(0)("LCL_Sigma").ToString() + "_LCL"), lcl.ToString())
                End If

            End If

            Chart.ChartAreas("Default").AxisY.Interval = 10
            Chart.ChartAreas("Default").AxisY.Maximum = 100
            Chart.ChartAreas("Default").AxisY.Minimum = 0
            
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Public Sub Chart_SPEC(ByRef dt As DataTable, ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal LineType As String, ByVal LineValue As Double)

        Dim series As Series
        series = Chart.Series.Add(LineType)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Red
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        Dim objPoint As DataPoint
        For i As Integer = 0 To (dt.Rows.Count - 1)
            objPoint = New DataPoint(i, LineValue)
            series.Points.Add(objPoint)
            objPoint.ToolTip = String.Format((LineType + ":{0:##0.###}"), LineValue)
        Next

        series.Points(series.Points.Count - 1).Label = String.Format((LineType + ":{0:##0.###}"), LineValue)
        series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
        series.SmartLabels.Enabled = True
        series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
        series.SmartLabels.MarkerOverlapping = False
        series.SmartLabels.MinMovingDistance = 15
        Chart.Series(LineType).LegendText = LineType

    End Sub

End Class
