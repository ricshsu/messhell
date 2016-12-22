Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing
Imports Dundas.Charting.WebControl
Imports System.IO
Imports System.Data.OleDb
Partial Class YieldLossAnalysis
    Inherits System.Web.UI.Page
    Private confTable As String = "Customer_Prodction_Mapping_BU_Rename"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not (Me.IsPostBack) Then
            pageInit()
        End If
    End Sub
   

    Private Sub pageInit()
        Try
           
            Customer()
            Product()
            Yieldloss_FailMode("")
            ExtenType()
            'ddlExtenBumpingType.SelectedValue = "CSP"
            txtDateFrom.Text = Date.Now.AddDays(-14).ToString("yyyy/MM/dd")
            txtDateTo.Text = Date.Now.AddDays(-0).ToString("yyyy/MM/dd")

        Catch ex As Exception
            Dim sError As String = ex.ToString()      
        End Try

    End Sub
    Private Function GetRemoveRebuilt(ByVal Allstation As String) As String
        Dim sValue As String = ""
        Dim station() As String
        Dim temp As String = ""
        station = Allstation.Split(",")
        If station.Length > 0 Then
            For i As Integer = 0 To station.Length - 1
                temp = station(i)

                If Left(temp, 1).ToLower <> "r" Then
                    If i = 0 Or temp = "" Then
                        sValue = temp
                    Else
                        sValue += "," + temp
                    End If
                End If

            Next
        End If

        Return sValue
    End Function

    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim customStr As String = ""
        Dim plantStr As String = ""
        Dim partStr As String = ""
        Dim weekStr As String = ""
        Dim itemStr As String = ""
        Dim topStr As String = ""
        Dim myAdapter As SqlDataAdapter
        Dim workTable As DataTable = New DataTable
        Dim x_axle As DataTable = New DataTable
        Dim workTable3 As DataTable = New DataTable

        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        Dim sGetPartID As String = Get_PartID()
        Dim Failmodeitem As String = ""
        For i As Integer = 0 To (lst_Yieldloss_Target.Items.Count - 1)
            Failmodeitem += ((lst_Yieldloss_Target.Items(i).Value).Replace("'", "''")) + ","
        Next
        If (Failmodeitem.Length > 0 AndAlso lst_Yieldloss_Target.Items.Count > 0) Then
            Failmodeitem = Failmodeitem.Substring(0, (Failmodeitem.Length - 1))
        End If

        Dim FailmodeitemTest As String = ""
        For i As Integer = 0 To (lst_Yieldloss_Target.Items.Count - 1)
            FailmodeitemTest += ((lst_Yieldloss_Target.Items(i).Text).Replace("'", "''")) + ","
        Next
        If (FailmodeitemTest.Length > 0 AndAlso lst_Yieldloss_Target.Items.Count > 0) Then
            FailmodeitemTest = FailmodeitemTest.Substring(0, (FailmodeitemTest.Length - 1))
        End If



        Dim yl As New YieldlossInfo
        yl.Part_ID = Replace(sGetPartID, "'", "")
        yl.dtStart = Date.Parse(txtDateFrom.Text)
        yl.dtEnd = Date.Parse(txtDateTo.Text)
        yl.Fail_item = Replace(Failmodeitem, "'", "")
        yl.Fail_itemTest = Replace(FailmodeitemTest, "'", "")
        yl.nFailMode = RadioButtonList2.SelectedIndex
        yl.nStation = rbl_Station.SelectedIndex
        yl.nType = rb_ProductPart.SelectedIndex
        yl.sStation = ddlStation.SelectedValue
        yl.sExtenBumpingType = Replace(Get_AD_PartID(), "'", "")
        'yl.nExtenWeek = ddlExtenWeek.SelectedValue

        If chkParallel.Checked = False And chkRebuilt.Checked = False Then
            yl.sStation = yl.sStation.Substring(0, 3)
        Else
            If chkRebuilt.Checked = False Then
                yl.sStation = GetRemoveRebuilt(yl.sStation)
            End If
        End If


            'If chkParallel.Checked = False Then

            '    yl.sStation = yl.sStation.Substring(0, 3)
            'Else

            '    If chkRebuilt.Checked = False Then
            '        yl.sStation = GetRemoveRebuilt(yl.sStation)
            '    End If

            'End If



            If yl.sStation = "" Then
                yl.sStationC = "FI0 FVI"
            Else
                yl.sStationC = ddlStation.SelectedItem.Text
            End If

            yl.nLotList = RadioButtonList1.SelectedIndex
            yl.sLotList = TextBox2.Text

            Try

            Dim tempSQL As String = ""
            Dim tempXSQL As String = ""
            If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                tempSQL = getRowDataSQL2(yl)
                tempXSQL = getRowDataSQL2_X(yl)

                myAdapter = New SqlDataAdapter(tempXSQL, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myAdapter.Fill(x_axle)
           
            Else
                tempSQL = getRowDataSQL22(yl)
            End If

                myAdapter = New SqlDataAdapter(tempSQL, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myAdapter.Fill(workTable)




                Dim expression As String = ""
                Dim foundRows() As DataRow
                If workTable.Rows.Count > 0 Then

                If RadioButtonList1.SelectedIndex = 1 And ckAdvanced.Checked = True And yl.sExtenBumpingType <> Nothing Then


                    If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                        tempSQL = getRowDataSQL3(yl, workTable)
                    Else
                        tempSQL = getRowDataSQL33(yl, workTable)
                    End If

                    myAdapter = New SqlDataAdapter(tempSQL, conn)
                    myAdapter.SelectCommand.CommandTimeout = 3600
                    myAdapter.Fill(workTable3)
                    gv_rowdata.DataSource = workTable3

                    Dim col = workTable3.Columns("TargetLot")
                    Dim lot = workTable3.Columns("Lot_ID")

                    For Each row As DataRow In workTable3.Rows
                        Dim temp As String = row(lot)

                        expression = "Lot_ID = '" & temp & "'"
                        foundRows = workTable.Select(expression)
                        If foundRows.Length > 0 Then
                            row(col) = "Y"
                        End If
                    Next



                Else
                    'workTable3 = workTable.Copy
                    workTable3 = generateDataSource(workTable, x_axle)
                    gv_rowdata.DataSource = workTable3

                    workTable.Clear()
                End If

                    gv_rowdata.DataBind()
                    UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)
                    tr_gvDisplay.Visible = True
                    tr_chartDisplay.Visible = True

                    Dim chart As New Dundas.Charting.WebControl.Chart()

                    If lst_Yieldloss_Target.Items.Count > 1 Then
                        DrawBarChart2(chart, workTable, workTable3, yl)
                    Else
                        If ckMachine.Checked = True Then
                            DrawBarChart3_ByTool(chart, workTable, workTable3, yl)
                        Else
                            DrawBarChart(chart, workTable, workTable3, yl, x_axle)
                        End If
                    End If



                    Chart_Panel.Controls.Add(chart)
                    Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                    lab_wait.Text = ""
                Else
                    tr_gvDisplay.Visible = False
                    tr_chartDisplay.Visible = False

                    lab_wait.Text = "無資料"
                End If

            Catch ex As Exception
                lab_wait.Text = "資料異常，請重新選取項目！！"


                If RadioButtonList1.SelectedIndex = 1 And TextBox2.Text = "" Then
                    lab_wait.Text = "資料異常，請載入LotID！！"
                End If
            End Try

    End Sub
    Private Function generateDataSource(ByVal DTSource As DataTable, ByVal x_axle As DataTable) As DataTable
        Dim expression As String = ""
        Dim foundRows() As DataRow
        Dim workTable As DataTable = New DataTable()
        Dim FailCNT As Double
        'If cb_SF.Checked = True Then
        '    ReDim FailCNT As Double

        'End If

        workTable.Columns.Add("Category", Type.GetType("System.String"))
        workTable.Columns.Add("BumpingType", Type.GetType("System.String"))
        workTable.Columns.Add("Part_ID", Type.GetType("System.String"))
        workTable.Columns.Add("Lot_ID", Type.GetType("System.String"))
        workTable.Columns.Add("WW", Type.GetType("System.String"))
        workTable.Columns.Add("Station_Out_DateTime", Type.GetType("System.String"))
        workTable.Columns.Add("DefectCode", Type.GetType("System.String"))
        workTable.Columns.Add("Fail_Mode", Type.GetType("System.String"))

        If cb_SF.Checked = True Then
            workTable.Columns.Add("Original_Input_QTY", Type.GetType("System.Double"))
            workTable.Columns.Add("Fail_Count", Type.GetType("System.Double"))
        Else
            workTable.Columns.Add("Original_Input_QTY", Type.GetType("System.Int32"))
            workTable.Columns.Add("Fail_Count", Type.GetType("System.Int32"))
        End If

        'If cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
        '    workTable.Columns.Add("Original_Input_QTY", Type.GetType("System.Double"))
        '    workTable.Columns.Add("Fail_Count", Type.GetType("System.Double"))
        'Else
        '    workTable.Columns.Add("Original_Input_QTY", Type.GetType("System.Int32"))
        '    workTable.Columns.Add("Fail_Count", Type.GetType("System.Int32"))
        'End If

        workTable.Columns.Add("Fail_Rate", Type.GetType("System.Double"))
        workTable.Columns.Add("FVI WW", Type.GetType("System.String"))
        workTable.Columns.Add("FVI Datatime", Type.GetType("System.String"))

        If ckMachine.Checked = True And rbl_Station.SelectedIndex = 1 Then
            workTable.Columns.Add("Machine_Id", Type.GetType("System.String"))
        End If

        Dim workRow As DataRow
        For x As Integer = 0 To (x_axle.Rows.Count - 1)
            workRow = workTable.NewRow

            workRow(0) = x_axle.Rows(x).Item("Category").ToString
            workRow(1) = x_axle.Rows(x).Item("Bumping_Type").ToString
            workRow(2) = x_axle.Rows(x).Item("Part_ID").ToString
            workRow(3) = x_axle.Rows(x).Item("Lot_ID").ToString
            workRow(4) = x_axle.Rows(x).Item("WW").ToString

            workRow(8) = x_axle.Rows(x).Item("Original_Input_QTY")
            expression = "Lot_ID = '" & x_axle.Rows(x).Item("Lot_ID").ToString & "'"
            If x_axle.Rows(x).Item("Lot_ID").ToString = "" Then
                Dim alfie As Integer
                alfie += 1
            End If
            foundRows = DTSource.Select(expression)
            If foundRows.Length > 0 Then
                workRow(5) = x_axle.Rows(x).Item("Station_Out_DateTime").ToString '***************************
                workRow(6) = foundRows(0).Item("DefectCode").ToString
                workRow(7) = foundRows(0).Item("Fail_Mode").ToString
                FailCNT = 0
                'If ckMachine.Checked = True And rbl_Station.SelectedIndex = 1 Then

                'Else
                If workRow(7) = "Inline異常報廢" Or workRow(7) = "匹配報廢" Then
                    For j As Integer = 0 To foundRows.Length - 1
                        FailCNT += CType(foundRows(j).Item("Fail_Count"), Double)
                    Next
                Else
                    FailCNT = CType(foundRows(0).Item("Fail_Count"), Double)
                End If

                'End If

                workRow(9) = FailCNT
                workRow(10) = FailCNT / CType(x_axle.Rows(x).Item("Original_Input_QTY"), Double) * 100


            Else
                If IsDBNull(x_axle.Rows(x).Item("Station_Out_DateTime")) = True Then
                    workRow(5) = x_axle.Rows(x).Item("FVI DataTime").ToString
                Else
                    workRow(5) = x_axle.Rows(x).Item("Station_Out_DateTime").ToString
                End If


                workRow(6) = ""
                workRow(7) = ""
                workRow(9) = 0
                workRow(10) = 0
            End If

            workRow(11) = x_axle.Rows(x).Item("FVI WW").ToString
            workRow(12) = x_axle.Rows(x).Item("FVI DataTime").ToString

            If ckMachine.Checked = True And rbl_Station.SelectedIndex = 1 Then
                workRow(13) = x_axle.Rows(x).Item("Machine_Id").ToString
            End If

            workTable.Rows.Add(workRow)

        Next
        Return workTable
    End Function


    Private Structure YieldlossInfo
        Dim BumpingType As String
        Dim Part_ID As String
        Dim TimePeriod As Integer
        Dim TimeRange As String
        Dim TotalOriginal As Integer

        Dim dtStart As Date
        Dim dtEnd As Date

        Dim nStation As Integer
        Dim sStation As String
        Dim sStationC As String

        Dim nLotList As Integer
        Dim sLotList As String

        Dim nTop As Integer
        Dim Fail_item As String
        Dim xoutscrape As Boolean
        Dim category As String
        Dim customer As String
        Dim isPart As Boolean
        Dim nFailMode As Integer
        Dim lotlist As String
        Dim nType As Integer
        Dim sExtenBumpingType As String
        Dim nExtenWeek As Integer

        Dim Fail_itemTest As String

    End Structure

    Private Const gChartH As Integer = 600
    Private Const gChartW As Integer = 1080
    Dim aryColor() As Color = {Color.DodgerBlue, Color.Olive, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.Blue, Color.Firebrick, Color.Green, Color.DarkSlateBlue, Color.DarkSlateGray, Color.Khaki, Color.Thistle}

    ' Draw Bar Chart
    Private Sub DrawBarChart(ByRef Chart As Chart, ByRef DtTarget As DataTable, ByRef DtSource As DataTable, ByVal yl As YieldlossInfo, ByVal x_axle As DataTable)
        'Extersion功能，需要一個Target，來至Lotlist，另一個就是Source，來至查詢條件
        Try
            Chart.ImageUrl = "temp/YieldLossAnalysis_#SEQ(10000,1)"
            Chart.ImageType = ChartImageType.Png
            Chart.Palette = ChartColorPalette.Dundas
            Chart.Height = Unit.Pixel(gChartH)
            Chart.Width = Unit.Pixel(gChartW)

            If yl.nStation = 0 Then
                Chart.Titles.Add(yl.Fail_itemTest + " Fail Ratio By FVI")
            Else
                If yl.sStation.Split(",").Length > 0 Then
                    Dim pstation() As String = yl.sStation.Split(",")
                    Dim MS As String = pstation(0)
                    Dim PS As String = ""
                    For i As Integer = 0 To pstation.Length - 1
                        If i = 0 Then

                        Else
                            If i = 1 Then
                                PS = pstation(i)
                            Else
                                PS += "," + pstation(i)
                            End If

                        End If
                    Next


                    Chart.Titles.Add(yl.Fail_item + " Fail Ratio By " + yl.sStationC + " Parallel Station :" + PS)
                Else
                    Chart.Titles.Add(yl.Fail_item + " Fail Ratio By " + yl.sStationC)
                End If

            End If

            If yl.nLotList = 1 Then

                If rbl_Station.SelectedIndex = 0 Then
                    Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("FVI Datatime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("FVI Datatime")).ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
                Else
                    Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
                End If

            Else
                Chart.Titles.Add("(" + yl.dtStart.ToString("yyyy/MM/dd") + "~" + yl.dtEnd.ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
            End If

            If rb_ProductPart.SelectedIndex = 2 Then
                Chart.Titles.Add("Part:" + yl.Part_ID)
            Else
                Chart.Titles.Add("BumpingType:" + yl.Part_ID)
            End If

            Chart.Titles(0).Font = New Font("Arial", 14, FontStyle.Bold)
            Chart.Titles(0).Color = Color.DarkBlue
            Chart.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(1).Color = Color.DarkBlue
            Chart.Titles(2).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(2).Color = Color.DarkBlue

            Chart.Palette = ChartColorPalette.Dundas
            ' Chart.BackColor = Color.White
            'Chart.BackGradientEndColor = Color.Peru
            'Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
            Chart.BorderStyle = ChartDashStyle.Solid
            'Chart.BorderWidth = 3
            'Chart.BorderColor = Color.Pink

            Chart.ChartAreas.Add("Default")
            Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
            Chart.ChartAreas("Default").AxisY2.LabelStyle.Format = "P2"
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
            Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -35 '文字對齊
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold)
            'Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
            Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            Chart.ChartAreas("Default").AxisY2.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            'Chart.ChartAreas("Default").AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
            Dim nInterval As Integer = DtSource.Rows.Count / 10
            Chart.ChartAreas("Default").AxisX.Interval = nInterval
            'Chart.ChartAreas("Default").AxisX.Title = "【" + yl.Fail_item + "】"
            Chart.ChartAreas("Default").AxisX.MajorGrid.Enabled = False
            Chart.ChartAreas("Default").AxisY.MajorGrid.Enabled = False
            Chart.ChartAreas("Default").AxisY2.MajorGrid.Enabled = False
            ' Chart.ChartAreas("Default").AxisY2.Minimum = -40
            ' Chart.ChartAreas("Default").AxisY.MinorGrid.Enabled = True
            Chart.UI.Toolbar.Enabled = False
            Chart.UI.ContextMenu.Enabled = True

            Dim series As Series

            Dim dtFilter As DataTable
            Dim dr As DataRow
            Dim foundRows() As DataRow
            Dim insideRows() As DataRow
            Dim lot_id As String
            Dim part_id As String
            Dim sdate As String

            Dim newfailMode As String
            Dim failValue As Double
            Dim weekStr As String

            Dim colorInx As Integer = 0
            Dim scriptStr As String = ""


            series = Chart.Series.Add("Fail")
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Column

            Dim AvgFail_Series As Series
            AvgFail_Series = Chart.Series.Add("FailAvg")
            AvgFail_Series.ChartArea = "Default"
            AvgFail_Series.Type = SeriesChartType.Line
            AvgFail_Series.Color = Color.DarkGoldenrod
            AvgFail_Series.BorderWidth = 3

            AvgFail_Series.YAxisType = AxisType.Secondary
            'AvgFail_Series.ShowInLegend = False
            'AvgFail_Series.Legend.e()


            If RadioButtonList1.SelectedIndex = 1 And ckAdvanced.Checked = False Then
                series.Color = Color.OrangeRed
            Else

                If DtTarget.Rows.Count > 0 Then
                    series.Color = aryColor(colorInx)
                Else
                    series.Color = Color.OrangeRed
                End If

            End If

            'series.BorderColor = Color.White
            series.BorderWidth = 1

            ' series.ShowInLegend = False
            series.LegendText = "Fail Rate"
            AvgFail_Series.LegendText = "Fail Rate Avg."

            Chart.Legends(0).LegendStyle = LegendStyle.Row
            Chart.Legends(0).BackColor = Color.White
            Chart.Legends(0).Alignment = StringAlignment.Center
            Chart.Legends(0).Docking = LegendDocking.Top
            Chart.Legends(0).FontColor = Color.DarkBlue
            Chart.Legends(0).LegendStyle = LegendStyle.Table



            Dim acolor As Integer = 30
            Dim bcolor As Integer = 144
            Dim ccolor As Integer = 255
            Dim expression As String = ""
            Dim htAvg As New Hashtable

            Dim sStartWW As String = ""
            Dim sEndWW As String = ""
            Dim nWWpointCNT As Integer = 0
            Dim WWpointCNT As Integer = 0
            Dim nWWAVG As Double = 0
            Dim WWAVGCNT As Double = 0
            Dim iMax As Double = 0
            Dim xValue As String = ""

            For i As Integer = 0 To (x_axle.Rows.Count - 1)
                xValue = ""
                lot_id = x_axle.Rows(i)("lot_id").ToString.Trim()
                ' If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                part_id = x_axle.Rows(i)("Part_id").ToString.Trim()

                If lot_id = "P63T86110041" Then
                    Dim gg As Integer
                    gg += 1
                End If

                'DT Source
                expression = "Lot_ID = '" & lot_id & "'"
                foundRows = DtSource.Select(expression)
                Dim arrDistinct As New ArrayList
                WWAVGCNT = 0
                If foundRows.Length > 0 Then
                    'failValue = CType(DtSource.Rows(i)("Fail_Rate"), Double)
                    For j As Integer = 0 To foundRows.Length - 1
                        If arrDistinct.Contains(foundRows(j).Item("DefectCode")) = False Then
                            arrDistinct.Add(foundRows(j).Item("DefectCode"))
                            WWAVGCNT += CType(foundRows(j).Item("Fail_Count"), Double)

                        End If
                    Next


                    'For j As Integer = 0 To foundRows.Length - 1

                    '    WWAVGCNT = CType(foundRows(j).Item("Fail_Count"), Double)
                    'Next
                    failValue = WWAVGCNT / CType(x_axle.Rows(i)("Original_Input_Qty"), Double) * 100
                Else
                    WWAVGCNT = 0
                    failValue = 0
                End If



                sdate = (CDate(x_axle.Rows(i)("FVI Datatime").ToString).Year - 1911).ToString + CDate(x_axle.Rows(i)("FVI Datatime")).ToString("MMdd")
                'End If
                xValue = part_id + "/" + lot_id + "/" + sdate
                Chart.Series(0).Points.AddXY(xValue, failValue)
                If yl.nStation = 0 Then
                    Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & x_axle.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & x_axle.Rows(i)("FVI Datatime").ToString
                Else
                    Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & x_axle.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & x_axle.Rows(i)("Station_Out_DateTime").ToString
                End If


                'DT Target
                expression = "Lot_ID = '" & lot_id & "'"
                foundRows = DtTarget.Select(expression)

                If foundRows.Length > 0 Then
                    Chart.Series(0).Points(i).Color = Color.OrangeRed
                End If

                If i = 0 Then
                    sStartWW = x_axle.Rows(i)("WW").ToString
                    sEndWW = sStartWW

                    nWWpointCNT += CType(x_axle.Rows(i)("Original_Input_Qty"), Double)
                    nWWAVG += WWAVGCNT

                Else
                    sEndWW = x_axle.Rows(i)("WW").ToString

                    If sStartWW <> sEndWW Then
                        nWWAVG = Math.Round(nWWAVG / nWWpointCNT * 100, 2)
                        htAvg.Add(sStartWW, nWWAVG)

                        If nWWAVG > iMax Then
                            iMax = nWWAVG
                        End If

                        sStartWW = sEndWW
                        nWWpointCNT = 0
                        nWWAVG = 0

                        nWWpointCNT += CType(x_axle.Rows(i)("Original_Input_Qty"), Double)
                        nWWAVG += WWAVGCNT
                    Else

                        nWWpointCNT += CType(x_axle.Rows(i)("Original_Input_Qty"), Double)
                        nWWAVG += WWAVGCNT
                    End If

                End If

            Next





            '***********************************************************************************************************************************
            'For i As Integer = 0 To (DtSource.Rows.Count - 1)
            '    xValue = ""
            '    lot_id = DtSource.Rows(i)("lot_id").ToString.Trim()
            '    ' If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
            '    part_id = DtSource.Rows(i)("Part_id").ToString.Trim()

            '    failValue = CType(DtSource.Rows(i)("Fail_Rate"), Double)

            '    sdate = (CDate(DtSource.Rows(i)("FVI Datatime").ToString).Year - 1911).ToString + CDate(DtSource.Rows(i)("FVI Datatime")).ToString("MMdd")
            '    'End If
            '    xValue = part_id + "/" + lot_id + "/" + sdate
            '    Chart.Series(0).Points.AddXY(xValue, failValue)
            '    If yl.nStation = 0 Then
            '        Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("FVI Datatime").ToString
            '    Else
            '        Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("Station_Out_DateTime").ToString
            '    End If


            '    'DT Target

            '    expression = "Lot_ID = '" & lot_id & "'"
            '    foundRows = DtTarget.Select(expression)
            '    If foundRows.Length > 0 Then
            '        Chart.Series(0).Points(i).Color = Color.OrangeRed
            '    End If

            '    If i = 0 Then
            '        sStartWW = DtSource.Rows(i)("WW").ToString
            '        sEndWW = sStartWW

            '        'nWWpointCNT += 1
            '        'nWWAVG += failValue
            '        nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
            '        nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)

            '    Else
            '        sEndWW = DtSource.Rows(i)("WW").ToString

            '        If sStartWW <> sEndWW Then
            '            nWWAVG = nWWAVG / nWWpointCNT * 100
            '            htAvg.Add(sStartWW, nWWAVG)

            '            If nWWAVG > iMax Then
            '                iMax = nWWAVG
            '            End If

            '            sStartWW = sEndWW
            '            nWWpointCNT = 0
            '            nWWAVG = 0
            '            'nWWpointCNT += 1
            '            'nWWAVG += failValue
            '            nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
            '            nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)
            '        Else
            '            'nWWpointCNT += 1
            '            'nWWAVG += failValue
            '            nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
            '            nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)
            '        End If

            '    End If

            'Next
            '***************************************************************************************************************************

            nWWAVG = Math.Round(nWWAVG / nWWpointCNT * 100, 2)
            htAvg.Add(sStartWW, nWWAVG)
            If nWWAVG > iMax Then
                iMax = nWWAVG
            End If

            Dim StartWW As String = ""
            Dim WW As String = ""

            'aryColor(colorInx)

            'For i As Integer = 0 To (DtSet.Rows.Count - 1)

            '    If i = 0 Then
            '        StartWW = DtSet.Rows(i)("WW")
            '        WW = DtSet.Rows(i)("WW")
            '        Chart.Series(0).Points(i).Label = "WW:" + StartWW
            '        Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
            '        Chart.Series(0).Points(i).ShowLabelAsValue = True
            '        Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
            '        Chart.Series(0).Points(i).FontAngle = -90
            '    Else
            '        WW = DtSet.Rows(i)("WW")
            '    End If

            '    If WW <> StartWW Then
            '        colorInx += 1
            '        If colorInx > 11 Then
            '            colorInx = 0
            '        End If
            '        StartWW = WW
            '        Chart.Series(0).Points(i).Label = "WW:" + StartWW
            '        Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
            '        Chart.Series(0).Points(i).ShowLabelAsValue = True
            '        Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
            '        Chart.Series(0).Points(i).FontAngle = -90
            '    End If

            '    Chart.Series(0).Points(i).Color = aryColor(colorInx)
            'Next

            Dim stripMed As New StripLine()
            Dim bWhile As Boolean = True
            Dim nStripMedCNT As Integer = 0
            For i As Integer = 0 To (x_axle.Rows.Count - 1)


                lot_id = x_axle.Rows(i)("lot_id").ToString.Trim()
                failValue = CDbl(htAvg(x_axle.Rows(i)("WW").ToString()))
                Chart.Series(1).Points.AddXY(lot_id, Math.Round(failValue, 4))

                If i = 0 Then
                    StartWW = x_axle.Rows(i)("WW")
                    WW = x_axle.Rows(i)("WW")
                    stripMed.IntervalOffset = 0

                    Chart.Series(0).Points(i).Label = "WW:" + StartWW
                    Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                    Chart.Series(0).Points(i).ShowLabelAsValue = True
                    Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                    Chart.Series(0).Points(i).FontAngle = -90

                    Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
                    Chart.Series(1).Font = New Font("Arial", 12, FontStyle.Bold)
                Else
                    WW = x_axle.Rows(i)("WW")
                End If

                If WW <> StartWW Then

                    Chart.Series(0).Points(i).Label = "WW:" + WW
                    Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                    Chart.Series(0).Points(i).ShowLabelAsValue = True
                    Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                    Chart.Series(0).Points(i).FontAngle = -90

                    stripMed.StripWidth = nStripMedCNT + 0.5
                    nStripMedCNT = 0
                    If bWhile = False Then
                        stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
                    Else
                        stripMed.BackColor = Color.White
                    End If
                    Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)
                    StartWW = WW
                    stripMed = New StripLine()
                    bWhile = Not bWhile
                    stripMed.IntervalOffset = i + 0.5

                    Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
                End If
                nStripMedCNT = nStripMedCNT + 1



            Next

            stripMed.StripWidth = nStripMedCNT + 1
            nStripMedCNT = 0
            If bWhile = False Then
                stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
            Else
                stripMed.BackColor = Color.White
            End If
            Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)


            'Chart.ChartAreas("Default").AxisY2.Minimum = iMax * -5
            'Chart.ChartAreas("Default").AxisY2.Maximum = iMax * 1.1

            Chart.ChartAreas("Default").AxisY2.Minimum = 0
            If iMax = 0 Then

            Else
                Chart.ChartAreas("Default").AxisY2.Maximum = iMax * 1.1
                Chart.ChartAreas("Default").AxisY2.Interval = iMax * 1.1 / 10
            End If



        Catch ex As Exception
            Dim sError As String = ex.ToString()
        End Try
    End Sub

    Private Sub DrawBarChart2(ByRef Chart As Chart, ByRef DtTarget As DataTable, ByRef DtSource As DataTable, ByVal yl As YieldlossInfo)
        'Try

        Dim sFailmode() As String = yl.Fail_item.Split(",")
        Dim sFail1 As String = sFailmode(0)
        Dim sFail2 As String = sFailmode(1)

        Dim sFailmodeC() As String = yl.Fail_itemTest.Split(",")
        Dim sFail1C As String = sFailmodeC(0)
        Dim sFail2C As String = sFailmodeC(1)




        Chart.ImageUrl = "temp/YieldLossAnalysis_#SEQ(10000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH)
        Chart.Width = Unit.Pixel(gChartW)


        If yl.nStation = 0 Then
            Chart.Titles.Add(yl.Fail_itemTest + " Fail Ratio By FVI")
        Else
            If yl.sStation.Split(",").Length > 0 Then
                Dim pstation() As String = yl.sStation.Split(",")
                Dim MS As String = pstation(0)
                Dim PS As String = ""
                For i As Integer = 0 To pstation.Length - 1
                    If i = 0 Then

                    Else
                        If i = 1 Then
                            PS = pstation(i)
                        Else
                            PS += "," + pstation(i)
                        End If

                    End If
                Next


                Chart.Titles.Add(yl.Fail_itemTest + " Fail Ratio By " + yl.sStationC + " Parallel Station :" + PS)
            Else
                Chart.Titles.Add(yl.Fail_itemTest + " Fail Ratio By " + yl.sStationC)
            End If

        End If


        Dim acolor As Integer = 30
        Dim bcolor As Integer = 144
        Dim ccolor As Integer = 255
        Dim expression As String = ""
        Dim htAvg As New Hashtable

        Dim sStartWW As String = ""
        Dim sEndWW As String = ""
        Dim nWWpointCNT As Integer = 0
        Dim nWWAVG As Double = 0
        Dim iMax As Double = 0
        Dim xValue As String = ""
        Dim foundRows() As DataRow
        Dim insideRows() As DataRow

        If yl.nFailMode = 0 Then
            expression = "Fail_Mode = '" & sFail1 & "'"

        Else
            expression = "DefectCode = '" & sFail1 & "'"
        End If
        foundRows = DtSource.Select(expression)

        If yl.nLotList = 1 Then

            If rbl_Station.SelectedIndex = 0 Then
                Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("FVI Datatime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("FVI Datatime")).ToString("yyyy/MM/dd") + ") (" + foundRows.Length.ToString + ")")
            Else
                Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + ") (" + foundRows.Length.ToString + ")")
            End If

        Else
            Chart.Titles.Add("(" + yl.dtStart.ToString("yyyy/MM/dd") + "~" + yl.dtEnd.ToString("yyyy/MM/dd") + ") (" + foundRows.Length.ToString + ")")
        End If

        If rb_ProductPart.SelectedIndex = 2 Then
            Chart.Titles.Add("Part:" + yl.Part_ID)
        Else
            Chart.Titles.Add("BumpingType:" + yl.Part_ID)
        End If

        Chart.Titles(0).Font = New Font("Arial", 14, FontStyle.Bold)
        Chart.Titles(0).Color = Color.DarkBlue
        Chart.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(1).Color = Color.DarkBlue
        Chart.Titles(2).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(2).Color = Color.DarkBlue

        Chart.Palette = ChartColorPalette.Dundas
        ' Chart.BackColor = Color.White
        'Chart.BackGradientEndColor = Color.Peru
        'Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        'Chart.BorderWidth = 3
        'Chart.BorderColor = Color.Pink

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisY2.LabelStyle.Format = "P2"
        'Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -35 '文字對齊
        'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold)
        'Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
        Chart.ChartAreas("Default").AxisY2.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
        'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
        'Chart.ChartAreas("Default").AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
        Dim nInterval As Integer = foundRows.Length / 10
        Chart.ChartAreas("Default").AxisX.Interval = nInterval
        'Chart.ChartAreas("Default").AxisX.Title = "【" + yl.Fail_item + "】"
        Chart.ChartAreas("Default").AxisX.MajorGrid.Enabled = False
        Chart.ChartAreas("Default").AxisY.MajorGrid.Enabled = False
        Chart.ChartAreas("Default").AxisY2.MajorGrid.Enabled = False
        ' Chart.ChartAreas("Default").AxisY2.Minimum = -40
        ' Chart.ChartAreas("Default").AxisY.MinorGrid.Enabled = True
        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series

        Dim dtFilter As DataTable
        Dim dr As DataRow

        Dim lot_id As String
        Dim part_id As String
        Dim sdate As String

        Dim newfailMode As String
        Dim failValue As Double
        Dim failValue2 As Double
        Dim weekStr As String

        Dim colorInx As Integer = 0
        Dim scriptStr As String = ""


        series = Chart.Series.Add("Fail")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Column

        Dim AvgFail_Series As Series
        AvgFail_Series = Chart.Series.Add("Fail2")
        AvgFail_Series.ChartArea = "Default"
        AvgFail_Series.Type = SeriesChartType.Line
        AvgFail_Series.Color = Color.Blue
        AvgFail_Series.BorderWidth = 3
        ' AvgFail_Series.EmptyPointStyle.Color = Color.Transparent


        AvgFail_Series.YAxisType = AxisType.Secondary
        'AvgFail_Series.ShowInLegend = False
        'AvgFail_Series.Legend.e()


        If RadioButtonList1.SelectedIndex = 1 And ckAdvanced.Checked = False Then
            series.Color = Color.OrangeRed
        Else

            If DtTarget.Rows.Count > 0 Then
                series.Color = aryColor(colorInx)
            Else
                series.Color = Color.OrangeRed
            End If

        End If

        'series.BorderColor = Color.White
        series.BorderWidth = 1

        ' series.ShowInLegend = False
        series.LegendText = "Fail:" + sFail1C
        AvgFail_Series.LegendText = "Fail:" + sFail2C

        Chart.Legends(0).LegendStyle = LegendStyle.Row
        Chart.Legends(0).BackColor = Color.White
        Chart.Legends(0).Alignment = StringAlignment.Center
        Chart.Legends(0).Docking = LegendDocking.Top
        Chart.Legends(0).FontColor = Color.DarkBlue
        Chart.Legends(0).LegendStyle = LegendStyle.Table





        For i As Integer = 0 To (foundRows.Length - 1)
            xValue = ""
            lot_id = foundRows(i)("lot_id").ToString.Trim()

            part_id = foundRows(i)("Part_id").ToString.Trim()

            failValue = CType(foundRows(i)("Fail_Rate"), Double)

            sdate = (CDate(foundRows(i)("FVI Datatime").ToString).Year - 1911).ToString + CDate(foundRows(i)("FVI Datatime")).ToString("MMdd")

            xValue = part_id + "/" + lot_id + "/" + sdate
            Chart.Series(0).Points.AddXY(xValue, failValue)


            If yl.nStation = 0 Then
                Chart.Series(0).Points(i).ToolTip = "Part_ID:" & part_id & vbCrLf & "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("FVI Datatime").ToString
            Else
                Chart.Series(0).Points(i).ToolTip = "Part_ID:" & part_id & vbCrLf & "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("Station_Out_DateTime").ToString
            End If


            If yl.nFailMode = 0 Then
                expression = "Fail_Mode = '" & sFail1 & "'"
                expression = "Lot_ID = '" & lot_id & "' and Fail_Mode = '" & sFail2 & "'"

            Else
                expression = "DefectCode = '" & sFail1 & "'"
                expression = "Lot_ID = '" & lot_id & "' and DefectCode = '" & sFail2 & "'"
            End If


            insideRows = DtSource.Select(expression)

            If insideRows.Length > 0 Then
                failValue2 = CType(insideRows(0)("Fail_Rate"), Double)

                Chart.Series(1).Points.AddXY(lot_id, Math.Round(failValue2, 4))
            Else

                Chart.Series(1).Points.AddXY(lot_id, 0)
            End If

        Next



        Dim StartWW As String = ""
        Dim WW As String = ""



        Dim stripMed As New StripLine()
        Dim bWhile As Boolean = True
        Dim nStripMedCNT As Integer = 0

        'If yl.nFailMode = 0 Then
        '    expression = "Fail_Mode = '" & sFail1 & "'"

        'Else
        '    expression = "DefectCode = '" & sFail1 & "'"
        'End If
        'foundRows = DtSource.Select(expression)

        For i As Integer = 0 To (foundRows.Length - 1)


            lot_id = foundRows(i)("lot_id").ToString.Trim()
            failValue = CDbl(htAvg(foundRows(i)("WW").ToString()))
            'Chart.Series(1).Points.AddXY(lot_id, Math.Round(failValue, 4))

            If i = 0 Then
                StartWW = foundRows(i)("WW")
                WW = foundRows(i)("WW")
                stripMed.IntervalOffset = 0

                Chart.Series(0).Points(i).Label = "WW:" + StartWW
                Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                Chart.Series(0).Points(i).ShowLabelAsValue = True
                Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                Chart.Series(0).Points(i).FontAngle = -90

                'Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
            Else
                WW = foundRows(i)("WW")
            End If

            If WW <> StartWW Then

                Chart.Series(0).Points(i).Label = "WW:" + WW
                Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                Chart.Series(0).Points(i).ShowLabelAsValue = True
                Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                Chart.Series(0).Points(i).FontAngle = -90

                stripMed.StripWidth = nStripMedCNT + 0.5
                nStripMedCNT = 0
                If bWhile = False Then
                    stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
                Else
                    stripMed.BackColor = Color.White
                End If
                Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)
                StartWW = WW
                stripMed = New StripLine()
                bWhile = Not bWhile
                stripMed.IntervalOffset = i + 0.5

                'Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
            End If
            nStripMedCNT = nStripMedCNT + 1



        Next

        stripMed.StripWidth = nStripMedCNT + 1
        nStripMedCNT = 0
        If bWhile = False Then
            stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
        Else
            stripMed.BackColor = Color.White
        End If
        Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)

        'Chart.ChartAreas("Default").AxisY2.Minimum = 0
        'Chart.ChartAreas("Default").AxisY2.Maximum = iMax * 1.1
        'Chart.ChartAreas("Default").AxisY2.Interval = iMax * 1.1 / 10


        'Catch ex As Exception
        '    Dim sError As String = ex.ToString()
        'End Try
    End Sub

    Private Sub DrawBarChart3_ByTool(ByRef Chart As Chart, ByRef DtTarget As DataTable, ByRef DtSource As DataTable, ByVal yl As YieldlossInfo)
        Try
            Chart.ImageUrl = "temp/YieldLossAnalysis_#SEQ(10000,1)"
            Chart.ImageType = ChartImageType.Png
            Chart.Palette = ChartColorPalette.Dundas
            Chart.Height = Unit.Pixel(gChartH)
            Chart.Width = Unit.Pixel(gChartW)

            If yl.nStation = 0 Then
                Chart.Titles.Add(yl.Fail_itemTest + " Fail Ratio By FVI")
            Else
                If yl.sStation.Split(",").Length > 0 Then
                    Dim pstation() As String = yl.sStation.Split(",")
                    Dim MS As String = pstation(0)
                    Dim PS As String = ""
                    For i As Integer = 0 To pstation.Length - 1
                        If i = 0 Then

                        Else
                            If i = 1 Then
                                PS = pstation(i)
                            Else
                                PS += "," + pstation(i)
                            End If

                        End If
                    Next


                    Chart.Titles.Add(yl.Fail_item + " Fail Ratio By " + yl.sStationC + " Parallel Station :" + PS)
                Else
                    Chart.Titles.Add(yl.Fail_item + " Fail Ratio By " + yl.sStationC)
                End If

            End If

            If yl.nLotList = 1 Then

                If rbl_Station.SelectedIndex = 0 Then
                    Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("FVI Datatime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("FVI Datatime")).ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
                Else
                    Chart.Titles.Add("(" + CDate(DtSource.Rows(0).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + "~" + CDate(DtSource.Rows(DtSource.Rows.Count - 1).Item("Station_Out_DateTime")).ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
                End If

            Else
                Chart.Titles.Add("(" + yl.dtStart.ToString("yyyy/MM/dd") + "~" + yl.dtEnd.ToString("yyyy/MM/dd") + ") (" + DtSource.Rows.Count.ToString + ")")
            End If

            If rb_ProductPart.SelectedIndex = 2 Then
                Chart.Titles.Add("Part:" + yl.Part_ID)
            Else
                Chart.Titles.Add("BumpingType:" + yl.Part_ID)
            End If

            Chart.Titles(0).Font = New Font("Arial", 14, FontStyle.Bold)
            Chart.Titles(0).Color = Color.DarkBlue
            Chart.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(1).Color = Color.DarkBlue
            Chart.Titles(2).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(2).Color = Color.DarkBlue

            Chart.Palette = ChartColorPalette.Dundas
            Chart.BorderStyle = ChartDashStyle.Solid


            Chart.ChartAreas.Add("Default")
            Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
            Chart.ChartAreas("Default").AxisY2.LabelStyle.Format = "P2"
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
            Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -35 '文字對齊
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 12, FontStyle.Bold)
            'Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
            Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            Chart.ChartAreas("Default").AxisY2.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            'Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font("Arial", 11, FontStyle.Bold)
            'Chart.ChartAreas("Default").AxisX.IntervalAutoMode = IntervalAutoMode.VariableCount
            Dim nInterval As Integer = DtSource.Rows.Count / 10
            Chart.ChartAreas("Default").AxisX.Interval = nInterval
            'Chart.ChartAreas("Default").AxisX.Title = "【" + yl.Fail_item + "】"
            Chart.ChartAreas("Default").AxisX.MajorGrid.Enabled = False
            Chart.ChartAreas("Default").AxisY.MajorGrid.Enabled = False
            Chart.ChartAreas("Default").AxisY2.MajorGrid.Enabled = False
            ' Chart.ChartAreas("Default").AxisY2.Minimum = -40
            ' Chart.ChartAreas("Default").AxisY.MinorGrid.Enabled = True
            Chart.UI.Toolbar.Enabled = False
            Chart.UI.ContextMenu.Enabled = True

            Dim series As Series

            Dim dtFilter As DataTable
            Dim dr As DataRow
            Dim foundRows() As DataRow
            Dim insideRows() As DataRow
            Dim lot_id As String
            Dim part_id As String
            Dim sdate As String

            Dim newfailMode As String
            Dim failValue As Double
            Dim weekStr As String

            Dim colorInx As Integer = 0
            Dim scriptStr As String = ""


            series = Chart.Series.Add("Fail")
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Column

            Dim AvgFail_Series As Series
            AvgFail_Series = Chart.Series.Add("FailAvg")
            AvgFail_Series.ChartArea = "Default"
            AvgFail_Series.Type = SeriesChartType.Line
            AvgFail_Series.Color = Color.DarkGoldenrod
            AvgFail_Series.BorderWidth = 3
            AvgFail_Series.Font = New Font("Arial", 15, FontStyle.Bold)

            AvgFail_Series.YAxisType = AxisType.Secondary
            'AvgFail_Series.ShowInLegend = False
            'AvgFail_Series.Legend.e()


            If RadioButtonList1.SelectedIndex = 1 And ckAdvanced.Checked = False Then
                series.Color = Color.OrangeRed
            Else

                If DtTarget.Rows.Count > 0 Then
                    series.Color = aryColor(colorInx)
                Else
                    series.Color = Color.OrangeRed
                End If

            End If

            'series.BorderColor = Color.White
            series.BorderWidth = 1

            ' series.ShowInLegend = False
            series.LegendText = "Fail Rate"
            AvgFail_Series.LegendText = "Fail Rate Avg."

            Chart.Legends(0).LegendStyle = LegendStyle.Row
            Chart.Legends(0).BackColor = Color.White
            Chart.Legends(0).Alignment = StringAlignment.Center
            Chart.Legends(0).Docking = LegendDocking.Top
            Chart.Legends(0).FontColor = Color.DarkBlue
            Chart.Legends(0).LegendStyle = LegendStyle.Table



            Dim acolor As Integer = 30
            Dim bcolor As Integer = 144
            Dim ccolor As Integer = 255
            Dim expression As String = ""
            Dim htAvg As New Hashtable

            Dim sStartTool As String = ""
            Dim sEndTool As String = ""
            Dim nWWpointCNT As Integer = 0
            Dim nWWAVG As Double = 0
            Dim iMax As Double = 0
            Dim xValue As String = ""
            For i As Integer = 0 To (DtSource.Rows.Count - 1)
                xValue = ""
                lot_id = DtSource.Rows(i)("lot_id").ToString.Trim()
                ' If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                part_id = DtSource.Rows(i)("Part_id").ToString.Trim()

                failValue = CType(DtSource.Rows(i)("Fail_Rate"), Double)

                sdate = (CDate(DtSource.Rows(i)("FVI Datatime").ToString).Year - 1911).ToString + CDate(DtSource.Rows(i)("FVI Datatime")).ToString("MMdd")
                'End If
                xValue = part_id + "/" + lot_id + "/" + sdate
                Chart.Series(0).Points.AddXY(xValue, failValue)
                If yl.nStation = 0 Then
                    Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("FVI Datatime").ToString
                Else
                    Chart.Series(0).Points(i).ToolTip = "Lot_ID:" & lot_id & vbCrLf & "Fail_Rate:" & failValue & vbCrLf & "Machine_Id:" & DtSource.Rows(i)("Machine_Id").ToString & vbCrLf & "WW:" & DtSource.Rows(i)("WW").ToString & vbCrLf & "Station_Out_DateTime:" & DtSource.Rows(i)("Station_Out_DateTime").ToString
                End If


                'DT Target

                expression = "Lot_ID = '" & lot_id & "'"
                foundRows = DtTarget.Select(expression)
                If foundRows.Length > 0 Then
                    Chart.Series(0).Points(i).Color = Color.OrangeRed
                End If

                If i = 0 Then
                    sStartTool = DtSource.Rows(i)("Machine_Id").ToString
                    sEndTool = sStartTool

                    nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
                    nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)

                Else
                    sEndTool = DtSource.Rows(i)("Machine_Id").ToString

                    If sStartTool <> sEndTool Then
                        nWWAVG = Math.Round(nWWAVG / nWWpointCNT * 100, 2)
                        htAvg.Add(sStartTool, nWWAVG)

                        If nWWAVG > iMax Then
                            iMax = nWWAVG
                        End If

                        sStartTool = sEndTool
                        nWWpointCNT = 0
                        nWWAVG = 0

                        nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
                        nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)
                    Else

                        nWWpointCNT += CType(DtSource.Rows(i)("Original_Input_Qty"), Double)
                        nWWAVG += CType(DtSource.Rows(i)("Fail_Count"), Double)
                    End If

                End If

            Next

            nWWAVG = Math.Round(nWWAVG / nWWpointCNT * 100, 2)
            htAvg.Add(sStartTool, nWWAVG)
            If nWWAVG > iMax Then
                iMax = nWWAVG
            End If

            Dim StartTool As String = ""
            Dim Machine_Id As String = ""



            Dim stripMed As New StripLine()
            Dim bWhile As Boolean = True
            Dim nStripMedCNT As Integer = 0
            For i As Integer = 0 To (DtSource.Rows.Count - 1)


                lot_id = DtSource.Rows(i)("lot_id").ToString.Trim()
                failValue = CDbl(htAvg(DtSource.Rows(i)("Machine_Id").ToString()))
                Chart.Series(1).Points.AddXY(lot_id, Math.Round(failValue, 4))

                If i = 0 Then
                    StartTool = DtSource.Rows(i)("Machine_Id")
                    Machine_Id = DtSource.Rows(i)("Machine_Id")
                    stripMed.IntervalOffset = 0

                    Chart.Series(0).Points(i).Label = "Machine_Id:" + StartTool
                    Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                    Chart.Series(0).Points(i).ShowLabelAsValue = True
                    Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                    Chart.Series(0).Points(i).FontAngle = -90

                    Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
                Else
                    Machine_Id = DtSource.Rows(i)("Machine_Id")
                End If

                If Machine_Id <> StartTool Then

                    ' Chart.Series(0).Points(i).Label = "Machine_Id:" + StartTool
                    Chart.Series(0).Points(i).Label = "Machine_Id:" + Machine_Id
                    Chart.Series(0).Points(i).LabelBackColor = Color.AliceBlue
                    Chart.Series(0).Points(i).ShowLabelAsValue = True
                    Chart.Series(0).Points(i).LabelBorderStyle = ChartDashStyle.Dash
                    Chart.Series(0).Points(i).FontAngle = -90

                    stripMed.StripWidth = nStripMedCNT + 0.5
                    nStripMedCNT = 0
                    If bWhile = False Then
                        stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
                    Else
                        stripMed.BackColor = Color.White
                    End If
                    Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)
                    StartTool = Machine_Id
                    stripMed = New StripLine()
                    bWhile = Not bWhile
                    stripMed.IntervalOffset = i + 0.5

                    Chart.Series(1).Points(i).Label = Math.Round(failValue, 4)
                End If
                nStripMedCNT = nStripMedCNT + 1



            Next

            stripMed.StripWidth = nStripMedCNT + 1
            nStripMedCNT = 0
            If bWhile = False Then
                stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD         
            Else
                stripMed.BackColor = Color.White
            End If
            Chart.ChartAreas("Default").AxisX.StripLines.Add(stripMed)

            Chart.ChartAreas("Default").AxisY2.Minimum = 0
            Chart.ChartAreas("Default").AxisY2.Maximum = iMax * 1.1
            Chart.ChartAreas("Default").AxisY2.Interval = iMax * 1.1 / 10


        Catch ex As Exception
            Dim sError As String = ex.ToString()
            lab_wait.Text = "Please select Station No"
        End Try
    End Sub



    Private Function getRowDataSQL2(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nFailMode = 1 Then
            Dim bincode() As String
            bincode = yl.Fail_item.Split("_")
            yl.Fail_item = bincode(0)
        End If

        If yl.nStation = 0 Then
            '處理FVI
            '開頭
            'tempSQL = "SELECT Category,a.Bumptype as BumpingType, a.Part_Id, a.Lot_Id,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
            '        & "FROM WB_BinCode_Daily_Lot a "

            tempSQL = "SELECT Category,a.Bumptype as BumpingType, a.Part_Id, a.Lot_Id, WW, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                   & "FROM WB_BinCode_Daily_Lot a with (nolock) "

            'Default FailMode=0

            If yl.nLotList = 0 Then
                If yl.nType = 2 Then
                    tempSQL += "WHERE a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ") AND a.Fail_Mode in ('" + yl.Fail_item + "') " _
                                                                     & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "

                ElseIf yl.nType = 1 Then
                    tempSQL += "WHERE a.Bumptype in (" + ConvertStr2AddMark2(yl.Part_ID) + ") AND a.Fail_Mode in ('" + yl.Fail_item + "') " _
                                                                    & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "

                End If
            Else
                tempSQL += "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND a.Fail_Mode in ('" + yl.Fail_item + "') "

            End If

            'FAI
            If ckFAI.Checked = False Then
                tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 "
            End If

            If cb_CR.Checked = False Then
                tempSQL += "and substring(lot_id,9,1)<>'Y' "
                tempSQL += "and substring(lot_id,9,1)<>'Z' "
                tempSQL += "and substring(a.Part_id,7,1)<>'V' "
            End If
            If Cb_Oversea.Checked = True Then
                tempSQL += "and substring (a.part_id,3,1)='W '"
            End If

            '結尾
            tempSQL += "ORDER BY fail_mode, DataTime "

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")

            '開頭
            tempSQL = "SELECT Category,a.Bumptype as BumpingType, a.Part_Id, a.Lot_Id,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  "


            If ckMachine.Checked = True Then
                tempSQL += ",Machine_Id "
            End If

            tempSQL += "FROM WB_BinCode_Daily_Lot a with (nolock) ,MES.dbo.Lot_History b with (nolock)"
            'Default FailMode=0


            If yl.nLotList = 0 Then
                tempSQL += "WHERE a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ") AND a.Fail_Mode in ('" + yl.Fail_item + "') " _
                                                 & "and a.Lot_Id =b.Lot_Id  " _
                                                 & "and b.Station_Id in (" + ConvertStr2AddMark2(station(0)) + ") " _
                                                 & "and a.DataTime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and a.DataTime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "


            Else
                tempSQL += "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND a.Fail_Mode in ('" + yl.Fail_item + "')  " _
                                                 & "and a.Lot_Id =b.Lot_Id  " _
                                                 & "and b.Station_Id in (" + ConvertStr2AddMark2(station(0)) + ") "


            End If


            'FAI
            If ckFAI.Checked = False Then
                tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 "
            End If

            If cb_CR.Checked = False Then

                tempSQL += "and substring(a.lot_id,9,1)<>'Y' "
                tempSQL += "and substring(a.lot_id,9,1)<>'Z' "
                tempSQL += "and substring(a.Part_id,7,1)<>'V' "
            End If

            If Cb_Oversea.Checked = True Then
                tempSQL += "and substring (a.part_id,3,1)='W '"
            End If

            If txtMachine.Text <> "" And ckMachine.Checked = True Then
                tempSQL += "and Machine_Id in (" + ConvertStr2AddMark2(txtMachine.Text) + ") "
            End If


            '替換Category
            tempSQL = tempSQL.Replace("Category", "'PPS' as Category")

            '結尾
            tempSQL += "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id,DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime "

            If ckMachine.Checked = True Then
                tempSQL += ",Machine_Id ORDER BY Machine_Id,fail_mode,b.Station_Out_DateTime "
            Else
                tempSQL += "ORDER BY fail_mode,b.Station_Out_DateTime "
            End If



            '處理不同Bumping Type
            If yl.nType = 1 And yl.nLotList = 0 Then
                tempSQL = tempSQL.Replace("a.Part_Id in", "a.Bumptype in")
            End If


        End If

        '置換failmode
        If yl.nFailMode = 1 Then
            tempSQL = tempSQL.Replace("a.Fail_Mode in", "a.DefectCode in")
        End If

        '處理Lot Merge
        If ddlProduct.SelectedValue = "PPS" Then
            If cb_Lot_Merge.Checked = False Then
                tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
                If cb_SF.Checked = True Then
                    tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
                End If
                tempSQL = tempSQL.Replace("a.Bumptype as BumpingType", "a.BumpingType")
                tempSQL = tempSQL.Replace(",a.Bumptype,", ",a.Bumpingtype,")
                tempSQL = tempSQL.Replace("a.Bumptype in", "a.Bumpingtype in")

            Else
                If cb_SF.Checked = True Then
                    tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "vw_WB_BinCode_Daily_Lot_SF")
                End If

            End If




        ElseIf ddlProduct.SelectedValue = "PCB" Then
            cb_Lot_Merge.Checked = False
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
            If cb_SF.Checked = True Then
                tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
            End If
            tempSQL = tempSQL.Replace("a.Bumptype as BumpingType", "a.BumpingType")
            tempSQL = tempSQL.Replace(",a.Bumptype,", ",a.Bumpingtype,")
            tempSQL = tempSQL.Replace("a.Bumptype in", "a.Bumpingtype in")
        End If



        Return tempSQL
    End Function
    Private Function getRowDataSQL2_X(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nFailMode = 1 Then
            Dim bincode() As String
            bincode = yl.Fail_item.Split("_")
            yl.Fail_item = bincode(0)
        End If

        If yl.nStation = 0 Then
            '處理FVI
            '開頭

            tempSQL = "SELECT WW,WW as 'FVI WW',DataTime as 'FVI DataTime',DataTime as 'Station_Out_DateTime', a.Part_Id, Lot_Id,  Original_Input_QTY,c.Bumping_type,c.category from [dbo].[WB_BinCode_Daily_RawData] a with (nolock) ,[dbo].[Customer_Prodction_Mapping_BU_Rename] c with (nolock)   "

            'Default FailMode=0

            If yl.nLotList = 0 Then
                If yl.nType = 2 Then
                    tempSQL += "WHERE a.Part_Id=c.Part_Id and a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ")  " _
                                                                     & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "

                ElseIf yl.nType = 1 Then
                    tempSQL += "WHERE a.Part_Id=c.Part_Id and c.Bumping_type in (" + ConvertStr2AddMark2(yl.Part_ID) + ")  " _
                                                                    & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "

                End If
            Else
                tempSQL += "WHERE a.Part_Id=c.Part_Id and a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ")  "

            End If

            'FAI
            If ckFAI.Checked = False Then
                tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 "
            End If

            If cb_CR.Checked = False Then
                tempSQL += "and substring(a.lot_id,9,1)<>'Y' "
                tempSQL += "and substring(a.lot_id,9,1)<>'Z' "
                tempSQL += "and substring(a.Part_id,7,1)<>'V' "
            End If
            If Cb_Oversea.Checked = True Then
                tempSQL += "and substring (a.part_id,3,1)='W '"
            End If

            If ddlProduct.SelectedValue = "PCB" Then
                tempSQL += "and Yield_Category ='Total Yield'"
            End If
            '結尾
            tempSQL += "ORDER BY DataTime "

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")

            '開頭
            tempSQL = "SELECT a.Part_Id, a.Lot_Id,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, a.Original_Input_QTY, WW as 'FVI WW',DataTime as 'FVI DataTime' ,c.Bumping_Type, c.Category "

            If ckMachine.Checked = True Then
            tempSQL += ",Machine_Id "
            End If
            tempSQL += "FROM WB_BinCode_Daily_RawData a with (nolock) ,MES.dbo.Lot_History b with (nolock) ,Customer_Prodction_Mapping_BU_Rename c with (nolock) "
        'Default FailMode=0


        If yl.nLotList = 0 Then
            tempSQL += "WHERE a.Part_Id=c.Part_Id and a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ")  " _
                                             & "and a.Lot_Id =b.Lot_Id  " _
                                             & "and b.Station_Id in (" + ConvertStr2AddMark2(Station(0)) + ") " _
                                             & "and a.DataTime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and a.DataTime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' "


        Else
            tempSQL += "WHERE a.Part_Id=c.Part_Id and a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ")  " _
                                             & "and a.Lot_Id =b.Lot_Id  " _
                                             & "and b.Station_Id in (" + ConvertStr2AddMark2(Station(0)) + ") "


        End If


        'FAI
        If ckFAI.Checked = False Then
            tempSQL += "and ISNUMERIC(SUBSTRING(a.Part_id,2,1))=0 "
        End If

            If cb_CR.Checked = False Then
                tempSQL += "and substring(a.lot_id,9,1)<>'Y' "
                tempSQL += "and substring(a.lot_id,9,1)<>'Z' "
                tempSQL += "and substring(a.Part_id,7,1)<>'V' "
            End If

            If Cb_Oversea.Checked = True Then
                tempSQL += "and substring (a.part_id,3,1)='W '"
            End If

            If txtMachine.Text <> "" And ckMachine.Checked = True Then
                tempSQL += "and Machine_Id in (" + ConvertStr2AddMark2(txtMachine.Text) + ") "
            End If

            If ddlProduct.SelectedValue = "PCB" Then
                tempSQL += "and Yield_Category ='Total Yield'"
            End If


        '結尾
        tempSQL += "group by WW, DataTime, a.Part_Id, a.Lot_Id,a.Original_Input_QTY, b.station_id,b.Station_Out_DateTime,c.Bumping_Type,c.category  "


        If ckMachine.Checked = True Then
            tempSQL += ",Machine_Id ORDER BY Machine_Id,b.Station_Out_DateTime "
        Else
            tempSQL += "ORDER BY b.Station_Out_DateTime "
        End If



        '處理不同Bumping Type
        If yl.nType = 1 And yl.nLotList = 0 Then
            tempSQL = tempSQL.Replace("a.Part_Id in", "c.Bumping_type in")
        End If


        End If



        '處理Lot Merge
        If ddlProduct.SelectedValue = "PPS" Then
            If cb_Lot_Merge.Checked = False Then
                tempSQL = tempSQL.Replace("WB_BinCode_Daily_RawData", "WB_BinCode_Daily_RawData_NotMerge")

                If cb_SF.Checked = True Then
                    tempSQL = tempSQL.Replace("WB_BinCode_Daily_RawData_NotMerge", "vw_WB_BinCode_Daily_RawData_NotMerge_SF")
                End If
            Else
                If cb_SF.Checked = True Then
                    tempSQL = tempSQL.Replace("WB_BinCode_Daily_RawData", "vw_WB_BinCode_Daily_RawData_SF")
                End If
            End If
        ElseIf ddlProduct.SelectedValue = "PCB" Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_RawData", "PCB_Yield_Daily_RawData_NotMerge_Storage")

            If cb_SF.Checked = True Then
                tempSQL = tempSQL.Replace("PCB_Yield_Daily_RawData_NotMerge_Storage", "vw_PCB_Yield_Daily_RawData_NotMerge_Storage_SF")
            End If
        End If



        Return tempSQL
    End Function
    Private Function getRowDataSQL22(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nFailMode = 1 Then
            Dim bincode() As String
            bincode = yl.Fail_item.Split("_")
            yl.Fail_item = bincode(0)
        End If

        If yl.nStation = 0 Then
            '處理FVI
            '開頭
            tempSQL = "SELECT Category,a.Production_Type, a.Part_Id, a.Lot_Id,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                    & "FROM BinCode_Daily_Lot a "

            'Default FailMode=0

            If yl.nLotList = 0 Then
                If yl.nType = 2 Then
                    tempSQL += "WHERE a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ") )" _
                                                                     & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' AND (a.Fail_Mode = N'" + yl.Fail_item + "') "

                ElseIf yl.nType = 0 Then
                    tempSQL += "WHERE a.Production_Type = N'" + yl.Part_ID + "'  " _
                                                                    & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59'  AND (a.Fail_Mode = N'" + yl.Fail_item + "')"

                End If
            Else
                tempSQL += "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') "

            End If

            '結尾
            tempSQL += "ORDER BY DataTime "

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")

            '開頭
            tempSQL = "SELECT Category,Production_Type, a.Part_Id, a.Lot_Id,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                    & "FROM BinCode_Daily_Lot a ,MES.dbo.Lot_History b "
            'Default FailMode=0


            If yl.nLotList = 0 Then
                tempSQL += "WHERE a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + "))  " _
                                                 & "and a.Lot_Id =b.Lot_Id  " _
                                                 & "and b.Station_Id ='" + station(0) + "' " _
                                                 & "and Station_Out_DateTime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Station_Out_DateTime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                 & "group by WW, DataTime, Category, a.Part_Id,a.Production_Type, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime "

            Else
                tempSQL += "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                 & "and a.Lot_Id =b.Lot_Id  " _
                                                 & "and b.Station_Id ='" + station(0) + "' " _
                                                 & "group by WW, DataTime, Category, a.Part_Id,a.Production_Type, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime "

            End If



            '結尾
            tempSQL += "ORDER BY b.Station_Out_DateTime "


            '處理不同Bumping Type
            If yl.nType = 0 And yl.nLotList = 0 Then
                tempSQL = tempSQL.Replace("a.Part_Id in", "a.Production_Type in")
            End If


        End If

        '置換failmode
        If yl.nFailMode = 1 Then
            tempSQL = tempSQL.Replace("a.Fail_Mode = N", "a.DefectCode = N")
        End If

        '處理Lot Merge
        'If cb_Lot_Merge.Checked = False Then
        '    tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
        '    tempSQL = tempSQL.Replace("a.Bumptype as BumpingType", "a.BumpingType")
        '    tempSQL = tempSQL.Replace(",a.Bumptype,", ",a.Bumpingtype,")
        '    tempSQL = tempSQL.Replace("a.Bumptype =", "a.Bumpingtype =")
        'End If


        Return tempSQL
    End Function

    Private Function getRowDataSQL3(ByVal yl As YieldlossInfo, ByVal dt As DataTable) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nFailMode = 1 Then
            Dim bincode() As String
            bincode = yl.Fail_item.Split("_")
            yl.Fail_item = bincode(0)
        End If
        yl.Part_ID = yl.sExtenBumpingType

        Dim nStart As Double = -ddlADStart.SelectedValue * 7
        Dim nEnd As Double = ddlADEnd.SelectedValue * 7

        If yl.nStation = 0 Then
            '處理FVI
            '開頭
            tempSQL = "SELECT '' as TargetLot,Category,a.Bumptype as BumpingType, a.Part_Id, a.Lot_Id,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                    & "FROM WB_BinCode_Daily_Lot a "

            'Default FailMode=0

            tempSQL += "WHERE a.Bumptype in (" + ConvertStr2AddMark2(yl.Part_ID) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                            & "and Datatime >='" + CDate(dt.Rows(0).Item("FVI DataTime")).AddDays(nStart).ToString("yyyy/MM/dd") + "' and Datatime<'" + CDate(dt.Rows(dt.Rows.Count - 1).Item("FVI DataTime")).AddDays(nEnd).ToString("yyyy/MM/dd") + " 23:59:59' "


            '結尾
            tempSQL += "ORDER BY DataTime "

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")

            '開頭
            tempSQL = "SELECT '' as TargetLot,Category,a.Bumptype as BumpingType, a.Part_Id, a.Lot_Id,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                    & "FROM WB_BinCode_Daily_Lot a ,MES.dbo.Lot_History b "
            'Default FailMode=0

            tempSQL += "WHERE (a.Part_Id in (" + ConvertStr2AddMark2(yl.Part_ID) + ")) AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                             & "and a.Lot_Id =b.Lot_Id  " _
                                             & "and b.Station_Id in (" + ConvertStr2AddMark2(station(0)) + ") " _
                                             & "and Station_Out_DateTime >='" + CDate(dt.Rows(0).Item("Station_Out_DateTime")).AddDays(nStart).ToString("yyyy/MM/dd") + "' and Station_Out_DateTime<'" + CDate(dt.Rows(dt.Rows.Count - 1).Item("Station_Out_DateTime")).AddDays(nEnd).ToString("yyyy/MM/dd") + " 23:59:59' " _
                                             & "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime "



            '結尾
            tempSQL += "ORDER BY b.Station_Out_DateTime "


            '處理不同Bumping Type

            tempSQL = tempSQL.Replace("a.Part_Id in", "a.Bumptype in")



        End If

        '置換failmode
        If yl.nFailMode = 1 Then
            tempSQL = tempSQL.Replace("a.Fail_Mode = N", "a.DefectCode = N")
        End If

        '處理Lot Merge
        If cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
            tempSQL = tempSQL.Replace("a.Bumptype as BumpingType", "a.BumpingType")
            tempSQL = tempSQL.Replace(",a.Bumptype,", ",a.Bumpingtype,")
            tempSQL = tempSQL.Replace("a.Bumptype in", "a.Bumpingtype in")
        End If


        Return tempSQL
    End Function

    Private Function getRowDataSQL33(ByVal yl As YieldlossInfo, ByVal dt As DataTable) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nFailMode = 1 Then
            Dim bincode() As String
            bincode = yl.Fail_item.Split("_")
            yl.Fail_item = bincode(0)
        End If
        yl.Part_ID = yl.sExtenBumpingType

        Dim nStart As Double = -ddlADStart.SelectedValue * 7
        Dim nEnd As Double = ddlADEnd.SelectedValue * 7

        If yl.nStation = 0 Then
            '處理FVI
            '開頭
            tempSQL = "SELECT '' as TargetLot,Category,a.Production_Type, a.Part_Id, a.Lot_Id,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                    & "FROM BinCode_Daily_Lot a "

            'Default FailMode=0

            tempSQL += "WHERE (a.Production_Type = N'" + yl.Part_ID + "') AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                            & "and Datatime >='" + CDate(dt.Rows(0).Item("FVI DataTime")).AddDays(nStart).ToString("yyyy/MM/dd") + "' and Datatime<'" + CDate(dt.Rows(dt.Rows.Count - 1).Item("FVI DataTime")).AddDays(nEnd).ToString("yyyy/MM/dd") + " 23:59:59' "


            '結尾
            tempSQL += "ORDER BY DataTime "

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")

            '開頭
            tempSQL = "SELECT '' as TargetLot,Category,a.Production_Type, a.Part_Id, a.Lot_Id,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                    & "FROM BinCode_Daily_Lot a ,MES.dbo.Lot_History b "
            'Default FailMode=0

            tempSQL += "WHERE (a.Part_Id = N'" + yl.Part_ID + "') AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                             & "and a.Lot_Id =b.Lot_Id  " _
                                             & "and b.Station_Id in (" + ConvertStr2AddMark2(station(0)) + ") " _
                                             & "and Station_Out_DateTime >='" + CDate(dt.Rows(0).Item("Station_Out_DateTime")).AddDays(nStart).ToString("yyyy/MM/dd") + "' and Station_Out_DateTime<'" + CDate(dt.Rows(dt.Rows.Count - 1).Item("Station_Out_DateTime")).AddDays(nEnd).ToString("yyyy/MM/dd") + " 23:59:59' " _
                                             & "group by WW, DataTime, Category, a.Part_Id,a.Production_Type, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime "



            '結尾
            tempSQL += "ORDER BY b.Station_Out_DateTime "


            If yl.nType = 0 And yl.nLotList = 0 Then
                tempSQL = tempSQL.Replace("a.Part_Id =", "a.Production_Type =")
            End If



        End If

        '置換failmode
        If yl.nFailMode = 1 Then
            tempSQL = tempSQL.Replace("a.Fail_Mode = N", "a.DefectCode = N")
        End If



        Return tempSQL
    End Function

    Private Function getRowDataSQL(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        If yl.nStation = 0 Then
            '處理FVI

            'FailMode
            If yl.nFailMode = 0 Then
                '& ",Fail_Count_ByXoutScrap, Fail_ratio_ByXoutScrap  " _

                If yl.nLotList = 0 Then

                    If yl.nType = 2 Then
                        tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                                         & "FROM WB_BinCode_Daily_Lot a " _
                                                                         & "WHERE (a.Part_Id = N'" + yl.Part_ID + "') AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                                         & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                                         & "ORDER BY DataTime "
                    ElseIf yl.nType = 1 Then
                        tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                                        & "FROM WB_BinCode_Daily_Lot a " _
                                                                        & "WHERE (a.Bumptype = N'" + yl.Part_ID + "') AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                                        & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                                        & "ORDER BY DataTime "
                    End If

                Else
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                   & "FROM WB_BinCode_Daily_Lot a " _
                                                   & "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                   & "ORDER BY DataTime "
                End If

            Else
                'BinCode
                Dim bincode() As String
                bincode = yl.Fail_item.Split("_")

                If yl.nLotList = 0 Then

                    If yl.nType = 2 Then
                        tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                                         & "FROM WB_BinCode_Daily_Lot a " _
                                                                         & "WHERE (a.Part_Id = N'" + yl.Part_ID + "') AND (a.DefectCode = N'" + bincode(0) + "') " _
                                                                         & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                                         & "ORDER BY DataTime "
                    ElseIf yl.nType = 1 Then
                        tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                                         & "FROM WB_BinCode_Daily_Lot a " _
                                                                         & "WHERE (a.Bumptype = N'" + yl.Part_ID + "') AND (a.DefectCode = N'" + bincode(0) + "') " _
                                                                         & "and Datatime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Datatime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                                         & "ORDER BY DataTime "
                    End If

                Else
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,datepart(yyyy,DataTime)*100+datepart(ww, DataTime) as WW, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                  & "FROM WB_BinCode_Daily_Lot a " _
                                                  & "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                  & "ORDER BY DataTime "
                End If

            End If

        Else
            '處理其他Process
            Dim station() As String
            station = yl.sStation.Split(" ")
            'FailMode
            If yl.nFailMode = 0 Then

                If yl.nLotList = 0 Then
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                     & "FROM WB_BinCode_Daily_Lot a ,MES.dbo.Lot_History b " _
                                                     & "WHERE (a.Part_Id = N'" + yl.Part_ID + "') AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                     & "and a.Lot_Id =b.Lot_Id  " _
                                                     & "and b.Station_Id ='" + station(0) + "' " _
                                                     & "and Station_Out_DateTime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Station_Out_DateTime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                     & "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime " _
                                                     & "ORDER BY b.Station_Out_DateTime "
                Else
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                     & "FROM WB_BinCode_Daily_Lot a ,MES.dbo.Lot_History b " _
                                                     & "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.Fail_Mode = N'" + yl.Fail_item + "') " _
                                                     & "and a.Lot_Id =b.Lot_Id  " _
                                                     & "and b.Station_Id ='" + station(0) + "' " _
                                                     & "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime " _
                                                     & "ORDER BY b.Station_Out_DateTime "
                End If


                If yl.nType = 1 And yl.nLotList = 0 Then
                    tempSQL = tempSQL.Replace("a.Part_Id =", "a.Bumptype =")
                End If


            Else
                Dim bincode() As String
                bincode = yl.Fail_item.Split("_")

                If yl.nLotList = 0 Then
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime'  " _
                                                     & "FROM WB_BinCode_Daily_Lot a ,MES.dbo.Lot_History b " _
                                                     & "WHERE (a.Part_Id = N'" + yl.Part_ID + "') AND (a.DefectCode = N'" + bincode(0) + "') " _
                                                     & "and a.Lot_Id =b.Lot_Id  " _
                                                     & "and b.Station_Id ='" + station(0) + "' " _
                                                     & "and Station_Out_DateTime >='" + yl.dtStart.ToString("yyyy/MM/dd") + "' and Station_Out_DateTime<'" + yl.dtEnd.ToString("yyyy/MM/dd") + " 23:59:59' " _
                                                     & "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime " _
                                                     & "ORDER BY b.Station_Out_DateTime "
                Else
                    tempSQL = "SELECT Category, a.Part_Id,a.Bumptype as BumpingType,b.Station_Id,'" + yl.sStationC + "' as Station_C,datepart(yyyy,b.Station_Out_DateTime)*100+datepart(ww, b.Station_Out_DateTime)  as WW,b.Station_Out_DateTime, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio as Fail_Rate,WW as 'FVI WW',DataTime as 'FVI DataTime' " _
                                                      & "FROM WB_BinCode_Daily_Lot a ,MES.dbo.Lot_History b " _
                                                      & "WHERE a.Lot_Id in (" + ConvertStr2AddMark(yl.sLotList) + ") AND (a.DefectCode = N'" + bincode(0) + "') " _
                                                      & "and a.Lot_Id =b.Lot_Id  " _
                                                      & "and b.Station_Id ='" + station(0) + "' " _
                                                      & "group by WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio  ,b.station_id,b.Station_Out_DateTime " _
                                                      & "ORDER BY b.Station_Out_DateTime "
                End If

                If yl.nType = 1 And yl.nLotList = 0 Then
                    tempSQL = tempSQL.Replace("a.Part_Id =", "a.Bumptype =")
                End If

            End If

        End If


        If cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
            tempSQL = tempSQL.Replace("a.Bumptype as BumpingType", "a.BumpingType")
            tempSQL = tempSQL.Replace(",a.Bumptype,", ",a.Bumpingtype,")
            tempSQL = tempSQL.Replace("a.Bumptype =", "a.Bumpingtype =")
        End If


        Return tempSQL
    End Function

    Private Function getRowDataSQL1(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""


        tempSQL = "SELECT WW, DataTime, Category, a.Part_Id,a.Bumptype, a.Lot_Id, DefectCode, Fail_Mode, BinCode, a.Original_Input_QTY, Fail_Count, Fail_ratio,  " _
        & "Fail_Count_ByXoutScrap, Fail_ratio_ByXoutScrap "

        If yl.nStation = 1 Then
            tempSQL += ",b.station_id,b.Station_Out_DateTime  "
        End If

        tempSQL += "FROM WB_BinCode_Daily_Lot a "


        If yl.nStation = 1 Then
            tempSQL += " ,MES.dbo.Lot_History b "
        End If

        tempSQL += "WHERE ("
        tempSQL += "a.Part_Id = '" + yl.Part_ID + "' "
        tempSQL += ") AND (a.BinCode = '" + yl.Fail_item + "') "

        If yl.nStation = 1 Then
            tempSQL += "and a.Lot_Id =b.Lot_Id "
            tempSQL += "and b.Station_Id ='" + yl.sStation + "' "
        End If


        tempSQL += "ORDER BY DataTime "



        Return tempSQL
    End Function



    Private Function Get_PartID() As String

        Dim oPartID As New StringBuilder
        'oPartID.Append("'" & listB_PartSource.Items(0).Text & "'")
        Try
            If listB_PartShow.Items.Count > 0 Then
                oPartID.Clear()
                For I As Integer = 0 To listB_PartShow.Items.Count - 1
                    If I = 0 Then
                        oPartID.Append("'" & listB_PartShow.Items(I).Text & "'")
                    Else
                        oPartID.Append(",'" & listB_PartShow.Items(I).Text & "'")
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

        Return oPartID.ToString()
    End Function

    Private Function Get_AD_PartID() As String

        Dim oPartID As New StringBuilder
        'oPartID.Append("'" & listB_PartSource.Items(0).Text & "'")
        Try
            If lst_Ad_Target.Items.Count > 0 Then
                oPartID.Clear()
                For I As Integer = 0 To lst_Ad_Target.Items.Count - 1
                    If I = 0 Then
                        oPartID.Append("'" & lst_Ad_Target.Items(I).Text & "'")
                    Else
                        oPartID.Append(",'" & lst_Ad_Target.Items(I).Text & "'")
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

        Return oPartID.ToString()
    End Function

#Region "Control event"
    Protected Sub ddlCustomer_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlCustomer.SelectedIndexChanged
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False

        Try
            rb_ProductPart.Items(0).Enabled = True
            rb_ProductPart.Items(1).Enabled = True

            conn.Open()

            If ddlProduct.Text = "WB" Then
                ' -- Part ID --
                rb_ProductPart.SelectedIndex = 1
                rb_ProductPart.Items(0).Enabled = False
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                If ddlCustomer.Text <> "All" Then
                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                End If
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by part_id order by part_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                listB_PartSource.Items.Clear()
                listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                Next
            Else
                ' -- Production ID --
                rb_ProductPart.SelectedIndex = 0
                rb_ProductPart.Items(1).Enabled = False
                'rb_ProductPart.Items(0).Enabled = True
                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by production_id order by production_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                listB_PartSource.Items.Clear()
                listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
                Next
            End If


            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click

        Dim sfilter As String = txt_YieldlossFilter.Text

        If RadioButtonList2.SelectedIndex = 0 Then
            Yieldloss_FailMode(sfilter)
        Else
            Yieldloss_MISDefect(sfilter)
        End If



        'Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        'Dim sqlStr As String = ""
        'Dim myDT As DataTable
        'Dim myAdapter As SqlDataAdapter

        'Try
        '    conn.Open()

        '    If rb_ProductPart.SelectedIndex = 1 And txt_YieldlossFilter.Text <> "" Then
        '        sqlStr = "select distinct FailMode from WB_BinCode_FailMode_Customer_Mapping where FailMode like '%" + txt_YieldlossFilter.Text + "%'"

        '    Else
        '        sqlStr = "select distinct FailMode from WB_BinCode_FailMode_Customer_Mapping "

        '    End If
        '    myAdapter = New SqlDataAdapter(sqlStr, conn)
        '    myDT = New DataTable
        '    myAdapter.Fill(myDT)

        '    lst_Yieldloss_Source.Items.Clear()
        '    lst_Yieldloss_Source.Items.Clear()
        '    For i As Integer = 0 To myDT.Rows.Count - 1
        '        lst_Yieldloss_Source.Items.Add(myDT.Rows(i)("FailMode").ToString())
        '    Next

        '    conn.Close()
        'Catch ex As Exception

        'Finally
        '    If conn.State = ConnectionState.Open Then
        '        conn.Close()
        '    End If
        'End Try
    End Sub


    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles btn_Filter_Part.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sfilter As String = TextBox1.Text

        sqlStr = "select distinct Part_id from " + confTable + " where 1=1 and Category='" + ddlProduct.SelectedValue + "'"
        If ddlCustomer.SelectedValue <> "All" Then
            sqlStr += " and customer_id ='" + ddlCustomer.SelectedValue + "' "
        End If

        If sfilter <> "" Then
            sqlStr += " and Part_id like '%" + sfilter + "%' "
        End If

        sqlStr += " order by Part_id"

        If rb_ProductPart.SelectedIndex = 0 Then

            sqlStr = sqlStr.Replace("Part_id", "Production_Id")
        End If

        If rb_ProductPart.SelectedIndex = 1 Then

            sqlStr = sqlStr.Replace("Part_id", "Bumping_Type")
        End If

        myAdapter = New SqlDataAdapter(sqlStr, conn)
        myAdapter.SelectCommand.CommandTimeout = 3600
        myDT = New DataTable
        myAdapter.Fill(myDT)

        listB_PartSource.Items.Clear()
        listB_PartShow.Items.Clear()
        For i As Integer = 0 To myDT.Rows.Count - 1
            If rb_ProductPart.SelectedIndex = 0 Then
                listB_PartSource.Items.Add(myDT.Rows(i)("Production_Id").ToString())
            ElseIf rb_ProductPart.SelectedIndex = 1 Then
                listB_PartSource.Items.Add(myDT.Rows(i)("Bumping_Type").ToString())
            Else
                listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
            End If

        Next

        Exit Sub





        If ddlProduct.SelectedValue = "WB" Then



        End If


        ' --- 檢查有無 Group Part, 沒有就呈現 Part ID ---
        sqlStr = "select production_id from " + confTable + " where 1=1 "
        If ddlCustomer.Text <> "All" Then
            sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
        End If
        sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "

        If TextBox1.Text <> "" Then
            If rb_ProductPart.SelectedIndex = 0 Then
                sqlStr += "and production_id like '%" + TextBox1.Text + "%' "
            ElseIf rb_ProductPart.SelectedIndex = 1 Then
                sqlStr += "and Part_id like '%" + TextBox1.Text + "%' "
            End If
        End If
        sqlStr += "and yield_function=1 "
        'sqlStr += "and production_id <> '' "
        sqlStr += "group by production_id"
        myAdapter = New SqlDataAdapter(sqlStr, conn)
        myAdapter.SelectCommand.CommandTimeout = 3600
        myDT = New DataTable
        myAdapter.Fill(myDT)

        ' --- Group ---
        If myDT.Rows.Count > 0 Then
            'If Not IsDBNull(myDT.Rows(0)(0)) Then
            If myDT.Rows(0)(0).ToString() <> "" Then
                If ddlProduct.SelectedValue.ToUpper() = "WB" Then
                    rb_ProductPart.SelectedIndex = 1
                    rb_ProductPart.Items(0).Enabled = False
                    ' 轉向 Part ID
                    rb_ProductPart.SelectedIndex = 1
                    sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
                    End If
                    sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "

                    sqlStr += "and yield_function=1 "
                    If TextBox1.Text <> "" Then
                        If rb_ProductPart.SelectedIndex = 0 Then
                            sqlStr += "and production_id like '%" + TextBox1.Text + "%' "
                        ElseIf rb_ProductPart.SelectedIndex = 1 Then
                            sqlStr += "and Part_id like '%" + TextBox1.Text + "%' "
                        End If
                    End If
                    sqlStr += "group by Part_id, Memo order by Part_id"
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myAdapter.SelectCommand.CommandTimeout = 3600
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0, "Part_id", "Memo")
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                    Next
                Else
                    rb_ProductPart.SelectedIndex = 0
                    rb_ProductPart.Items(0).Enabled = True
                    ' 有就秀 Group
                    rb_ProductPart.SelectedIndex = 0
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
                    Next
                End If
            Else
                ' 轉向 Part ID
                rb_ProductPart.SelectedIndex = 1
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                If ddlCustomer.Text <> "All" Then
                    sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
                End If
                sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "

                sqlStr += "and yield_function=1 "
                If TextBox1.Text <> "" Then
                    If rb_ProductPart.SelectedIndex = 0 Then
                        sqlStr += "and production_id like '%" + TextBox1.Text + "%' "
                    ElseIf rb_ProductPart.SelectedIndex = 1 Then
                        sqlStr += "and Part_id like '%" + TextBox1.Text + "%' "
                    End If
                End If
                sqlStr += "group by Part_id, Memo order by Part_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0, "Part_id", "Memo")
                listB_PartSource.Items.Clear()
                listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                Next
            End If
        Else
            ' 轉向 Part ID
            rb_ProductPart.SelectedIndex = 1
            rb_ProductPart.Items(0).Enabled = False
            sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
            sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "
            If ddlCustomer.Text <> "All" Then
                sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
            End If

            sqlStr += "and yield_function=1 "
            If TextBox1.Text <> "" Then
                If rb_ProductPart.SelectedIndex = 0 Then
                    sqlStr += "and production_id like '%" + TextBox1.Text + "%' "
                ElseIf rb_ProductPart.SelectedIndex = 1 Then
                    sqlStr += "and Part_id like '%" + TextBox1.Text + "%' "
                End If
            End If
            sqlStr += "group by Part_id, Memo order by Part_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            'UtilObj.FillController(myDT, ddlPart, 1, "Part_id", "Memo")
            listB_PartSource.Items.Clear()
            listB_PartShow.Items.Clear()
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
            Next
        End If
    End Sub

    Protected Sub RadioButtonList2_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList2.SelectedIndexChanged
        lst_Yieldloss_Target.Items.Clear()
        Button3.Enabled = True
        If RadioButtonList2.SelectedIndex = 0 Then
            Yieldloss_FailMode("")
        Else
            Yieldloss_MISDefect("")
        End If
    End Sub

    Protected Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        moveList(lst_Yieldloss_Source, lst_Yieldloss_Target)

        ListBoxSort(lst_Yieldloss_Source, True)
        ListBoxSort(lst_Yieldloss_Target, True)

        If lst_Yieldloss_Target.Items.Count > 1 Then
            Button3.Enabled = False
        End If


    End Sub

    Protected Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click
        moveList(lst_Yieldloss_Target, lst_Yieldloss_Source)

        ListBoxSort(lst_Yieldloss_Source, True)
        ListBoxSort(lst_Yieldloss_Target, True)

        If lst_Yieldloss_Target.Items.Count <= 1 Then
            Button3.Enabled = True
        End If

    End Sub

    Protected Sub but_PartRight_Click(sender As Object, e As System.EventArgs) Handles but_PartRight.Click
        moveList(listB_PartSource, listB_PartShow)

        ListBoxSort(listB_PartSource, True)
        ListBoxSort(listB_PartShow, True)

        If listB_PartShow.Items.Count > 1 And rb_ProductPart.SelectedIndex = 1 Then
            but_PartRight.Enabled = False
        ElseIf listB_PartShow.Items.Count > 10 And rb_ProductPart.SelectedIndex = 2 Then
            but_PartRight.Enabled = False
        End If
    End Sub

    Protected Sub but_PartLeft_Click(sender As Object, e As System.EventArgs) Handles but_PartLeft.Click
        moveList(listB_PartShow, listB_PartSource)

        ListBoxSort(listB_PartSource, True)
        ListBoxSort(listB_PartShow, True)

        If listB_PartShow.Items.Count < 2 And rb_ProductPart.SelectedIndex = 1 Then
            but_PartRight.Enabled = True
        ElseIf listB_PartShow.Items.Count < 10 And rb_ProductPart.SelectedIndex = 2 Then
            but_PartRight.Enabled = True
        End If
    End Sub

#End Region

#Region "Function"
    Private Sub Customer()
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()
            'sqlStr = "select customer_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
            'sqlStr += "and fail_function=1 "
            'sqlStr += "and Category='" + ddlProduct.SelectedValue + "' "
            'sqlStr += "group by customer_id order by customer_id"

            If ddlProduct.SelectedValue = "PPS" Then
                sqlStr = "select [Assfct] from MES.[dbo].[ProductInfo] where bu='PPS' group by assfct order by assfct"
            ElseIf ddlProduct.SelectedValue = "PCB" Then
                sqlStr = "select customer_name from MES.[dbo].[ProductInfo] where bu='pcb' "
                sqlStr += "and customer_name <> 'NULL' "
                sqlStr += "group by customer_name order by customer_name "
            Else
                sqlStr = "select customer_name from MES.[dbo].[ProductInfo] where bu='abfs' "
                sqlStr += "and customer_name <> 'NULL' "
                sqlStr += "group by customer_name order by customer_name "
            End If

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlCustomer, 1)

            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub

    Private Sub ExtenType()
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()
            If ddlProduct.SelectedValue = "PPS" Then
                RadioButtonList3.SelectedIndex = 1
                RadioButtonList3.Items(0).Enabled = False
                RadioButtonList3.Items(1).Enabled = True

                sqlStr = "select Bumping_Type from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                If ddlCustomer.Text <> "All" Then
                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                End If
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by Bumping_Type "
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                lst_Ad_Source.Items.Clear()
                lst_Ad_Target.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    lst_Ad_Source.Items.Add(myDT.Rows(i)("Bumping_Type").ToString())
                Next
            Else
                RadioButtonList3.SelectedIndex = 0
                RadioButtonList3.Items(0).Enabled = True
                RadioButtonList3.Items(1).Enabled = False


                sqlStr = "select Production_ID from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                If ddlCustomer.Text <> "All" Then
                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                End If
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by Production_ID order by Production_ID "
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                lst_Ad_Source.Items.Clear()
                lst_Ad_Target.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    lst_Ad_Source.Items.Add(myDT.Rows(i)("Production_ID").ToString())
                Next
            End If








            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub

    Private Sub Product()
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter


        Try
            conn.Open()
            If ddlProduct.SelectedValue = "PPS" Then
                rb_ProductPart.Items(0).Enabled = False
                rb_ProductPart.Items(1).Enabled = True


                If rb_ProductPart.SelectedIndex = 2 Then

                    sqlStr = "select part_no as part_id from Customer_Prodction_Mapping_BU_Rename a ,MES.[dbo].[ProductInfo] b where bu='PPS'  and a.Part_Id=b.Part_No  "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and Assfct='" + ddlCustomer.SelectedValue + "' "
                    End If
                    If Cb_Oversea.Checked = True Then
                        sqlStr += "and substring (part_no,3,1)='W' "
                    End If
                    sqlStr += "group by part_no order by part_no"


                    'sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                    'sqlStr += "and fail_function=1 "
                    'If ddlCustomer.Text <> "All" Then
                    '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    'End If
                    'sqlStr += "and category='" + ddlProduct.SelectedValue + "' "

                    'If Cb_Oversea.Checked = True Then
                    '    sqlStr += "and substring (part_id,3,1)='W' "
                    'End If
                    'sqlStr += "group by part_id order by part_id"
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()

                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                    Next
                ElseIf rb_ProductPart.SelectedIndex = 1 Then
                    sqlStr = "select Bumping_Type from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                    sqlStr += "and fail_function=1 "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    End If
                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                    sqlStr += "group by Bumping_Type "
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("Bumping_Type").ToString())
                    Next
                End If

            ElseIf ddlProduct.SelectedValue = "PCB" Then
                rb_ProductPart.Items(0).Enabled = False
                rb_ProductPart.Items(1).Enabled = True


                If rb_ProductPart.SelectedIndex = 2 Then

                    sqlStr = "select part_no as part_id from Customer_Prodction_Mapping_BU_Rename a ,MES.[dbo].[ProductInfo] b where bu='PCB'  and a.Part_Id=b.Part_No  "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_name='" + ddlCustomer.SelectedValue + "' "
                    End If
                    'If Cb_Oversea.Checked = True Then
                    '    sqlStr += "and substring (part_no,3,1)='W' "
                    'End If
                    sqlStr += "group by part_no order by part_no"

                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()

                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                    Next
                ElseIf rb_ProductPart.SelectedIndex = 1 Then
                    sqlStr = "select Bumping_Type from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                    sqlStr += "and fail_function=1 "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    End If
                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                    sqlStr += "group by Bumping_Type "
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("Bumping_Type").ToString())
                    Next
                End If

            Else
                rb_ProductPart.Items(0).Enabled = True
                rb_ProductPart.Items(1).Enabled = False


                If rb_ProductPart.SelectedIndex = 2 Then
                    sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                    sqlStr += "and fail_function=1 "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    End If
                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                    sqlStr += "group by part_id order by part_id"
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                    Next
                ElseIf rb_ProductPart.SelectedIndex = 0 Then
                    sqlStr = "select Production_ID from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                    sqlStr += "and fail_function=1 "
                    If ddlCustomer.Text <> "All" Then
                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    End If
                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                    sqlStr += "group by Production_ID order by production_id "
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("Production_ID").ToString())
                    Next
                End If
            End If


            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub

    Private Sub Yieldloss_FailMode(ByVal Filter As String)
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()

            If ddlProduct.SelectedValue = "PPS" Then

            End If
            If rb_ProductPart.SelectedIndex >= 1 Then
                If rb_ProductPart.SelectedIndex >= 1 And Filter <> "" Then
                    sqlStr = "select distinct FailMode from WB_BinCode_FailMode_Customer_Mapping where FailMode like '%" + Filter + "%'"

                Else
                    sqlStr = "select distinct FailMode from WB_BinCode_FailMode_Customer_Mapping "

                End If


                If rb_ProductPart.SelectedIndex >= 1 And Filter <> "" Then
                    sqlStr = "select distinct FailMode from vw_BinCode_FailMode_Customer_Mapping_BU_Rename where category='" + ddlProduct.SelectedValue + "' and FailMode like '%" + Filter + "%'"

                Else
                    sqlStr = "select distinct FailMode from vw_BinCode_FailMode_Customer_Mapping_BU_Rename where category='" + ddlProduct.SelectedValue + "'"

                End If

                'sqlStr += "order by FailMode_Id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                lst_Yieldloss_Source.Items.Clear()
                lst_Yieldloss_Source.Items.Clear()
                'Dim a As ListItem
                'a.Text = "sss"
                'a.Value = "ccc"

                For i As Integer = 0 To myDT.Rows.Count - 1
                    lst_Yieldloss_Source.Items.Add(myDT.Rows(i)("FailMode").ToString())

                    'lst_Yieldloss_Source.Items.Add(a)
                Next
            End If

            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub

    Private Sub Yieldloss_MISDefect(ByVal Filter As String)
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()
            If rb_ProductPart.SelectedIndex >= 1 Then

                If ddlProduct.SelectedValue = "PPS" Then
                    If rb_ProductPart.SelectedIndex >= 1 And Filter <> "" Then
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where  category ='WB' and MF_Stage <>'FVI' and DefectCode_Id+'_'+ Bincode like '%" + Filter + "%' order by DefectCode_Id "
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where  category ='WB' and MF_Stage ='inline' and DefectCode_Id+'_'+ Bincode like '%" + Filter + "%' order by DefectCode_Id "

                    Else
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where category ='WB' and MF_Stage <>'FVI' order by DefectCode_Id "
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where category ='WB' and MF_Stage ='inline' order by DefectCode_Id "
                    End If
                ElseIf ddlProduct.SelectedValue = "PCB" Then
                    If rb_ProductPart.SelectedIndex >= 1 And Filter <> "" Then
                        'sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where  category ='PCB' and MF_Stage <>'FVI' and DefectCode_Id+'_'+ Bincode like '%" + Filter + "%' order by DefectCode_Id "
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where  category ='PCB'  and DefectCode_Id+'_'+ Bincode like '%" + Filter + "%' order by DefectCode_Id "

                    Else
                        'sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where category ='PCB' and MF_Stage <>'FVI' order by DefectCode_Id "
                        sqlStr = "select Bincode, DefectCode_Id, DefectCode_Id+'_'+ Bincode  as MISDefect from dbo.view_BinCode where category ='PCB'  order by DefectCode_Id "
                    End If

                End If


                'sqlStr += "order by FailMode_Id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
                lst_Yieldloss_Source.Items.Clear()
                'lst_Yieldloss_Target.Items.Clear()

                Dim misdefect As ListItem

                For i As Integer = 0 To myDT.Rows.Count - 1
                    misdefect = New ListItem
                    misdefect.Text = myDT.Rows(i)("MISDefect").ToString()
                    misdefect.Value = myDT.Rows(i)("DefectCode_Id").ToString()

                    'lst_Yieldloss_Source.Items.Add(myDT.Rows(i)("MISDefect").ToString())
                    lst_Yieldloss_Source.Items.Add(misdefect)
                Next
            End If

            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub

    Private Sub Station(ByVal Part As String)
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()
            If rb_ProductPart.SelectedIndex = 2 Then  'By料號
                If rb_ProductPart.SelectedIndex = 2 And Part <> "" Then
                    sqlStr = "SELECT Station_Id,Station_Name_C FROM MES.dbo.PartStationMap WHERE (Part_No =" + Part + ")  ORDER BY SEQ"
                    sqlStr = "SELECT Station_Id,Station_Name_C,Parallel_Station_Id FROM Satation_SEQ_Para_Config WHERE Category='WB'  ORDER BY SEQ"

                Else
                    If ddlProduct.SelectedValue = "PPS" Then
                        sqlStr = "SELECT Station_Id,Station_Name_C,Parallel_Station_Id FROM Satation_SEQ_Para_Config WHERE Category='WB'  ORDER BY SEQ"
                    Else
                        If ddlProduct.SelectedValue = "CPU" Then
                            sqlStr = "SELECT Station_Id,Station_Name_C,'' as Parallel_Station_Id FROM MES.dbo.PartStationMap WHERE (Part_No ='FCS116A')  ORDER BY SEQ"
                        Else
                            sqlStr = "SELECT Station_Id,Station_Name_C,'' as Parallel_Station_Id FROM MES.dbo.PartStationMap WHERE (Part_No ='FCB305A')  ORDER BY SEQ"
                        End If
                    End If


                End If
            ElseIf rb_ProductPart.SelectedIndex = 1 Then 'By產品

                If ddlProduct.SelectedValue = "PPS" Then
                    sqlStr = "SELECT Station_Id,Station_Name_C,Parallel_Station_Id FROM Satation_SEQ_Para_Config WHERE Category='WB'  ORDER BY SEQ"
                Else
                    If ddlProduct.SelectedValue = "CPU" Then
                        sqlStr = "SELECT Station_Id,Station_Name_C,'' as Parallel_Station_Id FROM MES.dbo.PartStationMap WHERE (Part_No ='FCS116A')  ORDER BY SEQ"
                    Else
                        sqlStr = "SELECT Station_Id,Station_Name_C,'' as Parallel_Station_Id FROM MES.dbo.PartStationMap WHERE (Part_No ='FCB305A')  ORDER BY SEQ"
                    End If

                End If

            End If


            sqlStr = "select Station_Id,Station_Name_C,parallel_station as Parallel_Station_Id from [MES].[dbo].[ParalleStage] order by Station_Id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
                'UtilObj.FillController(myDT, ddlPart, 0)
            ddlStation.Items.Clear()
            Dim MS As String
            Dim PS As String

            For i As Integer = 0 To myDT.Rows.Count - 1
                MS = myDT.Rows(i)("Station_Id").ToString()
                PS = myDT.Rows(i)("Parallel_Station_Id").ToString()
                If PS <> "" Then
                    If Right(PS, 1) = "," Then
                        PS = PS.Substring(0, PS.Length - 1)
                        End If
                    MS = MS + "," + PS
                    End If

                Dim li As New ListItem(myDT.Rows(i)("Station_Id").ToString() + " " + myDT.Rows(i)("Station_Name_C").ToString(), MS)
                    'ddlStation.Items.Add(myDT.Rows(i)("Station_Id").ToString() + " " + myDT.Rows(i)("Station_Name_C").ToString())
                ddlStation.Items.Add(li)
            Next

            If myDT.Rows.Count > 1 Then
                ddlStation.Enabled = True
            Else
                ddlStation.Enabled = False
                End If


            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        ' -- Customer ID --

    End Sub


    Private Sub moveList(ByRef sourceList As ListBox, ByRef destList As ListBox)

        Dim sourceAry_text As New ArrayList()
        Dim sourceAry_value As New ArrayList()

        Dim DestAry_text As New ArrayList()
        Dim DestAry_value As New ArrayList()

        For i As Integer = 0 To sourceList.Items.Count - 1
            If sourceList.Items(i).Selected Then
                'DestAry.Add(sourceList.Items(i).Value)
                DestAry_value.Add(sourceList.Items(i).Value)
                DestAry_text.Add(sourceList.Items(i).Text)
            Else
                sourceAry_text.Add(sourceList.Items(i).Text)
                sourceAry_value.Add(sourceList.Items(i).Value)
            End If
        Next
        sourceList.Items.Clear()

        Dim sList As ListItem
        Dim tList As ListItem

        For i As Integer = 0 To sourceAry_text.Count - 1
            sList = New ListItem
            sList.Text = sourceAry_text(i).ToString()
            sList.Value = sourceAry_value(i).ToString()
            sourceList.Items.Add(sList)
        Next
        For i As Integer = 0 To DestAry_text.Count - 1
            tList = New ListItem
            tList.Text = DestAry_text(i).ToString()
            tList.Value = DestAry_value(i).ToString()
            destList.Items.Add(tList)
        Next

    End Sub

    Private Sub ListBoxSort(ByVal lbx As ListBox, ByVal ASC As Boolean)
        '利用sortedlist 類為listbox排序 
        Dim slist As SortedList
        If ASC = True Then
            slist = New SortedList()
        Else
            slist = New SortedList(New DecComparer())
        End If


        For i As Integer = 0 To lbx.Items.Count - 1
            '將listbox內容逐項複製到sortedlist物件中
            If slist.Contains(lbx.Items(i).Text) = False Then
                slist.Add(lbx.Items(i).Text, lbx.Items(i).Value)
            End If
        Next

        lbx.Items.Clear()
        '清空原listbox
        For Each obj As DictionaryEntry In slist
            Dim myit As New ListItem()
            myit.Text = obj.Key.ToString()
            myit.Value = obj.Value.ToString()
            '再重新將sortlist集合複製回listbox，這樣，複製回來的陣列是按值排序過的
            lbx.Items.Add(myit)
        Next
    End Sub

    Public Class DecComparer
        Implements IComparer
        Dim myComapar As CaseInsensitiveComparer = New CaseInsensitiveComparer()
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Return myComapar.Compare(y, x)
        End Function
    End Class

    Public Function ConvertStr2AddMark2(ByVal temp As String) As String
        Dim sValue As String = ""

        If temp Is Nothing Then
            Return ""
        End If
        Dim sSetting As String() = temp.Split(",")
        For i As Integer = 0 To sSetting.Length - 1
            If sValue = "" Then
                sValue = "'" + Trim(sSetting(i)) + "'"
            Else
                sValue += ",'" + Trim(sSetting(i)) + "'"
            End If
        Next
        Return sValue
    End Function

    Public Function ConvertStr2AddMark(ByVal temp As String) As String
        Dim sValue As String = ""
        Dim sTemp As String

        If temp Is Nothing Then
            Return ""
        End If

        If temp <> "" Then
            Dim sSetting As String() = temp.Split(",")
            For i As Integer = 0 To sSetting.Length - 1

                If cb_Lot_Merge.Checked = True Then
                    sTemp = sSetting(i)
                    sTemp = sTemp.Substring(0, 11)
                    sTemp = sTemp + "1"
                Else
                    sTemp = sSetting(i)
                End If


                If sValue = "" Then

                    sValue = "'" + sTemp + "'"
                Else
                    sValue += ",'" + sTemp + "'"
                End If
            Next
        End If

        Return sValue
    End Function

#End Region



    Protected Sub rbl_Station_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rbl_Station.SelectedIndexChanged
        If rbl_Station.SelectedIndex = 1 Then
            Dim sPart As String = Get_PartID()
            Station(sPart)
            TextBox3.Enabled = True
            Button1.Enabled = True
            chkParallel.Enabled = True
            chkRebuilt.Enabled = True

        Else
            ddlStation.Enabled = False
            TextBox3.Enabled = False
            Button1.Enabled = False

            chkParallel.Enabled = False
            chkRebuilt.Enabled = False
        End If

    End Sub

    Protected Sub but_Uupload_Click(sender As Object, e As System.EventArgs) Handles but_Uupload.Click
        Dim file1 As HttpPostedFile = uf_UfilePath.PostedFile
        Dim sPath As String = ""
        Try


            If file1.ContentLength <> 0 Then
                Dim filesplit() As String = Split(file1.FileName, "\")
                Dim filename As String = filesplit(filesplit.Length - 1)
                Dim apppqth As String = Request.PhysicalApplicationPath + "upload\"

                'Dim sw As StreamWriter = File.CreateText(apppqth + filename)
                If File.Exists(apppqth + filename) Then
                    File.Delete(apppqth + filename)
                End If

                file1.SaveAs(apppqth + filename)
                'file1.SaveAs(Server.MapPath(filename))
                sPath = apppqth + filename 'Server.MapPath(filename)

                Dim temp As String = ""

                TextBox2.Text = LoadExcelToText(sPath)
                'txtLotID.Text = GetSettingValue(sPath, "Lot_Info", "Lot_ID")

                File.Delete(sPath)
                lab_wait.Text = "File Load Success！！"
                '   ShowAlert(webMessage, "File Load Success！！ ")
            Else
                lab_wait.Text = "File Load Fail！！"
                '  ShowAlert(webMessage, "File Load Fail, Please try it again！！ ")
            End If
        Catch ex As Exception
            TextBox2.Text = ""
            lab_wait.Text = ex.Message
            'ShowAlert(webMessage, ex.Message)
        End Try
    End Sub
    Private Function LoadExcelToText(ByVal FilePath As String) As String
        Dim sResult As String = ""
        Dim dt As DataTable
        Dim myAdpt As OleDb.OleDbDataAdapter
        Dim conn As New OleDb.OleDbConnection(GetExcelConnstring(FilePath))
        conn.Open()
        myAdpt = New OleDb.OleDbDataAdapter("select * from [Sheet1$]", conn)
        dt = New DataTable
        myAdpt.Fill(dt)
        For i As Integer = 0 To dt.Rows.Count - 1
            If String.IsNullOrEmpty(dt.Rows(i)(0).ToString) = False Then
                If sResult = "" Then
                    sResult = dt.Rows(i)(0).ToString
                Else
                    sResult += "," + dt.Rows(i)(0).ToString
                End If
            End If

        Next
        conn.Close()
        Return sResult
    End Function

    Private Function GetExcelConnstring(ByVal excelfile As String) As String
        Dim sResult As String = String.Empty
        If Path.GetExtension(excelfile).Equals(".xls") = True Then
            sResult = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                   & "Data Source=" + excelfile + ";Extended Properties='Excel 8.0;HDR=No;IMEX=1'"
        ElseIf Path.GetExtension(excelfile).Equals(".xlsx") = True Then
            sResult = "Provider=Microsoft.ACE.OLEDB.12.0;" _
                   & "Data Source=" + excelfile + ";Extended Properties='Excel 12.0;HDR=No;IMEX=1'"
        End If
        Return sResult
    End Function

    Protected Sub RadioButtonList1_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RadioButtonList1.SelectedIndexChanged
        If RadioButtonList1.SelectedIndex = 0 Then
            but_Uupload.Enabled = False
            TextBox2.Enabled = False
            uf_UfilePath.Enabled = False

            rb_ProductPart.Enabled = True
            TextBox1.Enabled = True
            btn_Filter_Part.Enabled = True
            listB_PartSource.Enabled = True
            but_PartRight.Enabled = True
            but_PartLeft.Enabled = True
            listB_PartShow.Enabled = True
            Calendar1.Enabled = True
            Calendar2.Enabled = True
            txtDateFrom.Enabled = True
            txtDateTo.Enabled = True

            tr_time.Visible = True
            CategoryType.Visible = True
            CustomerType.Visible = True
            'ProductType.Visible = True
            Advanced.Visible = False
            ckAdvanced.Visible = False
        Else
            TextBox2.Enabled = True
            but_Uupload.Enabled = True
            uf_UfilePath.Enabled = True

            rb_ProductPart.Enabled = False
            TextBox1.Enabled = False
            btn_Filter_Part.Enabled = False
            listB_PartSource.Enabled = False
            but_PartRight.Enabled = False
            but_PartLeft.Enabled = False
            listB_PartShow.Enabled = False
            Calendar1.Enabled = False
            Calendar2.Enabled = False

            txtDateFrom.Enabled = False
            txtDateTo.Enabled = False


            tr_time.Visible = False
            CategoryType.Visible = False
            CustomerType.Visible = False
            'ProductType.Visible = False
            'Advanced.Visible = True
            ckAdvanced.Visible = True
            ckAdvanced.Checked = False
        End If
    End Sub

    Protected Sub rb_ProductPart_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rb_ProductPart.SelectedIndexChanged
        but_PartRight.Enabled = True
        Product()
    End Sub

    Protected Sub Button1_Click1(sender As Object, e As System.EventArgs) Handles Button1.Click
        For i As Integer = 0 To ddlStation.Items.Count - 1
            Dim sValus As String = ddlStation.Items(i).Text

            Dim eei As Integer = sValus.IndexOf(TextBox3.Text.ToUpper)
            If eei >= 0 Then
                ddlStation.SelectedIndex = i
                Exit Sub
            End If
        Next
        ddlStation.SelectedIndex = 0

    End Sub


    Protected Sub ckAdvanced_CheckedChanged(sender As Object, e As System.EventArgs) Handles ckAdvanced.CheckedChanged
        If ckAdvanced.Checked = True Then
            Advanced.Visible = True
        Else
            Advanced.Visible = False
        End If
    End Sub

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs) Handles btn_Ad_Right.Click
        moveList(lst_Ad_Source, lst_Ad_Target)

        ListBoxSort(lst_Ad_Source, True)
        ListBoxSort(lst_Ad_Target, True)

        If lst_Ad_Target.Items.Count > 0 Then
            btn_Ad_Right.Enabled = False
        End If

    End Sub

    Protected Sub Button7_Click(sender As Object, e As System.EventArgs) Handles Button7.Click
        moveList(lst_Ad_Target, lst_Ad_Source)

        ListBoxSort(lst_Ad_Source, True)
        ListBoxSort(lst_Ad_Target, True)


        If lst_Ad_Target.Items.Count = 0 Then
            btn_Ad_Right.Enabled = True
        End If
    End Sub

    Protected Sub but_Export_Click(sender As Object, e As System.EventArgs) Handles but_Export.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim customStr As String = ""
        Dim plantStr As String = ""
        Dim partStr As String = ""
        Dim weekStr As String = ""
        Dim itemStr As String = ""
        Dim topStr As String = ""
        Dim myAdapter As SqlDataAdapter
        Dim workTable As DataTable = New DataTable
        Dim workTable3 As DataTable = New DataTable
        Dim x_axle As DataTable = New DataTable

        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        Dim sGetPartID As String = Get_PartID()
        Dim Failmodeitem As String = ""
        For i As Integer = 0 To (lst_Yieldloss_Target.Items.Count - 1)
            Failmodeitem += ((lst_Yieldloss_Target.Items(i).Value).Replace("'", "''")) + ","
        Next
        If (Failmodeitem.Length > 0 AndAlso lst_Yieldloss_Target.Items.Count > 0) Then
            Failmodeitem = Failmodeitem.Substring(0, (Failmodeitem.Length - 1))
        End If


        Dim yl As New YieldlossInfo
        yl.Part_ID = Replace(sGetPartID, "'", "")
        yl.dtStart = Date.Parse(txtDateFrom.Text)
        yl.dtEnd = Date.Parse(txtDateTo.Text)
        yl.Fail_item = Replace(Failmodeitem, "'", "")
        yl.nFailMode = RadioButtonList2.SelectedIndex
        yl.nStation = rbl_Station.SelectedIndex
        yl.nType = rb_ProductPart.SelectedIndex
        yl.sStation = ddlStation.SelectedValue
        yl.sExtenBumpingType = Replace(Get_AD_PartID(), "'", "")
        'yl.nExtenWeek = ddlExtenWeek.SelectedValue

        If yl.sStation = "" Then
            yl.sStationC = "FI0 FVI"
        Else
            yl.sStationC = ddlStation.SelectedItem.Text
        End If

        yl.nLotList = RadioButtonList1.SelectedIndex
        yl.sLotList = TextBox2.Text

        Try

            Dim tempSQL As String = ""
            Dim tempXSQL As String = ""
            If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                tempSQL = getRowDataSQL2(yl)
                tempXSQL = getRowDataSQL2_X(yl)

                myAdapter = New SqlDataAdapter(tempXSQL, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myAdapter.Fill(x_axle)


            Else
                tempSQL = getRowDataSQL22(yl)
            End If

            myAdapter = New SqlDataAdapter(tempSQL, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myAdapter.Fill(workTable)



            Dim expression As String = ""
            Dim foundRows() As DataRow
            If workTable.Rows.Count > 0 Then

                If RadioButtonList1.SelectedIndex = 1 And ckAdvanced.Checked = True And yl.sExtenBumpingType <> Nothing Then

                    If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                        tempSQL = getRowDataSQL3(yl, workTable)
                    Else
                        tempSQL = getRowDataSQL33(yl, workTable)
                    End If

                    myAdapter = New SqlDataAdapter(tempSQL, conn)
                    myAdapter.SelectCommand.CommandTimeout = 3600
                    myAdapter.Fill(workTable3)
                    gv_rowdata.DataSource = workTable3

                    Dim col = workTable3.Columns("TargetLot")
                    Dim lot = workTable3.Columns("Lot_ID")

                    For Each row As DataRow In workTable3.Rows
                        Dim temp As String = row(lot)

                        expression = "Lot_ID = '" & temp & "'"
                        foundRows = workTable.Select(expression)
                        If foundRows.Length > 0 Then
                            row(col) = "Y"
                        End If
                    Next



                Else
                    workTable3 = generateDataSource(workTable, x_axle)
                    gv_rowdata.DataSource = workTable3
                    workTable.Clear()
                End If

                gv_rowdata.DataBind()
                UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)
                tr_gvDisplay.Visible = True
                tr_chartDisplay.Visible = True

                'Dim chart As New Dundas.Charting.WebControl.Chart()

                'DrawBarChart(chart, workTable, workTable3, yl)

                'Chart_Panel.Controls.Add(chart)
                'Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                ' workTable3.DataSet.Tables(0).Columns.Add()
                ExportToExcel(Page, workTable3, "YieldLoss")
                lab_wait.Text = ""
            Else
                tr_gvDisplay.Visible = False
                tr_chartDisplay.Visible = False

                lab_wait.Text = "無資料"
            End If

        Catch ex As Exception
            lab_wait.Text = "資料異常，請重新選取項目！！"


            If RadioButtonList1.SelectedIndex = 1 And TextBox2.Text = "" Then
                lab_wait.Text = "資料異常，請載入LotID！！"
            End If
        End Try
    End Sub
    Public Sub ExportToExcel(ByVal page As System.Web.UI.Page, ByVal dt As DataTable, ByRef FileName As String)
        Response.ClearContent()
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        If FileName = "" Then
            FileName = "IPP"
        End If

        Response.AddHeader("content-disposition", "attachment;filename=" & Server.UrlEncode(FileName & ".xls"))
        Response.ContentType = "application/excel"
        'Dim stringWrite As New System.IO.StringWriter()
        'Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)
        Dim tw As StringWriter = New System.IO.StringWriter()
        tw.Write("<html><table cellPadding=0 border=1 width=100%><tr bgcolor='Khaki'>")
        For i As Integer = 0 To dt.Columns.Count - 1
            tw.Write("<td align=center><b>" + dt.Columns(i).ColumnName.ToString + "</b></td>")
        Next
        tw.Write("</tr>")

        For i As Integer = 0 To dt.Rows.Count - 1
            tw.Write("<tr>")
            For j As Integer = 0 To dt.Columns.Count - 1
                Dim ColumnName As String = dt.Columns(j).ColumnName.Trim()
                Dim data As String = dt.Rows(i)(j).ToString().Trim()
                If data.IndexOf("上午") > 0 Or data.IndexOf("下午") > 0 Then
                    data = DateTime.Parse(data).ToString("yyyy/MM/dd hh:mm:ss")
                Else
                    data = data.Replace(" ", "&nbsp;")
                End If
                tw.Write("<td>" & data & "</td>")
            Next
            tw.Write("</tr>")
        Next
        tw.Write("</table></html>")

        'Console.WriteLine()
        Response.Write(tw.ToString())

        'Response.WriteFile(Server.UrlEncode(FileName & ".xls"))
        Response.[End]()
        'Response.BufferOutput = True
    End Sub

    Protected Sub ddlProduct_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlProduct.SelectedIndexChanged


        listB_PartShow.Items.Clear()
        ddlStation.Items.Clear()
        rbl_Station.SelectedIndex = 0

        lst_Yieldloss_Target.Items.Clear()
        lst_Ad_Target.Items.Clear()
        RadioButtonList2.SelectedIndex = 0
        but_PartRight.Enabled = True
        Button3.Enabled = True
        btn_Ad_Right.Enabled = True

        If ddlProduct.SelectedValue = "WB" Then
            cb_Lot_Merge.Visible = True
            RadioButtonList3.SelectedIndex = 1
        Else
            RadioButtonList3.SelectedIndex = 0
            cb_Lot_Merge.Visible = False
        End If

        Customer()
        Product()
        Yieldloss_FailMode("")
        ExtenType()
    End Sub

    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sfilter As String = TextBox4.Text

        sqlStr = "select distinct Part_id from " + confTable + " where 1=1 and Category='" + ddlProduct.SelectedValue + "'"
        If ddlCustomer.SelectedValue <> "All" Then
            sqlStr += " and customer_id ='" + ddlCustomer.SelectedValue + "' "
        End If

        If sfilter <> "" Then
            sqlStr += " and Part_id like '%" + sfilter + "%' "
        End If

        sqlStr += " order by Part_id"

        If RadioButtonList3.SelectedIndex = 0 Then

            sqlStr = sqlStr.Replace("Part_id", "Production_Id")
        End If

        If RadioButtonList3.SelectedIndex = 1 Then

            sqlStr = sqlStr.Replace("Part_id", "Bumping_Type")
        End If

        myAdapter = New SqlDataAdapter(sqlStr, conn)
        myAdapter.SelectCommand.CommandTimeout = 3600
        myDT = New DataTable
        myAdapter.Fill(myDT)

        lst_Ad_Source.Items.Clear()
        lst_Ad_Target.Items.Clear()
        For i As Integer = 0 To myDT.Rows.Count - 1
            If RadioButtonList3.SelectedIndex = 0 Then
                lst_Ad_Source.Items.Add(myDT.Rows(i)("Production_Id").ToString())
            ElseIf RadioButtonList3.SelectedIndex = 1 Then
                lst_Ad_Source.Items.Add(myDT.Rows(i)("Bumping_Type").ToString())
            Else
                lst_Ad_Source.Items.Add(myDT.Rows(i)("Part_id").ToString())
            End If

        Next
    End Sub

    Protected Sub Cb_Oversea_CheckedChanged(sender As Object, e As System.EventArgs) Handles Cb_Oversea.CheckedChanged
        Product()
    End Sub

    Protected Sub chkParallel_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkParallel.CheckedChanged

    End Sub

    Protected Sub chkRebuilt_CheckedChanged(sender As Object, e As System.EventArgs) Handles chkRebuilt.CheckedChanged

    End Sub

    Protected Sub ckMachine_CheckedChanged(sender As Object, e As System.EventArgs) Handles ckMachine.CheckedChanged
        If ckMachine.Checked = True Then
            txtMachine.Visible = True
        Else
            txtMachine.Visible = False
        End If
    End Sub
End Class
