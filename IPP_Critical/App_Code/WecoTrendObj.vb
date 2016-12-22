Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Data
Imports Dundas.Charting.WebControl

Public Class WecoTrendObj

    Public FunctionType As String = "Critical_Lot"
    Public Customer_ID As String = "INTEL"
    Public Product_Category As String = "CPU"
    Public Parameter_ID As String = "0"
    Public MAIN_ID As String = ""
    Public SUB_ID As String = ""
    Public KPP_Part As String = ""
    Public KPP_IPP As String = ""
    Public KPP_YieldImpact As String = ""
    Public KPP_KeyModule As String = ""
    Public KPP_CriticalItem As String = ""

    Public chartH As Integer = 400
    Public chartW As Integer = 400
    Public valueType As String = ""
    Public txtDateFrom As String = ""
    Public txtDateTo As String = ""
    Public dataSource As String = ""
    Public notDetail As Boolean = True
    Public isHighlight As Boolean = False
    Public linkToPoint As Boolean = False
    Public showProcess As Boolean = False    ' 對於量測的OOC 要呈現 Process Time.
    Public specialOldData As Boolean = False ' 對於過 WECO Rule 的資料要標示為特別
    Public ChartByTool As Boolean = False    ' 畫 Chart By 機台 Process 順序
    Public highlightLot As String = ""       ' 要標示星號的點
    Public HL_Day As String = ""

    Public Structure IPPInfo
        Dim sParts As String
        Dim sSatrtTime As String
        Dim sEndTime As String
        Dim sMIS As String
        Dim sLayer As String
        Dim sPARA As String
        Dim sLot As String
        Dim sSLI As String
    End Structure

    Public Structure IPPChart

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

        Dim xMeanMax As Double
        Dim xMeanMin As Double
        Dim sStdMax As Double
        Dim sStdMin As Double

        Dim yImpact As String
        Dim kModule As String
        Dim Critical As String
        Dim partId As String
        Dim edaItem As String

    End Structure

    Public Function Call_DrawChart(ByRef dt As DataTable, ByRef ChartObj As Dundas.Charting.WebControl.Chart, ByVal bLocation As Boolean) As Boolean

        Dim status As Boolean = False
        Dim sql As String = ""
        Dim s1 As String = ""
        Dim t1 As New DataTable
        Dim dt3 As New DataTable

        Dim new_column0 As DataColumn = New DataColumn
        new_column0.ColumnName = "SLI"
        new_column0.DataType = GetType(String)
        dt3.Columns.Add(new_column0)

        Dim new_column1 As DataColumn = New DataColumn
        new_column1.ColumnName = "Lot"
        new_column1.DataType = GetType(String)
        dt3.Columns.Add(new_column1)

        Dim new_column2 As DataColumn = New DataColumn
        new_column2.ColumnName = "meanval"
        new_column2.DataType = dt.Columns("meanval").DataType
        dt3.Columns.Add(new_column2)

        Dim new_column3 As DataColumn = New DataColumn
        new_column3.ColumnName = "std"
        new_column3.DataType = dt.Columns("std").DataType
        dt3.Columns.Add(new_column3)

        Dim new_column4 As DataColumn = New DataColumn
        new_column4.ColumnName = "WECO_Rule1"
        new_column4.DataType = dt.Columns("WECO_Rule1").DataType
        dt3.Columns.Add(new_column4)

        Dim new_column5 As DataColumn = New DataColumn
        new_column5.ColumnName = "WECO_Rule2"
        new_column5.DataType = dt.Columns("WECO_Rule2").DataType
        dt3.Columns.Add(new_column5)

        Dim new_column6 As DataColumn = New DataColumn
        new_column6.ColumnName = "WECO_Rule3"
        new_column6.DataType = dt.Columns("WECO_Rule3").DataType
        dt3.Columns.Add(new_column6)

        Dim new_column7 As DataColumn = New DataColumn
        new_column7.ColumnName = "WECO_Rule4"
        new_column7.DataType = dt.Columns("WECO_Rule4").DataType
        dt3.Columns.Add(new_column7)

        Dim new_column8 As DataColumn = New DataColumn
        new_column8.ColumnName = "WECO_Rule5"
        new_column8.DataType = dt.Columns("WECO_Rule5").DataType
        dt3.Columns.Add(new_column8)

        Dim new_column9 As DataColumn = New DataColumn
        new_column9.ColumnName = "WECO_Rule6"
        new_column9.DataType = dt.Columns("WECO_Rule6").DataType
        dt3.Columns.Add(new_column9)

        Dim new_column10 As DataColumn = New DataColumn
        new_column10.ColumnName = "WECO_Rule7"
        new_column10.DataType = dt.Columns("WECO_Rule7").DataType
        dt3.Columns.Add(new_column10)

        Dim new_column11 As DataColumn = New DataColumn
        new_column11.ColumnName = "WECO_Rule8"
        new_column11.DataType = dt.Columns("WECO_Rule8").DataType
        dt3.Columns.Add(new_column11)

        Dim new_column12 As DataColumn = New DataColumn
        new_column12.ColumnName = "ViolateRule"
        new_column12.DataType = GetType(String)
        dt3.Columns.Add(new_column12)

        Dim new_column13 As DataColumn = New DataColumn
        new_column13.ColumnName = "XLableValue"
        new_column13.DataType = GetType(String)
        dt3.Columns.Add(new_column13)

        Dim new_column14 As DataColumn = New DataColumn
        new_column14.ColumnName = "trtm"
        new_column14.DataType = GetType(String)
        dt3.Columns.Add(new_column14)

        Dim new_column15 As DataColumn = New DataColumn
        new_column15.ColumnName = "xLCL"
        new_column15.DataType = dt.Columns("xLCL").DataType
        dt3.Columns.Add(new_column15)

        Dim new_column16 As DataColumn = New DataColumn
        new_column16.ColumnName = "xUCL"
        new_column16.DataType = dt.Columns("xUCL").DataType
        dt3.Columns.Add(new_column16)

        Dim new_column17 As DataColumn = New DataColumn
        new_column17.ColumnName = "xCL"
        new_column17.DataType = GetType(Double)
        dt3.Columns.Add(new_column17)

        Dim new_column18 As DataColumn = New DataColumn
        new_column18.ColumnName = "xSigma"
        new_column18.DataType = GetType(Double)
        dt3.Columns.Add(new_column18)

        Dim new_column19 As DataColumn = New DataColumn
        new_column19.ColumnName = "sUCL"
        new_column19.DataType = dt.Columns("sUCL").DataType
        dt3.Columns.Add(new_column19)

        Dim new_column20 As DataColumn = New DataColumn
        new_column20.ColumnName = "sLCL"
        new_column20.DataType = dt.Columns("sLCL").DataType
        dt3.Columns.Add(new_column20)

        Dim new_column21 As DataColumn = New DataColumn
        new_column21.ColumnName = "sCL"
        new_column21.DataType = GetType(Double)
        dt3.Columns.Add(new_column21)

        Dim new_column22 As DataColumn = New DataColumn
        new_column22.ColumnName = "sSigma"
        new_column22.DataType = GetType(Double)
        dt3.Columns.Add(new_column22)

        ' Leon 2012/11/15
        Dim new_column23 As DataColumn = New DataColumn
        new_column23.ColumnName = "HL_Day"
        new_column23.DataType = GetType(String)
        dt3.Columns.Add(new_column23)

        ' Leon 2013/01/22
        If specialOldData Then
            Dim new_column24 As DataColumn = New DataColumn
            new_column24.ColumnName = "specialOldData"
            new_column24.DataType = GetType(String)
            dt3.Columns.Add(new_column24)
        End If

        ' Leon 2013/01/29 Chart By Tool
        If ChartByTool Then
            Dim new_column25 As DataColumn = New DataColumn
            new_column25.ColumnName = "MPID_Machine"
            new_column25.DataType = GetType(String)
            dt3.Columns.Add(new_column25)
        End If

        Dim expression As String = "1=1"
        Dim sortOrder As String = ""
        Dim foundRows() As DataRow = dt.Select(expression, sortOrder)
        Dim firPart As String = ""
        Dim lstPart As String = ""
        Dim temp1 As Double = -99
        Dim temp2 As Double = -99
        Dim temp3 As Double = -99
        Dim temp4 As Double = -99

        If foundRows.Length > 0 Then

            dt3.Clear()
            Dim ChartInfo As New WecoTrendObj.IPPChart
            ChartInfo.dt = dt3

            For k As Integer = 0 To (foundRows.Length - 1)

                Try
                    Dim pDROW_Row As DataRow = dt3.NewRow
                    pDROW_Row("SLI") = ""
                    pDROW_Row("XLableValue") = foundRows(k).Item("Lot") + "[" + Convert.ToDateTime(foundRows(k).Item("trtm")).ToString("yyyy/MM/dd") + "]"
                    pDROW_Row("Lot") = foundRows(k).Item("Lot")
                    pDROW_Row("meanval") = foundRows(k).Item("meanval")
                    pDROW_Row("std") = foundRows(k).Item("std")
                    pDROW_Row("trtm") = foundRows(k).Item("trtm")
                    pDROW_Row("WECO_Rule1") = foundRows(k).Item("WECO_Rule1")
                    pDROW_Row("WECO_Rule2") = foundRows(k).Item("WECO_Rule2")
                    pDROW_Row("WECO_Rule3") = foundRows(k).Item("WECO_Rule3")
                    pDROW_Row("WECO_Rule4") = foundRows(k).Item("WECO_Rule4")
                    pDROW_Row("WECO_Rule5") = foundRows(k).Item("WECO_Rule5")
                    pDROW_Row("WECO_Rule6") = foundRows(k).Item("WECO_Rule6")
                    pDROW_Row("WECO_Rule7") = foundRows(k).Item("WECO_Rule7")
                    pDROW_Row("WECO_Rule8") = foundRows(k).Item("WECO_Rule8")
                    pDROW_Row("ViolateRule") = TraslateWECORuleTip(foundRows(k))
                    pDROW_Row("xLCL") = foundRows(k).Item("xLCL")
                    pDROW_Row("xUCL") = foundRows(k).Item("xUCL")
                    pDROW_Row("sLCL") = foundRows(k).Item("sLCL")
                    pDROW_Row("sUCL") = foundRows(k).Item("sUCL")
                    pDROW_Row("HL_Day") = (Convert.ToDateTime(foundRows(k).Item("trtm"))).ToString("yyyy-MM-dd")

                    If specialOldData Then
                        If (foundRows(k).Item("NCW").ToString().ToUpper) = "TRUE" Then
                            pDROW_Row("specialOldData") = "N"
                        Else
                            pDROW_Row("specialOldData") = "Y"
                        End If
                    End If

                    ' Leon 2013/01/29 Chart By Tool
                    If ChartByTool Then
                        pDROW_Row("MPID_Machine") = (IIf(IsDBNull(foundRows(k).Item("MPID")), "", foundRows(k).Item("MPID")) + "_" + IIf(IsDBNull(foundRows(k).Item("EqpID")), "", foundRows(k).Item("EqpID")))
                    End If

                    dt3.Rows.Add(pDROW_Row)

                    If temp1 = -99 Then
                        temp1 = IIf(IsDBNull(foundRows(k).Item("xLCL")), -99, foundRows(k).Item("xLCL"))
                        ChartInfo.xLCL = temp1
                        ChartInfo.xMeanMin = temp1
                    End If

                    If IsDBNull(pDROW_Row("xLCL")) = True Then
                        pDROW_Row("xLCL") = temp1
                    End If

                    If pDROW_Row("xLCL") <> ChartInfo.xLCL Then
                        ChartInfo.xLCL = pDROW_Row("xLCL")
                    End If

                    If ChartInfo.xMeanMin > pDROW_Row("xLCL") Then
                        ChartInfo.xMeanMin = pDROW_Row("xLCL")
                    End If

                    If temp2 = -99 Then
                        temp2 = IIf(IsDBNull(foundRows(k).Item("xUCL")), -99, foundRows(k).Item("xUCL"))
                        ChartInfo.xUCL = temp2
                        ChartInfo.xMeanMax = temp2
                    End If

                    If IsDBNull(pDROW_Row("xUCL")) = True Then
                        pDROW_Row("xUCL") = ChartInfo.xUCL
                    End If

                    If pDROW_Row("xUCL") <> ChartInfo.xUCL Then
                        ChartInfo.xLCL = pDROW_Row("xUCL")
                    End If

                    If ChartInfo.xMeanMax < pDROW_Row("xUCL") Then
                        ChartInfo.xMeanMax = pDROW_Row("xUCL")
                    End If

                    pDROW_Row("xCL") = (pDROW_Row("xUCL") + pDROW_Row("xLCL")) / 2
                    pDROW_Row("xSigma") = (pDROW_Row("xUCL") - pDROW_Row("xLCL")) / 6

                    If temp3 = -99 Then
                        temp3 = IIf(IsDBNull(foundRows(k).Item("sLCL")), -99, foundRows(k).Item("sLCL"))
                        ChartInfo.sLCL = temp3
                        ChartInfo.sStdMin = temp3
                    End If
                    If IsDBNull(pDROW_Row("sLCL")) = True Then
                        pDROW_Row("sLCL") = ChartInfo.sLCL
                    End If

                    If pDROW_Row("sLCL") <> ChartInfo.sLCL Then
                        ChartInfo.sLCL = pDROW_Row("sLCL")
                    End If

                    If ChartInfo.sStdMin > pDROW_Row("sLCL") Then
                        ChartInfo.sStdMin = pDROW_Row("sLCL")
                    End If

                    If temp4 = -99 Then
                        temp4 = IIf(IsDBNull(foundRows(k).Item("sUCL")), -99, foundRows(k).Item("sUCL"))
                        ChartInfo.sUCL = temp4
                        ChartInfo.sStdMax = temp4
                    End If
                    If IsDBNull(pDROW_Row("sUCL")) = True Then
                        pDROW_Row("sUCL") = ChartInfo.sUCL
                    End If

                    If pDROW_Row("sUCL") <> ChartInfo.sUCL Then
                        ChartInfo.sUCL = pDROW_Row("sUCL")
                    End If

                    If ChartInfo.sStdMax < pDROW_Row("sUCL") Then
                        ChartInfo.sStdMax = pDROW_Row("sUCL")
                    End If

                    pDROW_Row("sCL") = (pDROW_Row("sUCL") + pDROW_Row("sLCL")) / 2
                    pDROW_Row("sSigma") = (pDROW_Row("sUCL") - pDROW_Row("sLCL")) / 6
                Catch ex As Exception

                End Try


            Next

            'ChartInfo.sMIS = foundRows(0).Item("MIS_OP")
            ChartInfo.sLot = ""
            ChartInfo.sSLI = ""
            ChartInfo.ParametricItem = foundRows(0).Item("Parametric_Measurement")
            ChartInfo.XLableValue = "Lot"
            ChartInfo.nType = "oldm"
            ChartInfo.YValue = valueType
            ChartInfo.partId = foundRows(0).Item("Part")
            CallDundals(ChartObj, ChartInfo, bLocation)
            status = True

        End If

        Return status
    End Function

    Public Sub CallDundals(ByRef ChartObj As Dundas.Charting.WebControl.Chart, ByRef cinfo As WecoTrendObj.IPPChart, ByVal bLocation As Boolean)

        ChartType(ChartObj, cinfo)
        If ChartByTool Then
            ChartDataByTool(ChartObj, cinfo)
        Else
            ChartData(ChartObj, cinfo)
        End If


        If valueType = "meanval" Then
            Try
                Chart_xLCL(ChartObj, cinfo, bLocation)
                Chart_xUCL(ChartObj, cinfo, bLocation)
                Chart_xCL(ChartObj, cinfo, bLocation)
            Catch ex As Exception
            End Try
            'Chart_xSigma(ChartObj, cinfo)
        Else
            Try
                Chart_sLCL(ChartObj, cinfo)
                Chart_sUCL(ChartObj, cinfo)
                Chart_sCL(ChartObj, cinfo)
            Catch ex As Exception
            End Try
        End If

    End Sub

    Public Sub ChartType(ByRef ChartObj As Dundas.Charting.WebControl.Chart, ByVal cinfo As WecoTrendObj.IPPChart)

        ChartObj.Palette = ChartColorPalette.Dundas
        ChartObj.Height = chartH
        ChartObj.Width = chartW
        ChartObj.Palette = ChartColorPalette.Dundas
        ChartObj.BackColor = Color.White
        ChartObj.BackGradientEndColor = Color.Peru
        ChartObj.ChartAreas.Add("Default")
        ChartObj.UI.Toolbar.Enabled = False
        ChartObj.UI.ContextMenu.Enabled = True

        If FunctionType = "Critical_Lot" Then
            If valueType = "meanval" Then
                ChartObj.Titles.Add(cinfo.ParametricItem + "(Mean)")
            Else
                ChartObj.Titles.Add(cinfo.ParametricItem + "(Std)")
            End If
        Else
            If valueType = "meanval" Then
                ChartObj.Titles.Add(KPP_Part + ":" + KPP_IPP + ":" + KPP_YieldImpact + ":" + KPP_KeyModule + ":" + KPP_CriticalItem + "\n" + cinfo.ParametricItem + "(Mean)")
            Else
                ChartObj.Titles.Add(KPP_Part + ":" + KPP_IPP + ":" + KPP_YieldImpact + ":" + KPP_KeyModule + ":" + KPP_CriticalItem + "\n" + cinfo.ParametricItem + "(Std)")
            End If
        End If
        

        ChartObj.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        ChartObj.Titles(0).Color = Color.DarkBlue
        ' Set AxisX			
        ChartObj.ChartAreas("Default").AxisX.LabelStyle.Enabled = True

        ' Set AxisY 
        ChartObj.ChartAreas("Default").AxisX.MajorGrid.Enabled = False
        ChartObj.ChartAreas("Default").AxisY.MajorGrid.Enabled = False
        'ChartObj.ChartAreas("Default").AxisX.MajorGrid.LineColor = Color.LightGray
        'ChartObj.ChartAreas("Default").AxisY.MajorGrid.LineColor = Color.LightGray
        'ChartObj.ChartAreas("Default").AxisY.MajorGrid.LineStyle = ChartDashStyle.Dash

        'ChartObj.ChartAreas("Default").AxisY.Title = "Data"
        ChartObj.ChartAreas("Default").AxisY.TitleFont = New Font("Arial", 8, FontStyle.Regular)
        ChartObj.ChartAreas("Default").AxisY.TitleColor = Color.Black
        ChartObj.ChartAreas("Default").AxisY.LabelsAutoFit = True
        ChartObj.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        ChartObj.ChartAreas("Default").AxisY.LabelStyle.FontColor = Color.Black
        ChartObj.ChartAreas("Default").AxisX.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        ChartObj.ChartAreas("Default").AxisX.LabelStyle.FontColor = Color.Black
        ChartObj.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -60
        ChartObj.ChartAreas("Default").BackColor = Color.White
        ChartObj.ChartAreas("Default").AxisX.LineColor = Color.Black
        ChartObj.ChartAreas("Default").AxisY.LineColor = Color.Black

        'ChartObj.ChartAreas("Default").AxisX.Title = "SLI"
        'ChartObj.ChartAreas("Default").AxisX.TitleFont = New Font("Arial", 10, FontStyle.Regular)
        'ChartObj.ChartAreas("Default").AxisX.TitleColor = Color.White

        ChartObj.Legends(0).LegendStyle = LegendStyle.Row
        ChartObj.Legends(0).BackColor = Color.White
        ChartObj.Legends(0).Alignment = StringAlignment.Center
        ChartObj.Legends(0).Docking = LegendDocking.Top
        ChartObj.Legends(0).FontColor = Color.DarkBlue
        ChartObj.Legends(0).LegendStyle = LegendStyle.Table
        ChartObj.ImageUrl = "temp/CriticalLot_#SEQ(1000,1)"
        ChartObj.ImageType = ChartImageType.Png

        '2012/10/17 modified by Chery
        'add the Legends Item for point marked Color description
        Dim objWECOLegend As Legend = New Legend
        objWECOLegend.LegendStyle = LegendStyle.Table
        Dim iLengentPointSize As Integer = 10
        Dim iLengendPointSize As Integer = 10
        Dim objPoint As LegendItem = New LegendItem()
        objPoint.MarkerSize = iLengentPointSize
        objPoint.Name = cinfo.YValue
        objPoint.Style = LegendImageStyle.Marker
        objPoint.MarkerColor = Color.Blue
        ChartObj.Legends(0).CustomItems.Add(objPoint)

        If specialOldData Then
            Dim objReIntry As LegendItem = New LegendItem()
            objReIntry.MarkerSize = iLengendPointSize
            objReIntry.Name = "Measure For Today"
            objReIntry.Style = LegendImageStyle.Marker
            objReIntry.MarkerColor = Color.Black
            objReIntry.BorderColor = Color.Black
            objReIntry.ToolTip = "計算完 WECO Rule 後進點的資料"
            ChartObj.Legends(0).CustomItems.Add(objReIntry)
        End If

        Dim objWECORule1 As LegendItem = New LegendItem()
        objWECORule1.MarkerSize = iLengendPointSize
        objWECORule1.Name = "WECO Rule 1"
        objWECORule1.Style = LegendImageStyle.Marker
        objWECORule1.MarkerColor = Color.Red
        objWECORule1.BorderColor = Color.Red
        objWECORule1.ToolTip = "單點超出管制界限" & vbCrLf & "A single point beyond either control limit"
        ChartObj.Legends(0).CustomItems.Add(objWECORule1)

        Dim objWECORule2 As LegendItem = New LegendItem()
        objWECORule2.MarkerSize = iLengendPointSize
        objWECORule2.Name = "WECO Rule 2"
        objWECORule2.Style = LegendImageStyle.Marker
        objWECORule2.MarkerColor = Color.LimeGreen
        objWECORule2.BorderColor = Color.LimeGreen
        objWECORule2.ToolTip = "連續九點在同一邊" & vbCrLf & "9 consecutive points on the same side of the centerline"
        ChartObj.Legends(0).CustomItems.Add(objWECORule2)

        Dim objWECORule3 As LegendItem = New LegendItem()
        objWECORule3.MarkerSize = iLengendPointSize
        objWECORule3.Name = "WECO Rule 3"
        objWECORule3.Style = LegendImageStyle.Marker
        objWECORule3.MarkerColor = Color.LightSalmon
        objWECORule3.BorderColor = Color.LightSalmon
        objWECORule3.ToolTip = "連續六點上升或是連續六點下降" & vbCrLf & "6 consecutive points steadily increasing or decreasing"
        ChartObj.Legends(0).CustomItems.Add(objWECORule3)

        Dim objWECORule4 As LegendItem = New LegendItem()
        objWECORule4.MarkerSize = iLengendPointSize
        objWECORule4.Name = "WECO Rule 4"
        objWECORule4.Style = LegendImageStyle.Marker
        objWECORule4.MarkerColor = Color.IndianRed
        objWECORule4.BorderColor = Color.IndianRed
        objWECORule4.ToolTip = "連續14點上下跳動" & vbCrLf & "14 (or more) consecutive points are alternating up and down"
        ChartObj.Legends(0).CustomItems.Add(objWECORule4)

        Dim objWECORule5 As LegendItem = New LegendItem()
        objWECORule5.MarkerSize = iLengendPointSize
        objWECORule5.Name = "WECO Rule 5"
        objWECORule5.Style = LegendImageStyle.Marker
        objWECORule5.MarkerColor = Color.Fuchsia
        objWECORule5.BorderColor = Color.Fuchsia
        objWECORule5.ToolTip = "連續3點在同側且有2點至少落在超過2倍標準差外" & vbCrLf & "2 out of 3 consecutive points at least 2 std dev beyond the centerline, on the same side"
        ChartObj.Legends(0).CustomItems.Add(objWECORule5)

        Dim objWECORule6 As LegendItem = New LegendItem()
        objWECORule6.MarkerSize = iLengendPointSize
        objWECORule6.Name = "WECO Rule 6"
        objWECORule6.Style = LegendImageStyle.Marker
        objWECORule6.MarkerColor = Color.Gray
        objWECORule6.BorderColor = Color.Gray
        objWECORule6.ToolTip = "連續5點在同側且有4點至少落在超過1倍標準差外" & vbCrLf & "4 out of 5 consecutive points on the chart are more than 1 std dev away from the CL"
        ChartObj.Legends(0).CustomItems.Add(objWECORule6)

        Dim objWECORule7 As LegendItem = New LegendItem()
        objWECORule7.MarkerSize = iLengendPointSize
        objWECORule7.Name = "WECO Rule 7"
        objWECORule7.Style = LegendImageStyle.Marker
        objWECORule7.MarkerColor = Color.MediumOrchid
        objWECORule7.BorderColor = Color.MediumOrchid
        objWECORule7.ToolTip = "連續15點在1倍標準差間" & vbCrLf & "15 (or more) consecutive points are within 1 std dev of the CL"
        ChartObj.Legends(0).CustomItems.Add(objWECORule7)

        Dim objWECORule8 As LegendItem = New LegendItem()
        objWECORule8.MarkerSize = iLengendPointSize
        objWECORule8.Name = "WECO Rule 8"
        objWECORule8.Style = LegendImageStyle.Marker
        objWECORule8.MarkerColor = Color.Gold
        objWECORule8.BorderColor = Color.Gold
        objWECORule8.ToolTip = "連續8點在中心2側，但皆不在1倍標準差內" & vbCrLf & "8 (or more) consecutive points are on both sides of the CL, but none are within 1 std dev of it."
        ChartObj.Legends(0).CustomItems.Add(objWECORule8)


      
        ChartObj.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        ChartObj.BorderStyle = ChartDashStyle.Solid
        ChartObj.BorderWidth = 3
        ChartObj.BorderColor = Color.DarkBlue
        ChartObj.ChartAreas("Default").AxisX.Interval = 1
        ChartObj.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

    End Sub

    Public Sub ChartData(ByRef ChartObj As Dundas.Charting.WebControl.Chart, ByVal cinfo As WecoTrendObj.IPPChart)

        Dim series As Series
        series = ChartObj.Series.Add("MQCS")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Black
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 8
        series.MarkerColor = Color.DarkBlue
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        'Series Data
        If cinfo.xMeanMin <> -99 Then
            cinfo.nMIN = cinfo.xMeanMin
        End If

        If cinfo.xMeanMax <> -99 Then
            cinfo.nMAX = cinfo.xMeanMax
        End If

        If cinfo.dt.Rows.Count > 0 Then

            Dim tmpDouble As Double
            Dim specMax As Double = 0
            Dim specMin As Double = 0
            Dim maxValue As Double = 0
            Dim minValue As Double = 0
            Dim sXaxis As String
            Dim nYvalue As Double
            Dim isHLStr As String = "N"
            Dim showProcessStr As String = "N"
            Dim HLLOT As String = ""
            If isHighlight Then
                isHLStr = "Y"
            End If
            If showProcess Then
                showProcessStr = "Y"
            End If

            ' --- 放值到 Chart 中 ---
            For i As Integer = 0 To (cinfo.dt.Rows.Count - 1)

                sXaxis = IIf(IsDBNull(cinfo.dt.Rows(i).Item("XLableValue")), "", cinfo.dt.Rows(i).Item("XLableValue"))
                nYvalue = IIf(IsDBNull(cinfo.dt.Rows(i).Item(cinfo.YValue)), 0, cinfo.dt.Rows(i).Item(cinfo.YValue))
                nYvalue = Math.Round(nYvalue, 4)
                HLLOT = IIf(IsDBNull(cinfo.dt.Rows(i).Item("LOT")), "", cinfo.dt.Rows(i).Item("LOT"))

                If valueType = "meanval" Then

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("xUCL")), 0, cinfo.dt.Rows(i)("xUCL"))
                    tmpDouble = Math.Round(tmpDouble, 4, MidpointRounding.AwayFromZero)

                    If tmpDouble <> -99 Then
                        specMax = tmpDouble
                    End If

                    If tmpDouble <> -99 Then
                        If tmpDouble > specMax Then
                            specMax = tmpDouble
                        End If
                    End If

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("xLCL")), 0, cinfo.dt.Rows(i)("xLCL"))
                    tmpDouble = Math.Round(tmpDouble, 4, MidpointRounding.AwayFromZero)

                    If tmpDouble <> -99 Then
                        specMin = tmpDouble
                    End If

                    If tmpDouble <> -99 Then
                        If tmpDouble < specMin Then
                            specMin = tmpDouble
                        End If
                    End If

                Else

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("sUCL")), 0, cinfo.dt.Rows(i)("sUCL"))
                    tmpDouble = Math.Round(tmpDouble, 4, MidpointRounding.AwayFromZero)

                    If tmpDouble <> -99 Then
                        specMax = tmpDouble
                    End If

                    If tmpDouble <> -99 Then
                        If tmpDouble > specMax Then
                            specMax = tmpDouble
                        End If
                    End If

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("sLCL")), 0, cinfo.dt.Rows(i)("sLCL"))
                    tmpDouble = Math.Round(tmpDouble, 4)

                    If tmpDouble <> -99 Then
                        specMin = tmpDouble
                    End If

                    If tmpDouble <> -99 Then
                        If tmpDouble < specMin Then
                            specMin = tmpDouble
                        End If
                    End If

                End If

                If i = 0 Then
                    maxValue = nYvalue
                    minValue = nYvalue
                End If

                ChartObj.Series(0).Points.AddXY(i, nYvalue)
                ChartObj.Series(0).Points(i).AxisLabel = sXaxis

                If linkToPoint Then
                    Dim linkLot As String = (cinfo.dt.Rows(i)("LOT")).ToString()
                    Dim linkTrtm As String = (cinfo.dt.Rows(i)("trtm")).ToString()
                    ChartObj.Series(0).Points(i).Href = "javascript:LinkPoint('" + (linkLot + linkTrtm) + "');"
                End If

                If notDetail Then
                    If FunctionType = "Critical_Lot" Then
                        ChartObj.Series(0).Points(i).Href = "javascript:openWin('" + (Me.Product_Category) + "','" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (Me.SUB_ID) + "','" + (valueType) + "','" + isHLStr + "','" + (HLLOT) + "')"
                    Else
                        ChartObj.Series(0).Points(i).Href = "javascript:openWin('" + (Me.Product_Category) + "','" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (KPP_Part) + "','" + (Me.MAIN_ID) + "','" + (Me.SUB_ID) + "','" + (isHLStr) + "','" + (showProcessStr) + "','" + (HLLOT) + "')"
                    End If
                End If

                ' Check Weco Rule
                ChartObj.Series(0).Points(i).MarkerColor = GetPointColor(i, cinfo.dt)
                If cinfo.dt.Rows(i)("ViolateRule").ToString().Length <= 0 Then
                    ChartObj.Series(0).Points(i).ToolTip = "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                Else
                    ChartObj.Series(0).Points(i).ToolTip = cinfo.dt.Rows(i)("ViolateRule") & vbCrLf & "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                End If

                ' --- 針對算過 WECO Rule 的資料又重新進點所要的標示 ---
                If specialOldData Then
                    If (cinfo.dt.Rows(i)("specialOldData").ToString() = "Y") Then
                        ChartObj.Series(0).Points(i).MarkerStyle = MarkerStyle.Star5
                        ChartObj.Series(0).Points(i).MarkerSize = 12
                        ChartObj.Series(0).Points(i).MarkerColor = Color.Black
                        ChartObj.Series(0).Points(i).BorderColor = Color.White
                        ChartObj.Series(0).Points(i).BorderWidth = 1
                    End If
                End If

                ' --- 前端使用者所點的點標示 --- 
                If highlightLot.Length > 0 And highlightLot.Equals(HLLOT) Then
                    ChartObj.Series(0).Points(i).MarkerStyle = MarkerStyle.Star5
                    ChartObj.Series(0).Points(i).MarkerSize = 18
                    ChartObj.Series(0).Points(i).MarkerColor = Color.Black
                    ChartObj.Series(0).Points(i).BorderColor = Color.White
                    ChartObj.Series(0).Points(i).BorderWidth = 1
                End If

                ' 2012/10/19 modified by Chery , display by WECO rule color.
                Select Case cinfo.nType
                    Case "olds"
                        If nYvalue > cinfo.sUCL And cinfo.sUCL <> -99 Then
                            ChartObj.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                        If nYvalue < cinfo.sLCL And cinfo.sLCL <> -99 Then
                            ChartObj.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                End Select

                If nYvalue > maxValue Then
                    maxValue = nYvalue
                End If

                If nYvalue < minValue Then
                    minValue = nYvalue
                End If

            Next
            ' --- E N D ---

            ' --- 先算上下界, Group 要用 ---
            If specMin < minValue And specMin <> -99 Then
                minValue = specMin
            End If

            If specMax > maxValue And specMin <> -99 Then
                maxValue = specMax
            End If

            Dim Maxntemp As Double
            Dim Minntemp As Double
            Dim nInterval As Double = Math.Round((maxValue - minValue) / 20, 2, MidpointRounding.AwayFromZero)

            If nInterval <> 0 Then
                ' --- Max ---
                Maxntemp = (maxValue + nInterval)
                Maxntemp = Math.Round(Maxntemp, 2)
                ' --- Min ---
                Minntemp = (minValue - nInterval)
                Minntemp = Math.Round(Minntemp, 2)
                If Maxntemp <> Minntemp Then
                    ChartObj.ChartAreas("Default").AxisY.Maximum = Maxntemp
                    ChartObj.ChartAreas("Default").AxisY.Minimum = Minntemp
                End If
            End If

            ' --- 如果使用選擇 Daily Event 要 HighLight Range ---
            If isHighlight Then

                Dim stripMed As New StripLine()
                Dim rowDataTime As String
                Dim d_offset As Double = 0
                Dim d_width As Double = 0
                ChartObj.BorderColor = Color.Red
                Try
                    Dim firstIn As Boolean = True
                    For i As Integer = 0 To (cinfo.dt.Rows.Count - 1)
                        rowDataTime = IIf(IsDBNull(cinfo.dt.Rows(i).Item("HL_Day")), "", cinfo.dt.Rows(i).Item("HL_Day"))
                        nYvalue = IIf(IsDBNull(cinfo.dt.Rows(i).Item(cinfo.YValue)), 0, cinfo.dt.Rows(i).Item(cinfo.YValue))
                        nYvalue = Math.Round(nYvalue, 3)
                        ' 如果使用選擇 Daily Event 要 HighLight Range
                        If (Me.HL_Day = rowDataTime) Then
                            If firstIn Then
                                d_offset = i
                                firstIn = False
                            End If
                            d_width += 1
                        End If
                    Next

                    stripMed.IntervalOffset = d_offset
                    stripMed.StripWidth = d_width
                    stripMed.BackColor = Color.FromArgb(255, 235, 205) 'FFEBCD
                    'stripMed.BackColor = Color.Blue
                    ChartObj.ChartAreas("Default").AxisX.StripLines.Add(stripMed)

                Catch ex As Exception
                End Try


            End If
            
            ' ----- Title -----
            If notDetail Then

                'If FunctionType = "Critical_Lot" Then
                '    ChartObj.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (cinfo.ParametricItem) + "','" + (valueType) + "','" + isHLStr + "')"
                'Else
                '    ChartObj.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','IPP','" + (KPP_Part) + "','" + (KPP_YieldImpact) + "','" + (KPP_KeyModule) + "','" + (KPP_CriticalItem) + "','" + (KPP_IPP) + "','" + (isHLStr) + "')"
                'End If

                If FunctionType = "Critical_Lot" Then
                    ChartObj.Titles(0).Href = "javascript:openWin('" + (Me.Product_Category) + "','" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (Me.SUB_ID) + "','" + (valueType) + "','" + isHLStr + "','" + (HLLOT) + "')"
                Else
                    ChartObj.Titles(0).Href = "javascript:openWin('" + (Me.Product_Category) + "','" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (KPP_Part) + "','" + (Me.MAIN_ID) + "','" + (Me.SUB_ID) + "','" + (isHLStr) + "','" + (showProcessStr) + "','" + (HLLOT) + "')"
                End If

            End If
            ChartObj.Titles.Add("Lot Count :" + cinfo.dt.Rows.Count.ToString)
            ChartObj.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            ChartObj.Titles(1).Color = Color.Black
            ' ----- Title -----

            ChartObj.ChartAreas("Default").AxisX.Maximum = cinfo.dt.Rows.Count
            If Math.Round(cinfo.dt.Rows.Count / 15) > 1 Then
                ChartObj.ChartAreas("Default").AxisX.Interval = Math.Round(cinfo.dt.Rows.Count / 10)
            End If
            ChartObj.Series(0).ShowInLegend = False

        End If

    End Sub

    Public Sub ChartDataByTool(ByRef ChartObj As Dundas.Charting.WebControl.Chart, ByVal cinfo As WecoTrendObj.IPPChart)

        Dim stripMed As StripLine
        Dim series As Series
        series = ChartObj.Series.Add("MQCS")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Black
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 8
        series.MarkerColor = Color.DarkBlue
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        'Series Data
        If cinfo.xMeanMin <> -99 Then
            cinfo.nMIN = cinfo.xMeanMin
        End If

        If cinfo.xMeanMax <> -99 Then
            cinfo.nMAX = cinfo.xMeanMax
        End If

        If cinfo.dt.Rows.Count > 0 Then

            Dim tmpDouble As Double
            Dim specMax As Double = 0
            Dim specMin As Double = 0
            Dim maxValue As Double = 0
            Dim minValue As Double = 0
            Dim sXaxis As String
            Dim nYvalue As Double
            Dim isHLStr As String = "N"
            Dim showProcessStr As String = "N"
            Dim HLLOT As String = ""
            Dim MPID_Machine As String = ""
            Dim newMPID_Machine As String = ""
            Dim oldMPID_Machine As String = ""
            If isHighlight Then
                isHLStr = "Y"
            End If
            If showProcess Then
                showProcessStr = "Y"
            End If

            ' --- 放值到 Chart 中 --- Start 
            Dim d_offset As Double = 0
            Dim d_width As Double = 0
            Dim colorChange As Integer = 0
            For i As Integer = 0 To (cinfo.dt.Rows.Count - 1)

                sXaxis = IIf(IsDBNull(cinfo.dt.Rows(i).Item("XLableValue")), "", cinfo.dt.Rows(i).Item("XLableValue"))
                nYvalue = IIf(IsDBNull(cinfo.dt.Rows(i).Item(cinfo.YValue)), 0, cinfo.dt.Rows(i).Item(cinfo.YValue))
                nYvalue = Math.Round(nYvalue, 4)
                HLLOT = IIf(IsDBNull(cinfo.dt.Rows(i).Item("LOT")), "", cinfo.dt.Rows(i).Item("LOT"))
                MPID_Machine = IIf(IsDBNull(cinfo.dt.Rows(i).Item("MPID_Machine")), "", cinfo.dt.Rows(i).Item("MPID_Machine"))
                newMPID_Machine = MPID_Machine

                If i = 0 Then
                    oldMPID_Machine = MPID_Machine
                End If

                ' --- 如果使用選擇 Daily Event 要 HighLight Range ---
                If newMPID_Machine <> oldMPID_Machine Then
                    stripMed = New StripLine()
                    ChartObj.BorderColor = Color.Red
                    Try
                        stripMed.Title = oldMPID_Machine
                        stripMed.ToolTip = ("Process : " + oldMPID_Machine)
                        stripMed.IntervalOffset = d_offset
                        stripMed.StripWidth = d_width
                        stripMed.TitleAlignment = StringAlignment.Far
                        stripMed.TitleLineAlignment = StringAlignment.Far
                        'stripMed.BorderColor = Color.Black
                        'stripMed.StripWidth = 0.1
                        If (colorChange Mod 2) = 0 Then
                            stripMed.BackColor = Color.FromArgb(99, 184, 255)
                        Else
                            stripMed.BackColor = Color.FromArgb(171, 171, 171)
                        End If
                        ChartObj.ChartAreas("Default").AxisX.StripLines.Add(stripMed)
                    Catch ex As Exception
                    End Try
                    colorChange += 1
                    d_width = 0
                    d_offset = i
                    oldMPID_Machine = newMPID_Machine
                End If

                d_width += 1

                If valueType = "meanval" Then

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("xUCL")), 0, cinfo.dt.Rows(i)("xUCL"))
                    tmpDouble = Math.Round(tmpDouble, 4, MidpointRounding.AwayFromZero)
                    specMax = tmpDouble
                    If tmpDouble > specMax Then
                        specMax = tmpDouble
                    End If

                    tmpDouble = IIf(IsDBNull(cinfo.dt.Rows(i)("xLCL")), 0, cinfo.dt.Rows(i)("xLCL"))
                    tmpDouble = Math.Round(tmpDouble, 4, MidpointRounding.AwayFromZero)
                    specMin = tmpDouble
                    If tmpDouble < specMin Then
                        specMin = tmpDouble
                    End If

                End If

                If i = 0 Then
                    maxValue = nYvalue
                    minValue = nYvalue
                End If

                ChartObj.Series(0).Points.AddXY(i, nYvalue)
                ChartObj.Series(0).Points(i).AxisLabel = sXaxis

                If linkToPoint Then
                    Dim linkLot As String = (cinfo.dt.Rows(i)("LOT")).ToString()
                    Dim linkTrtm As String = (cinfo.dt.Rows(i)("trtm")).ToString()
                    ChartObj.Series(0).Points(i).Href = "javascript:LinkPoint('" + (linkLot + linkTrtm) + "');"
                End If

                If notDetail Then
                    If FunctionType = "Critical_Lot" Then
                        ChartObj.Series(0).Points(i).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (cinfo.ParametricItem) + "','" + (valueType) + "','" + isHLStr + "','" + (HLLOT) + "')"
                    Else
                        ChartObj.Series(0).Points(i).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','IPP','" + (KPP_Part) + "','" + (KPP_YieldImpact) + "','" + (KPP_KeyModule) + "','" + (KPP_CriticalItem) + "','" + (KPP_IPP) + "','" + (isHLStr) + "','Y','" + (HLLOT) + "')"
                    End If
                End If

                ' --- Check Weco Rule
                ChartObj.Series(0).Points(i).MarkerColor = GetPointColor(i, cinfo.dt)
                If cinfo.dt.Rows(i)("ViolateRule").ToString().Length <= 0 Then
                    ChartObj.Series(0).Points(i).ToolTip = newMPID_Machine & vbCrLf & "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                Else
                    ChartObj.Series(0).Points(i).ToolTip = newMPID_Machine & vbCrLf & cinfo.dt.Rows(i)("ViolateRule") & vbCrLf & "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                End If

                ' --- 針對算過 WECO Rule 的資料又重新進點所要的標示 ---
                If specialOldData Then
                    If (cinfo.dt.Rows(i)("specialOldData").ToString() = "Y") Then
                        ChartObj.Series(0).Points(i).MarkerStyle = MarkerStyle.Star5
                        ChartObj.Series(0).Points(i).MarkerSize = 12
                        ChartObj.Series(0).Points(i).MarkerColor = Color.Black
                        ChartObj.Series(0).Points(i).BorderColor = Color.White
                        ChartObj.Series(0).Points(i).BorderWidth = 1
                    End If
                End If

                ' --- 前端使用者所點的點標示 --- 
                If highlightLot.Length > 0 And highlightLot.Equals(HLLOT) Then
                    ChartObj.Series(0).Points(i).MarkerStyle = MarkerStyle.Star5
                    ChartObj.Series(0).Points(i).MarkerSize = 18
                    ChartObj.Series(0).Points(i).MarkerColor = Color.Black
                    ChartObj.Series(0).Points(i).BorderColor = Color.White
                    ChartObj.Series(0).Points(i).BorderWidth = 1
                End If

                ' 2012/10/19 modified by Chery , display by WECO rule color.
                Select Case cinfo.nType
                    Case "olds"
                        If nYvalue > cinfo.sUCL And cinfo.sUCL <> -99 Then
                            ChartObj.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                        If nYvalue < cinfo.sLCL And cinfo.sLCL <> -99 Then
                            ChartObj.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                End Select

                If nYvalue > maxValue Then
                    maxValue = nYvalue
                End If

                If nYvalue < minValue Then
                    minValue = nYvalue
                End If

            Next
            ' --- 放值到 Chart 中 --- E N D

            stripMed = New StripLine()
            ChartObj.BorderColor = Color.Red
            Try
                stripMed.Title = oldMPID_Machine
                stripMed.ToolTip = ("Process : " + oldMPID_Machine)
                stripMed.IntervalOffset = d_offset
                stripMed.StripWidth = d_width
                stripMed.TitleAlignment = StringAlignment.Far
                stripMed.TitleLineAlignment = StringAlignment.Far
                If (colorChange Mod 2) = 0 Then
                    stripMed.BackColor = Color.FromArgb(99, 184, 255)
                Else
                    stripMed.BackColor = Color.FromArgb(171, 171, 171)
                End If
                ChartObj.ChartAreas("Default").AxisX.StripLines.Add(stripMed)
            Catch ex As Exception
            End Try

            ' --- 先算上下界, Group 要用 ---
            If specMin < minValue And specMin <> -99 Then
                minValue = specMin
            End If

            If specMax > maxValue And specMin <> -99 Then
                maxValue = specMax
            End If

            Dim Maxntemp As Double
            Dim Minntemp As Double
            Dim nInterval As Double = Math.Round((maxValue - minValue) / 5, 4)
            If nInterval <> 0 Then
                ' --- Max ---
                Maxntemp = (maxValue + nInterval)
                Maxntemp = Math.Round(Maxntemp, 2)
                ' --- Min ---
                Minntemp = (minValue - nInterval)
                Minntemp = Math.Round(Minntemp, 2)
                If Maxntemp <> Minntemp Then
                    ChartObj.ChartAreas("Default").AxisY.Maximum = Maxntemp
                    ChartObj.ChartAreas("Default").AxisY.Minimum = Minntemp
                End If
            End If

            ' ----- Title -----
            If notDetail Then
                If FunctionType = "Critical_Lot" Then
                    ChartObj.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (cinfo.partId) + "','" + (cinfo.ParametricItem) + "','" + (valueType) + "','" + isHLStr + "')"
                Else
                    ChartObj.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','IPP','" + (KPP_Part) + "','" + (KPP_YieldImpact) + "','" + (KPP_KeyModule) + "','" + (KPP_CriticalItem) + "','" + (KPP_IPP) + "','" + (isHLStr) + "')"
                End If
            End If
            ChartObj.Titles.Add("Lot Count :" + cinfo.dt.Rows.Count.ToString)
            ChartObj.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            ChartObj.Titles(1).Color = Color.Black
            ' ----- Title -----

            ChartObj.ChartAreas("Default").AxisX.Maximum = cinfo.dt.Rows.Count
            If Math.Round(cinfo.dt.Rows.Count / 15) > 1 Then
                ChartObj.ChartAreas("Default").AxisX.Interval = Math.Round(cinfo.dt.Rows.Count / 10)
            End If
            ChartObj.Series(0).ShowInLegend = False
            'ChartObj.ChartAreas("Default").AxisX2.Enabled = AxisEnabled.True
            'ChartObj.ChartAreas("Default").AxisX2.LineStyle = ChartDashStyle.DashDot
        End If

    End Sub

    Public Function GetPointColor(iRowIndex As Integer, ByRef objDT As DataTable) As System.Drawing.Color

        Dim iWECO_Rule1 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule1")
        Dim iWECO_Rule2 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule2")
        Dim iWECO_Rule3 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule3")
        Dim iWECO_Rule4 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule4")
        Dim iWECO_Rule5 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule5")
        Dim iWECO_Rule6 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule6")
        Dim iWECO_Rule7 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule7")
        Dim iWECO_Rule8 As Boolean = objDT.Rows(iRowIndex)("WECO_Rule8")

        If iWECO_Rule1 = True Then
            Return Color.Red
        ElseIf iWECO_Rule3 = True Then
            Return Color.LightSalmon
        ElseIf iWECO_Rule6 = True Then
            Return Color.Gray
        ElseIf iWECO_Rule5 = True Then
            Return Color.Fuchsia
        ElseIf iWECO_Rule4 = True Then
            Return Color.IndianRed
        ElseIf iWECO_Rule7 = True Then
            Return Color.MediumOrchid
        ElseIf iWECO_Rule2 = True Then
            Return Color.LimeGreen
        ElseIf iWECO_Rule8 = True Then
            Return Color.Gold
        Else
            Return Color.Blue
        End If

    End Function

    Public Function TraslateWECORuleTip(ByRef dataRow As DataRow) As String

        Dim sDesc As String = ""
        If dataRow("WECO_Rule1") = True Then
            sDesc = sDesc & "1 & "
        End If
        If dataRow("WECO_Rule2") = True Then
            sDesc = sDesc & "2 & "
        End If
        If dataRow("WECO_Rule3") = True Then
            sDesc = sDesc & "3 & "
        End If
        If dataRow("WECO_Rule4") = True Then
            sDesc = sDesc & "4 & "
        End If
        If dataRow("WECO_Rule5") = True Then
            sDesc = sDesc & "5 & "
        End If
        If dataRow("WECO_Rule6") = True Then
            sDesc = sDesc & "6 & "
        End If
        If dataRow("WECO_Rule7") = True Then
            sDesc = sDesc & "7 & "
        End If
        If dataRow("WECO_Rule8") = True Then
            sDesc = sDesc & "8 & "
        End If

        If sDesc.Length > 0 Then
            sDesc = Left(sDesc, sDesc.Length - 2)
            sDesc = "Violate WECO Rule " & sDesc
        End If
        Return sDesc

    End Function

    ' --- Mean 上下 3& ---
    Public Sub Chart_xUCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart, ByVal bLocation As Boolean)

        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("xUCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Red
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.xUCL = -99 Then
            Exit Sub
        End If

        If cinfo.dt.Rows.Count > 0 Then
            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                If cinfo.dt.Rows(i)("xUCL") <> -99 Then

                    tmpDouble = CType(cinfo.dt.Rows(i)("xUCL"), Double)
                    objPoint = New DataPoint(i, tmpDouble)
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("xUCL:{0:##0.###}", cinfo.dt.Rows(i)("xUCL"))

                End If
            Next
            If bLocation = False Then
                series.Points(series.Points.Count - 1).Label = String.Format("xUCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xUCL"))
                series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            Else
                series.Points(0).Label = String.Format("xUCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xUCL"))
                series.Points(0).LabelBackColor = Color.Yellow
            End If

            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15
        End If
        Chart.Series("xUCL").LegendText = "xUCL"

    End Sub
    Public Sub Chart_xLCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart, ByVal bLocation As Boolean)

        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("xLCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.DarkRed
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.xLCL = -99 Then
            Exit Sub
        End If

        'Series Data
        If cinfo.dt.Rows.Count > 0 Then

            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1

                If cinfo.dt.Rows(i)("xLCL") <> -99 Then

                    tmpDouble = CType(cinfo.dt.Rows(i)("xLCL"), Double)
                    objPoint = New DataPoint(i, tmpDouble)
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("xLCL:{0:##0.###}", cinfo.dt.Rows(i)("xLCL"))

                End If
            Next
            If bLocation = True Then
                series.Points(0).Label = String.Format("xLCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xLCL"))
                series.Points(0).LabelBackColor = Color.Yellow

            Else
                series.Points(series.Points.Count - 1).Label = String.Format("xLCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xLCL"))
                series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow

            End If
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("xLCL").LegendText = "xLCL"

    End Sub
    Public Sub Chart_xCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart, ByVal bLocation As Boolean)

        Dim series As Series
        series = Chart.Series.Add("xCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Coral
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.xUCL = -99 Or cinfo.xLCL = -99 Then
            series.ShowInLegend = False
            Exit Sub
        End If

        If cinfo.dt.Rows.Count > 0 Then

            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                If cinfo.dt.Rows(i)("xCL") <> -99 Then
                    objPoint = New DataPoint(i, CType(cinfo.dt.Rows(i)("xCL"), Double))
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("xCL:{0:##0.###}", cinfo.dt.Rows(i)("xCL"))
                End If
            Next

            If bLocation = True Then
                series.Points(0).Label = String.Format("xCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xCL"))
                series.Points(0).LabelBackColor = Color.Yellow
            Else
                series.Points(series.Points.Count - 1).Label = String.Format("xCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xCL"))
                series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            End If

            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("xCL").LegendText = "xCL"

    End Sub
    Public Sub Chart_xSigma(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)

        Dim series As Series
        Dim arrSigmaLabel As String() = {"xSigmaU1x", "xSigmaU2x", "xSigmaL1x", "xSigmaL2x"}
        Dim arrSerialValue(3) As Double
        arrSerialValue(0) = cinfo.xCL + cinfo.xSigma
        arrSerialValue(1) = cinfo.xCL + 2 * cinfo.xSigma
        arrSerialValue(2) = cinfo.xCL - cinfo.xSigma
        arrSerialValue(3) = cinfo.xCL - 2 * cinfo.xSigma

        For iSerialIndex As Integer = 0 To arrSigmaLabel.Length - 1

            series = Chart.Series.Add(arrSigmaLabel(iSerialIndex))
            series.ChartArea = "Default"
            series.Type = SeriesChartType.StepLine
            series.Color = Color.Pink
            series.BorderWidth = 1
            series.BorderStyle = ChartDashStyle.Dash
            series.MarkerStyle = MarkerStyle.None
            series.MarkerSize = 0
            series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
            series.ShowLabelAsValue = False
            series.ShowInLegend = False
            series("LabelStyle") = "Top"
            If cinfo.xUCL = -99 Then
                Exit Sub
            End If

            If cinfo.dt.Rows.Count > 0 Then
                For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                    Chart.Series(arrSigmaLabel(iSerialIndex)).Points.AddXY(i, arrSerialValue(iSerialIndex))
                    Chart.Series(arrSigmaLabel(iSerialIndex)).Points(i).ToolTip = arrSigmaLabel(iSerialIndex) & ": " + arrSerialValue(iSerialIndex).ToString()
                Next
            End If

        Next

    End Sub

    ' --- Std 上下 3& ---
    Public Sub Chart_sUCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)

        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("sUCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Red
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.xUCL = -99 Then
            Exit Sub
        End If

        If cinfo.dt.Rows.Count > 0 Then
            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                If (Not IsDBNull(cinfo.dt.Rows(i)("sUCL"))) Then

                    If (cinfo.dt.Rows(i)("sUCL") <> -99) Then
                        tmpDouble = CType(cinfo.dt.Rows(i)("sUCL"), Double)
                        objPoint = New DataPoint(i, tmpDouble)
                        series.Points.Add(objPoint)
                        objPoint.ToolTip = String.Format("sUCL:{0:##0.###}", cinfo.dt.Rows(i)("sUCL"))
                    End If

                End If
            Next
            series.Points(series.Points.Count - 1).Label = String.Format("sUCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("sUCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15
        End If
        Chart.Series("sUCL").LegendText = "sUCL"

    End Sub
    Public Sub Chart_sLCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)


        Dim tmpDouble As Double
        Dim series As Series
        series = Chart.Series.Add("sLCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.DarkRed
        series.BorderWidth = 2
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.xLCL = -99 Then
            Exit Sub
        End If

        'Series Data
        If cinfo.dt.Rows.Count > 0 Then

            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                If (Not IsDBNull(cinfo.dt.Rows(i)("sLCL"))) Then

                    If (cinfo.dt.Rows(i)("sLCL") <> -99) Then
                        tmpDouble = CType(cinfo.dt.Rows(i)("sLCL"), Double)
                        objPoint = New DataPoint(i, tmpDouble)
                        series.Points.Add(objPoint)
                        objPoint.ToolTip = String.Format("sLCL:{0:##0.###}", cinfo.dt.Rows(i)("sLCL"))
                    End If
                    

                End If
            Next
            series.Points(series.Points.Count - 1).Label = String.Format("sLCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("sLCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart.Series("sLCL").LegendText = "sLCL"

    End Sub
    Private Sub Chart_sCL(ByRef Chart3 As Dundas.Charting.WebControl.Chart, ByVal cinfo As IPPChart)
        'Series Type
        Dim series As Series
        series = Chart3.Series.Add("sCL")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = Color.Coral
        series.BorderWidth = 3
        series.MarkerStyle = MarkerStyle.None
        series.MarkerSize = 0
        series.Font = New Font("Times New Roman", 8, FontStyle.Regular)
        series.ShowLabelAsValue = False
        series("LabelStyle") = "Top"

        If cinfo.sUCL = -99 Or cinfo.sLCL = -99 Then
            series.ShowInLegend = False
            Exit Sub
        End If

        'Series Data
        If cinfo.dt.Rows.Count > 0 Then

            Dim objPoint As DataPoint
            For i As Integer = 0 To cinfo.dt.Rows.Count - 1
                If Not IsDBNull(cinfo.dt.Rows(i)("sCL")) Then
                    objPoint = New DataPoint(i, CType(cinfo.dt.Rows(i)("sCL"), Double))
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("sCL:{0:##0.###}", cinfo.dt.Rows(i)("sCL"))
                End If
            Next
            series.Points(series.Points.Count - 1).Label = String.Format("sCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("sCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If
        Chart3.Series("sCL").LegendText = "sCL"
        
    End Sub

End Class
