Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Data
Imports Dundas.Charting.WebControl

Public Class TrendObj

    Public txtDateFrom As String = ""
    Public txtDateTo As String = ""
    Public dataSource As String = ""
    Public notDetail As Boolean = True
    Public isHighlight As Boolean = False
    Public HL_Day As String = ""

    Public Structure IPPData
        Dim dt_data As DataTable
        Dim dt_item As DataTable
        Dim dt_yImpact As DataTable
    End Structure
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

    Public Sub Call_iPP_ByLot(ByRef ms As TrendObj.IPPData, ByRef Panel1 As Panel)

        Dim sql As String = ""
        Dim t1 As New DataTable
        Dim s1 As String = ""
        Dim dt As DataTable = ms.dt_data
        Dim dt_Item As DataTable = ms.dt_item
        Dim dt_yImpact As DataTable = ms.dt_yImpact

        Dim expression As String
        Dim sortOrder As String = ""
        Dim dt3 As New DataTable

        Dim new_column0 As DataColumn = New DataColumn
        new_column0.ColumnName = "SLI"
        new_column0.DataType = dt.Columns("SLI").DataType
        dt3.Columns.Add(new_column0)

        Dim new_column1 As DataColumn = New DataColumn
        new_column1.ColumnName = "Lot"
        new_column1.DataType = dt.Columns("Lot").DataType
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
        new_column14.DataType = GetType(DateTime)
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

        If notDetail Then

            For j As Integer = 0 To (dt_yImpact.Rows.Count - 1)

                expression = "part='" + (dt_yImpact.Rows(j).Item("Part").ToString) + "' and yield_impact_item='" + (dt_yImpact.Rows(j).Item("Yield_Impact_Item").ToString) + "' and EDA_Item='" + (dt_yImpact.Rows(j).Item("EDA_Item").ToString) + "'"
                sortOrder = "part, yield_impact_item, EDA_Item, trtm, lot asc"
                Dim foundRows() As DataRow
                foundRows = dt.Select(expression, sortOrder)

                Dim firPart As String = ""
                Dim lstPart As String = ""
                Dim temp1 As Double = -99
                Dim temp2 As Double = -99
                Dim temp3 As Double = -99
                Dim temp4 As Double = -99

                If foundRows.Length > 0 Then

                    dt3.Clear()
                    Dim ChartInfo As New TrendObj.IPPChart
                    ChartInfo.dt = dt3

                    For k As Integer = 0 To (foundRows.Length - 1)

                        Dim pDROW_Row As DataRow = dt3.NewRow
                        If foundRows(k).Item("Lot") <> foundRows(k).Item("SLI") Then
                            pDROW_Row("SLI") = foundRows(k).Item("SLI")
                            pDROW_Row("XLableValue") = foundRows(k).Item("Lot") + "-" + foundRows(k).Item("SLI") + "[" + Convert.ToDateTime(foundRows(k).Item("trtm")).ToString("yyyy/MM/dd") + "]"
                        Else
                            pDROW_Row("SLI") = ""
                            pDROW_Row("XLableValue") = foundRows(k).Item("Lot") + "[" + Convert.ToDateTime(foundRows(k).Item("trtm")).ToString("yyyy/MM/dd") + "]"
                        End If
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

                    Next

                    ChartInfo.sLayer = foundRows(0).Item("Layer")
                    ChartInfo.sMIS = foundRows(0).Item("MIS_OP")

                    ChartInfo.sLot = ""
                    ChartInfo.sSLI = ""

                    ChartInfo.ParametricItem = foundRows(0).Item("Parametric_Measurement")
                    ChartInfo.XLableValue = "Lot"

                    ChartInfo.nType = "oldm"
                    ChartInfo.YValue = "meanval"

                    ChartInfo.yImpact = foundRows(0).Item("Yield_Impact_Item")
                    ChartInfo.kModule = foundRows(0).Item("Key_Module")
                    ChartInfo.Critical = foundRows(0).Item("Critical_item")
                    ChartInfo.partId = foundRows(0).Item("Part_id")
                    ChartInfo.edaItem = (dt_yImpact.Rows(j).Item("EDA_Item").ToString)

                    Dim Chart1 As New Dundas.Charting.WebControl.Chart()
                    CallDundals(Chart1, ChartInfo, Panel1)

                End If

            Next

        Else

            expression = "1=1"
            sortOrder = ""
            Dim foundRows() As DataRow
            foundRows = dt.Select(expression, sortOrder)

            Dim firPart As String = ""
            Dim lstPart As String = ""
            Dim temp1 As Double = -99
            Dim temp2 As Double = -99
            Dim temp3 As Double = -99
            Dim temp4 As Double = -99

            If foundRows.Length > 0 Then

                dt3.Clear()
                Dim ChartInfo As New TrendObj.IPPChart
                ChartInfo.dt = dt3

                For k As Integer = 0 To foundRows.Length - 1

                    Dim pDROW_Row As DataRow = dt3.NewRow
                    If foundRows(k).Item("Lot") <> foundRows(k).Item("SLI") Then
                        pDROW_Row("SLI") = foundRows(k).Item("SLI")
                        pDROW_Row("XLableValue") = foundRows(k).Item("Lot") + "-" + foundRows(k).Item("SLI") + "[" + Convert.ToDateTime(foundRows(k).Item("trtm")).ToString("yyyy/MM/dd") + "]"
                    Else
                        pDROW_Row("SLI") = ""
                        pDROW_Row("XLableValue") = foundRows(k).Item("Lot") + "[" + Convert.ToDateTime(foundRows(k).Item("trtm")).ToString("yyyy/MM/dd") + "]"
                    End If
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

                Next

                ChartInfo.sLayer = foundRows(0).Item("Layer")
                ChartInfo.sMIS = foundRows(0).Item("MIS_OP")

                ChartInfo.sLot = ""
                ChartInfo.sSLI = ""

                ChartInfo.ParametricItem = foundRows(0).Item("Parametric_Measurement")
                ChartInfo.XLableValue = "Lot"

                ChartInfo.nType = "oldm"
                ChartInfo.YValue = "meanval"

                ChartInfo.yImpact = foundRows(0).Item("Yield_Impact_Item")
                ChartInfo.kModule = foundRows(0).Item("Key_Module")
                ChartInfo.Critical = foundRows(0).Item("Critical_item")
                ChartInfo.partId = foundRows(0).Item("Part_id")

                Dim Chart1 As New Dundas.Charting.WebControl.Chart()
                CallDundals(Chart1, ChartInfo, Panel1)

            End If

        End If

    End Sub
    Public Sub CallDundals(ByVal Chart3 As Dundas.Charting.WebControl.Chart, ByRef cinfo As TrendObj.IPPChart, ByRef Panel1 As Panel)

        ChartType(Chart3, cinfo)
        ChartData(Chart3, cinfo)
        Chart_xLCL(Chart3, cinfo)
        Chart_xUCL(Chart3, cinfo)
        Chart_xCL(Chart3, cinfo)
        Chart_xSigma(Chart3, cinfo)
        Panel1.Controls.Add(Chart3)
        Panel1.Controls.Add(New LiteralControl("<br>"))

    End Sub
    Public Sub ChartType(ByRef Chart3 As Dundas.Charting.WebControl.Chart, ByVal cinfo As TrendObj.IPPChart)

        Chart3.Palette = ChartColorPalette.Dundas
        Chart3.Height = Unit.Pixel(600) 'big
        Chart3.Width = Unit.Pixel(1080) 'midd
        Chart3.Palette = ChartColorPalette.Dundas
        Chart3.BackColor = Color.White
        Chart3.BackGradientEndColor = Color.Peru
        Chart3.ChartAreas.Add("Default")
        Chart3.UI.Toolbar.Enabled = False
        Chart3.UI.ContextMenu.Enabled = True
        Chart3.Titles.Add(cinfo.partId + ":" + cinfo.yImpact + ":" + cinfo.kModule + ":" + cinfo.Critical + "  --  " + cinfo.ParametricItem + "(Mean)")
        Chart3.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart3.Titles(0).Color = Color.DarkBlue
        ' Set AxisX			
        Chart3.ChartAreas("Default").AxisX.LabelStyle.Enabled = True
        ' Set AxisY 
        'Chart3.ChartAreas("Default").AxisX.MajorGrid.LineColor = Color.LightGray
        'Chart3.ChartAreas("Default").AxisY.MajorGrid.LineColor = Color.LightGray
        'Chart3.ChartAreas("Default").AxisY.MajorGrid.LineStyle = ChartDashStyle.Dash
        Chart3.ChartAreas("Default").AxisX.MajorGrid.Enabled = False
        Chart3.ChartAreas("Default").AxisY.MajorGrid.Enabled = False
        'Chart3.ChartAreas("Default").AxisY.Title = "Data"
        Chart3.ChartAreas("Default").AxisY.TitleFont = New Font("Arial", 8, FontStyle.Regular)
        Chart3.ChartAreas("Default").AxisY.TitleColor = Color.Black
        Chart3.ChartAreas("Default").AxisY.LabelsAutoFit = True
        Chart3.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        Chart3.ChartAreas("Default").AxisY.LabelStyle.FontColor = Color.Black
        Chart3.ChartAreas("Default").AxisX.LabelStyle.Font = New Font("Arial", 8, FontStyle.Regular)
        Chart3.ChartAreas("Default").AxisX.LabelStyle.FontColor = Color.Black
        Chart3.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -40
        Chart3.ChartAreas("Default").BackColor = Color.White
        Chart3.ChartAreas("Default").AxisX.LineColor = Color.Black
        Chart3.ChartAreas("Default").AxisY.LineColor = Color.Black

        Chart3.ChartAreas("Default").AxisX.Title = "SLI"
        Chart3.ChartAreas("Default").AxisX.TitleFont = New Font("Arial", 10, FontStyle.Regular)
        Chart3.ChartAreas("Default").AxisX.TitleColor = Color.White

        Chart3.Legends(0).LegendStyle = LegendStyle.Row
        Chart3.Legends(0).BackColor = Color.White
        Chart3.Legends(0).Alignment = StringAlignment.Center
        Chart3.Legends(0).Docking = LegendDocking.Top
        Chart3.Legends(0).FontColor = Color.DarkBlue
        Chart3.Legends(0).LegendStyle = LegendStyle.Table
        Chart3.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
        Chart3.ImageType = ChartImageType.Png
        
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
        Chart3.Legends(0).CustomItems.Add(objPoint)

        If cinfo.YValue = "meanval" Then

            Dim objWECORule1 As LegendItem = New LegendItem()
            objWECORule1.MarkerSize = iLengendPointSize
            objWECORule1.Name = "WECO Rule 1"
            objWECORule1.Style = LegendImageStyle.Marker
            objWECORule1.MarkerColor = Color.Red
            objWECORule1.BorderColor = Color.Red
            objWECORule1.ToolTip = "單點超出管制界限" & vbCrLf & "A single point beyond either control limit"
            Chart3.Legends(0).CustomItems.Add(objWECORule1)

            Dim objWECORule2 As LegendItem = New LegendItem()
            objWECORule2.MarkerSize = iLengendPointSize
            objWECORule2.Name = "WECO Rule 2"
            objWECORule2.Style = LegendImageStyle.Marker
            objWECORule2.MarkerColor = Color.LimeGreen
            objWECORule2.BorderColor = Color.LimeGreen
            objWECORule2.ToolTip = "連續九點在同一邊" & vbCrLf & "9 consecutive points on the same side of the centerline"
            Chart3.Legends(0).CustomItems.Add(objWECORule2)

            Dim objWECORule3 As LegendItem = New LegendItem()
            objWECORule3.MarkerSize = iLengendPointSize
            objWECORule3.Name = "WECO Rule 3"
            objWECORule3.Style = LegendImageStyle.Marker
            objWECORule3.MarkerColor = Color.LightSalmon
            objWECORule3.BorderColor = Color.LightSalmon
            objWECORule3.ToolTip = "連續六點上升或是連續六點下降" & vbCrLf & "6 consecutive points steadily increasing or decreasing"
            Chart3.Legends(0).CustomItems.Add(objWECORule3)

            Dim objWECORule4 As LegendItem = New LegendItem()
            objWECORule4.MarkerSize = iLengendPointSize
            objWECORule4.Name = "WECO Rule 4"
            objWECORule4.Style = LegendImageStyle.Marker
            objWECORule4.MarkerColor = Color.IndianRed
            objWECORule4.BorderColor = Color.IndianRed
            objWECORule4.ToolTip = "連續14點上下跳動" & vbCrLf & "14 (or more) consecutive points are alternating up and down"
            Chart3.Legends(0).CustomItems.Add(objWECORule4)

            Dim objWECORule5 As LegendItem = New LegendItem()
            objWECORule5.MarkerSize = iLengendPointSize
            objWECORule5.Name = "WECO Rule 5"
            objWECORule5.Style = LegendImageStyle.Marker
            objWECORule5.MarkerColor = Color.Fuchsia
            objWECORule5.BorderColor = Color.Fuchsia
            objWECORule5.ToolTip = "連續3點在同側且有2點至少落在超過2倍標準差外" & vbCrLf & "2 out of 3 consecutive points at least 2 std dev beyond the centerline, on the same side"
            Chart3.Legends(0).CustomItems.Add(objWECORule5)

            Dim objWECORule6 As LegendItem = New LegendItem()
            objWECORule6.MarkerSize = iLengendPointSize
            objWECORule6.Name = "WECO Rule 6"
            objWECORule6.Style = LegendImageStyle.Marker
            objWECORule6.MarkerColor = Color.Gray
            objWECORule6.BorderColor = Color.Gray
            objWECORule6.ToolTip = "連續5點在同側且有4點至少落在超過1倍標準差外" & vbCrLf & "4 out of 5 consecutive points on the chart are more than 1 std dev away from the CL"
            Chart3.Legends(0).CustomItems.Add(objWECORule6)

            Dim objWECORule7 As LegendItem = New LegendItem()
            objWECORule7.MarkerSize = iLengendPointSize
            objWECORule7.Name = "WECO Rule 7"
            objWECORule7.Style = LegendImageStyle.Marker
            objWECORule7.MarkerColor = Color.MediumOrchid
            objWECORule7.BorderColor = Color.MediumOrchid
            objWECORule7.ToolTip = "連續15點在1倍標準差間" & vbCrLf & "15 (or more) consecutive points are within 1 std dev of the CL"
            Chart3.Legends(0).CustomItems.Add(objWECORule7)

            Dim objWECORule8 As LegendItem = New LegendItem()
            objWECORule8.MarkerSize = iLengendPointSize
            objWECORule8.Name = "WECO Rule 8"
            objWECORule8.Style = LegendImageStyle.Marker
            objWECORule8.MarkerColor = Color.Gold
            objWECORule8.BorderColor = Color.Gold
            objWECORule8.ToolTip = "連續8點在中心2側，但皆不在1倍標準差內" & vbCrLf & "8 (or more) consecutive points are on both sides of the CL, but none are within 1 std dev of it."
            Chart3.Legends(0).CustomItems.Add(objWECORule8)

        End If

        Chart3.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart3.BorderStyle = ChartDashStyle.Solid
        Chart3.BorderWidth = 3
        Chart3.BorderColor = Color.DarkBlue
        Chart3.ChartAreas("Default").AxisX.Interval = 1
        Chart3.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

    End Sub
    Public Sub ChartData(ByRef Chart3 As Dundas.Charting.WebControl.Chart, ByVal cinfo As TrendObj.IPPChart)

        Dim series As Series
        series = Chart3.Series.Add("MQCS")
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

            Dim sXaxis As String
            Dim nYvalue As Double
            Dim isHLStr As String = "N"
            If isHighlight Then
                isHLStr = "Y"
            End If

            For i As Integer = 0 To (cinfo.dt.Rows.Count - 1)

                sXaxis = IIf(IsDBNull(cinfo.dt.Rows(i).Item("XLableValue")), "", cinfo.dt.Rows(i).Item("XLableValue"))
                nYvalue = IIf(IsDBNull(cinfo.dt.Rows(i).Item(cinfo.YValue)), 0, cinfo.dt.Rows(i).Item(cinfo.YValue))
                nYvalue = Math.Round(nYvalue, 3)

                Chart3.Series(0).Points.AddXY(i, nYvalue)
                Chart3.Series(0).Points(i).AxisLabel = sXaxis

                If notDetail Then
                    Chart3.Series(0).Points(i).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (dataSource) + "','" + (cinfo.partId) + "','" + (cinfo.yImpact) + "','" + (cinfo.kModule) + "','" + (cinfo.Critical) + "','" + (cinfo.edaItem) + "','" + isHLStr + "')"
                End If

                Chart3.Series(0).Points(i).MarkerColor = GetPointColor(i, cinfo.dt)
                If cinfo.dt.Rows(i)("ViolateRule").ToString().Length <= 0 Then
                    Chart3.Series(0).Points(i).ToolTip = "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & "SLI=" & cinfo.dt.Rows(i)("SLI") & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                Else
                    Chart3.Series(0).Points(i).ToolTip = cinfo.dt.Rows(i)("ViolateRule") & vbCrLf & "Lot_ID=" & cinfo.dt.Rows(i)("Lot") & vbCrLf & "SLI=" & cinfo.dt.Rows(i)("SLI") & vbCrLf & "Value=" & nYvalue.ToString & vbCrLf & "Date=" & Convert.ToDateTime(cinfo.dt.Rows(i)("trtm")).ToString("yyyy/MM/dd HH:mm:ss")
                End If

                '2012/10/19 modified by Chery , display by WECO rule color.
                Select Case cinfo.nType
                    Case "olds"
                        If nYvalue > cinfo.sUCL And cinfo.sUCL <> -99 Then
                            Chart3.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                        If nYvalue < cinfo.sLCL And cinfo.sLCL <> -99 Then
                            Chart3.Series(0).Points(i).MarkerColor = Color.Red
                        End If
                End Select

                If nYvalue > cinfo.nMAX Then
                    cinfo.nMAX = nYvalue
                End If

                If nYvalue < cinfo.nMIN Then
                    cinfo.nMIN = nYvalue
                End If

            Next

            ' --- 如果使用選擇 Daily Event 要 HighLight Range ---
            If isHighlight Then
                Chart3.BorderColor = Color.Red
                Try
                    Dim rowDataTime As String
                    Dim Rangeseries As Series
                    Rangeseries = Chart3.Series.Add("SplineRange")
                    Rangeseries.Type = SeriesChartType.Range
                    Rangeseries.Color = Color.FromArgb(70, 252, 180, 65)
                    Rangeseries.ShowInLegend = False
                    For i As Integer = 0 To (cinfo.dt.Rows.Count - 1)
                        rowDataTime = IIf(IsDBNull(cinfo.dt.Rows(i).Item("HL_Day")), "", cinfo.dt.Rows(i).Item("HL_Day"))
                        nYvalue = IIf(IsDBNull(cinfo.dt.Rows(i).Item(cinfo.YValue)), 0, cinfo.dt.Rows(i).Item(cinfo.YValue))
                        nYvalue = Math.Round(nYvalue, 3)
                        ' 如果使用選擇 Daily Event 要 HighLight Range
                        If (Me.HL_Day = rowDataTime) Then
                            Rangeseries.Points.AddXY(i, nYvalue)
                        Else
                            Rangeseries.Points.AddXY(i, 0)
                        End If
                    Next
                Catch ex As Exception

                End Try
            End If

            ' ----- Title -----
            If notDetail Then
                Chart3.Titles(0).Href = "javascript:openWin('" + (txtDateFrom) + "','" + (txtDateTo) + "','" + (dataSource) + "','" + (cinfo.partId) + "','" + (cinfo.yImpact) + "','" + (cinfo.kModule) + "','" + (cinfo.Critical) + "','" + (cinfo.edaItem) + "','" + isHLStr + "')"
            End If
            Chart3.Titles.Add("Lot Count :" + cinfo.dt.Rows.Count.ToString)
            Chart3.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart3.Titles(1).Color = Color.Black
            ' ----- Title -----

            Dim ntemp As Double
            Dim nInterval As Double = Math.Round((cinfo.nMAX - cinfo.nMIN) / 6, 4)
            If nInterval <> 0 Then
                ' --- Max ---
                ntemp = (cinfo.nMAX + nInterval)
                ntemp = Math.Round(ntemp, 2)
                Chart3.ChartAreas("Default").AxisY.Maximum = ntemp
                ' --- Min ---
                ntemp = (cinfo.nMIN - nInterval)
                ntemp = Math.Round(ntemp, 2)
                Chart3.ChartAreas("Default").AxisY.Minimum = ntemp
            End If

            Chart3.ChartAreas("Default").AxisX.Maximum = cinfo.dt.Rows.Count
            If Math.Round(cinfo.dt.Rows.Count / 15) > 1 Then
                Chart3.ChartAreas("Default").AxisX.Interval = Math.Round(cinfo.dt.Rows.Count / 10)
            End If
            Chart3.Series(0).ShowInLegend = False

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

        If iWECO_Rule8 = True Then
            Return Color.Gold
        ElseIf iWECO_Rule7 = True Then
            Return Color.MediumOrchid
        ElseIf iWECO_Rule6 = True Then
            Return Color.Gray
        ElseIf iWECO_Rule5 = True Then
            Return Color.Fuchsia
        ElseIf iWECO_Rule4 = True Then
            Return Color.IndianRed
        ElseIf iWECO_Rule3 = True Then
            Return Color.LightSalmon
        ElseIf iWECO_Rule2 = True Then
            Return Color.LimeGreen
        ElseIf iWECO_Rule1 = True Then
            Return Color.Red
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
    Public Sub Chart_xLCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)

        'Series Type
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
                    objPoint = New DataPoint(i, CType(cinfo.dt.Rows(i)("xLCL"), Double))
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("xLCL:{0:##0.###}", cinfo.dt.Rows(i)("xLCL"))
                End If
            Next
            series.Points(series.Points.Count - 1).Label = String.Format("xLCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xLCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15

        End If

        Chart.Series("xLCL").LegendText = "xLCL"

    End Sub
    Public Sub Chart_xUCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)

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
                    objPoint = New DataPoint(i, CType(cinfo.dt.Rows(i)("xUCL"), Double))
                    series.Points.Add(objPoint)
                    objPoint.ToolTip = String.Format("xUCL:{0:##0.###}", cinfo.dt.Rows(i)("xUCL"))
                End If
            Next
            series.Points(series.Points.Count - 1).Label = String.Format("xUCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xUCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
            series.SmartLabels.Enabled = True
            series.SmartLabels.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Partial
            series.SmartLabels.MarkerOverlapping = False
            series.SmartLabels.MinMovingDistance = 15
        End If

        Chart.Series("xUCL").LegendText = "xUCL"

    End Sub
    Public Sub Chart_xCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef cinfo As IPPChart)

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
            series.Points(series.Points.Count - 1).Label = String.Format("xCL:{0:##0.###}", cinfo.dt.Rows(cinfo.dt.Rows.Count - 1)("xCL"))
            series.Points(series.Points.Count - 1).LabelBackColor = Color.Yellow
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

End Class
