Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing
Imports Dundas.Charting.WebControl

Partial Class YieldLoss
    Inherits System.Web.UI.Page

    Private Const gChartH As Integer = 600
    Private Const gChartW As Integer = 1080
    Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.Green, Color.Firebrick, Color.DodgerBlue, Color.Olive, Color.DarkGreen, Color.Red, Color.Gold, Color.Gray, Color.Cyan}
    Dim plantAry() As String = {"1", "2", "3", "4", "5", "B", "T", "S", "All"}

    Private Structure FailObj

        Dim Fail_Mode As String
        Dim Fail_Value As Double

    End Structure

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById(""lab_wait"").innerText='Please wait ......';" & _
                                                 "javascript:document.getElementById(""but_Execute"").disabled=true;" & _
                                                  Me.Page.GetPostBackEventReference(but_Execute))
        If Not (Me.IsPostBack) Then
            pageInit()
        End If

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try

            conn.Open()

            ' -- Customer ID --
            sqlStr = "select customer_id from Customer_Prodction_Mapping where 1=1 "
            sqlStr += "group by customer_id order by customer_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlCustomer, 0)

            ' -- Product --
            sqlStr = "select Category from Customer_Prodction_Mapping where 1=1 "
            sqlStr += "group by Category order by Category"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlProduct, 0)

            ' -- Part ID --
            sqlStr = "select production_type from Customer_Prodction_Mapping where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
            sqlStr += "group by production_type order by production_type"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlPart, 0)

            ' -- Week --
            sqlStr = "select ww from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by ww order by ww desc"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)

            ' -- Yield Loss Item --
            sqlStr = "select fail_mode from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by fail_mode order by fail_mode"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' DDL --- Product Change
    Protected Sub ddlProduct_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlProduct.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False

        Try

            conn.Open()

            ' -- Part ID --
            sqlStr = "select production_type from Customer_Prodction_Mapping where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
            sqlStr += "group by production_type order by production_type"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlPart, 0)

            ' -- Week --
            sqlStr = "select ww from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by ww order by ww desc"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)

            ' -- Yield Loss Item --
            sqlStr = "select fail_mode from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by fail_mode order by fail_mode"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' DDL --- Part Change
    Protected Sub ddlPart_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlPart.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False

        Try

            conn.Open()
            ' -- Week --
            sqlStr = "select ww from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by ww order by ww desc"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)

            ' -- Yield Loss Item --
            sqlStr = "select fail_mode from dbo.BinCode_Summary where 1=1 "
            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
            sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            sqlStr += "group by fail_mode order by fail_mode"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Report Week Item
    Protected Sub rbl_week_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rbl_week.SelectedIndexChanged

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        lb_weekSource.Items.Clear()
        lb_weekShow.Items.Clear()
        tr_week.Visible = False

        If rbl_week.SelectedIndex = 1 Then

            Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            Dim sqlStr As String = ""
            Dim myDT As DataTable
            Dim myAdapter As SqlDataAdapter

            Try

                conn.Open()
                sqlStr = "select max(ww) from dbo.BinCode_Summary where 1=1 "
                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
                sqlStr += "group by ww order by ww desc"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)
                conn.Close()

            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try
            tr_week.Visible = True

        End If

    End Sub

    ' Week To >>
    Protected Sub but_weekTo_Click(sender As Object, e As System.EventArgs) Handles but_weekTo.Click

        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList
        For i As Integer = 0 To (lb_weekSource.Items.Count - 1)
            If lb_weekSource.Items(i).Selected Then
                DestAry.Add(lb_weekSource.Items(i).Value)
            Else
                sourceAry.Add(lb_weekSource.Items(i).Value)
            End If
        Next

        lb_weekSource.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_weekSource.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_weekShow.Items.Add(DestAry(i).ToString())
        Next

    End Sub

    ' Week Back <<
    Protected Sub but_weekBack_Click(sender As Object, e As System.EventArgs) Handles but_weekBack.Click

        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList

        For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
            If lb_weekShow.Items(i).Selected Then
                DestAry.Add(lb_weekShow.Items(i).Value)
            Else
                sourceAry.Add(lb_weekShow.Items(i).Value)
            End If
        Next

        lb_weekShow.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_weekShow.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_weekSource.Items.Add(DestAry(i).ToString())
        Next

    End Sub

    ' Yield Loss Item 
    Protected Sub rbl_lossItem_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rbl_lossItem.SelectedIndexChanged

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        txb_ylInput.Text = ""
        lb_LossSource.Items.Clear()
        lb_LossShow.Items.Clear()
        tr_lossItem.Visible = False

        If rbl_lossItem.SelectedIndex = 2 Then

            Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            Dim sqlStr As String = ""
            Dim myDT As DataTable
            Dim myAdapter As SqlDataAdapter

            Try

                conn.Open()
                sqlStr = "select fail_mode from dbo.BinCode_Summary where 1=1 "
                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
                sqlStr += "group by fail_mode order by fail_mode"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)
                conn.Close()

            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

            tr_lossItem.Visible = True
        End If

    End Sub

    ' Item To >>
    Protected Sub but_lossItemTo_Click(sender As Object, e As System.EventArgs) Handles but_lossItemTo.Click

        txb_ylInput.Text = ""
        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList
        For i As Integer = 0 To (lb_LossSource.Items.Count - 1)
            If lb_LossSource.Items(i).Selected Then
                DestAry.Add(lb_LossSource.Items(i).Value)
            Else
                sourceAry.Add(lb_LossSource.Items(i).Value)
            End If
        Next

        lb_LossSource.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_LossSource.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_LossShow.Items.Add(DestAry(i).ToString())
        Next

    End Sub

    ' Item Back <<
    Protected Sub but_lossItemBack_Click(sender As Object, e As System.EventArgs) Handles but_lossItemBack.Click

        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList

        For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
            If lb_LossShow.Items(i).Selected Then
                DestAry.Add(lb_LossShow.Items(i).Value)
            Else
                sourceAry.Add(lb_LossShow.Items(i).Value)
            End If
        Next

        lb_LossShow.Items.Clear()

        For i As Integer = 0 To (sourceAry.Count - 1)
            lb_LossShow.Items.Add(sourceAry(i).ToString())
        Next

        For i As Integer = 0 To (DestAry.Count - 1)
            lb_LossSource.Items.Add(DestAry(i).ToString())
        Next

    End Sub

    ' InQuery
    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        lab_wait.Text = "無資料 !"

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim customStr As String = " "
        Dim plantStr As String = ""
        Dim partStr As String = " "
        Dim weekStr As String = " "
        Dim itemStr As String = " "
        Dim topStr As String = " "
        Dim myAdapter As SqlDataAdapter
        Dim topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 12 Then
            ShowMessage("選擇週數最多為 12 週")
            Exit Sub
        End If

        Try
            ' --- Customer ID ---
            customStr = "and customer_id='" + (ddlCustomer.SelectedValue.Trim()) + "' "
            ' --- Part ID ---
            partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "

            If ddlProduct.SelectedValue <> "CPU" Then
                plantStr = "and Plant='All' "
            Else
                plantStr = ""
            End If

            ' --- Week ID ---
            Dim weekTemp As String = ""
            If rbl_week.SelectedIndex = 1 Then
                For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                    weekTemp += (lb_weekShow.Items(i).Value) + ","
                Next
                weekTemp = weekTemp.Substring(0, (weekTemp.Length - 1))
                weekStr = "and ww in (" + weekTemp + ") "
            End If

            If rbl_lossItem.SelectedIndex = 0 Then
                topStr = "top(10)"
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                topStr = "top(20)"
            End If

            conn.Open()
            ' --- Yield Loss ID ---
            Dim itemTemp As String = ""
            itemStr = ""

            If rbl_lossItem.SelectedIndex = 2 Then
                ' === Custom Item ===
                For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                    itemTemp += "'" + (lb_LossShow.Items(i).Value).Replace("'", "''") + "',"
                Next
                itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
                itemStr = "and fail_mode in (" + itemTemp + ") "

                sqlStr = "select Fail_Mode, Fail_Ratio from dbo.BinCode_Summary where 1=1 "
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                sqlStr += itemStr

                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and ww in (select max(ww) from BinCode_Summary) "
                Else
                    ' Custom
                    sqlStr += "and ww in (select max(ww) from BinCode_Summary where 1=1 " + weekStr + " ) "
                End If
                sqlStr += "group by Fail_Mode, Fail_Ratio "
                sqlStr += "order by fail_ratio desc"
            Else
                ' === Top N === ' 如果選 Custom 就要呈現選擇的 item
                sqlStr = "select " + topStr + " Fail_Mode, Fail_Ratio from dbo.BinCode_Summary where 1=1 "
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr

                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and ww in (select max(ww) from BinCode_Summary) "
                Else
                    ' Custom
                    sqlStr += "and ww in (select max(ww) from BinCode_Summary where 1=1 " + weekStr + " ) "
                End If
                sqlStr += "group by Fail_Mode, Fail_Ratio "
                sqlStr += "order by fail_ratio desc"
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.Fill(topDT)

            If topDT.Rows(0)("Fail_Ratio").ToString() <> "0" Then

                ' === Raw Data ===
                lab_wait.Text = ""
                sqlStr = "select WW, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID, (Fail_Mode+MF_Stage) AS 'NewFailMode' "
                sqlStr += "from dbo.BinCode_Summary where 1=1 "
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr

                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and ww in (select top(4) ww from BinCode_Summary group by ww order by ww desc) "
                Else
                    ' Custom
                    sqlStr += weekStr
                End If

                sqlStr += itemStr
                sqlStr += "group by WW, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID "
                sqlStr += "order by ww desc, fail_ratio desc"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                rawDT = New DataTable
                myAdapter.Fill(rawDT)

                ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
                chipSetRawDT = New DataTable
                If ddlProduct.SelectedValue = "CS" Then

                    sqlStr = "select ww, plant, Fail_Mode, Fail_Ratio "
                    sqlStr += "from dbo.BinCode_Summary where 1=1 "
                    sqlStr += customStr
                    sqlStr += partStr

                    If rbl_week.SelectedIndex = 0 Then
                        sqlStr += "and ww in (select top(1) ww from BinCode_Summary group by ww order by ww desc) "
                    Else
                        sqlStr += "and ww in (select max(ww) from BinCode_Summary where 1=1 " + weekStr + " ) "
                    End If

                    sqlStr += itemStr
                    sqlStr += "group by ww, plant, Fail_Mode, Fail_Ratio "
                    sqlStr += "order by plant desc, fail_ratio desc"
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myAdapter.Fill(chipSetRawDT)

                End If

                conn.Close()
                BarChart(rawDT, topDT)

                If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
                    Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                    Dim Chart As New Dundas.Charting.WebControl.Chart()
                    DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
                    Chart_Panel.Controls.Add(Chart)
                End If

                tr_chartDisplay.Visible = True
                If cb_DRowData.Checked Then
                    showRowData(rawDT, chipSetRawDT)
                    tr_gvDisplay.Visible = True
                End If


            End If

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Bar Chart
    Private Sub BarChart(ByRef raw As DataTable, ByRef setupDT As DataTable)

        Dim FAry As ArrayList = New ArrayList
        Dim BHash As Hashtable = New Hashtable

        Dim Chart As New Dundas.Charting.WebControl.Chart()
        DrawBarChart(Chart, raw, setupDT, FAry, BHash)
        Chart_Panel.Controls.Add(Chart)

        Chart_Panel.Controls.Add(New LiteralControl("<br>"))

        Chart = New Dundas.Charting.WebControl.Chart()
        DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash)
        Chart_Panel.Controls.Add(Chart)

    End Sub

    ' Draw Yield Loss Chart
    Private Sub DrawBarChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable)

        Chart.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH)
        Chart.Width = Unit.Pixel(gChartW)

        Chart.Titles.Add("Fail Mode By Week")
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
        'Chart.ChartAreas("Default").AxisX.Title = "【 Fail Mode 】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -60 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, "WW")
        weekGroupDT.DefaultView.Sort = "WW asc"
        weekGroupDT = weekGroupDT.DefaultView.ToTable
        Dim dtFilter As DataTable
        Dim dr As DataRow
        Dim foundRows() As DataRow
        Dim insideRows() As DataRow
        Dim failMode As String
        Dim failValue As Double
        Dim weekStr As String
        Dim failObj As FailObj
        Dim colorInx As Integer = 0
        Dim scriptStr As String = ""

        Dim pieChartWeek As String = ""
        For i As Integer = 0 To (weekGroupDT.Rows.Count - 1)
            pieChartWeek += (weekGroupDT.Rows(i)("WW")).ToString + ","
        Next
        pieChartWeek = pieChartWeek.Substring(0, (pieChartWeek.Length - 1))
        ViewState("pieChartWeek") = pieChartWeek

        colorInx = (weekGroupDT.Rows.Count - 1)
        For toolIndex As Integer = 0 To (weekGroupDT.Rows.Count - 1)

            weekStr = (weekGroupDT.Rows(toolIndex)("WW")).ToString
            foundRows = DtSet.Select("WW='" + weekStr + "'", "Fail_Ratio desc")
            dtFilter = DtSet.Clone
            For x = 0 To (foundRows.Length - 1)
                dr = foundRows(x)
                dtFilter.LoadDataRow(dr.ItemArray, False)
            Next
            dtFilter.CaseSensitive = True

            series = Chart.Series.Add(("WW" + weekStr))
            series.ChartArea = "Default"
            series.Type = SeriesChartType.Column
            series.Color = aryColor(colorInx)
            series.BorderColor = Color.White
            series.BorderWidth = 1
            
            For i As Integer = 0 To (setupDT.Rows.Count - 1)

                failMode = setupDT.Rows(i)("Fail_Mode").ToString.Trim()

                insideRows = DtSet.Select("WW='" + weekStr + "' and Fail_Mode='" + failMode.Replace("'", "''") + "'")

                If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                    failValue = CType(insideRows(0).Item("Fail_Ratio"), Double)
                Else
                    failValue = 0
                End If

                If toolIndex = (weekGroupDT.Rows.Count - 1) Then
                    failObj = New FailObj
                    failObj.Fail_Mode = failMode
                    failObj.Fail_Value = failValue
                    DiffFAry.Add(failObj)
                ElseIf toolIndex = (weekGroupDT.Rows.Count - 2) Then
                    DiffBHash.Add(failMode, failValue)
                End If

                scriptStr = "javascript:openWindowWithPost('YieldPieChart.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}')"
                scriptStr = String.Format(scriptStr, (ddlPart.SelectedValue.Trim()), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All")

                Chart.Series(("WW" + weekStr)).Points.AddXY(failMode, failValue)
                Chart.Series(("WW" + weekStr)).Points(i).ToolTip = "Week" & weekStr & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                Chart.Series(("WW" + weekStr)).Points(i).Href = scriptStr
                
            Next

            colorInx = (colorInx - 1)

        Next

    End Sub

    ' Draw Different Bar Chart
    Private Sub DrawDiffBarChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable)

        Chart.ImageUrl = "temp/Diff_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH)
        Chart.Width = Unit.Pixel(gChartW)

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.Titles.Add("Fail Mode Difference By Week")
        Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(0).Color = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        'Chart.ChartAreas("Default").AxisX.Title = "【 Fail Mode 】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        'Chart.ChartAreas("Default").AxisY.Interval = 20
        'Chart.ChartAreas("Default").AxisY.Minimum = -100
        'Chart.ChartAreas("Default").AxisY.Maximum = 100
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        series = Chart.Series.Add("Diff")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Column
        series.Color = Color.DodgerBlue
        'series.BorderColor = Color.White
        'series.BorderWidth = 1
        series.ShowInLegend = False

        Dim obj11 As LegendItem = New LegendItem()
        obj11.MarkerSize = 10
        obj11.Name = "Increase"
        obj11.Style = LegendImageStyle.Marker
        obj11.MarkerColor = Color.Red
        Chart.Legends(0).CustomItems.Add(obj11)

        Dim obj2 As LegendItem = New LegendItem()
        obj2.MarkerSize = 10
        obj2.Name = "Decrease"
        obj2.Style = LegendImageStyle.Marker
        obj2.MarkerColor = Color.DodgerBlue
        Chart.Legends(0).CustomItems.Add(obj2)

        Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, "WW")
        weekGroupDT.DefaultView.Sort = "WW Desc"
        weekGroupDT = weekGroupDT.DefaultView.ToTable
        Dim failMode As String
        Dim failValue, beforeValue, finalValue As Double
        Dim failObj As FailObj
        Dim desStr As String = "Increase"

        For i As Integer = 0 To (DiffFAry.Count - 1)

            failObj = CType(DiffFAry(i), FailObj)
            failMode = failObj.Fail_Mode
            failValue = failObj.Fail_Value
            If DiffBHash.Contains(failMode) Then
                beforeValue = CType(DiffBHash(failMode), Double)
            Else
                beforeValue = 0
            End If

            If (failValue > 0) And (beforeValue > 0) Then
                'If failValue > beforeValue Then ' %
                '    finalValue = Math.Round(((failValue - beforeValue) / failValue) * 100, 2)
                'Else
                '    finalValue = Math.Round(((failValue - beforeValue) / beforeValue) * 100, 2)
                'End If
                finalValue = Math.Round((beforeValue - failValue), 2)
            Else
                finalValue = 0
            End If
            Chart.Series("Diff").Points.AddXY(failMode, finalValue)

            'If finalValue >= 0 Then
            '    desStr = "Increase"
            '    Chart.Series("Diff").Points(i).Color = Color.Red
            'Else
            '    desStr = "Decrease"
            '    Chart.Series("Diff").Points(i).Color = Color.DodgerBlue
            'End If

            If finalValue <= 0 Then
                desStr = "Increase"
                Chart.Series("Diff").Points(i).Color = Color.Red
            Else
                desStr = "Decrease"
                Chart.Series("Diff").Points(i).Color = Color.DodgerBlue
            End If
            Chart.Series("Diff").Points(i).ToolTip = desStr & " " & (Math.Abs(finalValue)).ToString & "%"

        Next

    End Sub

    ' ChipSet Chart
    Private Sub DrawChipSetPlantBarChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable)

        Dim WeekStr As String = DtSet.Rows(0)("WW")
        Chart.ImageUrl = "temp/FailPlant_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(gChartH)
        Chart.Width = Unit.Pixel(gChartW)

        Chart.Titles.Add("Fail Mode By Plant (week : " + WeekStr + ")")
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
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)
        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        Dim dtFilter As DataTable
        Dim dr As DataRow
        Dim foundRows() As DataRow
        Dim insideRows() As DataRow
        Dim failMode As String
        Dim failValue As Double
        Dim colorInx As Integer = 0
        Dim scriptStr As String = ""

        colorInx = (plantAry.Length - 1)
        For toolIndex As Integer = 0 To (plantAry.Length - 1)

            foundRows = DtSet.Select("Plant='" + plantAry(toolIndex) + "'", "Fail_Ratio desc")
            If foundRows.Length > 0 Then

                series = Chart.Series.Add(("Plant" + plantAry(toolIndex)))
                series.ChartArea = "Default"
                series.Type = SeriesChartType.Column
                series.Color = aryColor(colorInx)
                series.BorderColor = Color.White
                series.BorderWidth = 1

                For i As Integer = 0 To (setupDT.Rows.Count - 1)

                    failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "''")
                    insideRows = DtSet.Select("Plant='" + plantAry(toolIndex) + "' and Fail_Mode='" + failMode + "'")

                    If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                        failValue = CType(insideRows(0).Item("Fail_Ratio"), Double)
                    Else
                        failValue = 0
                    End If

                    scriptStr = "javascript:openWindowWithPost('YieldPieChart.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}')"
                    scriptStr = String.Format(scriptStr, (ddlPart.SelectedValue.Trim()), failMode, WeekStr, WeekStr, (ddlProduct.SelectedValue), plantAry(toolIndex))

                    Chart.Series(("Plant" + plantAry(toolIndex))).Points.AddXY(failMode, failValue)
                    Chart.Series(("Plant" + plantAry(toolIndex))).Points(i).ToolTip = "Plant" & plantAry(toolIndex) & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                    Chart.Series(("Plant" + plantAry(toolIndex))).Points(i).Href = scriptStr

                Next
                colorInx = (colorInx - 1)

            End If

        Next

    End Sub

    ' Show RowData 'MF_Stage, MF_Area, DefectCode_ID
    Private Sub showRowData(ByRef sourceDT As DataTable, ByRef chipSetRawDT As DataTable)

        Dim plantAryList As ArrayList
        Dim allWeekDT, allPlantDT As DataTable
        Dim newDT As DataTable = sourceDT.Clone
        Dim nWeek As String = sourceDT.Rows(0)("WW")
        Dim dr As DataRow

        ' Step1. 取得最新週的所有 Fail_Mode Item
        Dim foundRows As DataRow() = sourceDT.Select("WW='" + nWeek + "'", "Fail_Ratio desc")
        For x = 0 To (foundRows.Length - 1)
            dr = foundRows(x)
            newDT.LoadDataRow(dr.ItemArray, False)
        Next
        newDT.CaseSensitive = True

        ' Step2. 取得所有週數
        allWeekDT = UtilObj.fun_DataTable_SelectDistinct(sourceDT, "WW")

        ' Step3. 取得所有廠別 (Chip Set 才需要)
        If chipSetRawDT.Rows.Count > 0 Then
            allPlantDT = UtilObj.fun_DataTable_SelectDistinct(chipSetRawDT, "Plant")
        Else
            allPlantDT = New DataTable()
        End If

        ' Step4. Create DataTable
        Dim workTable As DataTable = New DataTable()
        Dim workRow As DataRow
        Dim findRow As DataRow()
        workTable.Columns.Add("Fail Mode", Type.GetType("System.String"))
        workTable.Columns.Add("Stage", Type.GetType("System.String"))
        workTable.Columns.Add("Area", Type.GetType("System.String")) ' DefectCode_ID
        workTable.Columns.Add("ID", Type.GetType("System.String"))
        For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            workTable.Columns.Add("W" + (allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
        Next
        workTable.Columns.Add("Delta", Type.GetType("System.String"))
        ' 加入廠別欄位
        plantAryList = New ArrayList
        For x As Integer = 0 To (plantAry.Length - 1)
            For y As Integer = 0 To (allPlantDT.Rows.Count - 1)
                If (plantAry(x) = allPlantDT.Rows(y)("Plant")) And (plantAry(x).ToUpper <> "ALL") Then
                    workTable.Columns.Add(plantAry(x), Type.GetType("System.String"))
                    plantAryList.Add(plantAry(x))
                    Exit For
                End If
            Next
        Next

        
        ' Step4. 組合資料 WW, Fail_Mode, Fail_Ratio
        Dim fvalue, bvalue, rvalue As Double
        Dim newFailModeStr As String = ""
        Dim failModeStr As String = ""
        Dim colIndex = 0
        For i As Integer = 0 To (newDT.Rows.Count - 1)

            newFailModeStr = newDT.Rows(i)("NewFailMode").Replace("'", "''")
            failModeStr = newDT.Rows(i)("Fail_Mode").Replace("'", "''")
            rvalue = 0
            fvalue = 0
            bvalue = 0

            workRow = workTable.NewRow
            workRow(0) = newDT.Rows(i)("Fail_Mode")
            workRow(1) = newDT.Rows(i)("MF_Stage")
            workRow(2) = newDT.Rows(i)("MF_Area")
            workRow(3) = newDT.Rows(i)("DefectCode_ID")
            colIndex = 4

            ' === 加週數 ===
            For j As Integer = 0 To (allWeekDT.Rows.Count - 1)
                findRow = sourceDT.Select("NewFailMode='" + newFailModeStr + "' and WW='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                Try
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString() + "%"

                ' 最後 2 週的資料相減
                If j = (allWeekDT.Rows.Count - 1) Then
                    fvalue = rvalue
                End If

                If j = (allWeekDT.Rows.Count - 2) Then
                    bvalue = rvalue
                End If
                colIndex += 1
            Next
            rvalue = Math.Round((bvalue - fvalue), 2)
            workRow(colIndex) = (rvalue).ToString() + "%"
            colIndex += 1

            ' === 加廠別 [ 如果是 Chip Set 有資料才有 ] ===
            For j As Integer = 0 To (plantAryList.Count - 1)

                findRow = chipSetRawDT.Select("Fail_Mode='" + failModeStr + "' and plant='" + (plantAryList(j).ToString()) + "'")
                Try
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString() + "%"
                colIndex += 1
            Next
            workTable.Rows.Add(workRow)

        Next

        gv_rowdata.DataSource = workTable
        gv_rowdata.DataBind()
        UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)

    End Sub

    Public Sub ShowMessage(ByVal mesStr As String)

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("alert('" + mesStr + "');")
        sb.Append("</script>")
        Dim myCSManager As ClientScriptManager = Page.ClientScript
        myCSManager.RegisterStartupScript(Me.GetType(), "SetStatusScript", sb.ToString())

    End Sub

    ' Export Excel
    Protected Sub but_Excel_Click(sender As Object, e As System.EventArgs) Handles but_Excel.Click


    End Sub

    Protected Sub gv_rowdata_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_rowdata.RowDataBound


        If e.Row.RowType = DataControlRowType.Header Then
            Dim dayAry(e.Row.Cells.Count - 1) As String
            For i As Integer = 0 To (e.Row.Cells.Count - 1)
                dayAry(i) = e.Row.Cells(i).Text
            Next
        End If

        If e.Row.RowType = DataControlRowType.DataRow Then

            Dim scriptStr As String = ""
            Dim failMode As String = e.Row.Cells(0).Text.Trim
            Dim pieChartWeek As String = ViewState("pieChartWeek").ToString ' 所有選擇的 Week
            Dim startPlant As Boolean = False
            Dim headerStr As String = ""
            Dim newWeekInt As Integer = 0

            For i As Integer = 4 To (e.Row.Cells.Count - 1)

                headerStr = gv_rowdata.HeaderRow.Cells(i).Text

                If (headerStr.ToUpper) = "DELTA" Then
                    startPlant = True
                    newWeekInt = (i - 1)
                Else

                    If startPlant Then
                        ' 依廠別呈現資料
                        Dim newWeekStr As String = gv_rowdata.HeaderRow.Cells(newWeekInt).Text ' 要取得最新的 Week
                        scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""YieldPieChart.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"")'>"
                        scriptStr = String.Format(scriptStr, (ddlPart.SelectedValue.Trim()), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr)
                    Else
                        ' 依週數日期呈現資料
                        scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""YieldPieChart.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"")'>"
                        scriptStr = String.Format(scriptStr, (ddlPart.SelectedValue.Trim()), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All")
                    End If
                    e.Row.Cells(i).Text = scriptStr + (e.Row.Cells(i).Text) + "</a>"

                End If
                
            Next

            e.Row.Height = Unit.Pixel(30)
            e.Row.Cells(0).Font.Size = FontUnit.XXSmall
            e.Row.Cells(0).Font.Bold = True
            e.Row.Cells(0).ForeColor = Drawing.Color.Red

        End If

    End Sub

End Class
