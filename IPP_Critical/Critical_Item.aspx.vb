Imports System.Data.SqlClient
Imports System.Data
Imports Dundas.Charting.WebControl
Imports System.Drawing
Imports Microsoft.VisualBasic

Partial Class IPP_Critical
    Inherits System.Web.UI.Page

    ' 料號 &　參數由 Daily_CriticalItem_OOC_Monitor_Main_BU_Rename 控制
    ' 資料 FC 在 view_IPP_CriticalItem_Monitor, view_IPP_Process_CriticalItem_Monitor
    ' 資料 WB 在 view_IPP_CriticalItem_Monitor_wb, view_IPP_Process_CriticalItem_Monitor_wb

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait......';" & _
                                                 "javascript:document.getElementById('but_Execute').disabled=true;" & _
                                                  Me.Page.GetPostBackEventReference(but_Execute))
        If Not Me.IsPostBack Then
            pageInit()
            If Request("FUN") <> Nothing Then
                cb_DailyEvent.Checked = True
                Dim funStr As String = Request("FUN")
                Dim category As String = Request("CTYPE")
                Dim Main_id As String = Request("PRM")
                Dim partid As String = Request("PART")
                Dim STime As String = Request("S")
                If funStr = "MAIL" Then
                    mailInit(category, Main_id, partid, STime)
                End If
            End If
        End If

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim distinctDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try

            conn.Open()
            sqlStr = "select Category from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where customer_id = 'INTEL' Group by Category Order by Category"
            sqlStr = "select Category from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where 1=1 Group by Category Order by Category"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' -- Data Source --
            UtilObj.FillController(myDT, ddlDataSource, 0)

            sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where customer_id = 'INTEL' "
            sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where 1=1"
            sqlStr += "and Category='" + (ddlDataSource.SelectedItem.Value) + "'"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            ' -- Part ID --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"part_id"})
            UtilObj.FillController(distinctDT, ddlPartNo, 1, "part_id")
            ListBoxSort(ddlPartNo, True)


            ' -- Yield Impact --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"yield_impact_item"})
            UtilObj.FillController(distinctDT, ddlYImpact, 1, "yield_impact_item")

            ' -- Key Module --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"key_module"})
            UtilObj.FillController(distinctDT, ddlKModule, 1, "key_module")

            ' -- Critical Item --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"Critical_item"})
            UtilObj.FillController(distinctDT, ddlCItem, 1, "Critical_item")

            ' -- By Date --
            Dim sTime As String = Date.Now.AddDays(-14).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(0).ToString("yyyy-MM-dd")
            txtDateFrom.Value = sTime
            txtDateTo.Value = eTime
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub
    Private Sub ListBoxSort(ByVal lbx As DropDownList, ByVal ASC As Boolean)
        '利用sortedlist 類為listbox排序 
        Dim slist As SortedList
        If ASC = True Then
            slist = New SortedList()
        Else
            slist = New SortedList(New DecComparer())
        End If


        For i As Integer = 0 To lbx.Items.Count - 1
            '將listbox內容逐項複製到sortedlist物件中
            slist.Add(lbx.Items(i).Text, lbx.Items(i).Value)
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

    Private Sub mailInit(ByVal category As String, ByVal main_id As String, ByVal part_id As String, ByVal STime As String)

        ddlYImpact.SelectedIndex = 0
        ddlKModule.SelectedIndex = 0

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim Critical_Item As String = ""
        Dim myDT As DataTable
        Dim distinctDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try

            conn.Open()

            'sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where customer_id = 'INTEL' "
            sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where 1 = 1 "
            sqlStr += "and Category='" + category + "'"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' -- Part ID --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"part_id"})

            UtilObj.FillController(distinctDT, ddlPartNo, 1, "part_id")

            ' -- Yield Impact --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"yield_impact_item"})
            UtilObj.FillController(distinctDT, ddlYImpact, 1, "yield_impact_item")

            ' -- Key Module --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"key_module"})
            UtilObj.FillController(distinctDT, ddlKModule, 1, "key_module")

            ' -- Critical Item --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"Critical_item"})
            UtilObj.FillController(distinctDT, ddlCItem, 1, "Critical_item")

            'sqlStr = "select Critical_item from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where customer_id = 'INTEL' "
            sqlStr = "select Critical_item from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where 1 = 1 "
            sqlStr += "and Category='" + category + "' and main_id='" + main_id + "' group by Critical_item"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()
            Critical_Item = myDT.Rows(0)("Critical_item").ToString()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        If part_id = "0" Then
            ddlPartNo.SelectedIndex = 0
        Else
            ddlPartNo.SelectedValue = part_id
        End If

        ddlDataSource.SelectedValue = category
        ddlCItem.SelectedValue = Critical_Item
        cb_DailyEvent.Checked = True
        ddlYImpact.Enabled = False
        ddlKModule.Enabled = False

        ' 設定 Query Date Range
        Dim tmpSTime As Date = DateTime.ParseExact(STime, "yyyy-MM-dd", Nothing)
        txtDateFrom.Value = tmpSTime.AddDays(-14).ToString("yyyy-MM-dd")
        txtDateTo.Value = STime
        cb_DailyEvent.Text = "DailyEvent (Event Day : " + STime + ")"
        exeQueryMeasureTimeResult(True)

    End Sub

    ' Category 變動 CPU, CS, WB
    Protected Sub ddlDataSource_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlDataSource.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim distinctDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try

            conn.Open()
            'sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where customer_id = 'INTEL' "
            sqlStr = "select * from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename where 1=1 "
            sqlStr += "and Category='" + (ddlDataSource.SelectedItem.Value) + "' order by part_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            ' -- Part ID --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"part_id"})
            UtilObj.FillController(distinctDT, ddlPartNo, 1, "part_id")

            ' -- Yield Impact --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"yield_impact_item"})
            UtilObj.FillController(distinctDT, ddlYImpact, 1, "yield_impact_item")

            ' -- Key Module --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"key_module"})
            UtilObj.FillController(distinctDT, ddlKModule, 1, "key_module")

            ' -- Critical Item --
            distinctDT = myDT.DefaultView.ToTable(True, New String() {"Critical_item"})
            UtilObj.FillController(distinctDT, ddlCItem, 1, "Critical_item")

            ' -- By Date --
            Dim sTime As String = Date.Now.AddDays(-14).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(0).ToString("yyyy-MM-dd")
            txtDateFrom.Value = sTime
            txtDateTo.Value = eTime
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Chart Ganerate
    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click

        exeQueryMeasureTimeResult(False)
        If cb_DailyEvent.Checked Then
            cb_DailyEvent.Text = "DailyEvent ( Event Day : " + txtDateTo.Value.Trim() + ")"
        Else
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"
        End If

    End Sub

    Private Sub exeQueryMeasureTimeResult(ByVal showProcess As Boolean)

        Dim sqlColumn As String = ""
        Dim sqlTable As String = ""
        Dim sqlStr As String = ""
        Dim partSql As String = ""
        Dim yImpactSql As String = ""
        Dim kModuleSql As String = ""
        Dim cItemSql As String = ""
        Dim edaSql As String = ""
        Dim expression As String = ""
        Dim SortOrder As String = ""

        Dim myDt As New DataTable
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim wObj As WecoTrendObj
        Dim chartObj As Dundas.Charting.WebControl.Chart
        Me.tr_chartPanel.Visible = False
        lab_wait.Text = ""
        tr_chartPanel.Visible = False

        Try
            conn.Open()
            ' 找 EDA Item 
            If (cb_DailyEvent.Checked) Then
                sqlStr = getDailyEventSQL()
            Else
                If ddlDataSource.SelectedValue = "PPS" Then
                    sqlStr = "select distinct Part_id, Yield_Impact_Item, EDA_Item, MAIN_ID, ID as SUB_ID "
                    sqlStr += "from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename "
                    sqlStr += "where 1=1 and Category='" + (ddlDataSource.SelectedItem.Value) + "' "
                    sqlStr += "order by Part_id, Yield_Impact_Item, EDA_Item "
                Else
                    sqlStr = "select distinct Part, Yield_Impact_Item, EDA_Item, MAIN_ID, ID as SUB_ID "
                    sqlStr += "from dbo.Daily_CriticalItem_OOC_Monitor_Main_BU_Rename "
                    sqlStr += "where 1=1 and Category='" + (ddlDataSource.SelectedItem.Value) + "' "
                    sqlStr += "order by Part, Yield_Impact_Item, EDA_Item "
                End If


            End If
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            Dim yImpactDT As New DataTable
            myAdpt.Fill(yImpactDT)

            If ddlDataSource.SelectedItem.Value = "CPU" Then
                sqlColumn = "Lot, part, Part_Id, Yield_Impact_Item, Key_Module, Critical_item, EDA_Item, Parametric_Measurement, layer, Plant, trtm, meanval, maxval, minval, std, samplesize, oos, ooc, cpk, cp, usl, csl, lsl, xucl, xlcl, SUCL, SLCL, FUCL, FCCL, FLCL, FSTD, CIR, RUCL, RLCL, itemC,WECO_Rule1,WECO_Rule2, WECO_Rule3, WECO_Rule4, WECO_Rule5, WECO_Rule6, WECO_Rule7, WECO_Rule8, WECO_Rule9, SLI "
                sqlTable = "from view_IPP_CriticalItem_Monitor_BU_Rename "
            ElseIf ddlDataSource.SelectedItem.Value = "CS" Then
                sqlColumn = "Lot, part, Part_Id, Yield_Impact_Item, Key_Module, Critical_item, EDA_Item, Parametric_Measurement, layer, Plant, trtm, meanval, maxval, minval, std, samplesize, oos, ooc, cpk, cp, usl, csl, lsl, xucl, xlcl, SUCL, SLCL, FUCL, FCCL, FLCL, FSTD, CIR, RUCL, RLCL, itemC,WECO_Rule1,WECO_Rule2, WECO_Rule3, WECO_Rule4, WECO_Rule5, WECO_Rule6, WECO_Rule7, WECO_Rule8, WECO_Rule9, SLI "
                sqlTable = "from view_IPP_CriticalItem_Monitor_BU_Rename_CS "
            Else
                sqlColumn = "Lot, part, Part_Id, Yield_Impact_Item, Key_Module, Critical_item, EDA_Item, Parametric_Measurement, layer, Plant, trtm, meanval, maxval, minval, std, samplesize, oos, ooc, cpk, cp, usl, csl, lsl, xucl, xlcl, SUCL, SLCL, FUCL, FCCL, FLCL, FSTD, CIR, RUCL, RLCL, itemC,WECO_Rule1,WECO_Rule2, WECO_Rule3, WECO_Rule4, WECO_Rule5, WECO_Rule6, WECO_Rule7, WECO_Rule8, WECO_Rule9 "
                sqlTable = "from view_IPP_CriticalItem_Monitor_WB "
            End If

            ' -- Main SQL --
            sqlStr = "select " + sqlColumn
            sqlStr += sqlTable
            sqlStr += "where 1=1 "

            ' -- Date Range 
            sqlStr += "and trtm >= '" + (txtDateFrom.Value.Trim) + " 00:00:00' "
            sqlStr += "and trtm <= '" + (txtDateTo.Value.Trim) + " 23:59:59' "
            'sqlStr += "and  Data_Source = 'IPP' "

            ' -- Part NO--
            If ddlPartNo.SelectedIndex > 0 Then
                sqlStr += "and  Part_Id = '" + (ddlPartNo.SelectedValue).Trim() + "' "
            End If

            ' -- Yield Impact --
            If ddlYImpact.SelectedIndex > 0 Then
                sqlStr += "and Yield_Impact_Item = '" + (ddlYImpact.SelectedValue).Trim() + "' "
            End If

            ' -- Key Module --
            If ddlKModule.SelectedIndex > 0 Then
                sqlStr += "and  Key_Module = '" + (ddlKModule.SelectedValue).Trim() + "' "
            End If

            ' -- Criteria Item --
            If ddlCItem.SelectedIndex > 0 Then
                sqlStr += "and  Critical_Item = '" + (ddlCItem.SelectedValue).Trim() + "' "
            End If

            sqlStr += "group by " + sqlColumn
            sqlStr += "order by Part_Id, Yield_Impact_Item, EDA_Item, trtm,lot desc "

            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.SelectCommand.CommandTimeout = 300
            myAdpt.Fill(myDt)
            conn.Close()

            lab_wait.Text = ""
            If cb_DailyEvent.Checked Then
                If yImpactDT.Rows.Count <= 0 Then
                    lab_wait.Text = "Event Day 無異常 !"
                End If
            Else
                If myDt.Rows.Count <= 0 Then
                    lab_wait.Text = "無資料 !"
                End If
            End If

            If myDt.Rows.Count > 0 Then
                tr_chartPanel.Visible = True
            Else
                tr_chartPanel.Visible = False
            End If

            For i As Integer = 0 To (yImpactDT.Rows.Count - 1)

                If ddlDataSource.SelectedValue = "PPS" Then
                    expression = "part='" + (yImpactDT.Rows(i).Item("Part_id").ToString) + "' and yield_impact_item='" + (yImpactDT.Rows(i).Item("Yield_Impact_Item").ToString) + "' and EDA_Item='" + (yImpactDT.Rows(i).Item("EDA_Item").ToString) + "'"
                Else
                    expression = "part='" + (yImpactDT.Rows(i).Item("Part").ToString) + "' and yield_impact_item='" + (yImpactDT.Rows(i).Item("Yield_Impact_Item").ToString) + "' and EDA_Item='" + (yImpactDT.Rows(i).Item("EDA_Item").ToString) + "'"
                End If
                SortOrder = "part, yield_impact_item, EDA_Item, trtm, lot"
                Dim foundRows() As DataRow
                foundRows = myDt.Select(expression, SortOrder)

                If (foundRows.Length > 0) Then

                    Dim dtFilter As DataTable
                    Dim dr As DataRow
                    dtFilter = myDt.Clone
                    For x = 0 To (foundRows.Length - 1)
                        dr = foundRows(x)
                        dtFilter.LoadDataRow(dr.ItemArray, False)
                    Next
                    dtFilter.CaseSensitive = True

                    wObj = New WecoTrendObj()
                    wObj.FunctionType = "Critical_Item"
                    wObj.Product_Category = (ddlDataSource.SelectedItem.Value)
                    wObj.MAIN_ID = (yImpactDT.Rows(i)("MAIN_ID")).ToString()
                    wObj.SUB_ID = (yImpactDT.Rows(i)("SUB_ID")).ToString()
                    wObj.KPP_Part = (dtFilter.Rows(0)("Part_Id").Replace("'", "''"))
                    wObj.KPP_YieldImpact = (dtFilter.Rows(0)("Yield_Impact_Item").Replace("'", "''"))
                    wObj.KPP_KeyModule = (dtFilter.Rows(0)("Key_Module").Replace("'", "''"))
                    wObj.KPP_CriticalItem = (dtFilter.Rows(0)("Critical_item").Replace("'", "''"))
                    wObj.KPP_IPP = (dtFilter.Rows(0)("EDA_Item").Replace("'", "''"))
                    wObj.chartH = 600
                    wObj.chartW = 1090
                    wObj.valueType = "meanval"
                    wObj.txtDateFrom = txtDateFrom.Value.Trim()
                    wObj.txtDateTo = txtDateTo.Value.Trim()
                    wObj.notDetail = True
                    wObj.showProcess = showProcess
                    If cb_DailyEvent.Checked Then
                        wObj.isHighlight = True
                        wObj.HL_Day = (txtDateTo.Value.Trim)
                    End If

                    Dim bLocation As Boolean = cb_showlabel.Checked
                    chartObj = New Dundas.Charting.WebControl.Chart()
                    If (wObj.Call_DrawChart(dtFilter, chartObj, bLocation)) Then
                        Panel1.Controls.Add(New LiteralControl("<tr><td colspan=3 class='Table_Two_Title' valign='middle' align='left' style='width:500px;font-size:middle;font-weight:bold'>Measure Time Data : " & (dtFilter.Rows(0)("Part").ToString).Replace("'", "''") & ":" & (dtFilter.Rows(0)("EDA_Item").ToString) & ":" & (dtFilter.Rows(0)("Yield_Impact_Item").ToString) & ":" & (dtFilter.Rows(0)("Key_Module").ToString) & ":" & (dtFilter.Rows(0)("Critical_item").ToString) & "</td><td style='width:500px'></td></tr>"))
                        Panel1.Controls.Add(New LiteralControl("<tr><td colspan=4 valign=middle align='left' style='font-size:x-large;font-weight: bold'>"))
                        Panel1.Controls.Add(chartObj)
                        Panel1.Controls.Add(New LiteralControl("</td></tr>"))
                    End If

                End If

            Next

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Function getDailyEventSQL() As String

        Dim yImpact As String = ""
        Dim STime As String = ""
        Dim ETime As String = ""
        Dim criticalItem As String = ""

        If ddlCItem.SelectedIndex = 0 Then
            criticalItem = ""
        Else
            criticalItem = "and a.critical_item='" + (ddlCItem.SelectedValue) + "' "
        End If

        ' 如果是 Daily Event 就尋找傳來一天日期內的 Item 資訊
        STime = (txtDateTo.Value.Trim()) + " 00:00:00"
        ETime = (txtDateTo.Value.Trim()) + " 23:59:59"

        yImpact = "select a.Critical_Item, a.Part, b.Yield_Impact_Item, a.EDA_Item, b.MAIN_ID, b.ID as SUB_ID "
        If ddlDataSource.SelectedItem.Value = "CPU" Then
            yImpact += "from view_IPP_CriticalItem_Monitor_BU_Rename a, Daily_CriticalItem_OOC_Monitor_Main_BU_Rename b "
        ElseIf ddlDataSource.SelectedItem.Value = "CS" Then
            yImpact += "from view_IPP_CriticalItem_Monitor_BU_Rename_CS a, Daily_CriticalItem_OOC_Monitor_Main_BU_Rename b "
        Else
            yImpact += "from view_IPP_CriticalItem_Monitor_WB a, Daily_CriticalItem_OOC_Monitor_Main_BU_Rename b "
        End If
        yImpact += "where a.Critical_Item = b.Critical_Item "
        yImpact += "and a.Part_ID=b.Part_ID "
        yImpact += "and a.EDA_Item = b.EDA_Item "
        yImpact += "and b.Category='" + (ddlDataSource.SelectedItem.Value) + "' "
        yImpact += "and a.trtm >= '{0}' "
        yImpact += "and a.trtm <= '{1}' "
        yImpact += criticalItem
        yImpact += "and (WECO_Rule1=1 or WECO_Rule3=1) "
        yImpact += "group by a.Critical_Item, a.Part, b.Yield_Impact_Item, a.EDA_Item, b.MAIN_ID, b.ID "
        yImpact += "order by a.Critical_Item, a.Part, b.Yield_Impact_Item, a.EDA_Item "
        yImpact = String.Format(yImpact, STime, ETime)

        Return yImpact
    End Function

    ' 按下 DailyEvent
    Protected Sub cb_DailyEvent_CheckedChanged(sender As Object, e As System.EventArgs) Handles cb_DailyEvent.CheckedChanged

        If cb_DailyEvent.Checked Then
            Dim sTime As String = Date.Now.AddDays(-15).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
            txtDateFrom.Value = sTime
            txtDateTo.Value = eTime
            ddlYImpact.SelectedIndex = 0
            ddlKModule.SelectedIndex = 0
            ddlYImpact.Enabled = False
            ddlKModule.Enabled = False
            cb_DailyEvent.Text = "DailyEvent ( Event Day : " + eTime + ")"
        Else
            ddlYImpact.Enabled = True
            ddlKModule.Enabled = True
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"
        End If

    End Sub

End Class
