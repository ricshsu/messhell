Imports System.Data.SqlClient
Imports System.Data
Imports Dundas.Charting.WebControl
Imports System.Drawing
Imports Microsoft.VisualBasic

Partial Class IPP_Critical
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait......';" & _
                                                 "javascript:document.getElementById('but_Execute').disabled=true;" & _
                                                  Me.Page.GetPostBackEventReference(but_Execute))
        If Not Me.IsPostBack Then
            pageInit()
            If Request("FUN") <> Nothing Then
                cb_DailyEvent.Checked = True
                Dim funStr As String = Request("FUN")
                Dim Critical_Item As String = Request("PRM")
                Dim partid As String = Request("PART")
                Dim STime As String = Request("S")
                If funStr = "MAIL" Then
                    mailInit(Critical_Item, partid, STime)
                End If
            End If
        End If

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            ' -- Data Source --
            conn.Open()

            sqlStr = "select data_source from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by data_source"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlDataSource, 1)

            ' -- Part ID --
            sqlStr = "select part_id from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by part_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlPartNo, 1)

            ' -- Yield Impact --
            sqlStr = "select yield_impact_item from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by yield_impact_item"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlYImpact, 1)

            ' -- Key Module --
            sqlStr = "select key_module from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by key_module"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlKModule, 1)

            ' -- Critical Item --
            sqlStr = "select Critical_item from dbo.Daily_CriticalItem_OOC_Monitor_Summary group by Critical_item"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlCItem, 1)

            ' -- By Date --
            Dim sTime As String = Date.Now.AddDays(-14).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(0).ToString("yyyy-MM-dd")
            txtDateFrom.Text = sTime
            txtDateTo.Text = eTime
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"

            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub mailInit(ByVal Critical_Item As String, ByVal part_id As String, ByVal STime As String)

        ' IPP
        ddlDataSource.SelectedIndex = 1
        ddlYImpact.SelectedIndex = 0
        ddlKModule.SelectedIndex = 0

        If part_id = "0" Then
            ddlPartNo.SelectedIndex = 0
        Else
            ddlPartNo.SelectedValue = part_id
        End If

        If Critical_Item = "1" Then
            Critical_Item = "Line width"
        ElseIf Critical_Item = "2" Then
            Critical_Item = "Cu thickness"
        ElseIf Critical_Item = "3" Then
            Critical_Item = "SRO FB Nest"
        ElseIf Critical_Item = "4" Then
            Critical_Item = "SRT (Nest)"
        End If

        ddlCItem.SelectedValue = Critical_Item
        cb_DailyEvent.Checked = True
        ddlYImpact.Enabled = False
        ddlKModule.Enabled = False
        ' 設定 Query Date Range
        Dim tmpSTime As Date = DateTime.ParseExact(STime, "yyyy-MM-dd", Nothing)
        txtDateFrom.Text = tmpSTime.AddDays(-14).ToString("yyyy-MM-dd")
        txtDateTo.Text = STime
        cb_DailyEvent.Text = "DailyEvent ( Event Day : " + STime + ")"
        exeQueryResult()

    End Sub

    ' Chart Ganerate
    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click
        'exeQuery()
        exeQueryResult()
        If cb_DailyEvent.Checked Then
            cb_DailyEvent.Text = "DailyEvent ( Event Day : " + txtDateTo.Text.Trim() + ")"
        Else
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"
        End If
    End Sub

    Private Sub exeQuery()

        Dim ipp As New TrendObj.IPPInfo
        Dim ichart As New TrendObj.IPPData
        Dim mainSql As String = ""
        Dim itemSql As String = ""
        Dim yImpactSql As String = ""
        SQLCondition(mainSql, yImpactSql)
        Dim myDt As New DataTable
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Me.tr_chartPanel.Visible = False
        lab_wait.Text = ""

        Try

            conn.Open()
            myAdpt = New SqlDataAdapter(mainSql, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            ichart.dt_data = myDt

            If myDt.Rows.Count > 0 Then
                Me.tr_chartPanel.Visible = True
            Else
                lab_wait.Visible = True
                lab_wait.Text = "無資料 !"
            End If

            myAdpt = New SqlDataAdapter(yImpactSql, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            ichart.dt_yImpact = myDt
            conn.Close()

            If myDt.Rows.Count <= 0 Then
                Me.tr_chartPanel.Visible = False
            End If

            If cb_DailyEvent.Checked And myDt.Rows.Count <= 0 Then
                lab_wait.Visible = True
                lab_wait.Text = "Event Day 無異常!"
            End If

            Dim trendObj As New TrendObj
            Dim correlation As New DataTable
            trendObj.txtDateFrom = Me.txtDateFrom.Text.Trim()
            trendObj.txtDateTo = Me.txtDateTo.Text.Trim()
            trendObj.dataSource = ddlDataSource.SelectedValue
            trendObj.notDetail = True
            If cb_DailyEvent.Checked Then
                trendObj.isHighlight = True
                trendObj.HL_Day = (txtDateTo.Text.Trim)
            End If
            trendObj.Call_iPP_ByLot(ichart, Panel1)

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub exeQueryResult()

        Dim sqlColumn As String = "Lot, part, Part_Id, Yield_Impact_Item, Key_Module, Critical_item, EDA_Item, Parametric_Measurement, layer, Plant, trtm, meanval, maxval, minval, std, samplesize, oos, ooc, cpk, cp, usl, csl, lsl, xucl, xlcl, SUCL, SLCL, FUCL, FCCL, FLCL, FSTD, CIR, RUCL, RLCL, itemC,WECO_Rule1,WECO_Rule2, WECO_Rule3, WECO_Rule4, WECO_Rule5, WECO_Rule6, WECO_Rule7, WECO_Rule8, WECO_Rule9, SLI, NCW "
        Dim sqlStr As String = ""
        Dim mainSql As String = ""
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
            If cb_DailyEvent.Checked Then
                sqlStr = getDailyEventSQL()
            Else
                sqlStr = "select distinct Part,Yield_Impact_Item,EDA_Item "
                sqlStr += "from dbo.Daily_CriticalItem_OOC_Monitor_Summary "
                sqlStr += "order by Part, Yield_Impact_Item, EDA_Item "
            End If
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            Dim yImpactDT As New DataTable
            myAdpt.Fill(yImpactDT)

            ' -- Main SQL --
            sqlStr = ""
            mainSql = "select " + sqlColumn
            'mainSql += "from view_IPP_CriticalItem_Monitor " 
            mainSql += "from view_IPP_Process_CriticalItem_Monitor "
            mainSql += "where 1=1 "

            ' -- Date Range 
            sqlStr += "and trtm >= '" + (txtDateFrom.Text.Trim) + " 00:00:00' "
            sqlStr += "and trtm <= '" + (txtDateTo.Text.Trim) + " 23:59:59' "

            ' -- Data Source --
            If ddlDataSource.SelectedIndex > 0 Then
                sqlStr += "and  Data_Source = 'IPP' "
            End If

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

            mainSql += sqlStr
            mainSql += "group by " + sqlColumn
            mainSql += "order by Part_Id, Yield_Impact_Item, EDA_Item, trtm"

            myAdpt = New SqlDataAdapter(mainSql, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)

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

            For i As Integer = 0 To (yImpactDT.Rows.Count - 1)

                expression = "part='" + (yImpactDT.Rows(i).Item("Part").ToString) + "' and yield_impact_item='" + (yImpactDT.Rows(i).Item("Yield_Impact_Item").ToString) + "' and EDA_Item='" + (yImpactDT.Rows(i).Item("EDA_Item").ToString) + "'"
                SortOrder = "part, yield_impact_item, EDA_Item, trtm, lot asc"
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
                    wObj.FunctionType = "Critical_KPP"
                    wObj.KPP_Part = (dtFilter.Rows(0)("Part_Id").Replace("'", "''"))
                    wObj.KPP_YieldImpact = (dtFilter.Rows(0)("Yield_Impact_Item").Replace("'", "''"))
                    wObj.KPP_KeyModule = (dtFilter.Rows(0)("Key_Module").Replace("'", "''"))
                    wObj.KPP_CriticalItem = (dtFilter.Rows(0)("Critical_item").Replace("'", "''"))
                    wObj.KPP_IPP = (dtFilter.Rows(0)("EDA_Item").Replace("'", "''"))
                    wObj.chartH = 600
                    wObj.chartW = 1090
                    wObj.valueType = "meanval"
                    wObj.txtDateFrom = txtDateFrom.Text.Trim()
                    wObj.txtDateTo = txtDateTo.Text.Trim()
                    wObj.notDetail = True
                    wObj.specialOldData = True
                    If cb_DailyEvent.Checked Then
                        wObj.isHighlight = True
                        wObj.HL_Day = (txtDateTo.Text.Trim)
                    End If

                    chartObj = New Dundas.Charting.WebControl.Chart()
                    If (wObj.Call_DrawChart(dtFilter, chartObj, False)) Then
                        Panel1.Controls.Add(New LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>" & (dtFilter.Rows(0)("Part").ToString).Replace("'", "''") & ":" & (dtFilter.Rows(0)("EDA_Item").ToString) & ":" & (dtFilter.Rows(0)("Yield_Impact_Item").ToString) & ":" & (dtFilter.Rows(0)("Key_Module").ToString) & ":" & (dtFilter.Rows(0)("Critical_item").ToString) & "</td><td style='width:300px'></td></tr>"))
                        Panel1.Controls.Add(New LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"))
                        Panel1.Controls.Add(chartObj)
                        Panel1.Controls.Add(New LiteralControl("</td></tr>"))
                    End If
                End If

            Next

            tr_chartPanel.Visible = True
            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub SQLCondition(ByRef MainStr As String, ByRef yImpact As String)

        Dim sqlColumn As String = "Lot, part, Part_Id, Yield_Impact_Item, Key_Module, Critical_item, EDA_Item, Parametric_Measurement, layer, Plant, trtm, meanval, maxval, minval, std, samplesize, oos, ooc, cpk, cp, usl, csl, lsl, xucl, xlcl, SUCL, SLCL, FUCL, FCCL, FLCL, FSTD, CIR, RUCL, RLCL, itemC,WECO_Rule1,WECO_Rule2, WECO_Rule3, WECO_Rule4, WECO_Rule5, WECO_Rule6, WECO_Rule7, WECO_Rule8, WECO_Rule9, SLI, NCW "

        ' -- Main SQL --
        MainStr = "select " + sqlColumn
        'MainStr += "from view_IPP_CriticalItem_Monitor " 
        MainStr += "from view_IPP_Process_CriticalItem_Monitor "
        MainStr += "where 1=1 "

        ' -- Yield Impact --
        If cb_DailyEvent.Checked Then
            yImpact = getDailyEventSQL()
        Else
            yImpact = "select distinct Part,Yield_Impact_Item,EDA_Item "
            yImpact += "from dbo.Daily_CriticalItem_OOC_Monitor_Summary "
            yImpact += "order by Part,Yield_Impact_Item,EDA_Item "
        End If


        Dim sqlStr As String = ""
        ' -- Date Range 
        sqlStr += "and trtm >= '" + (txtDateFrom.Text.Trim) + " 00:00:00' "
        sqlStr += "and trtm <= '" + (txtDateTo.Text.Trim) + " 23:59:59' "

        ' -- Data Source --
        If ddlDataSource.SelectedIndex > 0 Then
            sqlStr += "and  Data_Source = '" + (ddlDataSource.SelectedValue).Trim() + "' "
        End If

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

        MainStr += sqlStr
        MainStr += "group by " + sqlColumn
        MainStr += "order by Part_Id, Yield_Impact_Item, EDA_Item, trtm"

    End Sub

    Private Function getDailyEventSQL() As String

        Dim yImpact As String = ""
        Dim STime As String = ""
        Dim ETime As String = ""
        Dim criticalItem As String = ""

        If ddlCItem.SelectedIndex = 0 Then
            criticalItem = ""
        Else
            criticalItem = "and critical_item='" + (ddlCItem.SelectedValue) + "' "
        End If

        ' 如果是 Daily Event 就尋找傳來一天日期內的 Item 資訊
        STime = (txtDateTo.Text.Trim()) + " 00:00:00"
        ETime = (txtDateTo.Text.Trim()) + " 23:59:59"

        yImpact = "select Critical_Item, Part, Yield_Impact_Item, EDA_Item "
        'yImpact += "from dbo.view_IPP_CriticalItem_Monitor where 1=1 "
        yImpact += "from dbo.view_IPP_Process_CriticalItem_Monitor where 1=1 "
        yImpact += "and trtm >= '{0}' "
        yImpact += "and trtm <= '{1}' "
        yImpact += criticalItem
        yImpact += "and (WECO_Rule1=1 or WECO_Rule3=1) "
        yImpact += "group by Critical_Item, Part, Yield_Impact_Item, EDA_Item "
        yImpact += "order by Critical_Item, Part, Yield_Impact_Item, EDA_Item "
        yImpact = String.Format(yImpact, STime, ETime)

        Return yImpact

    End Function

    Protected Sub cb_DailyEvent_CheckedChanged(sender As Object, e As System.EventArgs) Handles cb_DailyEvent.CheckedChanged

        lab_wait.Text = ""
        If cb_DailyEvent.Checked Then
            Dim sTime As String = Date.Now.AddDays(-15).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
            txtDateFrom.Text = sTime
            txtDateTo.Text = eTime
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
