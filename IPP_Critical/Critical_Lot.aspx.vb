Imports System.Data.SqlClient
Imports System.Data
Imports Dundas.Charting.WebControl
Imports System.Drawing
Imports Microsoft.VisualBasic

Partial Class Critical_Lot
    Inherits System.Web.UI.Page

    ' Part由 Customer_Prodction_Mapping_BU_Rename 控制
    ' 參數由 Critical_LOT_Params_BU_Rename 控制

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait......';" & _
                                                 "javascript:document.getElementById('but_Execute').disabled=true;" & _
                                                  Me.Page.GetPostBackEventReference(but_Execute))
        If Not Me.IsPostBack Then
            pageInit()
            If Request("FUN") <> Nothing Then
                cb_DailyEvent.Checked = True
                Dim funStr As String = Request("FUN")
                Dim Category As String = Request("CTYPE")
                Dim MainItem As String = Request("MAIN")
                Dim SublItem As String = Request("SUB")
                Dim partid As String = Request("PART")
                Dim STime As String = Request("S")
                If funStr = "MAIL" Then
                    mailInit(Category, MainItem, SublItem, partid, STime)
                End If
            End If
        End If
        txb_lotinput.Attributes.Add("onkeyup", String.Format("SearchList('{0}', '{1}')", (txb_lotinput.ClientID), (lb_lotSource.ClientID)))

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdpt As New SqlDataAdapter
        Dim sqlStr = ""
        Dim myDt As DataTable = New DataTable
        Dim FType As String = "critical_lot"

        Try

            conn.Open()
            ' -- Data Source --
            sqlStr = "Select param_Group FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "group by param_Group "
            sqlStr += "order by param_Group "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, c, 1)

            ' -- Critical Item --
            sqlStr = "Select param_Name FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "group by param_Name "
            sqlStr += "order by param_Name "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, ddlItem, 1)

            ' --- Part ID ---
            sqlStr = "SELECT Part_ID FROM Customer_Prodction_Mapping_BU_Rename "
            sqlStr += "WHERE 1=1 "
            'If ddl_Category.SelectedIndex = 0 Then
            '    sqlStr += "AND category='CPU' "
            'Else
            '    sqlStr += "AND category='PPS' "
            'End If
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND category='CPU' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND category='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND category='PPS' "
            End If

            sqlStr += "AND customer_id='INTEL' "
            sqlStr += "AND Shipping_Function='1' "
            sqlStr += "group by Part_ID "
            sqlStr += "order by Part_ID "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            conn.Close()
            UtilObj.FillLitsBoxController(myDt, lb_lotSource, 0)

            ' -- SetUp Date --
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

    Private Sub mailInit(ByVal PCategory As String, ByVal main_id As String, ByVal sub_id As String, ByVal partID As String, ByVal STime As String)

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdpt As New SqlDataAdapter
        Dim sqlStr = ""
        Dim paramList = ""
        Dim paramAry As ArrayList = New ArrayList
        Dim myDt As DataTable

        Try

            conn.Open()

            'CATEGORY [CPU, PPS]
            ddl_Category.SelectedValue = PCategory

            ' -- Data Source -- [RA, Bump ......]
            sqlStr = "Select param_Group FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "group by param_Group "
            sqlStr += "order by param_Group "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, c, 1)

            sqlStr = "Select param_Group FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "AND main_id='" + main_id + "' "
            sqlStr += "group by param_Group "
            sqlStr += "order by param_Group "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            c.SelectedValue = myDt.Rows(0)("param_Group").ToString()
            ' -- Data Source --

            ' -- Critical Item [HSS, Surface Energy(Contact Angle/MI) ...] --
            sqlStr = "Select param_Name FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "AND main_id='" + main_id + "' "
            sqlStr += "group by param_Name "
            sqlStr += "order by param_Name "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, ddlItem, 1)

            If sub_id <> "0" Then
                sqlStr = "Select param_Group, param_Name FROM Critical_LOT_Params_BU_Rename "
                sqlStr += "WHERE 1=1 "
                If ddl_Category.SelectedIndex = 0 Then
                    sqlStr += "AND FType='critical_lot' "
                ElseIf ddl_Category.SelectedIndex = 1 Then
                    sqlStr += "AND FType='CS' "
                ElseIf ddl_Category.SelectedIndex = 2 Then
                    sqlStr += "AND FType='PPS' "
                End If
                sqlStr += "AND main_id='" + main_id + "' "
                sqlStr += "and sub_id='" + sub_id + "' "
                sqlStr += "group by param_Group, param_Name "
                sqlStr += "order by param_Group, param_Name "
                myAdpt = New SqlDataAdapter(sqlStr, conn)
                myDt = New DataTable
                myAdpt.Fill(myDt)
                ddlItem.SelectedValue = myDt.Rows(0)("param_Name").ToString()
            End If
            ' -- Critical Item --

            ' --- Part ID ---
            sqlStr = "SELECT Part_ID FROM Customer_Prodction_Mapping_BU_Rename "
            sqlStr += "WHERE 1=1 "
            'If ddl_Category.SelectedIndex = 0 Then
            '    sqlStr += "AND category='CPU' "
            'Else
            '    sqlStr += "AND category='PPS' "
            'End If
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND category='CPU' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND category='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND category='PPS' "
            End If
            sqlStr += "AND customer_id='INTEL' "
            sqlStr += "AND Shipping_Function='1' "
            sqlStr += "group by Part_ID "
            sqlStr += "order by Part_ID "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillLitsBoxController(myDt, lb_lotSource, 0)
            conn.Close()
        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        cb_DailyEvent.Checked = True

        ' -- Part ID --
        If partID = "0" Then
            lb_lotShow.Items.Clear()
            For i As Integer = 0 To (lb_lotSource.Items.Count - 1)
                lb_lotShow.Items.Add(lb_lotSource.Items(i).Value)
            Next
            lb_lotSource.Items.Clear()
        Else
            lb_lotShow.Items.Add(partID)
            lb_lotSource.Items.Remove(partID)
        End If

        ' Query Date Range
        Dim tmpSTime As Date = DateTime.ParseExact(STime, "yyyy-MM-dd", Nothing)
        txtDateFrom.Value = tmpSTime.AddDays(-14).ToString("yyyy-MM-dd")
        txtDateTo.Value = STime
        cb_DailyEvent.Text = "DailyEvent ( Event Day : " + STime + ")"

        ExecuteFunction(paramAry)

    End Sub

    ' --- Category ---
    Protected Sub ddl_Category_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_Category.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdpt As New SqlDataAdapter
        Dim sqlStr = ""
        Dim myDt As DataTable = New DataTable

        Try
            conn.Open()

            ' -- Data Source --
            sqlStr = "Select param_Group FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "group by param_Group "
            sqlStr += "order by param_Group "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, c, 1)

            ' -- Critical Item --
            sqlStr = "Select param_Name FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If
            sqlStr += "group by param_Name "
            sqlStr += "order by param_Name "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, ddlItem, 1)

            ' --- Part ID ---
            sqlStr = "SELECT Part_ID FROM Customer_Prodction_Mapping_BU_Rename "
            sqlStr += "WHERE 1=1 "
            'If ddl_Category.SelectedIndex = 0 Then
            '    sqlStr += "AND category='CPU' "
            'Else
            '    sqlStr += "AND category='PPS' "
            'End If
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND category='CPU' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND category='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND category='PPS' "
            End If
            sqlStr += "AND customer_id='INTEL' "
            sqlStr += "AND Shipping_Function='1' "
            sqlStr += "group by Part_ID "
            sqlStr += "order by Part_ID "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            conn.Close()
            lb_lotSource.Items.Clear()
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDt, lb_lotSource, 0)

            ' -- SetUp Date --
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

    ' --- DataSource Change ---
    Protected Sub ddlDataSource_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles c.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdpt As New SqlDataAdapter
        Dim sqlStr = ""
        Dim itemStr As String = ""
        Dim myDt As DataTable

        Try
            ' -- Critical Item --
            conn.Open()
            sqlStr = "Select param_Name FROM Critical_LOT_Params_BU_Rename "
            sqlStr += "WHERE 1=1 "

            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND FType='critical_lot' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND FType='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND FType='PPS' "
            End If

            If c.SelectedIndex <> 0 Then
                sqlStr += "and param_group='" + (c.SelectedItem.Value.Trim()) + "' "
            End If

            sqlStr += "group by param_Name "
            sqlStr += "order by param_Name "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            UtilObj.FillController(myDt, ddlItem, 1)

            ' --- Part ID ---
            sqlStr = "SELECT Part_ID FROM Customer_Prodction_Mapping_BU_Rename  "
            sqlStr += "WHERE 1=1 "
            'If ddl_Category.SelectedIndex = 0 Then
            '    sqlStr += "AND category='CPU' "
            'Else
            '    sqlStr += "AND category='PPS' "
            'End If
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND category='CPU' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND category='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND category='PPS' "
            End If
            sqlStr += "AND customer_id='INTEL' "
            sqlStr += "AND Shipping_Function='1' "
            sqlStr += "group by Part_ID "
            sqlStr += "order by Part_ID "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            conn.Close()
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDt, lb_lotSource, 0)

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' --- Critical Item ---
    Protected Sub ddlItem_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlItem.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdpt As New SqlDataAdapter
        Dim sqlStr = ""
        Dim myDt As DataTable = New DataTable

        Try

            conn.Open()
            ' --- Part ID ---
            sqlStr = "SELECT Part_ID FROM Customer_Prodction_Mapping_BU_Rename  "
            sqlStr += "WHERE 1=1 "
            'If ddl_Category.SelectedIndex = 0 Then
            '    sqlStr += "AND category='CPU' "
            'Else
            '    sqlStr += "AND category='PPS' "
            'End If
            If ddl_Category.SelectedIndex = 0 Then
                sqlStr += "AND category='CPU' "
            ElseIf ddl_Category.SelectedIndex = 1 Then
                sqlStr += "AND category='CS' "
            ElseIf ddl_Category.SelectedIndex = 2 Then
                sqlStr += "AND category='PPS' "
            End If
            sqlStr += "AND customer_id='INTEL' "
            sqlStr += "AND Shipping_Function='1' "
            sqlStr += "group by Part_ID "
            sqlStr += "order by Part_ID "
            myAdpt = New SqlDataAdapter(sqlStr, conn)
            myDt = New DataTable
            myAdpt.Fill(myDt)
            conn.Close()
            lb_lotShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDt, lb_lotSource, 0)

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' -- InQuery --
    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click

        Dim paramList As New ArrayList
        ExecuteFunction(paramList)

    End Sub

    Private Sub ExecuteFunction(ByVal paramAry As ArrayList)

        lab_wait.Text = ""
        If lb_lotShow.Items.Count = 0 Then
            ShowMessage("請選擇 Part ID !")
        Else

            If cb_DailyEvent.Checked Then
                cb_DailyEvent.Text = "DailyEvent (Event Day : " + txtDateTo.Value.Trim() + ")"
            Else
                cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"
            End If

            Dim haveSomeThing As Boolean = False
            Dim partStr As String = ""
            Dim CriticalStr As String = ""
            Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            Dim myAdpt As New SqlDataAdapter
            Dim dt As DataTable
            Dim eventDT As DataTable
            Dim sqlStr = ""
            Dim wObj As WecoTrendObj
            Dim chartObj As Dundas.Charting.WebControl.Chart
            Dim itemAry As ArrayList = New ArrayList
            ChartPanel.Controls.Clear()

            Try
                ' 取出選擇的 Critical Item
                If paramAry.Count > 0 And cb_DailyEvent.Checked Then
                    itemAry = paramAry
                Else
                    If ddlItem.SelectedIndex = 0 Then
                        For x As Integer = 1 To (ddlItem.Items.Count - 1)
                            itemAry.Add(ddlItem.Items(x).Value)
                        Next
                    Else
                        itemAry.Add(ddlItem.SelectedValue)
                    End If
                End If

                conn.Open()
                ' --- 依 PartID 區分 ---
                For i As Integer = 0 To (lb_lotShow.Items.Count - 1)
                    partStr = (lb_lotShow.Items(i).Value).Substring(0, 6)
                    ' --- 依 Critical Item 區分 ---
                    For j As Integer = 0 To (itemAry.Count - 1)
                        ' --- Execute Method S ---
                        CriticalStr = itemAry(j).ToString()
                        Try
                            sqlStr = "select Lot,MeasureCount,PanelNo,part,Parametric_Measurement,MIS_OP,mchno,Plant,trtm,meanval,maxval,minval,std,samplesize,cpk,cp,usl,lsl,xucl,xlcl,SUCL,SLCL,WECO_Rule1,WECO_Rule2,WECO_Rule3,WECO_Rule4,WECO_Rule5,WECO_Rule6,WECO_Rule7,WECO_Rule8,WECO_Rule9, MAIN_ID, SUB_ID "
                            sqlStr += "From Critical_LOT_Data a, dbo.Critical_LOT_Params_BU_Rename b "
                            sqlStr += "where MeasureCount = 1 "
                            sqlStr += "and ftype='" + ddl_Category.SelectedValue + "' "
                            sqlStr += "and part='" + partStr + "' "
                            sqlStr += "and a.Parametric_Measurement = b.Param_Name "
                            sqlStr += "and a.Parametric_Measurement='" + CriticalStr + "' "
                            sqlStr += "and trtm >= '" + txtDateFrom.Value.Trim() + " 00:00:00' "
                            sqlStr += "and trtm <= '" + txtDateTo.Value.Trim() + " 23:59:59' "
                            sqlStr += "Group by Lot,MeasureCount,PanelNo,part,Parametric_Measurement,MIS_OP,mchno,Plant,trtm,meanval,maxval,minval,std,samplesize,cpk,cp,usl,lsl,xucl,xlcl,SUCL,SLCL,WECO_Rule1,WECO_Rule2,WECO_Rule3,WECO_Rule4,WECO_Rule5,WECO_Rule6,WECO_Rule7,WECO_Rule8,WECO_Rule9, MAIN_ID, SUB_ID "
                            sqlStr += "order by trtm, lot asc"
                            myAdpt = New SqlDataAdapter(sqlStr, conn)
                            myAdpt.SelectCommand.CommandTimeout = 3600
                            dt = New DataTable
                            myAdpt.Fill(dt)

                            ' Daily Event Check
                            If cb_DailyEvent.Checked Then
                                myAdpt = New SqlDataAdapter(getDailyEventSQL(partStr, CriticalStr), conn)
                                eventDT = New DataTable
                                myAdpt.Fill(eventDT)
                                If eventDT.Rows.Count <= 0 Then
                                    Exit Try
                                End If
                            End If

                            If dt.Rows.Count > 0 Then

                                ' --- Mean ---
                                wObj = New WecoTrendObj()
                                wObj.FunctionType = "Critical_Lot"
                                wObj.Customer_ID = "INTEL"
                                wObj.Product_Category = (ddl_Category.SelectedItem.Value) ' CPU, PPS
                                wObj.MAIN_ID = dt.Rows(0)("MAIN_ID").ToString()
                                wObj.SUB_ID = dt.Rows(ddl_Category.SelectedIndex)("SUB_ID").ToString()
                                wObj.chartH = 600
                                wObj.chartW = 1090
                                wObj.valueType = "meanval"
                                wObj.txtDateFrom = txtDateFrom.Value.Trim()
                                wObj.txtDateTo = txtDateTo.Value.Trim()
                                wObj.notDetail = True
                                If cb_DailyEvent.Checked Then
                                    wObj.isHighlight = True
                                    wObj.HL_Day = (txtDateTo.Value.Trim)
                                End If

                                chartObj = New Dundas.Charting.WebControl.Chart()
                                If (wObj.Call_DrawChart(dt, chartObj, False)) Then
                                    ChartPanel.Controls.Add(New LiteralControl("<tr><td class='Table_Two_Title' valign='middle' align='center' style='width:500px;font-size:middle;font-weight:bold'>" & partStr.Replace("'", "''") & "  " & CriticalStr & "</td><td style='width:300px'></td></tr>"))
                                    haveSomeThing = True
                                    ChartPanel.Controls.Add(New LiteralControl("<tr><td colspan=2 valign=middle align='center' style='font-size:x-large;font-weight: bold'>"))
                                    ChartPanel.Controls.Add(chartObj)
                                    ChartPanel.Controls.Add(New LiteralControl("</td></tr>"))
                                End If

                            End If

                        Catch ex As Exception
                            'lab_wait.Text = ex.Message
                        End Try
                        ' --- Execute Method E ---
                    Next

                Next
                conn.Close()

                If haveSomeThing = False Then
                    lab_wait.Visible = True
                    lab_wait.Text = "無資料 !"
                End If

                If (haveSomeThing = False) And (cb_DailyEvent.Checked) Then
                    lab_wait.Visible = True
                    lab_wait.Text = "Event Day 無異常!"
                End If

            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        End If

    End Sub

    ' -- Message --
    Public Sub ShowMessage(ByVal mesStr As String)

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("alert('" + mesStr + "');")
        sb.Append("</script>")
        ScriptManager.RegisterClientScriptBlock(Me, GetType(Page), "alert", sb.ToString(), False)

    End Sub

    ' >>
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

    ' <<
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

    Protected Sub cb_DailyEvent_CheckedChanged(sender As Object, e As System.EventArgs) Handles cb_DailyEvent.CheckedChanged

        lab_wait.Text = ""
        If cb_DailyEvent.Checked Then
            Dim sTime As String = Date.Now.AddDays(-15).ToString("yyyy-MM-dd")
            Dim eTime As String = Date.Now.AddDays(-1).ToString("yyyy-MM-dd")
            txtDateFrom.Value = sTime
            txtDateTo.Value = eTime
            cb_DailyEvent.Text = "DailyEvent (Event Day : " + eTime + ")"
        Else
            cb_DailyEvent.Text = "DailyEvent (以區間結束日期為 Event Day)"
        End If

    End Sub

    Private Function getDailyEventSQL(ByVal partStr As String, ByVal CriticalStr As String) As String

        Dim sqlStr As String = ""
        Dim STime As String = ""
        Dim ETime As String = ""

        ' 如果是 Daily Event 就尋找傳來一天日期內的 Item 資訊
        STime = (txtDateTo.Value.Trim()) + " 00:00:00"
        ETime = (txtDateTo.Value.Trim()) + " 23:59:59"

        sqlStr = "select * from Critical_LOT_Data "
        sqlStr += "where MeasureCount = 1 "
        'sqlStr += "and (WECO_Rule1=1 or WECO_Rule3=1 or cpk < 1.06) "
        sqlStr += "and (WECO_Rule1=1 or WECO_Rule3=1) "
        sqlStr += "and part='" + partStr + "' "
        sqlStr += "and Parametric_Measurement='" + CriticalStr + "' "
        sqlStr += "and trtm >= '" + STime + "' "
        sqlStr += "and trtm <= '" + ETime + "' "
        sqlStr += "order by trtm"

        Return sqlStr

    End Function

End Class
