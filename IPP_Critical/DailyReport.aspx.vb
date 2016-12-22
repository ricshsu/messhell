Imports System.Data.SqlClient
Imports System.Data
Imports Dundas.Charting.WebControl.Chart
Imports System.Drawing
Imports Dundas.Charting.WebControl

Partial Class DailyReport
    Inherits System.Web.UI.Page

    Private Structure yieldInfo

        Dim LotCount As String

        ' INTEL Column
        Dim TotalYieldWO_Inline As String
        Dim inLineyield As String
        Dim OSyield As String
        Dim Bumpyield As String
        Dim FVIyield As String
        Dim IPQC As String
        Dim Open4W As String
        Dim X6CAW As String
        Dim LE4W As String
        Dim Short4W As String
        Dim X24W As String
        Dim Land4W As String
        Dim Land As String

        ' Non-INTEL Column
        Dim InLine As String
        Dim CCPin As String
        Dim PinCC As String
        Dim BumpAOI As String
        Dim WOS2 As String
        Dim WOS4 As String
        Dim FEOS As String
        Dim FEFVI As String
        Dim BEOS As String
        Dim BEFVI As String
        Dim Bump As String
        Dim C4 As String
        Dim FETOTAL As String
        Dim BETOTAL As String
        Dim BumpAOI_FluxClean As String
        Dim Totalyield As String

    End Structure

    '*** 如果要加入新產品到此 ***
    ' 1. Customer_Prodction_Mapping_BU_Rename 要加入相關 CustomerID 與 PartID
    ' 2. Yield_CATEGORY_Mapping 要加入要呈現的 Yield Type
    ' 3. yieldInfo Structure 要加入yield
    ' 4. getDateRangeData 與 generateDataSource 要加入判斷式
    Private confTable As String = "Customer_Prodction_Mapping_BU_Rename"

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Me.but_Execute.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait ......';" & _
                                                 "javascript:document.getElementById('but_Execute').disabled=true;" & _
                                                 Me.Page.GetPostBackEventReference(but_Execute))
        If Not Me.IsPostBack Then
            PageInit()
        End If

    End Sub

    ' 程式初始化
    Private Sub PageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim BKMSql As String = "and BKM='N' "
        Dim AMD_FAILMODESql As String = "and AMD_FAILMODE='N' "

        Try
            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False
            txtDateFrom.Text = DateTime.Now.AddDays(-1).ToString("yyyyMMdd")
            txtDateTo.Text = DateTime.Now.ToString("yyyyMMdd")

            conn.Open()

            ' --- Product Type (產品種類) ---
            sqlStr = "select category from " + confTable + " where 1=1 "
            sqlStr += "and yield_function=1 "
            sqlStr += "group by category order by category"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddl_ProductType, 0)

            ' --- Customer ID (廠商) ---
            sqlStr = "select customer_id from " + confTable + " where 1=1 "
            sqlStr += "and yield_function=1 "
            sqlStr += "and category = '" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "group by customer_id order by customer_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddl_CustomerID, 0)

            ddl_YieldMode.Items.Clear()
            If ddl_CustomerID.SelectedValue = "INTEL" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By BKM")
                If ddl_YieldMode.SelectedIndex = 1 Then
                    BKMSql = "and BKM='Y' "
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By Station")
                If ddl_YieldMode.SelectedIndex = 1 = 0 Then
                    AMD_FAILMODESql = "and AMD_FAILMODE='Y' "
                End If
            Else
                If ddl_ProductType.SelectedValue.ToUpper() = "PPS" Then
                    ddl_YieldMode.Items.Add("By Fail Mode")
                Else
                    ddl_YieldMode.Items.Add("By Station")
                End If
            End If

            ' --- 檢查有無 Group Part, 沒有就呈現 Part ID ---
            sqlStr = "select production_id from " + confTable + " where 1=1 "
            sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and yield_function=1 "
            sqlStr += "group by production_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' --- Group ---
            If Not IsDBNull(myDT.Rows(0)(0)) Then
                ' 有就秀 Group
                rbl_BySource.SelectedIndex = 0
                UtilObj.FillController(myDT, ddlPart, 0)
            Else
                ' 轉向 Part ID
                rbl_BySource.SelectedIndex = 1
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and yield_function=1 "
                sqlStr += "group by Part_id, Memo order by Part_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillController(myDT, ddlPart, 1, "Part_id", "Memo")
            End If

            ' --- Yield Item ---
            sqlStr = "select yield_category, SEQ "
            sqlStr += "from dbo.Yield_CATEGORY_Mapping "
            sqlStr += "where customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and yield_type ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and SEQ != 1 "
            sqlStr += BKMSql
            sqlStr += AMD_FAILMODESql
            sqlStr += "group by yield_category, SEQ order by SEQ"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 0)

            ' 加入 LotCount -- 不可變動, DataTable 呈現的資料由此而來
            Dim addIndex As Integer = 0
            Dim gItem(myDT.Rows.Count) As String
            gItem(addIndex) = "LotCount"
            addIndex += 1
            For i As Integer = 0 To (myDT.Rows.Count - 1)
                gItem(addIndex) = myDT.Rows(i)(0)
                addIndex += 1
            Next
            ViewState("CPU_gItem") = gItem
            conn.Close()

            ' --- Add Hours ---
            Dim hourStr As String = ""
            For i As Integer = 0 To 23
                hourStr = (i.ToString).PadLeft(2, "0")
                ddlHourFrom.Items.Add(hourStr)
                ddlHourTo.Items.Add(hourStr)
            Next

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 選擇產品種類
    Protected Sub ddl_ProductType_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_ProductType.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim BKMSql As String = "and BKM='N' "
        Dim AMD_FAILMODESql As String = "and AMD_FAILMODE='N' "

        Try

            lb_weekSource.Items.Clear()
            lb_weekShow.Items.Clear()

            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False

            conn.Open()

            ' --- Customer ID (廠商) ---
            sqlStr = "select customer_id from " + confTable + " where 1=1 "
            sqlStr += "and yield_function=1 "
            sqlStr += "and category = '" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "group by customer_id order by customer_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddl_CustomerID, 0)

            ' --- 檢查有無 Group Part, 沒有就呈現 Part ID ---
            sqlStr = "select production_id from " + confTable + " where 1=1 "
            sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and yield_function=1 "
            sqlStr += "group by production_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ddl_YieldMode.Items.Clear()
            If ddl_CustomerID.SelectedValue = "INTEL" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By BKM")
                If ddl_YieldMode.SelectedIndex = 1 Then
                    BKMSql = "and BKM='Y' "
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By Station")
                If ddl_YieldMode.SelectedIndex = 0 Then
                    AMD_FAILMODESql = "and AMD_FAILMODE='Y' "
                End If
            Else
                If ddl_ProductType.SelectedValue.ToUpper() = "PPS" Then
                    ddl_YieldMode.Items.Add("By Fail Mode")
                Else
                    ddl_YieldMode.Items.Add("By Station")
                End If
            End If

            ' --- Group ---
            If Not IsDBNull(myDT.Rows(0)(0)) Then
                ' 有就秀 Group
                rbl_BySource.SelectedIndex = 0
                UtilObj.FillController(myDT, ddlPart, 0)
            Else
                ' 轉向 Part ID
                rbl_BySource.SelectedIndex = 1
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and yield_function=1 "
                sqlStr += "group by Part_id, Memo order by Part_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillController(myDT, ddlPart, 1, "Part_id", "Memo")
            End If

            ' Yield Item 
            sqlStr = "select yield_category, SEQ "
            sqlStr += "from dbo.Yield_CATEGORY_Mapping "
            sqlStr += "where customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and yield_type ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and SEQ != 1 "
            sqlStr += BKMSql
            sqlStr += AMD_FAILMODESql
            sqlStr += "group by yield_category, SEQ order by SEQ "
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_weekShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 0)

            ' 加入 LotCount -- 不可變動, DataTable 呈現的資料由此而來
            Dim addIndex As Integer = 0
            Dim gItem(myDT.Rows.Count) As String
            gItem(addIndex) = "LotCount"
            addIndex += 1
            For i As Integer = 0 To (myDT.Rows.Count - 1)
                gItem(addIndex) = myDT.Rows(i)(0)
                addIndex += 1
            Next
            ViewState("CPU_gItem") = gItem

            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 選擇廠商
    Protected Sub ddl_CustomerID_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_CustomerID.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim BKMSql As String = "and BKM='N' "
        Dim AMD_FAILMODESql As String = "and AMD_FAILMODE='N' "

        Try

            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False
            tr_dateRange.Visible = False
            rb_DataTimeCustor.SelectedIndex = 0
            conn.Open()

            ddl_YieldMode.Items.Clear()
            If ddl_CustomerID.SelectedValue = "INTEL" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By BKM")
                If ddl_YieldMode.SelectedIndex = 1 Then
                    BKMSql = "and BKM='Y' "
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                ddl_YieldMode.Items.Add("By Fail Mode")
                ddl_YieldMode.Items.Add("By Station")
                If ddl_YieldMode.SelectedIndex = 0 Then
                    AMD_FAILMODESql = "and AMD_FAILMODE='Y' "
                End If
            Else
                If ddl_ProductType.SelectedValue.ToUpper() = "PPS" Then
                    ddl_YieldMode.Items.Add("By Fail Mode")
                Else
                    ddl_YieldMode.Items.Add("By Station")
                End If
            End If

            ' --- 檢查有無 Group Part, 沒有就呈現 Part ID ---
            sqlStr = "select production_id from " + confTable + " where 1=1 "
            sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and yield_function=1 "
            sqlStr += "group by production_id"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' --- Group ---
            If Not IsDBNull(myDT.Rows(0)(0)) Then
                ' 有就秀 Group
                rbl_BySource.SelectedIndex = 0
                UtilObj.FillController(myDT, ddlPart, 0)
            Else
                ' 轉向 Part ID
                rbl_BySource.SelectedIndex = 1
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and yield_function=1 "
                sqlStr += "group by Part_id, Memo order by Part_id"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillController(myDT, ddlPart, 0, "Part_id", "Memo")
            End If

            ' Yield Item 
            sqlStr = "select yield_category, SEQ "
            sqlStr += "from dbo.Yield_CATEGORY_Mapping "
            sqlStr += "where customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and yield_type ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and SEQ != 1 "
            sqlStr += BKMSql
            sqlStr += AMD_FAILMODESql
            sqlStr += "group by yield_category, SEQ order by SEQ"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_weekShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 0)

            ' 加入 LotCount -- 不可變動, DataTable 呈現的資料由此而來
            Dim addIndex As Integer = 0
            Dim gItem(myDT.Rows.Count) As String
            gItem(addIndex) = "LotCount"
            addIndex += 1
            For i As Integer = 0 To (myDT.Rows.Count - 1)
                gItem(addIndex) = myDT.Rows(i)(0)
                addIndex += 1
            Next
            ViewState("CPU_gItem") = gItem

            conn.Close()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 選擇 Yield Mode
    Protected Sub ddl_YieldMode_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddl_YieldMode.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim BKMSql As String = "and BKM='N' "
        Dim AMD_FAILMODESql As String = "and AMD_FAILMODE='N' "

        Try

            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False

            ' --- INTEL ---
            If ddl_CustomerID.SelectedValue = "INTEL" And ddl_YieldMode.SelectedValue.ToUpper() = "BY BKM" Then
                BKMSql = "and BKM='Y' "
            End If

            ' --- AMD ---
            If ddl_CustomerID.SelectedValue = "AMD" And ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                AMD_FAILMODESql = "and AMD_FAILMODE='Y' "
            End If

            conn.Open()

            ' Yield Item 
            sqlStr = "select yield_category, SEQ "
            sqlStr += "from dbo.Yield_CATEGORY_Mapping "
            sqlStr += "where customer_id ='" + ddl_CustomerID.SelectedValue + "' "
            sqlStr += "and yield_type ='" + ddl_ProductType.SelectedValue + "' "
            sqlStr += "and SEQ != 1 "
            sqlStr += BKMSql
            sqlStr += AMD_FAILMODESql
            sqlStr += "group by yield_category, SEQ order by SEQ"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_weekShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 0)

            conn.Close()

            ' 加入 LotCount -- 不可變動, DataTable 呈現的資料由此而來
            Dim addIndex As Integer = 0
            Dim gItem(myDT.Rows.Count) As String
            gItem(addIndex) = "LotCount"
            addIndex += 1
            For i As Integer = 0 To (myDT.Rows.Count - 1)
                gItem(addIndex) = myDT.Rows(i)(0)
                addIndex += 1
            Next
            ViewState("CPU_gItem") = gItem

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 選擇 Group or Part ID
    Protected Sub rbl_BySource_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rbl_BySource.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        tr_gvDisplay.Visible = False
        Dim BKMSql As String = "and BKM='N' "
        Dim AMDFailModeSql As String = "and AMD_FAILMODE='N' "

        Try

            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False

            ' --- INTEL ---
            If ddl_CustomerID.SelectedValue = "INTEL" And ddl_YieldMode.SelectedValue.ToUpper() = "BY BKM" Then
                BKMSql = "and BKM='Y' "
            End If

            ' --- AMD ---
            If ddl_CustomerID.SelectedValue = "AMD" And ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                AMDFailModeSql = "and AMD_FAILMODE='Y' "
            End If

            conn.Open()

            If rbl_BySource.SelectedIndex = 0 Then
                ' Group 
                sqlStr = "select production_id from " + confTable + " where 1=1 "
                sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and yield_function=1 "
                sqlStr += "group by production_id"
            Else
                ' Part
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                sqlStr += "and customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and category ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and yield_function=1 "
                sqlStr += "group by Part_id, Memo order by Part_id"
            End If

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' --- Group ---
            ddlPart.Items.Clear()
            lb_weekSource.Items.Clear()
            lb_weekShow.Items.Clear()
            If IsDBNull(myDT.Rows(0)(0)) Then
                If rbl_BySource.SelectedIndex = 0 Then
                    ShowMessage("沒有 Group 的資訊!")
                Else
                    ShowMessage("沒有 Part 的資訊!")
                End If
            Else

                If rbl_BySource.SelectedIndex = 0 Then
                    UtilObj.FillController(myDT, ddlPart, 0)
                Else
                    If ddl_CustomerID.SelectedValue = "INTEL" Then
                        UtilObj.FillController(myDT, ddlPart, 0, "Part_id", "Memo")
                    Else
                        UtilObj.FillController(myDT, ddlPart, 1, "Part_id", "Memo")
                    End If
                End If

                ' Yield Item 
                sqlStr = "select yield_category, SEQ "
                sqlStr += "from dbo.Yield_CATEGORY_Mapping "
                sqlStr += "where customer_id ='" + ddl_CustomerID.SelectedValue + "' "
                sqlStr += "and yield_type ='" + ddl_ProductType.SelectedValue + "' "
                sqlStr += "and SEQ != 1 "
                sqlStr += BKMSql
                sqlStr += AMDFailModeSql
                sqlStr += "group by yield_category, SEQ order by SEQ"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                lb_weekShow.Items.Clear()
                UtilObj.FillLitsBoxController(myDT, lb_weekSource, 0)

            End If
            conn.Close()

            ' 加入 LotCount -- 不可變動, DataTable 呈現的資料由此而來
            Dim addIndex As Integer = 0
            Dim gItem(myDT.Rows.Count) As String
            gItem(addIndex) = "LotCount"
            addIndex += 1
            For i As Integer = 0 To (myDT.Rows.Count - 1)
                gItem(addIndex) = myDT.Rows(i)(0)
                addIndex += 1
            Next
            ViewState("CPU_gItem") = gItem

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Execute Query
    Protected Sub but_Execute_Click(sender As Object, e As System.EventArgs) Handles but_Execute.Click

        If (lb_weekShow.Items.Count = 0) Then
            ShowMessage("請選擇 Yield Type !")
            Return
        End If

        If rb_DataTimeCustor.SelectedIndex = 1 And listB_timeShow.Items.Count < 1 Then
            ShowMessage("請選擇時間區間 !")
            Return
        End If

        ' Step1. 尋找要的日, 週, 月
        Dim dayAry As ArrayList = New ArrayList
        Dim weekAry As ArrayList = New ArrayList
        Dim monthAry As ArrayList = New ArrayList
        Dim dyear As String = ""
        Dim dWeek As String = ""
        Dim dStr As String = ""
        Dim tmpString As String = ""

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim conditionStr As String = ""
        Dim myAdapter As SqlDataAdapter
        Dim yObj As yieldInfo
        Dim myDT, chartTable As DataTable
        Dim chartAry As ArrayList = New ArrayList
        Dim chartHash As Hashtable = New Hashtable
        Dim hAry As Hashtable = New Hashtable

        Try

            conn.Open()
            If rb_DataTimeCustor.SelectedIndex = 1 Then
                If rbl_lossItem.SelectedIndex = 0 Then
                    ' Day
                    For i As Integer = 0 To (listB_timeShow.Items.Count - 1)
                        dayAry.Add(listB_timeShow.Items(i).Text)
                    Next
                    dayAry.Reverse()
                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                    ' Week 
                    For i As Integer = 0 To (listB_timeShow.Items.Count - 1)
                        weekAry.Add(listB_timeShow.Items(i).Text)
                    Next
                    weekAry.Reverse()
                Else
                    ' Month
                    For i As Integer = 0 To (listB_timeShow.Items.Count - 1)
                        monthAry.Add(listB_timeShow.Items(i).Text)
                    Next
                    monthAry.Reverse()
                End If
            Else
                ' === Day === 取7天
                Dim i As Integer = 6
                While (i > -1)
                    If cb_ShowToday.Checked Then
                        ' 當天要納入
                        tmpString = Date.Now.AddDays(-(i)).ToString("yyyyMMdd")
                    ElseIf cb_customerDay.Checked Then
                        ' 自訂時間區間
                        If i = 0 Then
                            tmpString = ((txtDateFrom.Text) + (ddlHourFrom.SelectedValue)) + "~" + ((txtDateTo.Text) + (ddlHourTo.SelectedValue))
                        Else
                            tmpString = Date.Now.AddDays(-(i)).ToString("yyyyMMdd")
                        End If
                    Else
                        ' 一般,當天不要納入
                        tmpString = Date.Now.AddDays(-(i + 1)).ToString("yyyyMMdd")
                    End If
                    dayAry.Add(tmpString)
                    i -= 1
                End While
                ' === Week === 取5週  
                Dim weeksql As String = "select yearWW from SystemDateMapping where customer='YIP' and DateTime='{0}'"
                i = 4
                While (i > -1)
                    tmpString = Date.Now.AddDays(-(i * 7)).ToString("yyyy-MM-dd")
                    myAdapter = New SqlDataAdapter(String.Format(weeksql, tmpString), conn)
                    myDT = New DataTable
                    myAdapter.Fill(myDT)
                    weekAry.Add(myDT.Rows(0)("yearWW").ToString())
                    i -= 1
                End While
                ' === Month === 取3個月
                i = 2
                While (i > -1)
                    tmpString = Date.Now.AddMonths(-i).ToString("yyyy-MM")
                    monthAry.Add(tmpString)
                    i -= 1
                End While
            End If

            ' Step2. 先組合 Report 所要的 DataTable 
            Dim workTable As DataTable = New DataTable()
            workTable.Columns.Add("Type", Type.GetType("System.String"))
            workTable.Columns.Add("Customer ID", Type.GetType("System.String"))
            workTable.Columns.Add("Part ID", Type.GetType("System.String"))
            For x As Integer = 0 To (monthAry.Count - 1)
                workTable.Columns.Add((monthAry(x)).ToString(), Type.GetType("System.String"))
            Next
            For x As Integer = 0 To (weekAry.Count - 1)
                workTable.Columns.Add((weekAry(x)).ToString(), Type.GetType("System.String"))
            Next
            For x As Integer = 0 To (dayAry.Count - 1)
                workTable.Columns.Add((dayAry(x)).ToString(), Type.GetType("System.String"))
            Next

            ' 2012/12/20 如果有 All, 就呈現所有料號的 "Total yield Bar Chart!"
            If ddlPart.SelectedValue = "All" Then
                Dim gItem(1) As String
                gItem(0) = "LotCount"
                If lb_weekShow.Items(0).Text = "Total Yield" Then
                    gItem(1) = "Total Yield"
                Else
                    gItem(1) = "Total"
                End If
                ViewState("CPU_gItem") = gItem
            End If

            tr_chartDisplay.Visible = False
            tr_gvDisplay.Visible = False

            If (rbl_BySource.SelectedIndex = 0) Then
                ' --- Customer ID ---
                conditionStr += "and production_id='" + (Me.ddlPart.SelectedValue) + "' "
            Else
                ' --- Part ID ---
                conditionStr += "and Part_Id='" + (Me.ddlPart.SelectedValue) + "' "
            End If
            If (rbl_BySource.SelectedIndex = 1 And ddlPart.SelectedValue = "All") Then
                conditionStr = "and customer_id='" + ddl_CustomerID.SelectedValue + "' "
            End If
            sqlStr += "Select production_id, Part_Id, Memo "
            sqlStr += "from " + confTable + " "
            sqlStr += "where 1=1 "
            sqlStr += "and yield_function=1 "
            sqlStr += "and category='" + (ddl_ProductType.SelectedValue) + "' "
            sqlStr += conditionStr
            sqlStr += "group by production_id, Part_Id, Memo"

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            ' Step3. 取得 Part 下面的 RowData
            Dim customerStr, partStr, part_plant As String
            For partIndex As Integer = 0 To (myDT.Rows.Count - 1)

                If IsDBNull(myDT.Rows(partIndex)("production_id")) Then
                    customerStr = "N/A"
                Else
                    customerStr = myDT.Rows(partIndex)("production_id")
                End If
                partStr = myDT.Rows(partIndex)("Part_Id")
                part_plant = myDT.Rows(partIndex)("Memo")

                hAry = New Hashtable
                ' PART By Day
                For x As Integer = 0 To (dayAry.Count - 1)
                    yObj = New yieldInfo
                    If getDateRangeData(yObj, conn, partStr, (dayAry(x).ToString()), 0, False) Then
                        hAry.Add((dayAry(x).ToString()), yObj)
                    End If
                Next
                ' PART By Week
                For x As Integer = 0 To (weekAry.Count - 1)
                    yObj = New yieldInfo
                    If getDateRangeData(yObj, conn, partStr, (weekAry(x).ToString()), 1, False) Then
                        hAry.Add((weekAry(x).ToString()), yObj)
                    End If
                Next
                ' PART By Month
                For x As Integer = 0 To (monthAry.Count - 1)
                    yObj = New yieldInfo
                    If getDateRangeData(yObj, conn, partStr, (monthAry(x).ToString()), 2, False) Then
                        hAry.Add((monthAry(x).ToString()), yObj)
                    End If
                Next

                chartTable = New DataTable()
                chartTable = workTable.Clone()
                generateDataSource(customerStr, part_plant, chartTable, workTable, hAry)

                chartHash = New Hashtable
                chartHash.Add(partStr, chartTable)
                chartAry.Add(chartHash)

                '加入(Summary) ' 選擇 CustomerID, 最後一筆, Part有2筆以上 " * 非必要 "
                If (rbl_BySource.SelectedIndex = 0 Or ddlPart.SelectedValue = "All") And (partIndex = (myDT.Rows.Count - 1)) And (myDT.Rows.Count > 1) Then

                    Dim partStrtemp As String = ""
                    If (ddlPart.SelectedValue = "All") Then
                        partStrtemp = (ddl_CustomerID.SelectedValue)
                    Else
                        partStrtemp = (Me.ddlPart.SelectedValue)
                    End If
                    hAry = New Hashtable
                    ' PART By Day
                    For x As Integer = 0 To (dayAry.Count - 1)
                        yObj = New yieldInfo
                        If getDateRangeData(yObj, conn, partStrtemp, (dayAry(x).ToString()), 0, True) Then
                            hAry.Add((dayAry(x).ToString()), yObj)
                        End If
                    Next
                    ' PART By Week
                    For x As Integer = 0 To (weekAry.Count - 1)
                        yObj = New yieldInfo
                        If getDateRangeData(yObj, conn, partStrtemp, (weekAry(x).ToString()), 1, True) Then
                            hAry.Add((weekAry(x).ToString()), yObj)
                        End If
                    Next
                    ' PART By Month
                    For x As Integer = 0 To (monthAry.Count - 1)
                        yObj = New yieldInfo
                        If getDateRangeData(yObj, conn, partStrtemp, (monthAry(x).ToString()), 2, True) Then
                            hAry.Add((monthAry(x).ToString()), yObj)
                        End If
                    Next
                    chartTable = New DataTable()
                    chartTable = workTable.Clone()
                    generateDataSource(customerStr, ("Summary"), chartTable, workTable, hAry)

                    chartHash = New Hashtable
                    chartHash.Add(("Summary"), chartTable)
                    chartAry.Add(chartHash)

                End If

            Next
            conn.Close()

            ' --- Row Data Display ---
            Try
                gv_rowdata.DataSource = workTable
                gv_rowdata.DataBind()
                UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)
            Catch ex As Exception
            End Try

            ' --- 放 Chart 到 Chart_Panel ---
            Dim yield_item() As String = CType(ViewState("CPU_gItem"), String())
            Dim yieldexist As Boolean = False
            Dim chartYieldStr As String = ""
            For idx As Integer = 0 To (yield_item.Length - 1)
                chartYieldStr = yield_item(idx)
                yieldexist = False
                For idy As Integer = 0 To (lb_weekShow.Items.Count - 1)
                    If chartYieldStr = lb_weekShow.Items(idy).Text Then
                        yieldexist = True
                        Exit For
                    End If
                Next
                If yieldexist Then
                    If ddlPart.SelectedValue = "All" Then
                        For x As Integer = 0 To (chartAry.Count - 1)
                            Dim Chart As Dundas.Charting.WebControl.Chart = New Dundas.Charting.WebControl.Chart()
                            chartHash = CType(chartAry(x), Hashtable)
                            DrawChartBySingle(Chart, chartHash, idx, chartYieldStr)
                            Chart_Panel.Controls.Add(Chart)
                        Next
                    Else
                        Dim Chart As Dundas.Charting.WebControl.Chart = New Dundas.Charting.WebControl.Chart()
                        DrawChart(Chart, chartAry, idx, chartYieldStr)
                        Chart_Panel.Controls.Add(Chart)
                    End If
                End If
            Next

            tr_chartDisplay.Visible = True
            tr_gvDisplay.Visible = True

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 取得每一 Yield Type 資料  [主要資料來源]
    Private Function getDateRangeData(ByRef yObj As yieldInfo, ByRef conn As SqlConnection, ByVal partID As String, ByVal WD As String, ByVal timeType As Integer, ByVal is_summary As Boolean) As Boolean

        Dim status As Boolean = False
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim monthNextStr As String
        Dim ConditionStr As String = ""

        If is_summary Then
            If timeType = 0 Then
                ' --- Day ---
                ConditionStr += "and Production_id='{0}' "
                ConditionStr += "and DataTime >= '{1}' "
                ConditionStr += "and DataTime <= '{2}' "
                If (WD.IndexOf("~") >= 0) And cb_customerDay.Checked Then
                    ConditionStr = String.Format(ConditionStr, partID, ((txtDateFrom.Text.Trim) + " " + (ddlHourFrom.SelectedValue) + ":00:00"), ((txtDateTo.Text.Trim) + " " + (ddlHourTo.SelectedValue) + ":00:00"))
                Else
                    If cb_ShowToday.Checked And (WD = (DateTime.Now.ToString("yyyyMMdd"))) Then
                        Dim mWD As DateTime = DateTime.ParseExact(WD, "yyyyMMdd", Nothing)
                        ConditionStr = String.Format(ConditionStr, partID, ((mWD.AddDays(-1).ToString("yyyyMMdd")) + " 16:00:00"), (WD + " 16:00:00"))
                    Else
                        ConditionStr = String.Format(ConditionStr, partID, (WD + " 00:00:00"), (WD + " 23:59:59"))
                    End If
                End If
            ElseIf timeType = 1 Then
                ' --- Week --
                ConditionStr += "and Production_id='{0}' "
                ConditionStr += "and YearWW={1} "
                ConditionStr = String.Format(ConditionStr, partID, (WD))
            Else
                ' --- Month ---
                Dim monthNext As Date = (DateTime.ParseExact(WD, "yyyy-MM", Nothing))
                monthNextStr = monthNext.AddMonths(1).ToString("yyyy-MM")
                ConditionStr += "and Production_id='{0}' "
                ConditionStr += "and DataTime >='{1}' "
                ConditionStr += "and DataTime < '{2}' "
                ConditionStr = String.Format(ConditionStr, partID, (WD + "-01"), (monthNextStr + "-01"))
            End If
        Else
            If timeType = 0 Then
                ' --- Day ---
                ConditionStr += "and Part_Id='{0}' "
                ConditionStr += "and DataTime >= '{1}' "
                ConditionStr += "and DataTime <= '{2}' "
                If (WD.IndexOf("~") >= 0) And cb_customerDay.Checked Then
                    ConditionStr = String.Format(ConditionStr, partID, ((txtDateFrom.Text.Trim) + " " + (ddlHourFrom.SelectedValue) + ":00:00"), ((txtDateTo.Text.Trim) + " " + (ddlHourTo.SelectedValue) + ":00:00"))
                Else
                    If cb_ShowToday.Checked And (WD = (DateTime.Now.ToString("yyyyMMdd"))) Then
                        Dim mWD As DateTime = DateTime.ParseExact(WD, "yyyyMMdd", Nothing)
                        ConditionStr = String.Format(ConditionStr, partID, ((mWD.AddDays(-1).ToString("yyyyMMdd")) + " 16:00:00"), (WD + " 16:00:00"))
                    Else
                        ConditionStr = String.Format(ConditionStr, partID, (WD + " 00:00:00"), (WD + " 23:59:59"))
                    End If
                End If
            ElseIf timeType = 1 Then
                ' --- Week --
                ConditionStr += "and Part_Id='{0}' "
                ConditionStr += "and YearWW={1} "
                ConditionStr = String.Format(ConditionStr, partID, (WD))
            Else
                ' --- Month ---
                Dim monthNext As Date = (DateTime.ParseExact(WD, "yyyy-MM", Nothing))
                monthNextStr = monthNext.AddMonths(1).ToString("yyyy-MM")
                ConditionStr += "and Part_Id='{0}' "
                ConditionStr += "and DataTime >='{1}' "
                ConditionStr += "and DataTime < '{2}' "
                ConditionStr = String.Format(ConditionStr, partID, (WD + "-01"), (monthNextStr + "-01"))
            End If
        End If

        Try

            sqlStr += "select COUNT(*) AS LotCount, "
            sqlStr += "YIELD_CATEGORY, "
            sqlStr += "ROUND(convert(float,SUM(OUTPUT_QTY))/SUM(Input_QTY), 4) as YIELD_VALUE "

            ' --- INTEL 的產品 ---
            If ddl_CustomerID.SelectedValue = "INTEL" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    sqlStr += "from dbo.Yield_Daily_RawData "
                    sqlStr += "where Customer_Id='INTEL' "
                Else
                    sqlStr += "from dbo.BKM_Yield_Daily_RawData "
                    sqlStr += "where 1=1 "
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    sqlStr += "from dbo.Yield_Daily_RawData "
                    sqlStr += "where Customer_Id='AMD' "
                Else
                    sqlStr += "from dbo.Yield_Daily_RawData_NonIntel "
                    sqlStr += "where 1=1 "
                End If
            Else
                If ddl_ProductType.SelectedValue = "PPS" Then
                    sqlStr += "from dbo.WB_Yield_Daily_RawData "
                Else
                    sqlStr += "from dbo.Yield_Daily_RawData_NonIntel "
                End If
                sqlStr += "where 1=1 "
            End If

            sqlStr += ConditionStr
            sqlStr += "GROUP BY YIELD_CATEGORY"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)

            Dim itemStr As String = ""
            If (myDT.Rows.Count > 0) Then

                For i As Integer = 0 To (myDT.Rows.Count - 1)

                    If i = 0 Then
                        yObj.LotCount = myDT.Rows(i)("LotCount")
                    End If

                    itemStr = myDT.Rows(i)("YIELD_CATEGORY").ToString.Trim().ToLower

                    If itemStr = "inline yield" Then
                        yObj.inLineyield = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "total yield w/o inline" Then
                        yObj.TotalYieldWO_Inline = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "o/s yield" Then
                        yObj.OSyield = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "bump yield" Then
                        yObj.Bumpyield = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "fvi yield" Then
                        yObj.FVIyield = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "ipqc" Then
                        yObj.IPQC = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "4w open" Then
                        yObj.Open4W = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "4w short" Then
                        yObj.Short4W = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "4w x2" Then
                        yObj.X24W = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "4w le" Then
                        yObj.LE4W = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "4w land" Then
                        yObj.Land4W = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "x6 caw" Then
                        yObj.X6CAW = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "in line" Then
                        yObj.InLine = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "feos" Then
                        yObj.FEOS = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "fefvi" Then
                        yObj.FEFVI = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "beos" Then
                        yObj.BEOS = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "befvi" Then
                        yObj.BEFVI = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "bumpaoi" Then
                        yObj.Bump = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf itemStr = "c4" Then
                        yObj.C4 = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "fe total") Then
                        yObj.FETOTAL = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "be total") Then
                        yObj.BETOTAL = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "total" Or itemStr = "total yield") Then
                        yObj.Totalyield = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "ccpin") Then
                        yObj.CCPin = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "pincc") Then
                        yObj.PinCC = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "bumpaoi") Then
                        yObj.BumpAOI = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "2wos") Then
                        yObj.WOS2 = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "4wos") Then
                        yObj.WOS4 = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "land") Then
                        yObj.Land = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    ElseIf (itemStr = "fluxclean後bumpaoi") Then
                        yObj.BumpAOI_FluxClean = myDT.Rows(i)("YIELD_VALUE").ToString.Trim()
                    End If

                Next
                status = True

            End If

        Catch ex As Exception

        End Try

        Return status

    End Function

    ' DataTable 需要的資訊, 畫 Chart 資料在這組合
    Private Sub generateDataSource(ByVal customerID As String, ByVal partID As String, ByRef chartTable As DataTable, ByRef workTable As DataTable, ByRef hAry As Hashtable)

        Dim workRow As DataRow
        Dim yObj As yieldInfo
        Dim yield_item() As String = CType(ViewState("CPU_gItem"), String())

        For x As Integer = 0 To (yield_item.Length - 1) ' [列] : 產品 Yield Type

            workRow = workTable.NewRow

            For y As Integer = 0 To (workTable.Columns.Count - 1) ' [欄] : 日, 週, 月

                If y = 0 Then
                    workRow(y) = yield_item(x)
                ElseIf y = 1 Then
                    workRow(y) = customerID
                ElseIf y = 2 Then
                    workRow(y) = partID
                Else

                    If hAry.Contains(workTable.Columns(y).ColumnName) Then

                        yObj = CType(hAry(workTable.Columns(y).ColumnName), yieldInfo)

                        If yield_item(x).ToLower = "lotcount" Then
                            workRow(y) = yObj.LotCount
                        ElseIf yield_item(x).ToLower = "total yield w/o inline" Then
                            workRow(y) = yObj.TotalYieldWO_Inline
                        ElseIf yield_item(x).ToLower = "inline yield" Then
                            workRow(y) = yObj.inLineyield
                        ElseIf yield_item(x).ToLower = "o/s yield" Then
                            workRow(y) = yObj.OSyield
                        ElseIf yield_item(x).ToLower = "bump yield" Then
                            workRow(y) = yObj.Bumpyield
                        ElseIf yield_item(x).ToLower = "fvi yield" Then
                            workRow(y) = yObj.FVIyield
                        ElseIf yield_item(x).ToLower = "ipqc" Then
                            workRow(y) = yObj.IPQC
                        ElseIf yield_item(x).ToLower = "4w open" Then
                            workRow(y) = yObj.Open4W
                        ElseIf yield_item(x).ToLower = "4w short" Then
                            workRow(y) = yObj.Short4W
                        ElseIf yield_item(x).ToLower = "4w x2" Then
                            workRow(y) = yObj.X24W
                        ElseIf yield_item(x).ToLower = "4w le" Then
                            workRow(y) = yObj.LE4W
                        ElseIf yield_item(x).ToLower = "4w land" Then
                            workRow(y) = yObj.Land4W
                        ElseIf yield_item(x).ToLower = "x6 caw" Then
                            workRow(y) = yObj.X6CAW
                        ElseIf yield_item(x).ToLower = "in line" Then
                            workRow(y) = yObj.InLine
                        ElseIf yield_item(x).ToLower = "feos" Then
                            workRow(y) = yObj.FEOS
                        ElseIf yield_item(x).ToLower = "fefvi" Then
                            workRow(y) = yObj.FEFVI
                        ElseIf yield_item(x).ToLower = "beos" Then
                            workRow(y) = yObj.BEOS
                        ElseIf yield_item(x).ToLower = "befvi" Then
                            workRow(y) = yObj.BEFVI
                        ElseIf yield_item(x).ToLower = "bumpaoi" Then
                            workRow(y) = yObj.Bump
                        ElseIf yield_item(x).ToLower = "c4" Then
                            workRow(y) = yObj.C4
                        ElseIf yield_item(x).ToLower = "fe total" Then
                            workRow(y) = yObj.FETOTAL
                        ElseIf yield_item(x).ToLower = "be total" Then
                            workRow(y) = yObj.BETOTAL
                        ElseIf yield_item(x).ToLower = "ccpin" Then
                            workRow(y) = yObj.CCPin
                        ElseIf yield_item(x).ToLower = "pincc" Then
                            workRow(y) = yObj.PinCC
                        ElseIf yield_item(x).ToLower = "2wos" Then
                            workRow(y) = yObj.WOS2
                        ElseIf yield_item(x).ToLower = "4wos" Then
                            workRow(y) = yObj.WOS4
                        ElseIf yield_item(x).ToLower = "bumpaoi" Then
                            workRow(y) = yObj.BumpAOI
                        ElseIf yield_item(x).ToLower = "fluxclean後bumpaoi" Then
                            workRow(y) = yObj.BumpAOI_FluxClean
                        ElseIf (yield_item(x).ToLower = "total yield") Or (yield_item(x).ToLower = "total") Then
                            workRow(y) = yObj.Totalyield
                        ElseIf yield_item(x).ToLower = "land" Then
                            workRow(y) = yObj.Land
                        End If

                    Else
                        workRow(y) = ""
                    End If

                End If
            Next

            workTable.Rows.Add(workRow)
            chartTable.ImportRow(workRow)
        Next

    End Sub

    Private Function getGridView() As GridView

        Dim dataGrid As GridView = New GridView
        dataGrid.BackColor = Drawing.Color.White
        dataGrid.BorderColor = Drawing.Color.Black
        dataGrid.BorderWidth = Unit.Pixel(1)
        dataGrid.CellPadding = 3
        dataGrid.HeaderStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#5D7B9D")
        dataGrid.HeaderStyle.Font.Bold = True
        dataGrid.HeaderStyle.ForeColor = Drawing.Color.White
        dataGrid.AlternatingRowStyle.BackColor = System.Drawing.ColorTranslator.FromHtml("#DBEEFF")

        Return dataGrid

    End Function

    Protected Sub gv_rowdata_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_rowdata.RowDataBound

        If e.Row.RowType = DataControlRowType.Header Then

            For i As Integer = 0 To (e.Row.Cells.Count - 1)

                If i <= 2 Then
                    e.Row.Cells(i).Width = Unit.Pixel(90)
                End If

                If i >= 3 Then
                    e.Row.Cells(i).Font.Size = FontUnit.Point(8)
                    e.Row.Cells(i).Width = Unit.Pixel(60)
                End If

                If rb_DataTimeCustor.SelectedIndex = 1 Then
                    If i >= 3 Then
                        If rbl_lossItem.SelectedIndex = 0 Then
                            ' Day
                            e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>" + (e.Row.Cells(i).Text).Substring(4, 4)
                        ElseIf rbl_lossItem.SelectedIndex = 1 Then
                            ' Week
                            e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>W" + (e.Row.Cells(i).Text).Substring(4, 2)
                        Else
                            ' Month
                            e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>M" + (e.Row.Cells(i).Text).Substring(5, 2)
                        End If
                    End If
                Else
                    If i >= 3 And i <= 5 Then
                        e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>M" + (e.Row.Cells(i).Text).Substring(5, 2)
                    End If

                    If i >= 6 And i <= 10 Then
                        e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>W" + (e.Row.Cells(i).Text).Substring(4, 2)
                    End If

                    If i >= 11 Then
                        e.Row.Cells(i).Text = (e.Row.Cells(i).Text).Substring(0, 4) + "<br>" + (e.Row.Cells(i).Text).Substring(4, 4)
                    End If
                End If

            Next

        End If

        If e.Row.RowType = DataControlRowType.DataRow Then

            e.Row.Height = Unit.Pixel(30)
            e.Row.Cells(0).Font.Size = FontUnit.XXSmall
            e.Row.Cells(0).Font.Bold = True
            e.Row.Cells(0).ForeColor = Drawing.Color.Red
            e.Row.Cells(1).Font.Size = FontUnit.XXSmall
            e.Row.Cells(1).Font.Bold = True
            e.Row.Cells(2).Font.Size = FontUnit.XXSmall
            e.Row.Cells(2).Font.Bold = True

        End If

    End Sub

    Private Sub DrawChartBySingle(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef chartHash As Hashtable, ByVal yieldIndex As Integer, ByVal yieldName As String)

        Dim aryColor() As Color = {Color.DodgerBlue, Color.Olive, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.Blue, Color.Firebrick, Color.Green}
        Chart.ImageUrl = "temp/DailyYieldBihon_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(400)
        Chart.Width = Unit.Pixel(1050)

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.Title = "【" + yieldName + "】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim chartDT As DataTable
        Dim isMultiChart As Boolean = False

        Try
            For Each entry As DictionaryEntry In chartHash
                chartDT = CType(entry.Value, DataTable)
                If rb_DataTimeCustor.SelectedIndex = 1 Then
                    DrawBar_Range(Chart, aryColor(0), entry, chartDT, yieldIndex, isMultiChart)
                Else
                    DrawBar(Chart, aryColor(0), entry, chartDT, yieldIndex, isMultiChart)
                End If

            Next
        Catch ex As Exception
        End Try

        Chart.ChartAreas("Default").AxisY.Interval = 10
        If (yieldName.ToUpper).IndexOf("YIELD") >= 0 Then
            Chart.ChartAreas("Default").AxisY.Minimum = 50
            Chart.ChartAreas("Default").AxisY.Maximum = 100
        Else
            Chart.ChartAreas("Default").AxisY.Minimum = 0
            Chart.ChartAreas("Default").AxisY.Maximum = 100
        End If

    End Sub

    Private Sub DrawChart(ByRef Chart As Dundas.Charting.WebControl.Chart, ByRef chartAry As ArrayList, ByVal yieldIndex As Integer, ByVal yieldName As String)

        Dim aryColor() As Color = {Color.DodgerBlue, Color.Olive, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.Blue, Color.Firebrick, Color.Green}
        Chart.ImageUrl = "temp/DailyYieldBihon_#SEQ(1000,1)"
        Chart.ImageType = ChartImageType.Png
        Chart.Palette = ChartColorPalette.Dundas
        Chart.Height = Unit.Pixel(400)
        Chart.Width = Unit.Pixel(1050)

        Chart.Palette = ChartColorPalette.Dundas
        Chart.BackColor = Color.White
        Chart.BackGradientEndColor = Color.Peru
        Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
        Chart.BorderStyle = ChartDashStyle.Solid
        Chart.BorderWidth = 3
        Chart.BorderColor = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.Title = "【" + yieldName + "】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim chartDT As DataTable
        Dim chartHash As Hashtable
        Dim isMultiChart As Boolean = False

        If (chartAry.Count > 1) Then
            isMultiChart = True
        End If

        For i As Integer = 0 To (chartAry.Count - 1)

            Try
                chartHash = CType(chartAry(i), Hashtable)
                For Each entry As DictionaryEntry In chartHash
                    If entry.Key.ToString <> "Summary" Then
                        ' --- Bar Chart ---
                        chartDT = CType(entry.Value, DataTable)
                        If rb_DataTimeCustor.SelectedIndex = 1 Then
                            DrawBar_Range(Chart, aryColor(i), entry, chartDT, yieldIndex, isMultiChart)
                        Else
                            DrawBar(Chart, aryColor(i), entry, chartDT, yieldIndex, isMultiChart)
                        End If
                    Else
                        ' --- Thrend Chart ---
                        chartDT = CType(entry.Value, DataTable)
                        DrawThrend(Chart, Color.DarkRed, entry, chartDT, yieldIndex)
                    End If
                Next
            Catch ex As Exception
            End Try

        Next

        If (yieldName.ToUpper).IndexOf("YIELD") >= 0 Then
            Chart.ChartAreas("Default").AxisY.Interval = 10
            Chart.ChartAreas("Default").AxisY.Minimum = 0 '50
            Chart.ChartAreas("Default").AxisY.Maximum = 100
        Else
            Chart.ChartAreas("Default").AxisY.Interval = 10
            Chart.ChartAreas("Default").AxisY.Minimum = 0
            Chart.ChartAreas("Default").AxisY.Maximum = 100 '20
        End If

    End Sub

    Private Sub DrawBar(ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal chartColor As Color, ByRef entry As DictionaryEntry, ByRef chartDT As DataTable, ByVal yieldIndex As Integer, ByVal Ismultibar As Boolean)

        Dim series As Series
        Dim XStr As String
        Dim XFormat As String
        Dim Value As Double
        Dim nYvalue As Double = 0
        Dim part_id As String = entry.Key.ToString

        series = Chart.Series.Add(part_id)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Column
        series.Color = chartColor
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series("PointWidth") = "0.5"

        Dim addIndex As Integer = 0
        Dim cpu_item() As String = CType(ViewState("CPU_gItem"), String())
        ' M month, W week, D day, 如果之後要加日期區間選擇, 都只有 D
        Dim rangeType As String = "D"
        For rowIndex As Integer = 3 To (chartDT.Columns.Count - 1)

            XStr = chartDT.Columns(rowIndex).ColumnName
            XFormat = XStr
            If rowIndex >= 3 And rowIndex <= 5 Then
                rangeType = "M"
                Dim MDate As Date = (DateTime.ParseExact(XStr, "yyyy-MM", Nothing))
                XFormat = MDate.ToString("yyyy MMM", System.Globalization.CultureInfo.InvariantCulture)
            ElseIf rowIndex >= 6 And rowIndex <= 10 Then
                rangeType = "W"
                XFormat = XFormat.Substring(0, 4) + " W" + XFormat.Substring(4, 2)
            Else
                rangeType = "D"
                If (cb_customerDay.Checked) And (XFormat.IndexOf("~") >= 0) Then
                    XFormat = (txtDateFrom.Text) + (ddlHourFrom.SelectedValue) + vbCrLf + (txtDateTo.Text) + (ddlHourTo.SelectedValue)
                Else
                    XFormat = XFormat.Substring(0, 4) + " " + XFormat.Substring(4, 4)
                End If
            End If

            Dim temp As DataRow = chartDT.Rows(yieldIndex)

            If Not IsDBNull(chartDT.Rows(yieldIndex)(rowIndex)) Then
                If (chartDT.Rows(yieldIndex)(rowIndex)).ToString.Length > 0 Then
                    Value = (CType(chartDT.Rows(yieldIndex)(rowIndex), Double) * 100)
                Else
                    Value = 0
                End If
            Else
                Value = 0
            End If

            If Value > nYvalue Then
                nYvalue = Value
            End If

            Chart.Series(part_id).Points.AddXY(XFormat, Value)
            Chart.Series(part_id).Points((rowIndex - 3) + addIndex).ToolTip = (XFormat + " : " + (Value.ToString) + "%")
            If Not Ismultibar Then
                Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Label = ((Value.ToString) + "%")
            End If

            ' --- Intel ---
            'If rb_Type.SelectedIndex = 0 Then
            '    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','NORMAL')"
            'ElseIf rb_Type.SelectedIndex = 1 Then
            '    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
            'End If

            If ddl_CustomerID.SelectedValue = "INTEL" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','NORMAL')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','NORMAL')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            Else
                If ddl_ProductType.SelectedValue.ToUpper() = "PPS" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','PPS')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            End If

            Chart.Series(part_id).Points((rowIndex - 3) + addIndex).BorderWidth = 1

            If rowIndex = 5 Then
                Chart.Series(part_id).Points.AddXY("", 0)
                addIndex = 1
            ElseIf rowIndex = 10 Then
                Chart.Series(part_id).Points.AddXY("", 0)
                addIndex = 2
            End If

        Next

        If nYvalue < 10 Then
            Chart.ChartAreas("Default").AxisY.Maximum = 10
        End If

    End Sub

    Private Sub DrawBar_Range(ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal chartColor As Color, ByRef entry As DictionaryEntry, ByRef chartDT As DataTable, ByVal yieldIndex As Integer, ByVal Ismultibar As Boolean)

        Dim series As Series
        Dim XStr As String
        Dim XFormat As String
        Dim Value As Double
        Dim nYvalue As Double = 0
        Dim part_id As String = entry.Key.ToString

        series = Chart.Series.Add(part_id)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Column
        series.Color = chartColor
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series("PointWidth") = "0.5"

        Dim addIndex As Integer = 0
        Dim cpu_item() As String = CType(ViewState("CPU_gItem"), String())
        Dim rangeType As String = "D"
        For rowIndex As Integer = 3 To (chartDT.Columns.Count - 1)

            XStr = chartDT.Columns(rowIndex).ColumnName
            XFormat = XStr

            If rbl_lossItem.SelectedIndex = 0 Then
                ' Day 
                rangeType = "D"
                XFormat = XFormat.Substring(0, 4) + " " + XFormat.Substring(4, 4)
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                ' Week
                rangeType = "W"
                XFormat = XFormat.Substring(0, 4) + " W" + XFormat.Substring(4, 2)
            Else
                ' Month
                rangeType = "M"
                Dim MDate As Date = (DateTime.ParseExact(XStr, "yyyy-MM", Nothing))
                XFormat = MDate.ToString("yyyy MMM", System.Globalization.CultureInfo.InvariantCulture)
            End If

            If Not IsDBNull(chartDT.Rows(yieldIndex)(rowIndex)) Then
                If (chartDT.Rows(yieldIndex)(rowIndex)).ToString.Length > 0 Then
                    Value = (CType(chartDT.Rows(yieldIndex)(rowIndex), Double) * 100)
                Else
                    Value = 0
                End If
            Else
                Value = 0
            End If

            If Value > nYvalue Then
                nYvalue = Value
            End If

            Chart.Series(part_id).Points.AddXY(XFormat, Value)
            Chart.Series(part_id).Points((rowIndex - 3) + addIndex).ToolTip = (XFormat + " : " + (Value.ToString) + "%")
            If Not Ismultibar Then
                Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Label = ((Value.ToString) + "%")
            End If

            If ddl_CustomerID.SelectedValue = "INTEL" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','NORMAL')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','NORMAL')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            Else
                If ddl_ProductType.SelectedValue.ToUpper() = "PPS" Then
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','PPS')"
                Else
                    Chart.Series(part_id).Points((rowIndex - 3) + addIndex).Href = "javascript:openWindowWithPost('DailyReportDetail.aspx', 'WEB', '" + (ddl_CustomerID.SelectedValue) + "','" + (ddl_ProductType.SelectedValue) + "','" + part_id + "','" + cpu_item(yieldIndex) + "','" + XStr + "','" + rangeType + "','BKM')"
                End If
            End If

            Chart.Series(part_id).Points((rowIndex - 3) + addIndex).BorderWidth = 1

        Next

        If nYvalue < 10 Then
            Chart.ChartAreas("Default").AxisY.Maximum = 10
        End If

    End Sub

    Private Sub DrawThrend(ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal chartColor As Color, ByRef entry As DictionaryEntry, ByRef chartDT As DataTable, ByVal yieldIndex As Integer)

        Dim series As Series
        Dim XStr As String
        Dim Value As Double
        
        series = Chart.Series.Add(entry.Key.ToString)
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Line
        series.Color = chartColor
        series.BorderColor = Color.White
        series.BorderWidth = 1
        series.MarkerStyle = MarkerStyle.Circle
        series.MarkerSize = 8
        series.MarkerColor = Color.DarkRed

        Dim addIndex As Integer = 0
        For rowIndex As Integer = 3 To (chartDT.Columns.Count - 1)

            XStr = chartDT.Columns(rowIndex).ColumnName
            If Not IsDBNull(chartDT.Rows(yieldIndex)(rowIndex)) Then
                If (chartDT.Rows(yieldIndex)(rowIndex)).ToString.Length > 0 Then
                    Value = (CType(chartDT.Rows(yieldIndex)(rowIndex), Double) * 100)
                Else
                    Value = 0
                End If
            Else
                Value = 0
            End If

            Chart.Series(entry.Key.ToString).Points.AddXY(XStr, Value)
            Chart.Series(entry.Key.ToString).Points((rowIndex - 3) + addIndex).ToolTip = (XStr + " : " + (Value.ToString) + "%")
            Chart.Series(entry.Key.ToString).Points((rowIndex - 3) + addIndex).Label = ((Value.ToString) + "%")

            If rb_DataTimeCustor.SelectedIndex = 0 Then
                If rowIndex = 5 Then
                    Chart.Series(entry.Key.ToString).Points.AddXY("", 0)
                    addIndex = 1
                ElseIf rowIndex = 10 Then
                    Chart.Series(entry.Key.ToString).Points.AddXY("", 0)
                    addIndex = 2
                End If
            End If
        Next

    End Sub

    Private Sub getSPECLine(ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal Part_ID As String, ByVal yieldType As String)

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
       
        Try
            conn.Open()
            sqlStr = "SELECT "
            sqlStr += "CASE WHEN (UCL_Sigma = 3) THEN UCL_3S  WHEN (UCL_Sigma = 4) THEN UCL_4S WHEN (UCL_Sigma = 5) THEN UCL_5S ELSE NULL END as UCL,"
            sqlStr += "CASE WHEN (LCL_Sigma = 3) THEN LCL_3S  WHEN (LCL_Sigma = 4) THEN LCL_4S WHEN (LCL_Sigma = 5) THEN LCL_5S ELSE NULL END as LCL "
            sqlStr += "FROM dbo.Yield_SPEC "
            sqlStr += "WHERE 1=1 "
            sqlStr += "AND Part_Id='{0}' "
            sqlStr += "AND UPPER(OOC_Code)='{1}' "
            sqlStr = String.Format(sqlStr, Part_ID, (yieldType.ToUpper()))
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            If myDT.Rows.Count > 0 Then

                If Not IsDBNull(myDT.Rows(0)("UCL")) Then
                    Chart_UCL(Chart, "UCL", Math.Round((CType(myDT.Rows(0)("UCL").ToString(), Double) * 100), 2, MidpointRounding.AwayFromZero))
                End If

                If Not IsDBNull(myDT.Rows(0)("LCL")) Then
                    Chart_UCL(Chart, "LCL", Math.Round((CType(myDT.Rows(0)("LCL").ToString(), Double) * 100), 2, MidpointRounding.AwayFromZero))
                End If

            End If

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Public Sub Chart_UCL(ByRef Chart As Dundas.Charting.WebControl.Chart, ByVal LineType As String, ByVal LineValue As Double)

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
        For i As Integer = 0 To 18
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

    ' >>
    Protected Sub but_weekTo_Click(sender As Object, e As System.EventArgs) Handles but_weekTo.Click

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        Dim sourceAry As ArrayList = New ArrayList
        Dim DestAry As ArrayList = New ArrayList
        For i As Integer = 0 To (lb_weekSource.Items.Count - 1)

            If ddlPart.SelectedValue = "All" Then
                If lb_weekSource.Items(i).Value.ToUpper() = "TOTAL YIELD" Or lb_weekSource.Items(i).Value.ToUpper() = "TOTAL" Then
                    DestAry.Add(lb_weekSource.Items(i).Value)
                Else
                    sourceAry.Add(lb_weekSource.Items(i).Value)
                End If
            Else
                If lb_weekSource.Items(i).Selected Then
                    DestAry.Add(lb_weekSource.Items(i).Value)
                Else
                    sourceAry.Add(lb_weekSource.Items(i).Value)
                End If
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

    ' <<
    Protected Sub but_weekBack_Click(sender As Object, e As System.EventArgs) Handles but_weekBack.Click

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
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

    ' >>
    Protected Sub but_dateRight_Click(sender As Object, e As System.EventArgs) Handles but_dateRight.Click
        moveList(listB_timeSource, listB_timeShow)
    End Sub

    ' <<
    Protected Sub but_dateLeft_Click(sender As Object, e As System.EventArgs) Handles but_dateLeft.Click
        moveList(listB_timeShow, listB_timeSource)
    End Sub

    ' Alert 呈現
    Public Sub ShowMessage(ByVal mesStr As String)

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("alert('" + mesStr + "');")
        sb.Append("</script>")
        Dim myCSManager As ClientScriptManager = Page.ClientScript
        myCSManager.RegisterStartupScript(Me.GetType(), "SetStatusScript", sb.ToString())

    End Sub

    ' 自訂時間
    Protected Sub cb_customerDay_CheckedChanged(sender As Object, e As System.EventArgs) Handles cb_customerDay.CheckedChanged

        tr_dateRange.Visible = False
        rb_DataTimeCustor.SelectedIndex = 0
        If cb_customerDay.Checked Then
            txtDateFrom.Enabled = True
            txtDateTo.Enabled = True
            Calendar1.Enabled = True
            Calendar2.Enabled = True
            ddlHourFrom.Enabled = True
            ddlHourTo.Enabled = True
            cb_ShowToday.Checked = False
        Else
            txtDateFrom.Enabled = False
            txtDateTo.Enabled = False
            Calendar1.Enabled = False
            Calendar2.Enabled = False
            ddlHourFrom.Enabled = False
            ddlHourTo.Enabled = False
        End If

    End Sub

    ' 今天呈現
    Protected Sub cb_ShowToday_CheckedChanged(sender As Object, e As System.EventArgs) Handles cb_ShowToday.CheckedChanged
        If cb_ShowToday.Checked Then
            tr_dateRange.Visible = False
            cb_customerDay.Checked = False
            rb_DataTimeCustor.SelectedIndex = 0
        End If
    End Sub

    Private Sub moveList(ByRef sourceList As ListBox, ByRef destList As ListBox)

        Dim sourceAry As New ArrayList()
        Dim DestAry As New ArrayList()

        For i As Integer = 0 To sourceList.Items.Count - 1
            If sourceList.Items(i).Selected Then
                DestAry.Add(sourceList.Items(i).Value)
            Else
                sourceAry.Add(sourceList.Items(i).Value)
            End If
        Next
        sourceList.Items.Clear()

        For i As Integer = 0 To sourceAry.Count - 1
            sourceList.Items.Add(sourceAry(i).ToString())
        Next
        For i As Integer = 0 To DestAry.Count - 1
            destList.Items.Add(DestAry(i).ToString())
        Next

    End Sub

    Private Sub moveListOnlyTotal(ByRef sourceList As ListBox, ByRef destList As ListBox)

        Dim sourceAry As New ArrayList()
        Dim DestAry As New ArrayList()

        For i As Integer = 0 To sourceList.Items.Count - 1
            DestAry.Add(sourceList.Items(i).Value)
        Next
        sourceList.Items.Clear()

        For i As Integer = 0 To DestAry.Count - 1
            destList.Items.Add(DestAry(i).ToString())
        Next

        ' 移去 Total 與 加入 Total

    End Sub

    Protected Sub rbl_lossItem_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rbl_lossItem.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim tableStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        tr_gvDisplay.Visible = False

        Try

            conn.Open()
            If ddl_CustomerID.SelectedValue = "INTEL" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    tableStr += "from dbo.Yield_Daily_RawData "
                    tableStr += "where customer_id='INTEL' "
                Else
                    tableStr += "from dbo.BKM_Yield_Daily_RawData "
                End If
            ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                    tableStr += "from dbo.Yield_Daily_RawData "
                    tableStr += "where customer_id='AMD' "
                Else
                    tableStr += "from dbo.Yield_Daily_RawData_NonIntel "
                End If
            Else
                tableStr += "from dbo.Yield_Daily_RawData_NonIntel "
            End If

            If rbl_lossItem.SelectedIndex = 0 Then
                ' Daily 
                sqlStr = "select Top(120) convert(char(10), datatime, 112) as trtm "
                sqlStr += tableStr
                sqlStr += "group by convert(char(10), datatime, 112) "
                sqlStr += "order by convert(char(10), datatime, 112) desc "
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                ' Week 
                sqlStr = "select Top(52) yearWW as trtm "
                sqlStr += tableStr
                sqlStr += "group by yearWW "
                sqlStr += "order by yearWW desc "
            Else
                ' Month
                sqlStr = "select Top(12) convert(char(7), datatime, 120) as trtm "
                sqlStr += tableStr
                sqlStr += "group by convert(char(7), datatime, 120) "
                sqlStr += "order by convert(char(7), datatime, 120) desc "
            End If

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            listB_timeSource.Items.Clear()
            listB_timeShow.Items.Clear()
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_timeSource.Items.Add(myDT.Rows(i)("trtm").ToString())
            Next

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Protected Sub rb_DataTimeCustor_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles rb_DataTimeCustor.SelectedIndexChanged

        tr_dateRange.Visible = False
        If rb_DataTimeCustor.SelectedIndex = 1 Then

            tr_gvDisplay.Visible = False
            tr_dateRange.Visible = True
            cb_ShowToday.Checked = False
            cb_customerDay.Checked = False


            ' 取得時間區間
            Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            Dim sqlStr As String = ""
            Dim tableStr As String = ""
            Dim myDT As DataTable
            Dim myAdapter As SqlDataAdapter
            Dim categoryStr As String = ""

            Try

                conn.Open()

                If ddl_CustomerID.SelectedValue = "INTEL" Then
                    If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                        tableStr += "from dbo.Yield_Daily_RawData "
                        tableStr += "where customer_id='INTEL' "
                    Else
                        tableStr += "from dbo.BKM_Yield_Daily_RawData "
                    End If
                ElseIf ddl_CustomerID.SelectedValue = "AMD" Then
                    If ddl_YieldMode.SelectedValue.ToUpper() = "BY FAIL MODE" Then
                        tableStr += "from dbo.Yield_Daily_RawData "
                        tableStr += "where customer_id='AMD' "
                    Else
                        tableStr += "from dbo.Yield_Daily_RawData_NonIntel "
                    End If
                Else
                    tableStr += "from dbo.Yield_Daily_RawData_NonIntel "
                End If

                If rbl_lossItem.SelectedIndex = 0 Then
                    ' Daily 
                    sqlStr = "select convert(char(10), datatime, 112) as trtm "
                    sqlStr += tableStr
                    sqlStr += "group by convert(char(10), datatime, 112) "
                    sqlStr += "order by convert(char(10), datatime, 112) desc "
                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                    ' Week 
                    sqlStr = "select yearWW as trtm "
                    sqlStr += tableStr
                    sqlStr += "group by yearWW "
                    sqlStr += "order by yearWW desc "
                Else
                    ' Month
                    sqlStr = "select convert(char(7), datatime, 120) as trtm "
                    sqlStr += tableStr
                    sqlStr += "group by convert(char(7), datatime, 120) "
                    sqlStr += "order by convert(char(7), datatime, 120) desc "
                End If

                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                conn.Close()

                listB_timeSource.Items.Clear()
                listB_timeShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_timeSource.Items.Add(myDT.Rows(i)("trtm").ToString())
                Next

            Catch ex As Exception

            Finally
                If conn.State = ConnectionState.Open Then
                    conn.Close()
                End If
            End Try

        End If

    End Sub

End Class
