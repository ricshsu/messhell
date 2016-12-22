Imports System.Data.SqlClient
Imports System.Data
Imports System.Drawing
Imports Dundas.Charting.WebControl
Imports System.IO
Imports System.Data.OleDb

Partial Class YieldLoss_Test
    Inherits System.Web.UI.Page

    ' Week Table : BinCode_Summary
    ' Daily Table : BinCode_Daily_Lot
    Private confTable As String = "Customer_Prodction_Mapping_BU_Rename"
    Private Const gChartH As Integer = 600
    Private Const gChartW As Integer = 1080
    'Private Const gChartW As Integer = 2160
    'Dim aryColor() As Color = {Color.Blue, Color.DarkOrange, Color.Purple, Color.Green, Color.Firebrick, Color.DodgerBlue, Color.Olive, Color.DarkGreen, Color.Red, Color.Gold, Color.Gray, Color.Cyan}
    Dim aryColor() As Color = {Color.DodgerBlue, Color.Olive, Color.DarkOrange, Color.Purple, Color.DarkGreen, Color.Blue, Color.Firebrick, Color.Green, Color.DarkSlateBlue, Color.DarkSlateGray, Color.Khaki, Color.Thistle, Color.LightSkyBlue, Color.LightPink, Color.LightSalmon, Color.Lime, Color.LightSeaGreen, Color.LimeGreen, Color.ForestGreen, Color.DeepSkyBlue, Color.DarkTurquoise, Color.DarkViolet, Color.DarkKhaki, Color.Tan}
    'Dim aryColor() As Color = {Color.FromArgb(179, 29, 64), Color.FromArgb(239, 104, 38), Color.FromArgb(245, 218, 13), Color.FromArgb(134, 196, 63), Color.FromArgb(37, 170, 227), Color.FromArgb(16, 83, 164), Color.FromArgb(88, 55, 146), Color.FromArgb(216, 27, 91), Color.FromArgb(252, 181, 29), Color.FromArgb(247, 238, 47), Color.FromArgb(29, 157, 132), Color.FromArgb(24, 121, 190), Color.FromArgb(13, 84, 166), Color.FromArgb(187, 29, 106), Color.FromArgb(240, 149, 32), Color.FromArgb(204, 219, 40), Color.FromArgb(32, 127, 145), Color.FromArgb(25, 86, 166), Color.FromArgb(45, 57, 141), Color.FromArgb(150, 36, 147)}
    Dim plantAry() As String = {"1", "2", "3", "4", "5", "B", "T", "S", "All"}
    Private Structure YieldlossInfo
        Dim BumpingType As String
        Dim Part_ID As String
        Dim TimePeriod As Integer
        Dim TimeRange As String
        Dim TotalOriginal As Integer
        Dim BU As String
        Dim nTop As Integer
        Dim Fail_Mode As String
        Dim xoutscrape As Boolean
    End Structure


    Private Structure FailObj

        Dim OriFail_Mode As String
        Dim Fail_Mode As String
        Dim Fail_Value As Double

    End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
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
            listB_BumpingTypeSource.Enabled = False
            listB_BumpingTypeShow.Enabled = False

            rb_ProductPart.SelectedIndex = 1
            rb_ProductPart.Items(0).Enabled = True
            rb_ProductPart.Items(1).Enabled = True

            conn.Open()

            ' -- Product --
            sqlStr = "select Category from Customer_Prodction_Mapping_BU_Rename where 1=1 "
            'sqlStr += "and fail_function=1 "
            sqlStr += "group by Category order by Category"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillController(myDT, ddlProduct, 0)

            ' -- Customer ID --
            'sqlStr = "select customer_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
            'sqlStr += "and fail_function=1 "
            'sqlStr += "and Category='" + ddlProduct.SelectedValue + "' "
            'sqlStr += "group by customer_id order by customer_id"
            If ddlProduct.Text = "PPS" Then
                sqlStr = "select distinct [Assfct] from MES.[dbo].[ProductInfo] where bu='" + ddlProduct.SelectedValue + "' group by assfct order by assfct"

                'sqlStr = "select b.Assfct from Customer_Prodction_Mapping_BU_Rename a "
                'sqlStr += "LEFT JOIN MES.dbo.ProductInfo b ON a.Part_Id = b.Part_No "
                'sqlStr += "where (a.customer_id is not null and a.customer_id <> '')"

            ElseIf ddlProduct.Text = "PCB" Then
                sqlStr = "select distinct [Dircu] from MES.[dbo].[ProductInfo] where bu='" + ddlProduct.SelectedValue + "' group by Dircu order by Dircu"

                sqlStr = "select distinct b.Dircu from Customer_Prodction_Mapping_BU_Rename a "
                sqlStr += "LEFT JOIN MES.dbo.ProductInfo b ON a.Part_Id = b.Part_No "
                sqlStr += "where (a.customer_id is not null and a.customer_id <> '' and a.category='PCB' ) order by Dircu"

            Else

                sqlStr = "select customer_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                sqlStr += "and Category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by customer_id order by customer_id"
            End If


            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, listB_CustomerSource, 0)

            If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then
                ' -- Part ID --
                rb_ProductPart.Items(0).Enabled = False
                rb_ProductPart.SelectedIndex = 1
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                'sqlStr += "and fail_function=1 "
               
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

            Else
                ' -- Production ID --
                rb_ProductPart.SelectedIndex = 0
                rb_ProductPart.Items(0).Enabled = True
                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
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

            'ddl_BumpingType.Enabled = False
            listB_BumpingTypeSource.Items.Clear()
            listB_BumpingTypeShow.Items.Clear()

            If ddlProduct.Text = "PCB" Then
                sqlStr = "select Notes as Bumping_Type from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
                sqlStr += "and BU_type='" + ddlProduct.Text + "' and Notes is not null "
                sqlStr += "group by Notes order by Notes"

                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myDT = New DataTable
                myAdapter.Fill(myDT)
            Else
                sqlStr = "select Bumping_Type from " + confTable + " where 1=1 "
                sqlStr += "and BU_type='" + ddlProduct.Text + "' "
                sqlStr += "group by Bumping_Type order by Bumping_Type"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myDT = New DataTable
                myAdapter.Fill(myDT)
            End If

            For i As Integer = 0 To myDT.Rows.Count - 1
                If Not IsDBNull(myDT.Rows(i)("Bumping_Type")) Then
                    If myDT.Rows(i)("Bumping_Type") <> "" Then
                        listB_BumpingTypeSource.Items.Add(myDT.Rows(i)("Bumping_Type"))
                    End If
                End If
            Next


            listB_OLProcessSource.Items.Clear()
            listB_OLProcessShow.Items.Clear()
            sqlStr = "select OL_Process from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
            sqlStr += "and BU_type='" + ddlProduct.Text + "' and OL_Process is not null "
            sqlStr += "group by OL_Process order by OL_Process"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_OLProcessSource.Items.Add(myDT.Rows(i)("OL_Process"))
            Next


            listB_BackendSource.Items.Clear()
            listB_BackendShow.Items.Clear()
            sqlStr = "select backend from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
            sqlStr += "and BU_type='" + ddlProduct.Text + "' and backend is not null "
            sqlStr += "group by backend order by backend"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_BackendSource.Items.Add(myDT.Rows(i)("backend"))
            Next


            rb_dayType.SelectedIndex = 1   ' Daily OR Weekly
            rbl_week.SelectedIndex = 0     ' Day Type  Customer OR Default
            rbl_lossItem.SelectedIndex = 0 ' Loss Item Customer OR Default

            conn.Close()

            If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then
                listB_BumpingTypeSource.Enabled = True
                listB_BumpingTypeShow.Enabled = True
                tr_BumpingType.Visible = True
                If ddlProduct.Text = "PCB" Then
                    tr_BumpingType.Cells(0).InnerText = "Notes"
                    tr1.Cells(0).InnerText = "End User"
                Else
                    tr_BumpingType.Cells(0).InnerText = "Bumping Type"
                    tr1.Cells(0).InnerText = "Assembly Plant"
                End If

                tr_OLProcess.Visible = True
                tr_Backend.Visible = True

                ckFAI.Visible = True
                cb_Non8K.Visible = True
                cb_uploadLot.Visible = True
                cb_NonIPQC.Visible = False
                cb_Lot_Merge.Visible = True
                '昆山用
                Cb_SF.Visible = True
                Cb_Inline.Visible = True
            Else
                listB_BumpingTypeSource.Enabled = False
                listB_BumpingTypeShow.Enabled = False
                tr_BumpingType.Visible = False
                tr_OLProcess.Visible = False
                tr_Backend.Visible = False
                tr1.Cells(0).InnerText = "Customer_id"
                cb_Non8K.Visible = False
                'cb_uploadLot.Visible = False
                cb_NonIPQC.Visible = True
                Cb_Inline.Visible = False
                Cb_Inline.Checked = False
            End If
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' InQuery
    Protected Sub but_Execute_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_Execute.Click
        'but_Excel.Enabled = False
        ViewState("RowData") = Nothing
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False

        If rb_dayType.SelectedIndex = 0 Then
            lab_wait.Text = "最新一天無資料, 可使用天數自訂 !"
            If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 24 Then
                ShowMessage("選擇天數最多為 24 天")
                Exit Sub
            End If
            'daily_failMode()
        ElseIf rb_dayType.SelectedIndex = 1 Then
            lab_wait.Text = "最新一週無資料, 可使用週數自訂 !"
            If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 24 Then
                ShowMessage("選擇週數最多為 24 週")
                Exit Sub
            End If
            'weekly_failMode()
        ElseIf rb_dayType.SelectedIndex = 2 Then
            lab_wait.Text = "最新一月無資料, 可使用月數自訂 !"
            If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 24 Then
                ShowMessage("選擇月數最多為 24 月")
                Exit Sub
            End If
            'monthly_failMode()
        End If
        failMode()
    End Sub

    Private Sub failMode()
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim customStr As String = ""
        Dim plantStr As String = ""
        Dim partStr As String = ""
        Dim weekStr As String = ""
        Dim itemStr As String = ""
        Dim topStr As String = ""
        Dim myAdapter As SqlDataAdapter
        Dim topDT As DataTable = New DataTable
        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable


        ' --- Bumping Type ---
        Dim strBumpingType As String = ""
        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
            If n = 0 Then
                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
            Else
                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
            End If
        Next

        Dim sGetPartID As String = Get_PartID()
        ' --- DateTime ---

        Dim dateTimeTemp As String = ""
        If rbl_week.SelectedIndex = 1 Then
            For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                dateTimeTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
            Next
            If (dateTimeTemp <> "") Then
                dateTimeTemp = dateTimeTemp.Substring(0, (dateTimeTemp.Length - 1))
            End If


        Else
            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 1 Then
                dateTimeTemp = GetYearWW(Date.Now)
            End If

            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 0 Then
                dateTimeTemp = GetYearDay(Date.Now)
            End If

            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 2 Then
                dateTimeTemp = GetYearMM(Date.Now)
            End If


        End If

        ' --- Yield Loss ID ---
        Dim Failmodeitem As String = ""

        Dim nTop As Integer = 10

        If rbl_lossItem.SelectedIndex = 0 Then
            topStr = "top(10)"
            nTop = 10
        ElseIf rbl_lossItem.SelectedIndex = 1 Then
            topStr = "top(20)"
            nTop = 20
        ElseIf rbl_lossItem.SelectedIndex = 2 Then
            topStr = "top(30)"
            nTop = 30
        ElseIf rbl_lossItem.SelectedIndex = 3 Then
            topStr = "top(40)"
            nTop = 40
        ElseIf rbl_lossItem.SelectedIndex = 4 Then
            topStr = "top(100)"
            nTop = 100
        Else
            nTop = 50
            For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                If i = 0 Then
                    Failmodeitem = "'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "'"
                Else
                    Failmodeitem += ",'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "'"
                End If

            Next
            'If (Failmodeitem.Length > 0 AndAlso lb_LossShow.Items.Count > 0) Then
            '    Failmodeitem = Failmodeitem.Substring(0, (Failmodeitem.Length - 1))

            'End If
        End If


        'Alfie--------------------------------------------------------------------------------------------------------------
        Dim yl As New YieldlossInfo
        'yl.BumpingType = Replace(strBumpingType, "'", "")
        yl.BumpingType = strBumpingType
        yl.Part_ID = Replace(sGetPartID, "'", "")
        yl.TimePeriod = rb_dayType.SelectedIndex
        yl.TimeRange = Replace(dateTimeTemp, "'", "")
        yl.xoutscrape = cb_DRowData0.Checked
        yl.nTop = nTop
        'yl.Fail_Mode = Replace(Failmodeitem, "'", "")
        yl.Fail_Mode = Failmodeitem
        yl.BU = ddlProduct.SelectedValue
        Dim sTemp As String = getTotalOriginal_SQL(yl)



        yl.TotalOriginal = Get_WB_TotalOriginal(sTemp)


        'sTemp = getTopWBSQL(yl)
        'sTemp = getRowDataWBSQL2(yl)
        '-------------------------------------------------------------------------------------------------------------------

        Try
            lab_wait.Text = ""
            Dim BumpingType_Part As String = ""
            Dim itemTemp As String = ""
            Dim iTotal As String = "0"

            itemStr = ""
            conn.Open()
            sqlStr = getTopWBSQL(yl)

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myAdapter.Fill(topDT)

            If topDT.Rows.Count <> 0 Then


                If ddlProduct.SelectedValue = "PPS" Then
                    sqlStr = getRowDataWBSQL2(yl)
                ElseIf ddlProduct.SelectedValue = "PCB" Then
                    sqlStr = getRowDataPCBSQL2(yl)
                Else
                    sqlStr = getRowDataFCSQL2(yl)
                End If




                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                rawDT = New DataTable
                myAdapter.Fill(rawDT)

                ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
                chipSetRawDT = New DataTable


                conn.Close()
                lab_wait.Text = ""
                If rb_dayType.SelectedIndex = 0 Then
                    If rawDT.Rows.Count = 0 Then
                        lab_wait.Text = "查無資料 !"
                        ShowMessage("查無資料 !")
                        Exit Sub
                    End If
                    BarChart(rawDT, topDT, yl)
                Else
                    If rawDT.Rows.Count = 0 Then
                        lab_wait.Text = "查無資料 !"
                        ShowMessage("查無資料 !")
                        Exit Sub
                    End If
                    ' BarChart_FailModeByStageRatioSummary(rawDT, new_topDT)
                    BarChart_FailModeByStageRatioSummary(yl, rawDT, topDT)
                End If




                tr_chartDisplay.Visible = True
                If cb_DRowData.Checked Then
                    If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                        showDailyRowData(rawDT, chipSetRawDT)
                    Else
                        showDailyRowData_FC(rawDT, chipSetRawDT)
                    End If

                    tr_gvDisplay.Visible = True
                End If

                'Dim workTable As DataTable
                'sqlStr = getRawData_SQL(yl)
                'myAdapter = New SqlDataAdapter(sqlStr, conn)
                'myAdapter.SelectCommand.CommandTimeout = 360
                'workTable = New DataTable
                'myAdapter.Fill(workTable)
                'ViewState("RowData") = workTable
                but_Excel.Enabled = True
            Else
                lab_wait.Text = ""
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 
                    If rb_dayType.SelectedIndex = 0 Then
                        lab_wait.Text = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 無資料, 可使用天數自訂 !"
                    Else
                        lab_wait.Text = "最新一週無資料, 可使用週數自訂 !"
                    End If

                Else
                    ' Custom
                    lab_wait.Text = dateTimeTemp + " 無資料, 可使用天數自訂 !"
                End If
            End If
        Catch ex As Exception
            Dim sError As String = ex.ToString()
            If lab_wait.Text <> "" Then
                lab_wait.Text += "," + sError
            Else
                lab_wait.Text = sError
            End If

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub


    ' DDL --- CPU or CS or WB Change
    Protected Sub ddlProduct_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlProduct.SelectedIndexChanged

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        rbl_week.SelectedIndex = 0
        tr_week.Visible = False

        Try
            rb_ProductPart.Items(0).Enabled = True
            rb_ProductPart.Items(1).Enabled = True

            conn.Open()

            ' -- Customer ID --

            If ddlProduct.Text = "PPS" Then
                sqlStr = "select [Assfct] from MES.[dbo].[ProductInfo] where bu='" + ddlProduct.SelectedValue + "' group by assfct order by assfct"

                'sqlStr = "select distinct b.Assfct from Customer_Prodction_Mapping_BU_Rename a "
                'sqlStr += "LEFT JOIN MES.dbo.ProductInfo b ON a.Part_Id = b.Part_No "
                'sqlStr += "where (a.customer_id is not null and a.customer_id <> '') order by b.assfct"
            ElseIf ddlProduct.Text = "PCB" Then
                'sqlStr = "select [Dircu] from MES.[dbo].[ProductInfo] where bu='" + ddlProduct.SelectedValue + "' group by Dircu order by Dircu"
                sqlStr = "select distinct b.Dircu from Customer_Prodction_Mapping_BU_Rename a "
                sqlStr += "LEFT JOIN MES.dbo.ProductInfo b ON a.Part_Id = b.Part_No "
                sqlStr += "where (a.customer_id is not null and a.customer_id <> '' and a.category='PCB' ) order by Dircu"
            Else
                sqlStr = "select customer_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                sqlStr += "and Category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by customer_id order by customer_id"
            End If

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, listB_CustomerSource, 0)

            If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then
                ' -- Part ID --
                rb_ProductPart.SelectedIndex = 1
                rb_ProductPart.Items(0).Enabled = False
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                'sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
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
            Else
                ' -- Production ID --
                rb_ProductPart.SelectedIndex = 0
                rb_ProductPart.Items(0).Enabled = True
                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
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

            ' -- Yield Loss Item --
            If ddlProduct.Text <> "PPS" And ddlProduct.Text <> "PCB" Then


                sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                'sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
                Dim sGetPartID As String = Get_PartID()
                If sGetPartID <> "" Then
                    sqlStr += "and part_id in(" + sGetPartID + ") "
                    sqlStr += "group by fail_mode order by fail_mode"

                Else
                    sqlStr = "select Fail_Mode from dbo.VW_BinCode_Summary where 1=1 and WW='201401'  and part_id in('SNE135C') "
                End If

                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)
                UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

                conn.Close()
            End If


            listB_BumpingTypeSource.Items.Clear()
            listB_BumpingTypeShow.Items.Clear()

            If ddlProduct.Text = "PCB" Then
                sqlStr = "select Notes as Bumping_Type from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
                sqlStr += "and BU_type='" + ddlProduct.Text + "' and Notes is not null "
                sqlStr += "group by Notes order by Notes"
            Else
                sqlStr = "select Bumping_Type from " + confTable + " where 1=1 "
                sqlStr += "and BU_type='" + ddlProduct.Text + "' "
                sqlStr += "group by Bumping_Type order by Bumping_Type"

            End If

            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            For i As Integer = 0 To myDT.Rows.Count - 1
                If Not IsDBNull(myDT.Rows(i)("Bumping_Type")) Then
                    If myDT.Rows(i)("Bumping_Type") <> "" Then
                        listB_BumpingTypeSource.Items.Add(myDT.Rows(i)("Bumping_Type"))
                    End If
                End If
            Next


            listB_OLProcessSource.Items.Clear()
            listB_OLProcessShow.Items.Clear()
            sqlStr = "select OL_Process from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
            sqlStr += "and BU_type='" + ddlProduct.Text + "' and OL_Process is not null "
            sqlStr += "group by OL_Process order by OL_Process"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_OLProcessSource.Items.Add(myDT.Rows(i)("OL_Process"))
            Next


            listB_BackendSource.Items.Clear()
            listB_BackendShow.Items.Clear()
            sqlStr = "select backend from " + confTable + " a,MES.[dbo].[ProductInfo]b where 1=1 and a.Part_Id=b.Part_No "
            sqlStr += "and BU_type='" + ddlProduct.Text + "' and backend is not null "
            sqlStr += "group by backend order by backend"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            For i As Integer = 0 To myDT.Rows.Count - 1
                listB_BackendSource.Items.Add(myDT.Rows(i)("backend"))
            Next



            If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then
                listB_BumpingTypeSource.Enabled = True
                listB_BumpingTypeShow.Enabled = True
                tr_BumpingType.Visible = True
                If ddlProduct.Text = "PCB" Then
                    tr_BumpingType.Cells(0).InnerText = "Notes"
                    tr1.Cells(0).InnerText = "End User"
                Else
                    tr_BumpingType.Cells(0).InnerText = "Bumping Type"
                    tr1.Cells(0).InnerText = "Assembly Plant"
                End If
                tr_OLProcess.Visible = True
                tr_Backend.Visible = True

                cb_Non8K.Visible = True
                cb_uploadLot.Visible = True
                cb_NonIPQC.Visible = False
                cb_Lot_Merge.Visible = True
                ckFAI.Visible = True
                Cb_Inline.Visible = True

            Else
                listB_BumpingTypeSource.Enabled = False
                listB_BumpingTypeShow.Enabled = False
                tr_BumpingType.Visible = False
                tr_OLProcess.Visible = False
                tr_Backend.Visible = False
                tr1.Cells(0).InnerText = "Customer_id"
                cb_Non8K.Visible = False
                'cb_uploadLot.Visible = False
                cb_NonIPQC.Visible = True
                cb_Lot_Merge.Visible = False
                ckFAI.Visible = False
                Cb_Inline.Visible = False
                Cb_Inline.Checked = False
            End If
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' DDL --- Customer Change
    'Protected Sub ddlCustomer_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCustomer.SelectedIndexChanged

    '    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '    Dim sqlStr As String = ""
    '    Dim myDT As DataTable
    '    Dim myAdapter As SqlDataAdapter
    '    tr_chartDisplay.Visible = False
    '    tr_gvDisplay.Visible = False
    '    rbl_week.SelectedIndex = 0
    '    tr_week.Visible = False

    '    Try
    '        rb_ProductPart.Items(0).Enabled = True
    '        rb_ProductPart.Items(1).Enabled = True

    '        conn.Open()

    '        If ddlProduct.Text = "PPS" Then
    '            ' -- Part ID --
    '            rb_ProductPart.SelectedIndex = 1
    '            rb_ProductPart.Items(0).Enabled = False
    '            sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            If ddlCustomer.Text <> "All" Then
    '                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            End If
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            sqlStr += "group by part_id order by part_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
    '            Next
    '        Else
    '            ' -- Production ID --
    '            rb_ProductPart.SelectedIndex = 0
    '            rb_ProductPart.Items(0).Enabled = True
    '            sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            sqlStr += "group by production_id order by production_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '            Next
    '        End If

    '        ' -- Yield Loss Item --
    '        sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
    '        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '        'sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
    '        Dim sGetPartID As String = Get_PartID()
    '        If sGetPartID <> "" Then
    '            sqlStr += "and part_id in(" + sGetPartID + ") "
    '        End If
    '        sqlStr += "group by fail_mode order by fail_mode"
    '        myAdapter = New SqlDataAdapter(sqlStr, conn)
    '        myDT = New DataTable
    '        myAdapter.Fill(myDT)
    '        UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

    '        conn.Close()

    '    Catch ex As Exception

    '    Finally
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try

    'End Sub

    ' CheckBox Product / Part
    Protected Sub rb_ProductPart_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rb_ProductPart.SelectedIndexChanged
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim sCustomer_ID As String = ""
        Try

            conn.Open()
            For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
                If n = 0 Then
                    sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
                Else
                    sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
                End If
            Next

            ' -- Production_ID --
            If rb_ProductPart.SelectedIndex = 0 Then
                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If


                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by production_id order by production_id"
            Else
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If

                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "group by part_id order by part_id"
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            'UtilObj.FillController(myDT, ddlPart, 0)
            listB_PartSource.Items.Clear()
            listB_PartShow.Items.Clear()
            For i As Integer = 0 To myDT.Rows.Count - 1
                If rb_ProductPart.SelectedIndex = 0 Then
                    listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
                Else
                    listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                End If
            Next

            ' -- Yield Loss Item --
            sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
            If ddlProduct.Text = "WB" Then
                sqlStr += "and customer_id='All' "
            Else
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If
            End If
            'sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            Dim sGetPartID As String = Get_PartID()
            If sGetPartID <> "" Then
                sqlStr += "and part_id in(" + sGetPartID + ") "
            End If
            sqlStr += "group by fail_mode order by fail_mode"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

            conn.Close()
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub



    Protected Sub rb_dayType_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles rb_dayType.Init

    End Sub

    ' Change Daily OR Weekly
    Protected Sub rb_dayType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rb_dayType.SelectedIndexChanged
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        lab_wait.Text = ""

        Try

            conn.Open()
            If rbl_week.SelectedIndex = 1 Then
                'If rb_dayType.SelectedIndex = 0 Then
                '    ' Daily
                '    'sqlStr = "select top(90) Convert(char(10), Datetime, 120) "
                '    sqlStr = "SELECT TOP(90) SUBSTRING(CONVERT(VARCHAR, DATeTIME, 112),1,8) AS Datetime "
                '    sqlStr += "FROM SystemDateMapping a, dbo.VW_BinCode_Daily_Lot b "
                '    sqlStr += "WHERE 1=1 "
                '    sqlStr += "AND a.DateTime = b.trtm "
                '    sqlStr += "AND b.Category='" + ddlProduct.SelectedItem.Text + "' "

                '    If rb_ProductPart.SelectedIndex = 0 Then
                '        sqlStr += "AND b.Production_type='" + ddlPart.SelectedItem.Text + "' "
                '    Else
                '        sqlStr += "AND b.Part_ID='" + ddlPart.SelectedItem.Text + "' "
                '    End If

                '    sqlStr += "AND DateTime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '    sqlStr += "GROUP BY Datetime "
                '    sqlStr += "ORDER BY Datetime desc "
                'ElseIf rb_dayType.SelectedIndex = 1 Then
                '    If tr_BumpingType.Visible = False Then
                '        ' Weekly
                '        sqlStr = "select top(12) yearWW "
                '        sqlStr += "FROM SystemDateMapping a, dbo.VW_BinCode_Daily_Lot b "
                '        sqlStr += "WHERE 1=1 "
                '        sqlStr += "AND a.DateTime = b.trtm "
                '        sqlStr += "AND b.Category='" + ddlProduct.SelectedItem.Text + "' "

                '        If rb_ProductPart.SelectedIndex = 0 Then
                '            sqlStr += "AND b.Production_type='" + ddlPart.SelectedItem.Text + "' "
                '        Else
                '            sqlStr += "AND b.Part_ID='" + ddlPart.SelectedItem.Text + "' "
                '        End If

                '        If CStr(DatePart("w", DateTime.Now()) - 1) = "5" Then
                '            sqlStr += "AND DateTime < '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '        Else
                '            sqlStr += "AND DateTime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '        End If
                '        sqlStr += "GROUP BY yearWW "
                '        sqlStr += "ORDER BY yearWW desc "
                '    Else
                '        If listB_BumpingTypeShow.Items.Count = 0 Then
                '            ' Weekly
                '            sqlStr = "select top(12) yearWW "
                '            sqlStr += "FROM SystemDateMapping a, dbo.VW_BinCode_Daily_Lot b "
                '            sqlStr += "WHERE 1=1 "
                '            sqlStr += "AND a.DateTime = b.trtm "
                '            sqlStr += "AND b.Category='" + ddlProduct.SelectedItem.Text + "' "

                '            If rb_ProductPart.SelectedIndex = 0 Then
                '                sqlStr += "AND b.Production_type='" + ddlPart.SelectedItem.Text + "' "
                '            Else
                '                sqlStr += "AND b.Part_ID='" + ddlPart.SelectedItem.Text + "' "
                '            End If

                '            If CStr(DatePart("w", DateTime.Now()) - 1) = "5" Then
                '                sqlStr += "AND DateTime < '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            Else
                '                sqlStr += "AND DateTime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            End If
                '            sqlStr += "GROUP BY yearWW "
                '            sqlStr += "ORDER BY yearWW desc "
                '        Else
                '            Dim strBumpingType As String = ""
                '            For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                '                If n = 0 Then
                '                    strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                '                Else
                '                    strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                '                End If
                '            Next

                '            ' Weekly
                '            sqlStr = "SELECT top(12) WW FROM "
                '            sqlStr += "FROM WB_BinCode_Summary_ByBT "
                '            sqlStr += "WHERE 1=1 "
                '            sqlStr += "AND BumpingType_Id IN(" + strBumpingType + ") "
                '            If CStr(DatePart("w", DateTime.Now()) - 1) = "5" Then
                '                sqlStr += "AND DataTime < '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            Else
                '                sqlStr += "AND DataTime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            End If
                '            sqlStr += "GROUP BY WW ORDER BY WW DESC"
                '        End If
                '    End If
                'ElseIf rb_dayType.SelectedIndex = 2 Then
                '    If tr_BumpingType.Visible = False Then
                '        ' Weekly
                '        If ddlProduct.Text = "WB" Then
                '            sqlStr = "select top(12) MM "
                '            sqlStr += "FROM  dbo.WB_BinCode_Summary_Monthly "
                '            sqlStr += "WHERE 1=1 "
                '            'sqlStr += "AND Category='" + (ddlProduct.SelectedItem.Text) + "' "

                '            If rb_ProductPart.SelectedIndex = 0 Then
                '                sqlStr += "AND Production_type='" + (ddlPart.SelectedItem.Text) + "' "
                '            Else
                '                sqlStr += "AND Part_ID='" + (ddlPart.SelectedItem.Text) + "' "
                '            End If

                '            'If CStr(DatePart("w", DateTime.Now()) - 1) = "5" Then
                '            '    sqlStr += "AND DateTime < '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            'Else
                '            '    sqlStr += "AND DateTime <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' "
                '            'End If
                '            sqlStr += "AND MM <= '" + DateTime.Now.ToString("yyyyMM") + "' "
                '            sqlStr += "AND MM >= '" + (DateTime.Now.AddDays(-150).ToString("yyyyMM")) + "' "
                '            sqlStr += "GROUP BY MM "
                '            sqlStr += "ORDER BY MM desc "

                '            ' Weekly
                '            'sqlStr = "select top(12) WW as yearWW "
                '            'sqlStr += "FROM dbo.vw_BinCode_Summary "
                '            'sqlStr += "WHERE 1=1 "
                '            'sqlStr += "AND Part_ID='" + (ddlPart.SelectedItem.Text) + "' "
                '            'sqlStr += "GROUP BY WW ORDER BY WW desc "
                '        End If
                '    Else
                '        If listB_BumpingTypeShow.Items.Count = 0 Then
                '            ' Weekly
                '            If ddlProduct.Text = "WB" Then
                '                sqlStr = "select top(12) MM "
                '                sqlStr += "FROM  dbo.WB_BinCode_Summary_Monthly "
                '                sqlStr += "WHERE 1=1 "

                '                If rb_ProductPart.SelectedIndex = 0 Then
                '                    sqlStr += "AND Production_type='" + (ddlPart.SelectedItem.Text) + "' "
                '                Else
                '                    sqlStr += "AND Part_ID='" + (ddlPart.SelectedItem.Text) + "' "
                '                End If

                '                sqlStr += "AND MM <= '" + DateTime.Now.ToString("yyyyMM") + "' "
                '                sqlStr += "AND MM >= '" + (DateTime.Now.AddDays(-150).ToString("yyyyMM")) + "' "
                '                sqlStr += "GROUP BY MM "
                '                sqlStr += "ORDER BY MM desc "
                '            End If
                '        Else
                '            Dim strBumpingType As String = ""
                '            For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                '                If n = 0 Then
                '                    strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                '                Else
                '                    strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                '                End If
                '            Next

                '            ' Monthly
                '            sqlStr = "SELECT top(12) MM "
                '            sqlStr += "FROM WB_BinCode_Summary_Monthly_ByBT "
                '            sqlStr += "WHERE 1=1 "
                '            sqlStr += "AND BumpingType_Id IN(" + strBumpingType + ") "
                '            sqlStr += "AND MM <= '" + DateTime.Now.ToString("yyyyMM") + "' "
                '            sqlStr += "AND MM >= '" + (DateTime.Now.AddDays(-150).ToString("yyyyMM")) + "' "
                '            sqlStr += "GROUP BY MM ORDER BY MM DESC"
                '        End If
                '    End If
                'End If
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    'SELECT TOP(90) SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) AS Datetime FROM VW_BinCode_Daily_Lot WHERE 1=1 AND Category = 'WB' AND Part_ID = 'SPLP49H' GROUP BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) ORDER BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) DESC
                    sqlStr = "SELECT TOP(90) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datetime "
                    sqlStr += "FROM VW_BinCode_Daily_Lot "
                    sqlStr += "WHERE 1=1 "
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                    If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 Then
                        Dim strBumpingType As String = ""
                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                            If n = 0 Then
                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            Else
                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            End If
                        Next
                        sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    Else
                        Dim sGetPartID As String = Get_PartID()
                        If sGetPartID <> "" Then
                            If rb_ProductPart.SelectedIndex = 0 Then
                                sqlStr += "AND Production_type in(" + sGetPartID + ") "
                            Else
                                sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    sqlStr += "GROUP BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                    sqlStr += "ORDER BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) DESC"
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    'SELECT TOP(12) SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, DATATIME) AS NVARCHAR), 2) AS Datetime FROM VW_BinCode_Daily_Lot WHERE 1=1 AND Category = 'WB' AND Part_ID = 'SPLP49H' GROUP BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, DATATIME) AS NVARCHAR), 2) ORDER BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, DATATIME) AS NVARCHAR), 2) DESC
                    sqlStr = "SELECT TOP(12) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datetime "
                    sqlStr += "FROM VW_BinCode_Daily_Lot "
                    sqlStr += "WHERE 1=1 "
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                    If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 Then
                        Dim strBumpingType As String = ""
                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                            If n = 0 Then
                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            Else
                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            End If
                        Next
                        sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    Else
                        Dim sGetPartID As String = Get_PartID()
                        If sGetPartID <> "" Then
                            If rb_ProductPart.SelectedIndex = 0 Then
                                sqlStr += "AND Production_type in(" + sGetPartID + ") "
                            Else
                                sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    sqlStr += "GROUP BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) "
                    sqlStr += "ORDER BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) DESC"
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    'SELECT TOP(12) SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) AS Datetime FROM VW_BinCode_Daily_Lot WHERE 1=1 AND Category = 'WB' AND Part_ID = 'SPLP49H' GROUP BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) ORDER BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) DESC
                    sqlStr = "SELECT TOP(12) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datetime "
                    sqlStr += "FROM VW_BinCode_Daily_Lot "
                    sqlStr += "WHERE 1=1 "
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                    If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 Then
                        Dim strBumpingType As String = ""
                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                            If n = 0 Then
                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            Else
                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                            End If
                        Next
                        sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    Else
                        Dim sGetPartID As String = Get_PartID()
                        If sGetPartID <> "" Then
                            If rb_ProductPart.SelectedIndex = 0 Then
                                sqlStr += "AND Production_type in(" + sGetPartID + ") "
                            Else
                                sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    sqlStr += "GROUP BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                    sqlStr += "ORDER BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
                End If

                If cb_Lot_Merge.Checked = False Then
                    sqlStr = sqlStr.Replace("VW_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
                End If

                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 3600
                myDT = New DataTable
                myAdapter.Fill(myDT)
                lb_weekShow.Items.Clear()
                UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)
            End If

        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    ' Report Week Item
    Protected Sub rbl_week_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbl_week.SelectedIndexChanged

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        lb_weekSource.Items.Clear()
        lb_weekShow.Items.Clear()
        tr_week.Visible = False
        lab_wait.Text = ""

        'If rbl_week.SelectedIndex = 1 Then

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()
         
            If rb_dayType.SelectedIndex = 0 Then 'Daily
                'SELECT TOP(90) SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) AS Datetime FROM VW_BinCode_Daily_Lot WHERE 1=1 AND Category = 'WB' AND Part_ID = 'SPLP49H' GROUP BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) ORDER BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 8) DESC
                sqlStr = "SELECT TOP(90) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datetime "
                sqlStr += "FROM VW_BinCode_Daily_Lot with (nolock) "
                sqlStr += "WHERE 1=1 "

                If ddlProduct.SelectedItem.Text = "PPS" Then
                    sqlStr += "AND Category = 'WB' "
                Else
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                End If
                If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 Then
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        End If
                    Next
                    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                Else
                    Dim sGetPartID As String = Get_PartID()
                    If sGetPartID <> "" Then
                        If rb_ProductPart.SelectedIndex = 0 Then
                            sqlStr += "AND Production_type in(" + sGetPartID + ") "
                        Else
                            sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                        End If
                    End If
                End If
                sqlStr += "GROUP BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                sqlStr += "ORDER BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) DESC"
            ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly

                sqlStr = "SELECT TOP(24) WW AS Datetime "
                sqlStr += "FROM VW_BinCode_Daily_Lot  with (nolock) "
                sqlStr += "WHERE 1=1 "
                If ddlProduct.SelectedItem.Text = "PPS" Then
                    sqlStr += "AND Category = 'WB' "
                Else
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                End If

                If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 And listB_PartShow.Items.Count = 0 Then
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        End If
                    Next
                    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                Else
                    Dim sGetPartID As String = Get_PartID()
                    If sGetPartID <> "" Then
                        If rb_ProductPart.SelectedIndex = 0 Then
                            sqlStr += "AND Production_type in(" + sGetPartID + ") "
                        Else
                            sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                        End If
                    End If
                End If
                sqlStr += "GROUP BY WW "
                sqlStr += "ORDER BY WW DESC"



             ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                'SELECT TOP(12) SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) AS Datetime FROM VW_BinCode_Daily_Lot WHERE 1=1 AND Category = 'WB' AND Part_ID = 'SPLP49H' GROUP BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) ORDER BY SUBSTRING(CONVERT(VARCHAR, DATATIME, 112), 1, 6) DESC
                sqlStr = "SELECT TOP(24) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datetime "
                sqlStr += "FROM VW_BinCode_Daily_Lot  with (nolock) "
                sqlStr += "WHERE 1=1 "

                If ddlProduct.SelectedItem.Text = "PPS" Then
                    sqlStr += "AND Category = 'WB' "
                Else
                    sqlStr += "AND Category = '" + ddlProduct.SelectedItem.Text + "' "
                End If

                If tr_BumpingType.Visible = True AndAlso listB_BumpingTypeShow.Items.Count > 0 Then
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        End If
                    Next
                    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                Else
                    Dim sGetPartID As String = Get_PartID()
                    If sGetPartID <> "" Then
                        If rb_ProductPart.SelectedIndex = 0 Then
                            sqlStr += "AND Production_type in(" + sGetPartID + ") "
                        Else
                            sqlStr += "AND Part_ID in(" + sGetPartID + ") "
                        End If
                    End If
                End If
                sqlStr += "GROUP BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                sqlStr += "ORDER BY SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
            End If

            If ddlProduct.SelectedItem.Text = "PPS" Or ddlProduct.SelectedItem.Text = "PCB" Then
                If cb_Lot_Merge.Checked = False Then
                    sqlStr = sqlStr.Replace("VW_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
                End If
            End If


            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)
            lb_weekShow.Items.Clear()
            UtilObj.FillLitsBoxController(myDT, lb_weekSource, 1)
            conn.Close()
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
        tr_week.Visible = True

        'End If
        If rbl_week.Items(0).Selected = True Then
            tr_week.Visible = False
            lb_weekSource.Items.Clear()
            lb_weekShow.Items.Clear()
        End If
    End Sub

    ' Week To >>
    Protected Sub but_weekTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_weekTo.Click

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

        ListBoxSort(lb_weekSource)
        ListBoxSort(lb_weekShow)
    End Sub

    ' Week Back <<
    Protected Sub but_weekBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_weekBack.Click

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

        ListBoxSort(lb_weekSource)
        ListBoxSort(lb_weekShow)
    End Sub

    '' Yield Loss Item 
    'Protected Sub rbl_lossItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbl_lossItem.SelectedIndexChanged

    '    tr_chartDisplay.Visible = False
    '    tr_gvDisplay.Visible = False
    '    txb_ylInput.Text = ""
    '    lb_LossSource.Items.Clear()
    '    lb_LossShow.Items.Clear()
    '    tr_lossItem.Visible = False

    '    If rbl_lossItem.SelectedIndex = 5 Then

    '        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '        Dim sqlStr As String = ""
    '        Dim myDT As DataTable
    '        Dim myAdapter As SqlDataAdapter
    '        Dim sCustomer_ID As String = ""
    '        Try
    '            For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
    '                If n = 0 Then
    '                    sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
    '                Else
    '                    sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
    '                End If
    '            Next
    '            conn.Open()
    '            sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
    '            If ddlProduct.Text = "PPS" Then
    '                sqlStr += "and customer_id='All' "
    '            Else
    '                ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "

    '                If sCustomer_ID <> "" Then
    '                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
    '                End If


    '            End If
    '            Dim sGetPartID As String = Get_PartID()
    '            If sGetPartID <> "" Then
    '                sqlStr += "and part_id in(" + sGetPartID + ") "
    '            End If
    '            sqlStr += "group by fail_mode order by fail_mode"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)
    '            conn.Close()

    '        Catch ex As Exception

    '        Finally
    '            If conn.State = ConnectionState.Open Then
    '                conn.Close()
    '            End If
    '        End Try

    '        tr_lossItem.Visible = True
    '    End If
    'End Sub

    ' Yield Loss Item 
    Protected Sub rbl_lossItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbl_lossItem.SelectedIndexChanged

        tr_chartDisplay.Visible = False
        tr_gvDisplay.Visible = False
        txb_ylInput.Text = ""
        lb_LossSource.Items.Clear()
        lb_LossShow.Items.Clear()
        tr_lossItem.Visible = False

        If rbl_lossItem.SelectedIndex = 5 Then

            Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            Dim sqlStr As String = ""
            Dim myDT As DataTable
            Dim myAdapter As SqlDataAdapter
            Dim sCustomer_ID As String = ""
            Try
                For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
                    If n = 0 Then
                        sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
                    Else
                        sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
                    End If
                Next
                conn.Open()
                sqlStr = "select fail_mode from dbo.WB_BinCode_Daily_Lot where 1=1 "
                If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then
                    'sqlStr += "and customer_id='All' "
                Else
                    ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                    sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
                    If sCustomer_ID <> "" Then
                        sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                    End If


                End If
                Dim sGetPartID As String = Get_PartID()
                If sGetPartID <> "" Then
                    sqlStr += "and part_id in(" + sGetPartID + ") "
                End If
                sqlStr += "group by fail_mode order by fail_mode"


                If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" And cb_Lot_Merge.Checked = False Then
                    sqlStr = sqlStr.Replace("WB_BinCode_Daily_Lot", "WB_BinCode_Daily_Lot_NotMerge")
                End If


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
    Protected Sub but_lossItemTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_lossItemTo.Click

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

        ListBoxSort(lb_LossSource)
        ListBoxSort(lb_LossShow)
    End Sub

    ' Item Back <<
    Protected Sub but_lossItemBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_lossItemBack.Click

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

        ListBoxSort(lb_LossSource)
        ListBoxSort(lb_LossShow)
    End Sub


    Private Function Get_WB_TotalOriginal(ByVal sqlStr As String) As Integer
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim nValue As Integer = 0
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()


            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)

            If myDT.Rows.Count > 0 Then
                nValue = CInt(myDT.Rows(0).Item(0))
            End If


            Return nValue
        Catch ex As Exception

        Finally
            conn.Close()
        End Try

        Return Nothing
    End Function

    Private Function Get_BumpingType_PartID(ByVal sourceBumpingType As String) As String
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()

            ' --- Product Type (產品種類) ---
            If sourceBumpingType <> "All" Then
                sqlStr = "SELECT DISTINCT [Part_Id] FROM [EDA].[dbo].[Customer_Prodction_Mapping_BU_Rename] WHERE [Bumping_Type] in(" & sourceBumpingType & ")"
            Else
                Return Nothing
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)

            conn.Close()

            Dim oPartID As New StringBuilder
            For I As Integer = 0 To myDT.Rows.Count - 1
                If I = 0 Then
                    oPartID.Append("'" & myDT.Rows(I)("Part_Id").ToString() & "'")
                Else
                    oPartID.Append(",'" & myDT.Rows(I)("Part_Id").ToString() & "'")
                End If
            Next

            Return oPartID.ToString()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        Return Nothing
    End Function

    Private Function Get_OLProcess_PartID(ByVal sourceBumpingType As String) As String
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        Try
            conn.Open()

            ' --- Product Type (產品種類) ---
            If sourceBumpingType <> "All" Then
                ' sqlStr = "SELECT DISTINCT [Part_Id] FROM [EDA].[dbo].[Customer_Prodction_Mapping_BU_Rename] WHERE [Bumping_Type] in(" & sourceBumpingType & ")"
                sqlStr = "SELECT DISTINCT [Part_Id] FROM [EDA].[dbo].[Customer_Prodction_Mapping_BU_Rename] a,[MES].[dbo].[ProductInfo] b where a.Part_Id=b.Part_No and b.OL_Process in (" & sourceBumpingType & ")"



            Else
                Return Nothing
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myDT = New DataTable
            myAdapter.Fill(myDT)

            conn.Close()

            Dim oPartID As New StringBuilder
            For I As Integer = 0 To myDT.Rows.Count - 1
                If I = 0 Then
                    oPartID.Append("'" & myDT.Rows(I)("Part_Id").ToString() & "'")
                Else
                    oPartID.Append(",'" & myDT.Rows(I)("Part_Id").ToString() & "'")
                End If
            Next

            Return oPartID.ToString()
        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        Return Nothing
    End Function
    Private Function Get_PartID() As String
        'Dim oPartID As New StringBuilder
        'Try
        '    For I As Integer = 0 To ddlPart.Items.Count - 1
        '        If I = 0 Then
        '            oPartID.Append("'" & ddlPart.Items(I).Text & "'")
        '        Else
        '            oPartID.Append(",'" & ddlPart.Items(I).Text & "'")
        '        End If
        '    Next
        'Catch ex As Exception

        'End Try

        'Return oPartID.ToString()
        'Return "'" + ddlPart.Text + "'"

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

    Private Function getTotalOriginal_SQL_PCB(ByVal yl As YieldlossInfo) As String

        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        tempSQL = "select SUM(Original_Input_QTY) from " _
        & "( " _
        '& "select distinct Lot_Id ,Input_QTY Original_Input_QTY  " _
        '& "from "


        'tempSQL += "dbo.PCB_Yield_Daily_RawData_NotMerge "
        'tempSQL += "where 1=1  "

        'If yl.BumpingType <> "" And yl.Part_ID = "" Then
        '    tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        'ElseIf yl.Part_ID <> "" Then
        '    tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
        'End If

        ''FAI
        'If ckFAI.Checked = False Then
        '    tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
        'End If

        ''tempSQL += "and substring(Part_Id,7,1)<>'V' "
        'If Cb_CR.Checked = False Then
        '    tempSQL += "and substring(lot_id,9,1)<>'Y' "
        '    tempSQL += "and substring(lot_id,9,1)<>'Z' "
        '    tempSQL += "and substring(Part_id,7,1)<>'V' "
        'End If

        'If yl.TimePeriod = 0 Then
        '    tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in(" + ConvertStr2AddMark(yl.TimeRange) + ") "
        'ElseIf yl.TimePeriod = 1 Then
        '    tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        'Else

        '    Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
        '    Dim eDate As DateTime = sDate.AddMonths(1)

        '    tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "
        'End If

        'tempSQL += "union "
        tempSQL += "select Lot_Id ,Input_QTY Original_Input_QTY  "
        tempSQL += "from "


        tempSQL += "dbo.PCB_Yield_Daily_RawData_NotMerge "
        tempSQL += "where 1=1  "

        If yl.BumpingType <> "" And yl.Part_ID = "" Then
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        ElseIf yl.Part_ID <> "" Then
            tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
        End If

        'FAI
        If ckFAI.Checked = False Then
            tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
        End If

        'tempSQL += "and substring(Part_Id,7,1)<>'V' "
        If Cb_CR.Checked = False Then
            tempSQL += "and substring(lot_id,9,1)<>'Y' "
            tempSQL += "and substring(lot_id,9,1)<>'Z' "
            tempSQL += "and substring(Part_id,7,1)<>'V' "
        End If

        If yl.TimePeriod = 0 Then
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in(" + ConvertStr2AddMark(yl.TimeRange) + ") "
        ElseIf yl.TimePeriod = 1 Then
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        Else

            Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
            Dim eDate As DateTime = sDate.AddMonths(1)

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "
        End If


        tempSQL += " and yield_category='Total Yield' "

        tempSQL += ")a "


        If Cb_SF.Checked = True Then
            tempSQL = tempSQL.Replace("PCB_Yield_Daily_RawData_NotMerge", "vw_PCB_Yield_Daily_RawData_NotMerge_SF")
            tempSQL = tempSQL.Replace("PCB_Yield_Daily_RawData_NotMerge_Storage", "vw_PCB_Yield_Daily_RawData_NotMerge_Storage_SF")
        End If

        Return tempSQL
    End Function


    Private Function getTotalOriginal_SQL(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""


        If ddlProduct.SelectedValue = "PCB" Then
            tempSQL = getTotalOriginal_SQL_PCB(yl)
        Else






            tempSQL = "select SUM(Original_Input_QTY) from " _
            & "( " _
            & "select distinct Lot_Id ,Original_Input_QTY  " _
            & "from "


            If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                If cb_Lot_Merge.Checked = True Then
                    tempSQL += "dbo.VW_BinCode_Daily_Lot "
                Else
                    tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
                End If

            Else
                tempSQL += "dbo.VW_BinCode_Daily_Lot "
            End If

            tempSQL += "where 1=1  "


            If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                If yl.BumpingType <> "" And yl.Part_ID = "" Then
                    tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
                ElseIf yl.Part_ID <> "" Then
                    tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
                End If
            Else
                If rb_ProductPart.SelectedIndex = 0 Then
                    tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
                Else
                    tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
                End If

            End If


            tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
            & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  "


            If cb_Non8K.Checked = True Then
                tempSQL += "and Fail_Mode NOT LIKE '8K%'  "
            End If

            'FAI
            If ckFAI.Checked = False Then
                tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
            End If

            'tempSQL += "and substring(Part_Id,7,1)<>'V' "
            If Cb_CR.Checked = False Then
                tempSQL += "and substring(lot_id,9,1)<>'Y' "
                tempSQL += "and substring(lot_id,9,1)<>'Z' "
                tempSQL += "and substring(Part_id,7,1)<>'V' "
            End If



            If yl.TimePeriod = 0 Then
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in(" + ConvertStr2AddMark(yl.TimeRange) + ") "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
            Else

                Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
                Dim eDate As DateTime = sDate.AddMonths(1)

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "
            End If

            tempSQL += ")a "


            If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
                tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
            End If
        End If
        Return tempSQL
    End Function

    Private Function getRawData_SQL(ByVal yl As YieldlossInfo) As String

       

        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        tempSQL = "select *  " _
        & "from "


        If cb_Lot_Merge.Checked = False And (ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB") Then

            tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
        Else
            tempSQL += "dbo.VW_BinCode_Daily_Lot "
        End If

        'If cb_Lot_Merge.Checked = True Or ddlProduct.SelectedValue <> "PPS" Or ddlProduct.SelectedValue <> "PCB" Then
        '    tempSQL += "dbo.VW_BinCode_Daily_Lot "
        'Else
        '    tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
        'End If

        tempSQL += "where 1=1  "

        If yl.BumpingType <> "" And yl.Part_ID = "" Then
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        ElseIf yl.Part_ID <> "" Then
            tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
        End If

        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
        & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
        & "and Fail_Mode NOT LIKE '8K%'  "

        If yl.TimePeriod = 0 Then
            'tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in (" + ConvertStr2AddMark(yl.TimeRange) + " )"
        ElseIf yl.TimePeriod = 1 Then
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        Else

            Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
            Dim eDate As DateTime = sDate.AddMonths(1)

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "
        End If

        'FAI
        If ckFAI.Checked = False Then
            tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
        End If

        If Cb_CR.Checked = False Then
            tempSQL += "and substring(lot_id,9,1)<>'Y' "
            tempSQL += "and substring(lot_id,9,1)<>'Z' "
            tempSQL += "and substring(Part_id,7,1)<>'V' "
        End If



        If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
        End If

        If Cb_Inline.Checked = True Then
            tempSQL += " and MF_Stage='INLINE'"
        End If
        Return tempSQL
    End Function

    Public Function ConvertStr2AddMark(ByVal temp As String) As String
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
    Private Function getTopWBSQL(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        Dim sDateTime() As String = yl.TimeRange.Split(",")
        'yl.TimeRange = sDateTime(sDateTime.Length - 1)
        yl.TimeRange = sDateTime(0)
        yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))




        If ddlProduct.SelectedValue = "PPS" Then
            tempSQL = "select top(" + yl.nTop.ToString + ") Fail_Mode, Fail_Mode AS 'newFailMode', "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap) AS Fail_Count"
            Else
                tempSQL += "SUM(Fail_Count) AS Fail_Count"
            End If
        Else
            tempSQL = "select top(" + yl.nTop.ToString + ") Fail_Mode, Fail_Mode AS 'newFailMode', "

            tempSQL += "SUM(Fail_Count) AS Fail_Count"
        End If



        tempSQL += ", " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
        & "round((convert(float, "

        If ddlProduct.SelectedValue = "PPS" Then
            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If
        Else
            tempSQL += "SUM(Fail_Count)"
        End If


        tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "


        If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
            If cb_Lot_Merge.Checked = True Then
                tempSQL += "dbo.VW_BinCode_Daily_Lot "
            Else
                tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
            End If

        Else
            tempSQL += "dbo.VW_BinCode_Daily_Lot "
        End If


        tempSQL += "where 1=1  "


        If yl.BumpingType <> "" And yl.Part_ID = "" Then
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        ElseIf yl.Part_ID <> "" Then
            If rb_ProductPart.SelectedIndex = 0 Then
                tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            Else
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            End If

        End If


        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
        & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
        & "and Fail_Mode <>'' "


        If cb_Non8K.Checked = True Then
            tempSQL += "and Fail_Mode NOT LIKE '8K%'  "
        End If

        If cb_NonIPQC.Checked = True Then
            tempSQL += "AND Fail_Mode NOT LIKE 'IPQC%' "
        End If

        'FAI
        If ckFAI.Checked = False Then
            tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
        End If

        If Cb_CR.Checked = False Then
            tempSQL += "and substring(lot_id,9,1)<>'Y' "
            tempSQL += "and substring(lot_id,9,1)<>'Z' "
            tempSQL += "and substring(Part_id,7,1)<>'V' "
        End If



        If ddlProduct.SelectedValue = "WB" Then
            If yl.xoutscrape = True Then
                tempSQL += "and Fail_Mode <>'匹配報廢' "
            End If

        End If




        If yl.TimePeriod = 0 Then
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
        ElseIf yl.TimePeriod = 1 Then
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        Else
            Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
            Dim eDate As DateTime = sDate.AddMonths(1)

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

        End If

        If yl.Fail_Mode <> "" Then
            'tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
            tempSQL += "and Fail_Mode IN(" + yl.Fail_Mode + ") "
        End If

        If Cb_Inline.Checked = True Then
            tempSQL += "and MF_Stage='INLINE' "
        End If

        'tempSQL += "group by Fail_Mode order by round((convert(float, "
        If ddlProduct.SelectedValue = "PPS" Then
            tempSQL += "group by Fail_Mode order by  "
        Else
            tempSQL += "group by Fail_Mode order by  "
        End If


        If yl.xoutscrape = True Then
            tempSQL += "SUM(Fail_Count_ByXoutScrap)"
        Else
            tempSQL += "SUM(Fail_Count)"
        End If


        tempSQL += " DESC, Fail_Mode "

        If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
        End If


        If Cb_Inline.Checked = True Then
            tempSQL = tempSQL.Replace("Fail_Mode, Fail_Mode AS 'newFailMode'", "BinCode as Fail_Mode, BinCode AS 'newFailMode'")
            tempSQL = tempSQL.Replace("group by Fail_Mode order by  SUM(Fail_Count) DESC, Fail_Mode", "group by BinCode order by  SUM(Fail_Count) DESC, BinCode")
        End If


        Return tempSQL
    End Function

    Private Function getTopPCBSQL(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""

        Dim sDateTime() As String = yl.TimeRange.Split(",")
        yl.TimeRange = sDateTime(sDateTime.Length - 1)
        yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))




        'If ddlProduct.SelectedValue = "PPS" Then
        '    tempSQL = "select top(" + yl.nTop.ToString + ") Fail_Mode, Fail_Mode AS 'newFailMode', "

        '    If yl.xoutscrape = True Then
        '        tempSQL += "SUM(Fail_Count_ByXoutScrap) AS Fail_Count"
        '    Else
        '        tempSQL += "SUM(Fail_Count) AS Fail_Count"
        '    End If
        'Else
        tempSQL = "select top(" + yl.nTop.ToString + ") Fail_Mode, Fail_Mode AS 'newFailMode', "

        tempSQL += "SUM(Fail_Count) AS Fail_Count"
        'End If



        tempSQL += ", " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
        & "round((convert(float, "

        'If ddlProduct.SelectedValue = "PPS" Then
        '    If yl.xoutscrape = True Then
        '        tempSQL += "SUM(Fail_Count_ByXoutScrap)"
        '    Else
        '        tempSQL += "SUM(Fail_Count)"
        '    End If
        'Else
        tempSQL += "SUM(Fail_Count)"
        'End If


        tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "


        If ddlProduct.SelectedValue = "PCB" Then
            If cb_Lot_Merge.Checked = True Then
                tempSQL += "dbo.VW_BinCode_Daily_Lot "
            Else
                tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
            End If

        Else
            tempSQL += "dbo.VW_BinCode_Daily_Lot "
        End If


        tempSQL += "where 1=1  "


        If yl.BumpingType <> "" And yl.Part_ID = "" Then
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        ElseIf yl.Part_ID <> "" Then
            If rb_ProductPart.SelectedIndex = 0 Then
                tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            Else
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            End If

        End If


        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
        & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
        & "and Fail_Mode <>'' "


        If cb_Non8K.Checked = True Then
            tempSQL += "and Fail_Mode NOT LIKE '8K%'  "
        End If

        If cb_NonIPQC.Checked = True Then
            tempSQL += "AND Fail_Mode NOT LIKE 'IPQC%' "
        End If

        If Cb_CR.Checked = False Then
            tempSQL += "and substring(lot_id,9,1)<>'Y' "
            tempSQL += "and substring(lot_id,9,1)<>'Z' "
            tempSQL += "and substring(Part_id,7,1)<>'V' "
        End If



        'If ddlProduct.SelectedValue = "WB" Then
        '    If yl.xoutscrape = True Then
        '        tempSQL += "and Fail_Mode <>'匹配報廢' "
        '    End If

        'End If




        If yl.TimePeriod = 0 Then
            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
        ElseIf yl.TimePeriod = 1 Then
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        Else
            Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
            Dim eDate As DateTime = sDate.AddMonths(1)

            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

        End If

        If yl.Fail_Mode <> "" Then
            'tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
            tempSQL += "and Fail_Mode IN(" + yl.Fail_Mode + ") "
        End If

        'tempSQL += "group by Fail_Mode order by round((convert(float, "
        If ddlProduct.SelectedValue = "PPS" Then
            tempSQL += "group by Fail_Mode order by  "
        Else
            tempSQL += "group by Fail_Mode order by  "
        End If


        'If yl.xoutscrape = True Then
        '    tempSQL += "SUM(Fail_Count_ByXoutScrap)"
        'Else
        tempSQL += "SUM(Fail_Count)"
        'End If


        tempSQL += " DESC, Fail_Mode "

        'If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
        '    tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
        'End If

        Return tempSQL
    End Function

    Private Function getRowDataWBSQL(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        tempSQL = "select Datatime,Fail_Mode,'' as BumpingType " _
        & ",(select TOP 1 MF_Stage from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as MF_Stage " _
        & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE MF_Area END from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as MF_Area " _
        & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE DefectCode_id END from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as DefectCode " _
        & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
        & "from (select top(10) ww as Datatime,Fail_Mode, SUM(Fail_Count) AS Fail_Count, " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
        & "round((convert(float, SUM(Fail_Count))/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from dbo.VW_BinCode_Daily_Lot where 1=1  "

        If yl.BumpingType <> "" And yl.Part_ID = "" Then
            tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
        ElseIf yl.Part_ID <> "" Then
            tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
        End If

        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
        & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
        & "and Fail_Mode NOT LIKE '8K%'  " _
        & "and Fail_Mode <>'' "

        If yl.TimePeriod = 0 Then

        ElseIf yl.TimePeriod = 1 Then
            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
        Else

        End If

        tempSQL += "group by ww,Fail_Mode " _
        & "order by round((convert(float, SUM(Fail_Count))/" + yl.TotalOriginal.ToString + "), 5) * 100 DESC )a  "



        Return tempSQL
    End Function

    Private Function getRowDataWBSQL2(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        Dim sDateTime() As String = yl.TimeRange.Split(",")
        Dim descBumpingType As String = yl.BumpingType.Replace("'", "")
        Dim nTop2 As Integer = yl.nTop + 100

        tempSQL = "select * from ( " '<-最外層
        For i As Integer = 0 To sDateTime.Length - 1
            yl.TimeRange = sDateTime(i)
            yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))

            If i > 0 Then
                tempSQL += "union all "
            End If

            If Cb_Inline.Checked = True Then
                tempSQL += "select Datatime,Fail_Mode,'" + descBumpingType + "' as BumpingType " _
                              & ",'INLINE' as MF_Stage  " _
                              & ",'FE' as MF_Area  " _
                              & ",(select TOP 1  DefectCode_id  from view_BinCode where category='WB' and a.Fail_Mode=BinCode ) as DefectCode " _
                              & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
                              & "from ("
            Else
                tempSQL += "select Datatime,Fail_Mode,'" + descBumpingType + "' as BumpingType " _
                              & ",(select TOP 1 MF_Stage from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as MF_Stage " _
                              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE MF_Area END from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as MF_Area " _
                              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE DefectCode_id END from view_BinCode where category='WB' and a.Fail_Mode=failmode ) as DefectCode " _
                              & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
                              & "from ("
            End If


           
            If i = 0 Then
                If yl.TimePeriod = 0 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                ElseIf yl.TimePeriod = 1 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") ww as Datatime "
                Else
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                End If
            Else
                If yl.TimePeriod = 0 Then
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                ElseIf yl.TimePeriod = 1 Then
                    tempSQL += "select top(" + nTop2.ToString + ") ww as Datatime "
                Else
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                End If
            End If

            tempSQL += ",Fail_Mode, "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If


            tempSQL += " AS Fail_Count, " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
              & "round((convert(float, "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If

            tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "


            If cb_Lot_Merge.Checked = True Then
                tempSQL += "dbo.VW_BinCode_Daily_Lot "
            Else
                tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
            End If

            tempSQL += "where 1=1  "

            If yl.BumpingType <> "" And yl.Part_ID = "" Then
                tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
            ElseIf yl.Part_ID <> "" Then
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            End If

            tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
            & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
            & "and Fail_Mode <>'' "

            If cb_Non8K.Checked = True Then
                tempSQL += "and Fail_Mode NOT LIKE '8K%'  "
            End If
            'FAI
            If ckFAI.Checked = False Then
                tempSQL += "and ISNUMERIC(SUBSTRING(Part_id,2,1))=0 "
            End If

            If Cb_CR.Checked = False Then
                tempSQL += "and substring(lot_id,9,1)<>'Y' "
                tempSQL += "and substring(lot_id,9,1)<>'Z' "
                tempSQL += "and substring(Part_id,7,1)<>'V' "
            End If

            If yl.xoutscrape = True Then
                tempSQL += "and Fail_Mode <>'匹配報廢' "
            End If

            If yl.TimePeriod = 0 Then
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
            Else
                Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
                Dim eDate As DateTime = sDate.AddMonths(1)

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

            End If

            If yl.Fail_Mode <> "" Then
                'tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
                tempSQL += "and Fail_Mode IN(" + yl.Fail_Mode + ") "
            End If

            If Cb_Inline.Checked = True Then
                tempSQL += "and MF_Stage='INLINE' "
            End If

            If yl.TimePeriod = 0 Then
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ,Fail_Mode  "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "group by ww,Fail_Mode "
            Else
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ,Fail_Mode  "
            End If

            'tempSQL += "order by round((convert(float, "
            tempSQL += "order by  "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If

            'tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 DESC )a  "
            tempSQL += " DESC, Fail_Mode )a  "


        Next
        'tempSQL += ")c  order by Datatime desc,Fail_Ratio desc " '< -最外層
        tempSQL += ")c  order by Datatime desc,Fail_Count desc , Fail_Mode" '< -最外層


        If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
        End If

        If Cb_Inline.Checked = True Then
            tempSQL = tempSQL.Replace("Fail_Mode, SUM(Fail_Count) AS Fail_Count", "BinCode as Fail_Mode, SUM(Fail_Count) AS Fail_Count")
            tempSQL = tempSQL.Replace("group by ww,Fail_Mode order by  SUM(Fail_Count) DESC, Fail_Mode", "group by ww,BinCode order by  SUM(Fail_Count) DESC, BinCode")
        End If

        Return tempSQL
    End Function

    Private Function getRowDataPCBSQL2(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        Dim sDateTime() As String = yl.TimeRange.Split(",")
        Dim descBumpingType As String = yl.BumpingType.Replace("'", "")
        Dim nTop2 As Integer = yl.nTop + 100
        'yl.nTop = yl.nTop + 10


        tempSQL = "select * from ( " '<-最外層
        For i As Integer = 0 To sDateTime.Length - 1
            yl.TimeRange = sDateTime(i)
            yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))

            If i > 0 Then
                tempSQL += "union all "
            End If

            tempSQL += "select Datatime,Fail_Mode,'" + descBumpingType + "' as BumpingType " _
              & ",(select TOP 1 MF_Stage from view_BinCode where category='PCB' and a.Fail_Mode=failmode ) as MF_Stage " _
              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE MF_Area END from view_BinCode where category='PCB' and a.Fail_Mode=failmode ) as MF_Area " _
              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE DefectCode_id END from view_BinCode where category='PCB' and a.Fail_Mode=failmode ) as DefectCode " _
              & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
              & "from ("

            If i = 0 Then
                If yl.TimePeriod = 0 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                ElseIf yl.TimePeriod = 1 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") ww as Datatime "
                Else
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                End If
            Else
                If yl.TimePeriod = 0 Then
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                ElseIf yl.TimePeriod = 1 Then
                    tempSQL += "select top(" + nTop2.ToString + ") ww as Datatime "
                Else
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                End If
            End If





            tempSQL += ",Fail_Mode, "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If


            tempSQL += " AS Fail_Count, " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
              & "round((convert(float, "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If

            tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "


            If cb_Lot_Merge.Checked = True Then
                tempSQL += "dbo.VW_BinCode_Daily_Lot "
            Else
                tempSQL += "dbo.WB_BinCode_Daily_Lot_NotMerge "
            End If

            tempSQL += "where 1=1  "

            If yl.BumpingType <> "" And yl.Part_ID = "" Then
                tempSQL += "AND BumpingType in (" + yl.BumpingType + ")  "
            ElseIf yl.Part_ID <> "" Then
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            End If

            tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
            & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
            & "and Fail_Mode <>'' "

            If cb_Non8K.Checked = True Then
                tempSQL += "and Fail_Mode NOT LIKE '8K%'  "
            End If

            If Cb_CR.Checked = False Then
                tempSQL += "and substring(lot_id,9,1)<>'Y' "
                tempSQL += "and substring(lot_id,9,1)<>'Z' "
                tempSQL += "and substring(Part_id,7,1)<>'V' "
            End If

            If yl.xoutscrape = True Then
                tempSQL += "and Fail_Mode <>'匹配報廢' "
            End If

            If yl.TimePeriod = 0 Then
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
            Else
                Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
                Dim eDate As DateTime = sDate.AddMonths(1)

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

            End If

            If yl.Fail_Mode <> "" Then
                'tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
                tempSQL += "and Fail_Mode IN(" + yl.Fail_Mode + ") "
            End If


            If yl.TimePeriod = 0 Then
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ,Fail_Mode  "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "group by ww,Fail_Mode "
            Else
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ,Fail_Mode  "
            End If

            'tempSQL += "order by round((convert(float, "
            tempSQL += "order by  "

            If yl.xoutscrape = True Then
                tempSQL += "SUM(Fail_Count_ByXoutScrap)"
            Else
                tempSQL += "SUM(Fail_Count)"
            End If

            'tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 DESC )a  "
            tempSQL += " DESC , Fail_Mode)a  "


        Next
        'tempSQL += ")c  order by Datatime desc,Fail_Ratio desc " '< -最外層
        tempSQL += ")c  order by Datatime desc,Fail_Count desc , Fail_Mode " '< -最外層


        If Cb_SF.Checked = True And cb_Lot_Merge.Checked = False Then
            tempSQL = tempSQL.Replace("WB_BinCode_Daily_Lot_NotMerge", "vw_WB_BinCode_Daily_Lot_NotMerge_SF")
        End If

        Return tempSQL
    End Function

    Private Function getRowDataFCSQL2(ByVal yl As YieldlossInfo) As String
        Dim tempReplace As String = ""
        Dim tempSQL As String = ""
        Dim sDateTime() As String = yl.TimeRange.Split(",")
        Dim nTop2 As Integer = yl.nTop + 100

        tempSQL = "select * from ( " '<-最外層
        For i As Integer = 0 To sDateTime.Length - 1
            yl.TimeRange = sDateTime(i)
            yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))

            If i > 0 Then
                tempSQL += "union all "
            End If

            tempSQL += "select Datatime,Fail_Mode " _
              & ",(select TOP 1 MF_Stage from view_BinCode where category='" + ddlProduct.SelectedValue + "' and a.Fail_Mode=failmode ) as MF_Stage " _
              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE MF_Area END from view_BinCode where category='" + ddlProduct.SelectedValue + "' and a.Fail_Mode=failmode ) as MF_Area " _
              & ",(select TOP 1 CASE WHEN SUBSTRING(Fail_Mode,1,6)='Inline' THEN '' ELSE DefectCode_id END from view_BinCode where category='" + ddlProduct.SelectedValue + "' and a.Fail_Mode=failmode ) as DefectCode " _
              & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
              & "from ("

            If yl.TimePeriod = 0 Then
                If i = 0 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                Else
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
                End If

            ElseIf yl.TimePeriod = 1 Then
                If i = 0 Then
                    tempSQL += "select top(" + nTop2.ToString + ") ww as Datatime "

                Else
                    tempSQL += "select top(" + nTop2.ToString + ") ww as Datatime "
                End If

            Else
                If i = 0 Then
                    tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                Else
                    tempSQL += "select top(" + nTop2.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
                End If

            End If


            tempSQL += ",Fail_Mode, "


            tempSQL += "SUM(Fail_Count)"


            tempSQL += " AS Fail_Count, " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
              & "round((convert(float, "


            tempSQL += "SUM(Fail_Count)"


            tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "



            tempSQL += "dbo.VW_BinCode_Daily_Lot "


            tempSQL += "where 1=1  "



            If rb_ProductPart.SelectedIndex = 0 Then
                tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            Else
                tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
            End If



            tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
            & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
            & "and Fail_Mode NOT LIKE '8K%'  " _
            & "and Fail_Mode <>'' "

            If cb_NonIPQC.Checked = True Then
                'tempSQL += "AND Fail_Mode NOT LIKE 'IPQC%' "
                tempSQL += "AND Fail_Mode NOT LIKE 'IPQC defect%' "

            End If



            If yl.TimePeriod = 0 Then
                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
            Else
                Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
                Dim eDate As DateTime = sDate.AddMonths(1)

                tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

            End If

            If yl.Fail_Mode <> "" Then
                'tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
                tempSQL += "and Fail_Mode IN(" + yl.Fail_Mode + ") "
            End If


            If yl.TimePeriod = 0 Then
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ,Fail_Mode  "
            ElseIf yl.TimePeriod = 1 Then
                tempSQL += "group by ww,Fail_Mode "
            Else
                tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ,Fail_Mode  "
            End If

            tempSQL += "order by SUM(Fail_Count) DESC )a  "

        Next

        tempSQL += ")c  order by Datatime desc,fail_count desc " '< -最外層

        Return tempSQL
    End Function
    'Private Function getRowDataFCSQL2(ByVal yl As YieldlossInfo) As String
    '    Dim tempReplace As String = ""
    '    Dim tempSQL As String = ""
    '    Dim sDateTime() As String = yl.TimeRange.Split(",")


    '    tempSQL = "select * from ( " '<-最外層
    '    For i As Integer = 0 To sDateTime.Length - 1
    '        yl.TimeRange = sDateTime(i)
    '        yl.TotalOriginal = Get_WB_TotalOriginal(getTotalOriginal_SQL(yl))

    '        If i > 0 Then
    '            tempSQL += "union all "
    '        End If

    '        tempSQL += "select Datatime,Fail_Mode, MF_Stage,MF_Area ,DefectCode " _
    '          & ",Fail_Mode AS 'newFailMode',Fail_Count,Original_Input_QTY,Fail_Ratio " _
    '          & "from ("

    '        If yl.TimePeriod = 0 Then
    '            tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) as Datatime "
    '        ElseIf yl.TimePeriod = 1 Then
    '            tempSQL += "select top(" + yl.nTop.ToString + ") ww as Datatime "
    '        Else
    '            tempSQL += "select top(" + yl.nTop.ToString + ") SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) as Datatime "
    '        End If


    '        tempSQL += ",Fail_Mode,MF_Stage,MF_Area ,DefectCode, "


    '        tempSQL += "SUM(Fail_Count)"


    '        tempSQL += " AS Fail_Count, " + yl.TotalOriginal.ToString + " AS Original_Input_QTY, " _
    '          & "round((convert(float, "


    '        tempSQL += "SUM(Fail_Count)"


    '        tempSQL += ")/" + yl.TotalOriginal.ToString + "), 5) * 100 as Fail_Ratio from "



    '        tempSQL += "dbo.VW_BinCode_Daily_Lot "


    '        tempSQL += "where 1=1  "


    '        If rb_ProductPart.SelectedIndex = 0 Then
    '            tempSQL += "AND production_type in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
    '        Else
    '            tempSQL += "AND Part_Id in (" + ConvertStr2AddMark(yl.Part_ID) + ")  "
    '        End If



    '        tempSQL += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)  " _
    '        & "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)  " _
    '        & "and Fail_Mode NOT LIKE '8K%'  " _
    '        & "and Fail_Mode <>'' "

    '        If yl.TimePeriod = 0 Then
    '            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) =" + ConvertStr2AddMark(yl.TimeRange) + " "
    '        ElseIf yl.TimePeriod = 1 Then
    '            tempSQL += "and WW in (" + ConvertStr2AddMark(yl.TimeRange) + ") "
    '        Else
    '            Dim sDate As DateTime = Date.Parse(Left(yl.TimeRange.ToString, 4) + "/" + Right(yl.TimeRange.ToString, 2) + "/01")
    '            Dim eDate As DateTime = sDate.AddMonths(1)

    '            tempSQL += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) >=" + ConvertStr2AddMark(sDate.ToString("yyyyMMdd")) + " and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) <" + ConvertStr2AddMark(eDate.ToString("yyyyMMdd")) + " "

    '        End If

    '        If yl.Fail_Mode <> "" Then
    '            tempSQL += "and Fail_Mode IN(" + ConvertStr2AddMark(yl.Fail_Mode) + ") "
    '        End If


    '        If yl.TimePeriod = 0 Then
    '            tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) ,Fail_Mode,MF_Stage,DefectCode,MF_Area   "
    '        ElseIf yl.TimePeriod = 1 Then
    '            tempSQL += "group by ww,Fail_Mode,MF_Stage,DefectCode,MF_Area  "
    '        Else
    '            tempSQL += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) ,Fail_Mode,MF_Stage,DefectCode,MF_Area   "
    '        End If

    '        tempSQL += "order by SUM(Fail_Count) DESC )a  "

    '    Next

    '    tempSQL += ")c  order by Datatime desc,fail_count desc " '< -最外層

    '    Return tempSQL
    'End Function
    Private Function GetYearWW(ByVal td As DateTime) As String
        Dim dt As DataTable
        Dim dtAOI As DataTable
        Dim nValue As String = ""
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Try
            Dim tempSQL As String = ""
            tempSQL = "select top 4 yearww from SystemDateMapping where datetime<='" & td.ToString("yyyy/MM/dd") & "' group by yearww order by yearww desc "

            myAdpt = New SqlDataAdapter(tempSQL, conn)
            dt = New DataTable
            myAdpt.Fill(dt)
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    If i = 0 Then
                        nValue = dt.Rows(i).Item("yearWW").ToString
                    Else
                        nValue += "," + dt.Rows(i).Item("yearWW").ToString
                    End If

                Next
                'For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                '    If i = dt.Rows.Count - 1 Then
                '        nValue = dt.Rows(i).Item("yearWW").ToString
                '    Else
                '        nValue += "," + dt.Rows(i).Item("yearWW").ToString
                '    End If

                'Next

            End If


        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally

        End Try

        Return nValue
    End Function

    Private Function GetYearDay(ByVal td As DateTime) As String
        Dim dt As DataTable
        Dim dtAOI As DataTable
        Dim nValue As String = ""
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Try
            Dim tempSQL As String = ""
            tempSQL = "select top 4 datetime from SystemDateMapping where datetime<'" & td.ToString("yyyy/MM/dd") & "' group by datetime order by datetime desc "

            myAdpt = New SqlDataAdapter(tempSQL, conn)
            dt = New DataTable
            myAdpt.Fill(dt)
            If dt.Rows.Count > 0 Then
                For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                    If i = dt.Rows.Count - 1 Then
                        nValue = CDate(dt.Rows(i).Item("datetime")).ToString("yyyyMMdd")
                    Else
                        nValue += "," + CDate(dt.Rows(i).Item("datetime")).ToString("yyyyMMdd")
                    End If

                Next


            End If


        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally

        End Try

        Return nValue
    End Function

    Private Function GetYearMM(ByVal td As DateTime) As String
        Dim dt As DataTable
        Dim dtAOI As DataTable
        Dim nValue As String = ""
        Dim myAdpt As SqlDataAdapter
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Try
            Dim tempSQL As String = ""
            tempSQL = "select top 4  SUBSTRING(CONVERT(VARCHAR, datetime, 112), 1, 6) as Datatime  from SystemDateMapping  where datetime<='" & td.ToString("yyyy/MM/dd") & "' group by SUBSTRING(CONVERT(VARCHAR, datetime, 112), 1, 6) order by Datatime desc "

            myAdpt = New SqlDataAdapter(tempSQL, conn)
            dt = New DataTable
            myAdpt.Fill(dt)
            If dt.Rows.Count > 0 Then
                For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                    If i = dt.Rows.Count - 1 Then
                        nValue = dt.Rows(i).Item("Datatime").ToString
                    Else
                        nValue += "," + dt.Rows(i).Item("Datatime").ToString
                    End If

                Next


            End If


        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally

        End Try

        Return nValue
    End Function



    Private Sub daily_failMode()
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
        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        Try
            ' --- Bumping Type ---
            Dim strBumpingType As String = ""
            For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                If n = 0 Then
                    strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                Else
                    strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                End If
            Next

            Dim BumpingType_Part As String = ""
            If strBumpingType <> "" Then
                BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
            End If

            ' --- Part ID ---
            Dim sGetPartID As String = Get_PartID()
            If tr_BumpingType.Visible = False Then
                If rb_ProductPart.SelectedIndex = 0 Then
                    If listB_BumpingTypeShow.Enabled = False Then
                        If sGetPartID <> "" Then
                            partStr += "and Production_id={0} "
                        Else
                            partStr += "{0}"
                        End If
                    Else
                        If sGetPartID <> "" Then
                            partStr += "and Production_id={0} "
                        Else
                            partStr += "{0}"
                        End If
                    End If
                ElseIf rb_ProductPart.SelectedIndex = 1 Then
                    Dim oPartID = sGetPartID.Split(",")
                    If oPartID.Length = 1 Then
                        If listB_BumpingTypeShow.Enabled = False Then
                            If sGetPartID <> "" Then
                                partStr += "and Part_Id = {0} "
                            Else
                                partStr += "{0}"
                            End If
                        Else
                            If BumpingType_Part <> "" Then
                                partStr += "and Part_Id = {0} "
                            Else
                                partStr += "{0}"
                            End If
                        End If
                    ElseIf oPartID.Length > 1 Then
                        If listB_BumpingTypeShow.Enabled = False Then
                            If sGetPartID <> "" Then
                                partStr += "and Part_Id in({0}) "
                            Else
                                partStr += "{0}"
                            End If
                        Else
                            If BumpingType_Part <> "" Then
                                partStr += "and Part_Id in({0}) "
                            Else
                                partStr += "{0}"
                            End If
                        End If
                    End If
                End If
            Else
                If BumpingType_Part <> "" Then
                    partStr += "and Part_Id in (" & BumpingType_Part & ") "
                Else
                    Dim oPartID = sGetPartID.Split(",")
                    If oPartID.Length = 1 Then
                        If sGetPartID <> "" Then
                            partStr += "and Part_Id = " + sGetPartID
                        End If
                    ElseIf oPartID.Length > 1 Then
                        If sGetPartID <> "" Then
                            partStr += "and Part_Id in(" + sGetPartID + ") "
                        End If
                    End If
                End If
            End If

            If ddlProduct.SelectedValue <> "CPU" Then
                If rb_dayType.SelectedIndex <> 0 Then
                    plantStr = "and Plant='All' "
                End If
            Else
                plantStr = ""
            End If

            ' --- DateTime ---
            Dim dateTimeTemp As String = ""
            If rbl_week.SelectedIndex = 1 Then
                For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                    dateTimeTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
                Next
                dateTimeTemp = dateTimeTemp.Substring(0, (dateTimeTemp.Length - 1))
            End If

            If rbl_lossItem.SelectedIndex = 0 Then
                topStr = "top(10)"
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                topStr = "top(20)"
            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                topStr = "top(30)"
            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                topStr = "top(40)"
            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                topStr = "top(50)"
            End If

            conn.Open()
            ' --- Yield Loss ID ---
            Dim itemTemp As String = ""
            itemStr = ""

            For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                itemTemp += ((lb_LossShow.Items(i).Value).Replace("'", "''")) + ","
            Next
            If (itemTemp.Length > 0 AndAlso lb_LossShow.Items.Count > 0) Then
                itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
                itemStr = "and fail_mode in (" + itemTemp + ") "
            End If

            If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                sqlStr = "select Fail_Mode, Fail_Mode AS 'newFailMode', "
            Else
                sqlStr = "select " + topStr + " Fail_Mode, Fail_Mode AS 'newFailMode', "
            End If
            If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                'Fail_Count_byXoutScrap
                sqlStr += "round((convert(float, SUM(Fail_Count_byXoutScrap))/SUM(Original_Input_QTY))/COUNT(Original_Input_QTY), 4) * 100 as Fail_Ratio "
            Else
                sqlStr += "round((convert(float, SUM(Fail_Count))/SUM(Original_Input_QTY))/COUNT(Original_Input_QTY), 4) * 100 as Fail_Ratio "
            End If
            sqlStr += "from dbo.VW_BinCode_Daily_Lot where 1=1 "

            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
            sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
            sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

            If cb_NonIPQC.Checked Then
                sqlStr += "and Fail_Mode <> 'IPQC defect' "
            End If
            If cb_Non8K.Checked Then
                sqlStr += "and Fail_Mode NOT LIKE '8K%' "
            End If
            'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
            '    sqlStr += partStr
            'Else
            '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
            'End If
            If tr_BumpingType.Visible = False Then
                sqlStr += partStr
            Else
                If listB_PartShow.Items.Count = 0 Then
                    If strBumpingType <> "" Then
                        sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    End If
                Else
                    If strBumpingType <> "" Then
                        sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    End If
                    If sGetPartID <> "" Then
                        sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                    End If
                End If
            End If
            If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                sqlStr += itemStr
            End If
            sqlStr += customStr
            sqlStr += plantStr
            If rbl_week.SelectedIndex = 0 Then
                ' Defaule
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) = ("
                    sqlStr += "select top (1) SUBSTRING(CONVERT(VARCHAR, dateadd(day, -1, getdate()), 112), 1, 8) from VW_BinCode_Daily_Lot"
                    sqlStr += ") "
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) = ("
                    sqlStr += "select top (1) SUBSTRING(CONVERT(VARCHAR, dateadd(day, -1, getdate()), 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, dateadd(day, -1, getdate())) AS NVARCHAR), 2) from VW_BinCode_Daily_Lot"
                    sqlStr += ") "
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) = ("
                    sqlStr += "select top (1) SUBSTRING(CONVERT(VARCHAR, dateadd(day, -1, getdate()), 112), 1, 6) from VW_BinCode_Daily_Lot"
                    sqlStr += ") "
                End If
            Else
                ' Custom
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) = ("
                    sqlStr += "select Top(1) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime "
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) = ("
                    sqlStr += "select Top(1) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datatime "
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) = ("
                    sqlStr += "select Top(1) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datatime "
                End If
                sqlStr += "from VW_BinCode_Daily_Lot "
                sqlStr += "where 1=1 "
                sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                End If
                'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                '    sqlStr += partStr
                'Else
                '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += partStr
                Else
                    If listB_PartShow.Items.Count = 0 Then
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                    Else
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                        If sGetPartID <> "" Then
                            sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                        End If
                    End If
                End If
                If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                    sqlStr += itemStr
                End If
                sqlStr += customStr
                sqlStr += plantStr
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                    sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc"
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) "
                    sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc"
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                    sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
                End If
                sqlStr += ") "
            End If
            sqlStr += "group by Fail_Mode "
            sqlStr += "order by fail_ratio desc"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.Fill(topDT)

            If topDT.Rows.Count <> 0 Then
                lab_wait.Text = ""
                ' === Raw Data ===
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    sqlStr = "select SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime, "
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    sqlStr = "select SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datatime, "
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    sqlStr = "select SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datatime, "
                End If
                sqlStr += "Fail_Mode, "
                sqlStr += "' ' as MF_Area, "
                sqlStr += "' ' as DefectCode_ID, "
                sqlStr += "Fail_Mode as 'NewFailMode', "
                If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                    sqlStr += "round((convert(float, SUM(Fail_Count_byXoutScrap))/SUM(Original_Input_QTY))/COUNT(Original_Input_QTY), 4) * 100 as Fail_ratio_byXoutScrap "
                Else
                    sqlStr += "round((convert(float, SUM(Fail_Count))/SUM(Original_Input_QTY))/COUNT(Original_Input_QTY), 4) * 100 as Fail_Ratio "
                End If
                sqlStr += "from dbo.VW_BinCode_Daily_Lot "
                sqlStr += "where 1=1 "

                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                End If
                'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                '    sqlStr += partStr
                'Else
                '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += partStr
                Else
                    If listB_PartShow.Items.Count = 0 Then
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                    Else
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                        If sGetPartID <> "" Then
                            sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                        End If
                    End If
                End If
                If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                    sqlStr += itemStr
                End If
                sqlStr += customStr
                sqlStr += plantStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 days
                    'sqlStr += "and substring(Convert(varchar, Convert(DateTime, trtm), 112), 1, 8) in ("
                    'sqlStr += "select top(4) substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "from SystemDateMapping "
                    'sqlStr += "WHERE substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) <= '" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "' "
                    'sqlStr += "group by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "order by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) desc"
                    'sqlStr += ") "
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datatime "
                    End If
                    sqlStr += "from VW_BinCode_Daily_Lot "
                    sqlStr += "where 1 = 1 "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                    If cb_NonIPQC.Checked Then
                        sqlStr += "and Fail_Mode <> 'IPQC defect' "
                    End If
                    If cb_Non8K.Checked Then
                        sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                    End If
                    'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                    '    sqlStr += partStr
                    'Else
                    '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr += partStr
                    Else
                        If listB_PartShow.Items.Count = 0 Then
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                        Else
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                            If sGetPartID <> "" Then
                                sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                        sqlStr += itemStr
                    End If
                    sqlStr += customStr
                    sqlStr += plantStr
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc"
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc"
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
                    End If
                    sqlStr += ")"
                Else
                    ' Custom
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) in (" + dateTimeTemp + ") "
                    End If
                End If
                If rb_dayType.SelectedIndex = 0 Then 'Daily
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8), Fail_Mode "
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc, Fail_ratio_byXoutScrap desc"
                    Else
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc, fail_ratio desc"
                    End If
                ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2), Fail_Mode "
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc, Fail_ratio_byXoutScrap desc"
                    Else
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc, fail_ratio desc"
                    End If
                ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                    sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6), Fail_Mode "
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc, Fail_ratio_byXoutScrap desc"
                    Else
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc, fail_ratio desc"
                    End If
                End If
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                rawDT = New DataTable
                myAdapter.Fill(rawDT)

                ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
                chipSetRawDT = New DataTable
                If ddlProduct.SelectedValue = "CS" Then
                    sqlStr = "select ww, plant, Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
                    sqlStr += "from dbo.VW_BinCode_Summary where 1=1 "

                    ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                    sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                    sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                    sqlStr += customStr
                    'sqlStr += "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
                    If (sGetPartID <> "") Then
                        sqlStr += "and part_id in (" + sGetPartID + ") "
                    End If

                    If rbl_week.SelectedIndex = 0 Then
                        sqlStr += "and ww in (SELECT Top(4) yearWW FROM SystemDateMapping WHERE trtm <= '" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW ORDER BY yearWW DESC) "
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) = ("
                        sqlStr += "select top(1) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime from dbo.VW_BinCode_Daily_Lot "
                        sqlStr += "where 1=1 "
                        sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                        sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                        If cb_NonIPQC.Checked Then
                            sqlStr += "and Fail_Mode <> 'IPQC defect' "
                        End If
                        If cb_Non8K.Checked Then
                            sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                        End If
                        'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                        '    sqlStr += partStr
                        'Else
                        '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        'End If
                        If tr_BumpingType.Visible = False Then
                            sqlStr += partStr
                        Else
                            If listB_PartShow.Items.Count = 0 Then
                                If strBumpingType <> "" Then
                                    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                                End If
                            Else
                                If strBumpingType <> "" Then
                                    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                                End If
                                If sGetPartID <> "" Then
                                    sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                                End If
                            End If
                        End If
                        If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                            sqlStr += itemStr
                        End If
                        sqlStr += customStr
                        sqlStr += plantStr
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc"
                        sqlStr += ") "
                    Else
                        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW DESC"
                        sqlStr += ") "
                    End If

                    sqlStr += itemStr
                    'sqlStr += "group by ww, plant, Fail_Mode, Fail_Ratio, MF_Stage "
                    'sqlStr += "order by plant desc, fail_ratio desc"
                    sqlStr += "group by ww, plant, Fail_Mode, Fail_Ratio, MF_Stage "
                    sqlStr += "order by plant desc, fail_ratio desc"
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myAdapter.Fill(chipSetRawDT)
                End If

                'If rbl_lossItem.SelectedIndex = 5 Then ' [Yield Loss Item Rank] ==> Custom
                sqlStr = "select * from ("
                If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                    sqlStr &= "select a.Fail_Mode, "
                Else
                    sqlStr &= "select " + topStr + " a.Fail_Mode, "
                End If
                sqlStr &= "("
                sqlStr &= "select SUM(Fail_Ratio) "
                sqlStr &= "from VW_BinCode_Daily_Lot "
                sqlStr &= "where 1=1 "
                sqlStr &= "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                sqlStr &= "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                End If
                'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                '    sqlStr += partStr
                'Else
                '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += partStr
                Else
                    If listB_PartShow.Items.Count = 0 Then
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                    Else
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                        If sGetPartID <> "" Then
                            sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                        End If
                    End If
                End If
                If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                    sqlStr += itemStr
                End If
                sqlStr += customStr
                sqlStr += plantStr
                'If rbl_week.SelectedIndex = 0 Then
                '    If rb_dayType.SelectedIndex = 0 Then
                '        ' Defaule 4 day
                '        'sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) in ("
                '        'sqlStr += "select top 4 substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        'sqlStr += "from SystemDateMapping "
                '        'sqlStr += "group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        'sqlStr += "order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) desc"
                '        'sqlStr += ") "
                '    ElseIf rb_dayType.SelectedIndex = 1 Then
                '        ' Defaule 4 week
                '        'sqlStr += "and ww in (select top 4 yearWW from SystemDateMapping group by yearWW order by yearWW desc) "
                '    ElseIf rb_dayType.SelectedIndex = 2 Then
                '        ' Defaule 4 month
                '        'sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) in (select top 4 substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) from SystemDateMapping group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) desc) "
                '    End If
                'Else
                '    ' Custom
                '    If rb_dayType.SelectedIndex = 0 Then
                '        ' Defaule 4 day
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) in ("
                '        'sqlStr += "select top (1) substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        'sqlStr += "from SystemDateMapping "
                '        'sqlStr += "where 1=1 " + weekStr + " "
                '        'sqlStr += "group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        'sqlStr += "order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) desc"
                '        sqlStr += ") "
                '    ElseIf rb_dayType.SelectedIndex = 1 Then
                '        ' Defaule 4 week
                '        sqlStr += "and ww in ("
                '        'sqlStr += "select top (1) yearWW "
                '        'sqlStr += "from SystemDateMapping "
                '        'sqlStr += "where 1=1 " + weekStr + " "
                '        'sqlStr += "group by yearWW "
                '        'sqlStr += "order by yearWW desc"
                '        sqlStr += ") "
                '    ElseIf rb_dayType.SelectedIndex = 2 Then
                '        ' Defaule 4 month
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) in ("
                '        'sqlStr += "select top (1) substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) "
                '        'sqlStr += "from SystemDateMapping "
                '        'sqlStr += "where 1=1 " + weekStr + " "
                '        'sqlStr += "group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) "
                '        'sqlStr += "order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) desc"
                '        sqlStr += ") "
                '    End If
                'End If
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 days
                    'sqlStr += "and substring(Convert(varchar, Convert(DateTime, trtm), 112), 1, 8) in ("
                    'sqlStr += "select top(4) substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "from SystemDateMapping "
                    'sqlStr += "WHERE substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) <= '" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "' "
                    'sqlStr += "group by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "order by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) desc"
                    'sqlStr += ") "
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datatime "
                    End If
                    sqlStr += "from VW_BinCode_Daily_Lot "
                    sqlStr += "where 1 = 1 "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                    If cb_NonIPQC.Checked Then
                        sqlStr += "and Fail_Mode <> 'IPQC defect' "
                    End If
                    If cb_Non8K.Checked Then
                        sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                    End If
                    'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                    '    sqlStr += partStr
                    'Else
                    '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr += partStr
                    Else
                        If listB_PartShow.Items.Count = 0 Then
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                        Else
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                            If sGetPartID <> "" Then
                                sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                        sqlStr += itemStr
                    End If
                    sqlStr += customStr
                    sqlStr += plantStr
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc"
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc"
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
                    End If
                    sqlStr += ")"
                Else
                    ' Custom
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) in (" + dateTimeTemp + ") "
                    End If
                End If
                sqlStr &= "and Fail_Mode = a.Fail_Mode "
                sqlStr &= ") as Fail_Ratio "
                sqlStr &= "from dbo.VW_BinCode_Daily_Lot a "
                sqlStr &= "where 1=1 "
                sqlStr &= "AND a.DefectCode != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                sqlStr &= "AND a.DefectCode != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                End If
                'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                '    sqlStr += partStr
                'Else
                '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += partStr
                Else
                    If listB_PartShow.Items.Count = 0 Then
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                    Else
                        If strBumpingType <> "" Then
                            sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                        End If
                        If sGetPartID <> "" Then
                            sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                        End If
                    End If
                End If
                If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                    sqlStr += itemStr
                End If
                sqlStr += customStr
                sqlStr += plantStr
                'If rbl_week.SelectedIndex = 0 Then
                '    If rb_dayType.SelectedIndex = 0 Then
                '        ' Defaule 4 day
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) in ("
                '        sqlStr += "select top 4 substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        sqlStr += "from SystemDateMapping "
                '        sqlStr += "group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) "
                '        sqlStr += "order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) desc"
                '        sqlStr += ") "
                '    ElseIf rb_dayType.SelectedIndex = 1 Then
                '        ' Defaule 4 week
                '        sqlStr += "and ww in (select top 4 yearWW from SystemDateMapping group by yearWW order by yearWW desc) "
                '    ElseIf rb_dayType.SelectedIndex = 2 Then
                '        ' Defaule 4 month
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) in (select top 4 substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) from SystemDateMapping group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) desc) "
                '    End If
                'Else
                '    ' Custom
                '    If rb_dayType.SelectedIndex = 0 Then
                '        ' Defaule 4 day
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) in (select top (1) substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) from SystemDateMapping where 1=1 " + weekStr + " group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 8) desc) "
                '    ElseIf rb_dayType.SelectedIndex = 1 Then
                '        ' Defaule 4 week
                '        'sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                '        sqlStr += "and ww in (select top (1) yearWW from SystemDateMapping where 1=1 " + weekStr + " group by yearWW order by yearWW desc) "
                '    ElseIf rb_dayType.SelectedIndex = 2 Then
                '        ' Defaule 4 month
                '        sqlStr += "and substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) in (select top (1) substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) from SystemDateMapping where 1=1 " + weekStr + " group by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) order by substring(Convert(varchar, Convert(DateTime, [DataTime]), 112), 1, 6) desc) "
                '    End If
                'End If
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 days
                    'sqlStr += "and substring(Convert(varchar, Convert(DateTime, trtm), 112), 1, 8) in ("
                    'sqlStr += "select top(4) substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "from SystemDateMapping "
                    'sqlStr += "WHERE substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) <= '" + DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "' "
                    'sqlStr += "group by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) "
                    'sqlStr += "order by substring(CONVERT(varchar, Convert(datetime, [DateTime]), 112),1,8) desc"
                    'sqlStr += ") "
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) AS Datatime "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) IN ("
                        sqlStr += "select Top(4) SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) AS Datatime "
                    End If
                    sqlStr += "from VW_BinCode_Daily_Lot "
                    sqlStr += "where 1 = 1 "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) "
                    sqlStr += "AND DefectCode != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "
                    If cb_NonIPQC.Checked Then
                        sqlStr += "and Fail_Mode <> 'IPQC defect' "
                    End If
                    If cb_Non8K.Checked Then
                        sqlStr += "and Fail_Mode NOT LIKE '8K%' "
                    End If
                    'If (rb_ProductPart.Enabled = True AndAlso ddlPart.Enabled = True) Then
                    '    sqlStr += partStr
                    'Else
                    '    sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr += partStr
                    Else
                        If listB_PartShow.Items.Count = 0 Then
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                        Else
                            If strBumpingType <> "" Then
                                sqlStr += "AND BumpingType in (" + strBumpingType + ") "
                            End If
                            If sGetPartID <> "" Then
                                sqlStr += "AND Part_Id in (" + sGetPartID + ") "
                            End If
                        End If
                    End If
                    If rbl_lossItem.SelectedIndex = 5 Then '選擇CustomItem(非TopN)
                        sqlStr += itemStr
                    End If
                    sqlStr += customStr
                    sqlStr += plantStr
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) desc"
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) desc"
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "group by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) "
                        sqlStr += "order by SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) desc"
                    End If
                    sqlStr += ")"
                Else
                    ' Custom
                    If rb_dayType.SelectedIndex = 0 Then 'Daily
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 8) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 1 Then 'Weekly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 4) + RIGHT(REPLICATE('0', 2) + CAST(DATENAME(Week, Datatime) AS NVARCHAR), 2) in (" + dateTimeTemp + ") "
                    ElseIf rb_dayType.SelectedIndex = 2 Then 'Monthly
                        sqlStr += "and SUBSTRING(CONVERT(VARCHAR, Datatime, 112), 1, 6) in (" + dateTimeTemp + ") "
                    End If
                End If
                sqlStr &= "group by a.Fail_Mode" & Space(1)
                sqlStr &= ") sm" & Space(1)
                sqlStr &= "order by Fail_Ratio desc"
                'Else
                '    sqlStr = "select " + topStr + " * from (" & Space(1)
                '    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_Ratio)" & Space(1)
                '    sqlStr &= "from dbo.VW_BinCode_Summary" & Space(1)
                '    sqlStr &= "where 1=1" & Space(1)
                '    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                '    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                '    sqlStr &= customStr
                '    sqlStr &= partStr
                '    sqlStr &= plantStr
                '    If rbl_week.SelectedIndex = 0 Then
                '        ' Defaule 4 week
                '        sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                '    Else
                '        ' Custom
                '        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                '    End If
                '    sqlStr &= "and Fail_Mode = a.Fail_Mode" & Space(1)
                '    sqlStr &= ") as Fail_Ratio" & Space(1)
                '    sqlStr &= "from dbo.VW_BinCode_Summary a" & Space(1)
                '    sqlStr &= "where 1=1" & Space(1)
                '    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                '    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                '    If cb_NonIPQC.Checked Then
                '        sqlStr &= "and Fail_Mode <> 'IPQC defect' "
                '    End If
                '    If cb_Non8K.Checked Then
                '        sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                '    End If
                '    sqlStr &= customStr
                '    sqlStr &= partStr
                '    sqlStr &= plantStr
                '    If rbl_week.SelectedIndex = 0 Then
                '        ' Defaule 4 week
                '        sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW) "
                '    Else
                '        ' Custom
                '        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                '    End If
                '    sqlStr &= "group by a.Fail_Mode "
                '    sqlStr &= ") sm "
                '    sqlStr &= "order by Fail_Ratio desc"
                'End If
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(new_topDT)

                conn.Close()
                If rb_dayType.SelectedIndex = 0 Then
                    'BarChart(rawDT, topDT)
                Else
                    'BarChart_FailModeByStageRatioSummary(yl, rawDT, new_topDT)
                End If

                If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
                    Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                    Dim Chart As New Dundas.Charting.WebControl.Chart()
                    DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
                    Chart_Panel.Controls.Add(Chart)
                End If

                tr_chartDisplay.Visible = True
                If cb_DRowData.Checked Then
                    showDailyRowData(rawDT, chipSetRawDT)
                    tr_gvDisplay.Visible = True
                End If
            Else
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 
                    lab_wait.Text = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 無資料, 可使用天數自訂 !"
                Else
                    ' Custom
                    lab_wait.Text = dateTimeTemp + " 無資料, 可使用天數自訂 !"
                End If
            End If
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try
    End Sub

    'Private Sub weekly_failMode()

    '    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '    Dim sqlStr As String = ""
    '    Dim customStr As String = " "
    '    Dim plantStr As String = ""
    '    Dim partStr As String = " "
    '    Dim weekStr As String = " "
    '    Dim itemStr As String = " "
    '    Dim topStr As String = " "
    '    Dim myAdapter As SqlDataAdapter
    '    Dim topDT As DataTable = New DataTable
    '    Dim new_topDT As DataTable = New DataTable
    '    Dim rawDT, chipSetRawDT As DataTable

    '    If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 12 Then
    '        ShowMessage("選擇週數最多為 12 週")
    '        Exit Sub
    '    End If

    '    Try
    '        ' --- Customer ID ---
    '        ' customStr = "and customer_id='" + (ddlCustomer.SelectedValue.Trim()) + "' "
    '        ' --- Part ID ---
    '        'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
    '        Dim strBumpingType As String = ""
    '        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '            If n = 0 Then
    '                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '            Else
    '                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '            End If
    '        Next
    '        Dim BumpingType_Part As String = ""
    '        Dim sGetPartID As String = Get_PartID()
    '        If strBumpingType <> "" Then
    '            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '        End If
    '        If tr_BumpingType.Visible = False Then
    '            If rb_ProductPart.SelectedIndex = 0 Then
    '                If listB_BumpingTypeShow.Enabled = False Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Production_id={0} "
    '                    Else
    '                        partStr += "{0}"
    '                    End If
    '                Else
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Production_id={0} "
    '                    Else
    '                        partStr += "{0}"
    '                    End If
    '                End If
    '            ElseIf rb_ProductPart.SelectedIndex = 1 Then
    '                Dim oPartID = sGetPartID.Split(",")
    '                If oPartID.Length = 1 Then
    '                    If listB_BumpingTypeShow.Enabled = False Then
    '                        If sGetPartID <> "" Then
    '                            partStr += "and Part_Id = {0} "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    Else
    '                        If BumpingType_Part <> "" Then
    '                            partStr += "and Part_Id = {0} "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    End If
    '                ElseIf oPartID.Length > 1 Then
    '                    If listB_BumpingTypeShow.Enabled = False Then
    '                        If sGetPartID <> "" Then
    '                            partStr += "and Part_Id in({0}) "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    Else
    '                        If BumpingType_Part <> "" Then
    '                            partStr += "and Part_Id in({0}) "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Else
    '            If BumpingType_Part <> "" Then
    '                partStr += "and Part_Id in (" & BumpingType_Part & ") "
    '            Else
    '                Dim oPartID = sGetPartID.Split(",")
    '                If oPartID.Length = 1 Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Part_Id = " + sGetPartID
    '                    End If
    '                ElseIf oPartID.Length > 1 Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Part_Id in(" + sGetPartID + ") "
    '                    End If
    '                End If
    '            End If
    '        End If
    '        If ddlProduct.SelectedValue <> "CPU" Then
    '            plantStr = "and Plant='All' "
    '        Else
    '            plantStr = ""
    '        End If

    '        ' --- Week ID ---
    '        Dim weekTemp As String = ""
    '        If rbl_week.SelectedIndex = 1 Then
    '            For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
    '                weekTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
    '            Next
    '            If weekTemp.Length > 0 Then
    '                weekTemp = weekTemp.Substring(0, (weekTemp.Length - 1))
    '                weekStr = "and yearWW in (" + weekTemp + ") "
    '            End If
    '        End If

    '        If rbl_lossItem.SelectedIndex = 0 Then
    '            topStr = "top(10)"
    '        ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '            topStr = "top(20)"
    '        ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '            topStr = "top(30)"
    '        ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '            topStr = "top(40)"
    '        ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '            topStr = "top(50)"
    '        End If

    '        conn.Open()
    '        ' --- Yield Loss ID ---
    '        Dim itemTemp As String = ""
    '        itemStr = ""

    '        If rbl_lossItem.SelectedIndex = 5 Then
    '            ' === Custom Item ===
    '            For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
    '                itemTemp += "'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "',"
    '            Next
    '            itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
    '            itemStr += "and fail_mode in (" + itemTemp + ") "

    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr = "select Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            Else
    '                sqlStr = "select Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            End If
    '            sqlStr += "from dbo.VW_BinCode_Summary where 1=1 "

    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            sqlStr += itemStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 week
    '                sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '            Else
    '                ' Custom
    '                sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '            End If
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                sqlStr += "order by Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
    '                sqlStr += "order by fail_ratio desc"
    '            End If
    '        Else
    '            ' === Top N === ' 如果選 Custom 就要呈現選擇的 item
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr = "select " + topStr + " Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            Else
    '                sqlStr = "select " + topStr + " Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            End If
    '            sqlStr += "from dbo.VW_BinCode_Summary where 1=1 "

    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            If cb_NonIPQC.Checked Then
    '                sqlStr += "and Fail_Mode <> 'IPQC defect' "
    '            End If
    '            If cb_Non8K.Checked Then
    '                sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
    '            End If
    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 week
    '                sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '            Else
    '                ' Custom
    '                sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '            End If
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                sqlStr += "order by Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
    '                sqlStr += "order by fail_ratio desc"
    '            End If
    '        End If
    '        myAdapter = New SqlDataAdapter(sqlStr, conn)
    '        myAdapter.Fill(topDT)

    '        If (topDT.Rows.Count <> 0) Then

    '            ' === Raw Data ===
    '            lab_wait.Text = ""
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr = "select a.WW, a.Fail_Mode, a.Fail_ratio_byXoutScrap, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
    '            Else
    '                sqlStr = "select a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
    '            End If
    '            sqlStr += "from dbo.VW_BinCode_Summary a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
    '            ' --- WB 資料當時沒有 Customer_ID 2014/01/07 ---
    '            'sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and a.DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            If cb_NonIPQC.Checked Then
    '                sqlStr += "and Fail_Mode <> 'IPQC defect' "
    '            End If
    '            If cb_Non8K.Checked Then
    '                sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
    '            End If
    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 week
    '                sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
    '            Else
    '                ' Custom
    '                sqlStr += weekStr
    '            End If

    '            sqlStr += itemStr
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                'Fail_ratio_byXoutScrap
    '                sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_ratio_byXoutScrap, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
    '                sqlStr += "order by b.yearWW desc, a.Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
    '                sqlStr += "order by b.yearWW desc, a.fail_ratio desc"
    '            End If
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            rawDT = New DataTable
    '            myAdapter.Fill(rawDT)

    '            ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
    '            chipSetRawDT = New DataTable
    '            If ddlProduct.SelectedValue = "CS" Then
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr = "select ww, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
    '                Else
    '                    sqlStr = "select ww, plant, Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
    '                End If
    '                sqlStr += "from dbo.VW_BinCode_Summary where 1=1 "
    '                ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
    '                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '                sqlStr += customStr
    '                sqlStr += partStr

    '                If rbl_week.SelectedIndex = 0 Then
    '                    sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW ORDER BY yearWW DESC) "
    '                Else
    '                    sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW DESC) "
    '                End If

    '                sqlStr += itemStr
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr += "group by ww, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                    sqlStr += "order by plant desc, Fail_ratio_byXoutScrap desc"
    '                Else
    '                    sqlStr += "group by ww, plant, Fail_Mode, Fail_Ratio, MF_Stage "
    '                    sqlStr += "order by plant desc, fail_ratio desc"
    '                End If
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myAdapter.Fill(chipSetRawDT)

    '            End If

    '            If rbl_lossItem.SelectedIndex = 5 Then
    '                sqlStr = "select * from (" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_ratio_byXoutScrap)" & Space(1)
    '                    sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.VW_BinCode_Summary" & Space(1)
    '                Else
    '                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_Ratio)" & Space(1)
    '                    sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.VW_BinCode_Summary" & Space(1)
    '                End If
    '                sqlStr &= "where 1=1" & Space(1)
    '                sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= customStr
    '                sqlStr &= partStr
    '                sqlStr &= plantStr
    '                sqlStr &= itemStr
    '                If rbl_week.SelectedIndex = 0 Then
    '                    ' Defaule 4 week
    '                    sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                Else
    '                    ' Custom
    '                    sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                End If
    '                sqlStr &= "and Fail_Mode = a.Fail_Mode) sm1" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= ") as Fail_ratio_byXoutScrap" & Space(1)
    '                Else
    '                    sqlStr &= ") as Fail_Ratio" & Space(1)
    '                End If
    '                sqlStr &= "from dbo.VW_BinCode_Summary a" & Space(1)
    '                sqlStr &= "where 1=1" & Space(1)
    '                sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= customStr
    '                sqlStr &= partStr
    '                sqlStr &= plantStr
    '                sqlStr &= itemStr
    '                If rbl_week.SelectedIndex = 0 Then
    '                    ' Defaule 4 week
    '                    sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                Else
    '                    ' Custom
    '                    sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                End If
    '                sqlStr &= "group by a.Fail_Mode" & Space(1)
    '                sqlStr &= ") sm" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= "order by Fail_ratio_byXoutScrap desc" & Space(1)
    '                Else
    '                    sqlStr &= "order by Fail_Ratio desc" & Space(1)
    '                End If
    '            Else
    '                sqlStr = "select " + topStr + " * from (" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_ratio_byXoutScrap)" & Space(1)
    '                    sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.VW_BinCode_Summary" & Space(1)
    '                Else
    '                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_Ratio)" & Space(1)
    '                    sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.VW_BinCode_Summary" & Space(1)
    '                End If
    '                sqlStr &= "where 1=1" & Space(1)
    '                sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= customStr
    '                sqlStr &= partStr
    '                sqlStr &= plantStr
    '                If rbl_week.SelectedIndex = 0 Then
    '                    ' Defaule 4 week
    '                    sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                Else
    '                    ' Custom
    '                    sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                End If
    '                sqlStr &= "and Fail_Mode = a.Fail_Mode) sm1" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= ") as Fail_ratio_byXoutScrap" & Space(1)
    '                Else
    '                    sqlStr &= ") as Fail_Ratio" & Space(1)
    '                End If
    '                sqlStr &= "from dbo.VW_BinCode_Summary a" & Space(1)
    '                sqlStr &= "where 1=1" & Space(1)
    '                sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
    '                sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
    '                If cb_NonIPQC.Checked Then
    '                    sqlStr &= "and Fail_Mode <> 'IPQC defect' "
    '                End If
    '                If cb_Non8K.Checked Then
    '                    sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
    '                End If
    '                sqlStr &= customStr
    '                sqlStr &= partStr
    '                sqlStr &= plantStr
    '                If rbl_week.SelectedIndex = 0 Then
    '                    ' Defaule 4 week
    '                    sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                Else
    '                    ' Custom
    '                    sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                End If
    '                sqlStr &= "group by a.Fail_Mode" & Space(1)
    '                sqlStr &= ") sm" & Space(1)
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    'Fail_ratio_byXoutScrap
    '                    sqlStr &= "order by Fail_ratio_byXoutScrap desc" & Space(1)
    '                Else
    '                    sqlStr &= "order by Fail_Ratio desc" & Space(1)
    '                End If
    '            End If
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myAdapter.Fill(new_topDT)

    '            conn.Close()
    '            '***BarChart(rawDT, topDT)
    '            BarChart_FailModeByStageRatioSummary(rawDT, new_topDT)

    '            If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
    '                Chart_Panel.Controls.Add(New LiteralControl("<br>"))
    '                Dim Chart As New Dundas.Charting.WebControl.Chart()
    '                DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
    '                Chart_Panel.Controls.Add(Chart)
    '            End If

    '            tr_chartDisplay.Visible = True
    '            If cb_DRowData.Checked Then
    '                showWeeklyRowData(rawDT, chipSetRawDT)
    '                tr_gvDisplay.Visible = True
    '            End If

    '        Else

    '            Dim inDT As DataTable = New DataTable()
    '            sqlStr = "SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myAdapter.Fill(inDT)
    '            lab_wait.Text = inDT.Rows(0)("yearWW").ToString().Substring(0, 4) + "W" + inDT.Rows(0)("yearWW").ToString().Substring(4, 2) + " 無資料, 可使用週數自訂 !"

    '        End If

    '    Catch ex As Exception
    '        Dim sError As String = ex.ToString()
    '    Finally
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try

    'End Sub

    Private Sub weekly_failMode()

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
        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 12 Then
            ShowMessage("選擇週數最多為 12 週")
            Exit Sub
        End If

        Try
            ' --- Customer ID ---
            ' customStr = "and customer_id='" + (ddlCustomer.SelectedValue.Trim()) + "' "
            ' --- Part ID ---
            If tr_BumpingType.Visible = False Then
                'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
                Dim sGetPartID As String = Get_PartID()
                If (sGetPartID <> "") Then
                    partStr += "and part_id in (" + sGetPartID + ") "
                End If
            Else
                If listB_BumpingTypeShow.Items.Count = 0 Then
                    'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
                    Dim sGetPartID As String = Get_PartID()
                    If (sGetPartID <> "") Then
                        partStr += "and part_id in (" + sGetPartID + ") "
                    End If
                Else
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        End If
                    Next

                    partStr = "and BumpingType_Id IN(" + strBumpingType + ") "
                End If
            End If
            If ddlProduct.SelectedValue <> "CPU" Then
                plantStr = "and Plant='All' "
            Else
                plantStr = ""
            End If

            ' --- Week ID ---
            Dim weekTemp As String = ""
            If rbl_week.SelectedIndex = 1 Then
                For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                    weekTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
                Next
                If weekTemp.Length > 0 Then
                    weekTemp = weekTemp.Substring(0, (weekTemp.Length - 1))
                    weekStr = "and yearWW in (" + weekTemp + ") "
                End If
            End If

            If rbl_lossItem.SelectedIndex = 0 Then
                topStr = "top(10)"
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                topStr = "top(20)"
            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                topStr = "top(30)"
            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                topStr = "top(40)"
            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                topStr = "top(50)"
            End If

            conn.Open()
            ' --- Yield Loss ID ---
            Dim itemTemp As String = ""
            itemStr = ""

            If rbl_lossItem.SelectedIndex = 5 Then
                ' === Custom Item ===
                For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                    itemTemp += "'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "',"
                Next
                itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
                itemStr += "and fail_mode in (" + itemTemp + ") "

                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    'Fail_ratio_byXoutScrap
                    sqlStr = "select Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                Else
                    sqlStr = "select Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_ByBT where 1=1 "
                    End If
                End If

                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                sqlStr += itemStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                Else
                    ' Custom
                    sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                End If
                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    'Fail_ratio_byXoutScrap
                    sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                    sqlStr += "order by Fail_ratio_byXoutScrap desc"
                Else
                    sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
                    sqlStr += "order by fail_ratio desc"
                End If
            Else
                ' === Top N === ' 如果選 Custom 就要呈現選擇的 item
                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    'Fail_ratio_byXoutScrap
                    sqlStr = "select " + topStr + " Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                Else
                    sqlStr = "select " + topStr + " Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_ByBT where 1=1 "
                    End If
                End If

                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                End If
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                Else
                    ' Custom
                    sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                End If
                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    'Fail_ratio_byXoutScrap
                    sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                    sqlStr += "order by Fail_ratio_byXoutScrap desc"
                Else
                    sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
                    sqlStr += "order by fail_ratio desc"
                End If
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.Fill(topDT)

            If (topDT.Rows.Count <> 0) Then

                ' === Raw Data ===
                lab_wait.Text = ""
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                '    'Fail_ratio_byXoutScrap
                '    sqlStr = "select a.WW, a.Fail_Mode, a.Fail_ratio_byXoutScrap, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
                'Else
                sqlStr = "select a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_ByBT a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
                    End If
                End If

                ' --- WB 資料當時沒有 Customer_ID 2014/01/07 ---
                'sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and a.DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                End If
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
                Else
                    ' Custom
                    sqlStr += weekStr
                End If

                sqlStr += itemStr
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                '    'Fail_ratio_byXoutScrap
                '    sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_ratio_byXoutScrap, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
                '    sqlStr += "order by b.yearWW desc, a.Fail_ratio_byXoutScrap desc"
                'Else
                sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
                sqlStr += "order by b.yearWW desc, a.fail_ratio desc"
                'End If
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                rawDT = New DataTable
                myAdapter.Fill(rawDT)

                ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
                chipSetRawDT = New DataTable
                If ddlProduct.SelectedValue = "CS" Then
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr = "select ww, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
                    'Else
                    sqlStr = "select ww, plant, Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr += "from dbo.WB_BinCode_Summary where 1=1 "
                        Else
                            sqlStr += "from dbo.WB_BinCode_Summary_ByBT where 1=1 "
                        End If
                    End If
                    ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
                    ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                    'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                    sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                    sqlStr += customStr
                    sqlStr += partStr

                    If rbl_week.SelectedIndex = 0 Then
                        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW ORDER BY yearWW DESC) "
                    Else
                        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW DESC) "
                    End If

                    sqlStr += itemStr
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr += "group by ww, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                    '    sqlStr += "order by plant desc, Fail_ratio_byXoutScrap desc"
                    'Else
                    sqlStr += "group by ww, plant, Fail_Mode, Fail_Ratio, MF_Stage "
                    sqlStr += "order by plant desc, fail_ratio desc"
                    'End If
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myAdapter.Fill(chipSetRawDT)

                End If

                If rbl_lossItem.SelectedIndex = 5 Then
                    sqlStr = "select * from (" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_ratio_byXoutScrap)" & Space(1)
                    '    If tr_BumpingType.Visible = False Then
                    '        sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary" & Space(1)
                    '    Else
                    '        If listB_BumpingTypeShow.Items.Count = 0 Then
                    '            sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary" & Space(1)
                    '        Else
                    '            sqlStr &= "from (select DISTINCT ww,BumpingType_Id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary_ByBT" & Space(1)
                    '        End If
                    '    End If
                    'Else
                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_Ratio)" & Space(1)
                    If tr_BumpingType.Visible = False Then
                        sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary" & Space(1)
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary" & Space(1)
                        Else
                            sqlStr &= "from (select DISTINCT ww,BumpingType_Id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary_ByBT" & Space(1)
                        End If
                    End If
                    'End If
                    sqlStr &= "where 1=1" & Space(1)
                    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= customStr
                    sqlStr &= partStr
                    sqlStr &= plantStr
                    sqlStr &= itemStr
                    If rbl_week.SelectedIndex = 0 Then
                        ' Defaule 4 week
                        sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    Else
                        ' Custom
                        sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    End If
                    sqlStr &= "and Fail_Mode = a.Fail_Mode) sm1" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= ") as Fail_ratio_byXoutScrap" & Space(1)
                    'Else
                    sqlStr &= ") as Fail_Ratio" & Space(1)
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr &= "from dbo.WB_BinCode_Summary a" & Space(1)
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr &= "from dbo.WB_BinCode_Summary a" & Space(1)
                        Else
                            sqlStr &= "from dbo.WB_BinCode_Summary_ByBT a" & Space(1)
                        End If
                    End If
                    sqlStr &= "where 1=1" & Space(1)
                    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= customStr
                    sqlStr &= partStr
                    sqlStr &= plantStr
                    sqlStr &= itemStr
                    If rbl_week.SelectedIndex = 0 Then
                        ' Defaule 4 week
                        sqlStr += "and ww in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    Else
                        ' Custom
                        sqlStr += "and ww in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    End If
                    sqlStr &= "group by a.Fail_Mode" & Space(1)
                    sqlStr &= ") sm" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= "order by Fail_ratio_byXoutScrap desc" & Space(1)
                    'Else
                    sqlStr &= "order by Fail_Ratio desc" & Space(1)
                    'End If
                Else
                    sqlStr = "select " + topStr + " * from (" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_ratio_byXoutScrap)" & Space(1)
                    '    If tr_BumpingType.Visible = False Then
                    '        sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary" & Space(1)
                    '    Else
                    '        If listB_BumpingTypeShow.Items.Count = 0 Then
                    '            sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary" & Space(1)
                    '        Else
                    '            sqlStr &= "from (select DISTINCT ww,BumpingType_Id,Fail_Mode,Fail_ratio_byXoutScrap from dbo.WB_BinCode_Summary_ByBT" & Space(1)
                    '        End If
                    '    End If
                    'Else
                    sqlStr &= "select a.Fail_Mode, (select SUM(Fail_Ratio)" & Space(1)
                    If tr_BumpingType.Visible = False Then
                        sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary" & Space(1)
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr &= "from (select DISTINCT ww,part_id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary" & Space(1)
                        Else
                            sqlStr &= "from (select DISTINCT ww,BumpingType_Id,Fail_Mode,Fail_Ratio from dbo.WB_BinCode_Summary_ByBT" & Space(1)
                        End If
                    End If
                    'End If
                    sqlStr &= "where 1=1" & Space(1)
                    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= customStr
                    sqlStr &= partStr
                    sqlStr &= plantStr
                    If rbl_week.SelectedIndex = 0 Then
                        ' Defaule 4 week
                        sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    Else
                        ' Custom
                        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    End If
                    sqlStr &= "and Fail_Mode = a.Fail_Mode) sm1" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= ") as Fail_ratio_byXoutScrap" & Space(1)
                    'Else
                    sqlStr &= ") as Fail_Ratio" & Space(1)
                    'End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr &= "from dbo.WB_BinCode_Summary a" & Space(1)
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr &= "from dbo.WB_BinCode_Summary a" & Space(1)
                        Else
                            sqlStr &= "from dbo.WB_BinCode_Summary_ByBT a" & Space(1)
                        End If
                    End If
                    sqlStr &= "where 1=1" & Space(1)
                    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END)" & Space(1)
                    sqlStr &= "AND a.DefectCode_ID != (CASE WHEN a.MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END)" & Space(1)
                    If cb_NonIPQC.Checked Then
                        sqlStr &= "and Fail_Mode <> 'IPQC defect' "
                    End If
                    If cb_Non8K.Checked Then
                        sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                    End If
                    sqlStr &= customStr
                    sqlStr &= partStr
                    sqlStr &= plantStr
                    If rbl_week.SelectedIndex = 0 Then
                        ' Defaule 4 week
                        sqlStr += "and ww in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    Else
                        ' Custom
                        sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    End If
                    sqlStr &= "group by a.Fail_Mode" & Space(1)
                    sqlStr &= ") sm" & Space(1)
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                    '    'Fail_ratio_byXoutScrap
                    '    sqlStr &= "order by Fail_ratio_byXoutScrap desc" & Space(1)
                    'Else
                    sqlStr &= "order by Fail_Ratio desc" & Space(1)
                    'End If
                End If
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(new_topDT)

                conn.Close()
                '***BarChart(rawDT, topDT)
                ' BarChart_FailModeByStageRatioSummary(rawDT, new_topDT)

                If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
                    Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                    Dim Chart As New Dundas.Charting.WebControl.Chart()
                    DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
                    Chart_Panel.Controls.Add(Chart)
                End If

                tr_chartDisplay.Visible = True
                If cb_DRowData.Checked Then
                    showWeeklyRowData(rawDT, chipSetRawDT)
                    tr_gvDisplay.Visible = True
                End If

            Else

                Dim inDT As DataTable = New DataTable()
                sqlStr = "SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(inDT)
                lab_wait.Text = inDT.Rows(0)("yearWW").ToString().Substring(0, 4) + "W" + inDT.Rows(0)("yearWW").ToString().Substring(4, 2) + " 無資料, 可使用週數自訂 !"

            End If

        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Draw Chart
    Private Sub BarChart(ByRef raw As DataTable, ByRef setupDT As DataTable, ByVal yl As YieldlossInfo)
        Try
            Dim fary As ArrayList = New ArrayList
            Dim bhash As Hashtable = New Hashtable

            Dim chart As New Dundas.Charting.WebControl.Chart()
            'If rb_dayType.SelectedIndex = 0 Then
            '    DrawBarChart(chart, raw, setupDT, fary, bhash, "datatime")
            'Else
            '    DrawBarChart(chart, raw, setupDT, fary, bhash, "yearww")
            'End If
            DrawBarChart(yl, chart, raw, setupDT, fary, bhash, "datatime")

            Chart_Panel.Controls.Add(chart)
            Chart_Panel.Controls.Add(New LiteralControl("<br>"))

            chart = New Dundas.Charting.WebControl.Chart()
            'If rb_dayType.SelectedIndex = 0 Then
            '    DrawDiffBarChart(chart, raw, setupDT, fary, bhash, "datatime")
            'Else
            '    DrawDiffBarChart(chart, raw, setupDT, fary, bhash, "yearww")
            'End If
            DrawDiffBarChart(chart, raw, setupDT, fary, bhash, "datatime")
            Chart_Panel.Controls.Add(chart)
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        End Try
    End Sub

    ' Draw Chart
    Private Sub BarChart_MM(ByRef raw As DataTable, ByRef setupDT As DataTable, ByVal weekTemp As String)
        Try
            Dim FAry As ArrayList = New ArrayList
            Dim BHash As Hashtable = New Hashtable
            Dim yl As YieldlossInfo
            Dim Chart As New Dundas.Charting.WebControl.Chart()
            If rb_dayType.SelectedIndex = 0 Then
                DrawBarChart(yl, Chart, raw, setupDT, FAry, BHash, "DataTime", weekTemp)
            Else
                DrawBarChart(yl, Chart, raw, setupDT, FAry, BHash, "MM", weekTemp)
            End If

            Chart_Panel.Controls.Add(Chart)
            Chart_Panel.Controls.Add(New LiteralControl("<br>"))

            Chart = New Dundas.Charting.WebControl.Chart()
            If rb_dayType.SelectedIndex = 0 Then
                DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash, "DataTime")
            Else
                DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash, "MM")
            End If
            Chart_Panel.Controls.Add(Chart)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BarChart_FailModeByStageRatioSummary(ByVal yl As YieldlossInfo, ByRef raw As DataTable, ByRef setupDT As DataTable)
        Dim FAry As ArrayList = New ArrayList
        Dim BHash As Hashtable = New Hashtable

        Dim Chart As New Dundas.Charting.WebControl.Chart()
        'If rb_dayType.SelectedIndex = 0 Then
        '    DrawBarChart_FailModeByStageRatioSummary(Chart, raw, setupDT, FAry, BHash, "DataTime")
        'Else
        '    DrawBarChart_FailModeByStageRatioSummary(Chart, raw, setupDT, FAry, BHash, "yearWW")
        'End If
        DrawBarChart_FailModeByStageRatioSummary(yl, Chart, raw, setupDT, FAry, BHash, "DataTime")

        Chart_Panel.Controls.Add(Chart)
        Chart_Panel.Controls.Add(New LiteralControl("<br>"))

        Chart = New Dundas.Charting.WebControl.Chart()
        'If rb_dayType.SelectedIndex = 0 Then
        '    DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash, "DataTime")
        'Else
        '    DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash, "yearWW")
        'End If

        DrawDiffBarChart(Chart, raw, setupDT, FAry, BHash, "DataTime")
        Chart_Panel.Controls.Add(Chart)
    End Sub

    ' Draw Bar Chart
    Private Sub DrawBarChart(ByVal yl As YieldlossInfo, ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable, ByVal TimeColumn As String, Optional ByVal weekTemp As String = "")
        Try
            Chart.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
            Chart.ImageType = ChartImageType.Png
            Chart.Palette = ChartColorPalette.Dundas
            Chart.Height = Unit.Pixel(gChartH)
            Chart.Width = Unit.Pixel(gChartW)

            If rb_dayType.SelectedIndex = 0 Then
                Chart.Titles.Add("Fail Mode By Day")
            ElseIf rb_dayType.SelectedIndex = 1 Then
                Chart.Titles.Add("Fail Mode By Week")
            ElseIf rb_dayType.SelectedIndex = 2 Then
                Chart.Titles.Add("Fail Mode By Month")
            End If

            Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(0).Color = Color.DarkBlue


            If yl.BumpingType <> "" Then
                Chart.Titles.Add("BumpingType:" + yl.BumpingType.ToString)
            Else
                Chart.Titles.Add("Part:" + yl.Part_ID.ToString)
            End If

            Chart.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
            Chart.Titles(1).Color = Color.DarkBlue

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
            Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 14, FontStyle.Bold)
            Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
            Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)


            Chart.UI.Toolbar.Enabled = False
            Chart.UI.ContextMenu.Enabled = True

            Dim series As Series
            Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, TimeColumn)
            weekGroupDT.DefaultView.Sort = (TimeColumn + " asc")
            weekGroupDT = weekGroupDT.DefaultView.ToTable
            Dim dtFilter As DataTable
            Dim dr As DataRow
            Dim foundRows() As DataRow
            Dim insideRows() As DataRow
            Dim failMode As String
            Dim newfailMode As String
            Dim failValue As Double
            Dim weekStr As String
            Dim failObj As FailObj
            Dim colorInx As Integer = 0
            Dim scriptStr As String = ""

            Dim pieChartWeek As String = ""
            For i As Integer = 0 To (weekGroupDT.Rows.Count - 1)
                pieChartWeek += (weekGroupDT.Rows(i)(TimeColumn)).ToString + ","
            Next
            If pieChartWeek.Length > 0 Then
                pieChartWeek = pieChartWeek.Substring(0, (pieChartWeek.Length - 1))
            End If
            ViewState("pieChartWeek") = pieChartWeek

            colorInx = (weekGroupDT.Rows.Count - 1)
            For toolIndex As Integer = 0 To (weekGroupDT.Rows.Count - 1)

                weekStr = (weekGroupDT.Rows(toolIndex)(TimeColumn)).ToString
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                '    foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_ratio_byXoutScrap desc")
                'Else
                foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_Ratio desc")
                'End If
                dtFilter = DtSet.Clone
                For x = 0 To (foundRows.Length - 1)
                    dr = foundRows(x)
                    dtFilter.LoadDataRow(dr.ItemArray, False)
                Next
                dtFilter.CaseSensitive = True

                series = Chart.Series.Add((weekStr))
                series.ChartArea = "Default"
                series.Type = SeriesChartType.Column
                series.Color = aryColor(colorInx)
                series.BorderColor = Color.White
                series.BorderWidth = 1

                Dim product_part As String = "PRODUCT"
                If rb_ProductPart.SelectedIndex = 1 Then
                    product_part = "PART"
                End If

                For i As Integer = 0 To (setupDT.Rows.Count - 1)

                    newfailMode = (setupDT.Rows(i)("newFailMode").ToString.Trim()).Replace("'", "''")
                    failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "||")
                    insideRows = DtSet.Select(TimeColumn + "='" + weekStr + "' and newFailMode='" + newfailMode + "'")
                    failValue = 0

                    If insideRows.Length > 0 Then
                        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        '    If Not IsDBNull(insideRows(0).Item("Fail_ratio_byXoutScrap")) Then
                        '        failValue = CType(insideRows(0).Item("Fail_ratio_byXoutScrap"), Double)
                        '    End If
                        'Else
                        If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                            failValue = CType(insideRows(0).Item("Fail_Ratio"), Double)
                        End If
                        'End If
                    End If

                    If toolIndex = (weekGroupDT.Rows.Count - 1) Then
                        failObj = New FailObj
                        failObj.OriFail_Mode = failMode
                        failObj.Fail_Mode = newfailMode
                        failObj.Fail_Value = failValue
                        DiffFAry.Add(failObj)
                    ElseIf toolIndex = (weekGroupDT.Rows.Count - 2) Then
                        DiffBHash.Add(newfailMode, failValue)
                    End If


                    Dim sLotMerge As String = "True"

                    If cb_Lot_Merge.Checked = False Then
                        sLotMerge = "False"
                        If Cb_SF.Checked = True Then
                            sLotMerge += "_SF"
                        End If
                        If Cb_CR.Checked = False Then
                            sLotMerge += "_CR"
                        End If

                    Else
                        sLotMerge = "True"
                        If Cb_SF.Checked = True Then
                            sLotMerge += "_SF"
                        End If
                        If Cb_CR.Checked = False Then
                            sLotMerge += "_CR"
                        End If
                    End If

                    If ckFAI.Checked = True Then
                        sLotMerge += "_FAI"
                    End If

                    Dim sTimePeriod As String = "1"
                    If rb_dayType.SelectedIndex = 0 Then
                        sTimePeriod = "0"
                    ElseIf rb_dayType.SelectedIndex = 1 Then
                        sTimePeriod = "1"
                    Else
                        sTimePeriod = "2"
                    End If


                    'If rb_dayType.SelectedIndex = 0 Then
                    '    ' Daily
                    '    scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}')"
                    '    '=====================================================================================================================================================================================
                    '    'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,C}', '{1,CA}', '{2,P}', '{3,F}', '{4,D}', '{5,PLANT}', '{6,TYPE}', '{7,LotList}', '{8,TopN}')"
                    '    '=====================================================================================================================================================================================
                    '    If rbl_lossItem.SelectedIndex = 0 Then
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "10")
                    '    ElseIf rbl_lossItem.SelectedIndex = 1 Then
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "20")
                    '    ElseIf rbl_lossItem.SelectedIndex = 2 Then
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "30")
                    '    ElseIf rbl_lossItem.SelectedIndex = 3 Then
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "40")
                    '    ElseIf rbl_lossItem.SelectedIndex = 4 Then
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "50")
                    '    Else
                    '        scriptStr = String.Format(scriptStr, (ddlCustomer.SelectedValue.Trim()), (ddlProduct.SelectedValue.Trim()), (ddlPart.SelectedValue.Trim()), failMode, weekStr, "All", product_part, "", "")
                    '    End If
                    '    '("INTEL", "CPU", "SNB P22", "Bump fail", "2013-01-02", "All");
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += listB_BumpingTypeShow.Items(n).Text
                        Else
                            strBumpingType += "," & listB_BumpingTypeShow.Items(n).Text
                        End If
                    Next
                    Dim sGetPartID As String = Get_PartID().Replace("'", "")
                    If rb_dayType.SelectedIndex = 0 Then
                        ' Daily
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                        '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
                    ElseIf rb_dayType.SelectedIndex = 1 Then
                        ' Weekly
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}'), '{12}', '{13}'"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                        '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
                    ElseIf rb_dayType.SelectedIndex = 2 Then
                        ' Monthly
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Monthly_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                        '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
                    End If

                    Chart.Series((weekStr)).Points.AddXY(failMode, failValue)
                    Chart.Series((weekStr)).Points(i).ToolTip = weekStr & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                    'Chart.Series((weekStr)).Points.AddXY(newfailMode, failValue)
                    'Chart.Series((weekStr)).Points(i).ToolTip = weekStr & vbCrLf & "FailMode=" & newfailMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                    Chart.Series((weekStr)).Points(i).Href = scriptStr
                    If cb_DRowData1.Checked = True Then
                        Chart.Series((weekStr)).Points(i).Label = ((failValue.ToString("0.#")) + "%")
                    End If
                Next

                colorInx = (colorInx - 1)

            Next
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        End Try
    End Sub
    Private Sub DrawBarChart_FailModeByStageRatioSummary(ByVal yl As YieldlossInfo, ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable, ByVal TimeColumn As String, Optional ByVal weekTemp As String = "")
        Try
            Chart.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
            Chart.ImageType = ChartImageType.Png
            Chart.Palette = ChartColorPalette.Dundas
            Chart.Height = Unit.Pixel(gChartH)
            Chart.Width = Unit.Pixel(gChartW)

            If rb_dayType.SelectedIndex = 0 Then
                Chart.Titles.Add("Fail Mode By Day")
            ElseIf rb_dayType.SelectedIndex = 1 Then
                Chart.Titles.Add("Fail Mode By Week")
            Else
                Chart.Titles.Add("Fail Mode By Month")
            End If

            Chart.Titles(0).Font = New Font("Arial", 15, FontStyle.Bold)
            Chart.Titles(0).Color = Color.DarkBlue

            If yl.BumpingType <> "" Then
                Chart.Titles.Add("Bumping Type ： " + yl.BumpingType.ToString)
            Else

                If yl.Part_ID.Length > 100 Then
                    yl.Part_ID = yl.Part_ID.Substring(0, 150) + "...."
                End If
                Chart.Titles.Add("PartID ： " + yl.Part_ID.ToString)
            End If

            Chart.Titles(1).Font = New Font("Arial", 15, FontStyle.Bold)
            Chart.Titles(1).Color = Color.DarkBlue

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
            If setupDT.Rows.Count < 11 Then

                Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 18, FontStyle.Bold)
            Else
                Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 15, GraphicsUnit.Pixel)
            End If

            Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet


            Chart.UI.Toolbar.Enabled = False
            Chart.UI.ContextMenu.Enabled = True

            Dim series As Series
            Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, TimeColumn)
            weekGroupDT.DefaultView.Sort = (TimeColumn + " asc")
            weekGroupDT = weekGroupDT.DefaultView.ToTable
            Dim dtFilter As DataTable
            Dim dr As DataRow
            Dim foundRows() As DataRow
            Dim insideRows() As DataRow
            Dim failMode As String
            Dim defectcode As String
            Dim failValue As Double
            Dim weekStr As String
            Dim failObj As FailObj
            Dim colorInx As Integer = 0
            Dim scriptStr As String = ""

            Dim pieChartWeek As String = ""
            For i As Integer = 0 To (weekGroupDT.Rows.Count - 1)
                pieChartWeek += (weekGroupDT.Rows(i)(TimeColumn)).ToString + ","
            Next
            If pieChartWeek.Length > 0 Then
                pieChartWeek = pieChartWeek.Substring(0, (pieChartWeek.Length - 1))
            End If
            ViewState("pieChartWeek") = pieChartWeek

            colorInx = (weekGroupDT.Rows.Count - 1)
            For toolIndex As Integer = 0 To (weekGroupDT.Rows.Count - 1)
                weekStr = (weekGroupDT.Rows(toolIndex)(TimeColumn)).ToString
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                '    foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_ratio_byXoutScrap desc")
                'Else
                foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_Ratio desc")
                'End If
                dtFilter = DtSet.Clone
                For x = 0 To (foundRows.Length - 1)
                    dr = foundRows(x)
                    dtFilter.LoadDataRow(dr.ItemArray, False)
                Next
                dtFilter.CaseSensitive = True

                series = Chart.Series.Add((weekStr))
                series.ChartArea = "Default"
                series.Type = SeriesChartType.Column
                series.Color = aryColor(colorInx)
                series.BorderColor = Color.White
                series.BorderWidth = 1


                Dim product_part As String = "PRODUCT"
                If rb_ProductPart.SelectedIndex = 1 Then
                    product_part = "PART"
                End If

                For i As Integer = 0 To (setupDT.Rows.Count - 1)
                    failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "||")
                    'If yl.BU = "PPS" Then

                    insideRows = DtSet.Select(TimeColumn + "='" + weekStr + "' and Fail_Mode='" + failMode.Replace("||", "''") + "'")
                    'Else
                    '    defectcode = (setupDT.Rows(i)("DefectCode").ToString.Trim()).Replace("'", "||")
                    '    insideRows = DtSet.Select(TimeColumn + "='" + weekStr + "' and DefectCode='" + defectcode.Replace("||", "''") + "'")
                    'End If

                    failValue = 0

                    If insideRows.Length > 0 Then
                        For j As Integer = 0 To (insideRows.Length - 1)
                            'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                            '    If Not IsDBNull(insideRows(j).Item("Fail_ratio_byXoutScrap")) Then
                            '        failValue += CType(insideRows(j).Item("Fail_ratio_byXoutScrap"), Double)
                            '    End If
                            'Else
                            If Not IsDBNull(insideRows(j).Item("Fail_Ratio")) Then
                                failValue += CType(insideRows(j).Item("Fail_Ratio"), Double)
                            End If
                            'End If
                        Next
                    End If

                    If toolIndex = (weekGroupDT.Rows.Count - 1) Then
                        failObj = New FailObj
                        failObj.OriFail_Mode = failMode
                        failObj.Fail_Mode = failMode
                        failObj.Fail_Value = failValue
                        DiffFAry.Add(failObj)
                    ElseIf toolIndex = (weekGroupDT.Rows.Count - 2) Then

                        'If yl.BU = "PPS" Then
                        DiffBHash.Add(failMode, failValue)
                        'Else
                        '    DiffBHash.Add(failMode + defectcode, failValue)
                        'End If



                    End If


                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += listB_BumpingTypeShow.Items(n).Text
                        Else
                            strBumpingType += "," & listB_BumpingTypeShow.Items(n).Text
                        End If
                    Next

                    Dim sLotMerge As String = "True"

                    'If cb_Lot_Merge.Checked = False Then
                    '    sLotMerge = "False"
                    '    If Cb_SF.Checked = True Then
                    '        sLotMerge = "False_SF"
                    '    End If
                    'End If
                    If cb_Lot_Merge.Checked = False Then
                        sLotMerge = "False"
                        If Cb_SF.Checked = True Then
                            sLotMerge += "_SF"
                        End If
                        If Cb_CR.Checked = False Then
                            sLotMerge += "_CR"
                        End If

                    Else
                        sLotMerge = "True"
                        If Cb_SF.Checked = True Then
                            sLotMerge += "_SF"
                        End If
                        If Cb_CR.Checked = False Then
                            sLotMerge += "_CR"
                        End If
                    End If

                    If ckFAI.Checked = True Then
                        sLotMerge += "_FAI"
                    End If

                    Dim sTimePeriod As String = "1"
                    If rb_dayType.SelectedIndex = 0 Then
                        sTimePeriod = "0"
                    ElseIf rb_dayType.SelectedIndex = 1 Then
                        sTimePeriod = "1"
                    Else
                        sTimePeriod = "2"
                    End If

                    Dim failTemp As String = failMode
                    If Cb_Inline.Checked = True Then
                        failMode = "Inline異常報廢"
                    End If


                    Dim sGetPartID As String = Get_PartID().Replace("'", "")
                    If rb_dayType.SelectedIndex = 0 Then
                        ' Daily
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                    ElseIf rb_dayType.SelectedIndex = 1 Then
                        ' Weekly
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}'), '{10,IsXoutScrap}', '{11,BumpingType}'"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                        '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
                    ElseIf rb_dayType.SelectedIndex = 2 Then
                        ' Monthly
                        scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
                        '=====================================================================================================================================================================================
                        'javascript:openWindowWithPost('FailDetail_Monthly_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                        '=====================================================================================================================================================================================
                        If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        Else
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
                            End If
                        End If
                        '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
                    End If


                    failMode = failTemp
                    Chart.Series((weekStr)).Points.AddXY(failMode, failValue)
                    Chart.Series((weekStr)).Points(i).ToolTip = weekStr & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                    Chart.Series((weekStr)).Points(i).Href = scriptStr
                    If cb_DRowData1.Checked = True Then
                        Chart.Series((weekStr)).Points(i).Label = ((failValue.ToString("0.#")) + "%")
                    End If
                Next

                colorInx = (colorInx - 1)
            Next


            Chart.Legends(0).Font = New Font("Arial", 14, GraphicsUnit.Pixel)


        Catch ex As Exception
            Dim sError As String = ex.ToString()
        End Try
    End Sub
    'Private Sub DrawBarChart_FailModeByStageRatioSummary(ByVal yl As YieldlossInfo, ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable, ByVal TimeColumn As String, Optional ByVal weekTemp As String = "")
    '    'Try
    '    Chart.ImageUrl = "temp/Bihon_#SEQ(1000,1)"
    '    Chart.ImageType = ChartImageType.Png
    '    Chart.Palette = ChartColorPalette.Dundas
    '    Chart.Height = Unit.Pixel(gChartH)
    '    Chart.Width = Unit.Pixel(gChartW)

    '    If rb_dayType.SelectedIndex = 0 Then
    '        Chart.Titles.Add("Fail Mode By Day")
    '    Else
    '        Chart.Titles.Add("Fail Mode By Week")
    '    End If

    '    Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
    '    Chart.Titles(0).Color = Color.DarkBlue

    '    If yl.BumpingType <> "" Then
    '        Chart.Titles.Add("Bumping Type ： " + yl.BumpingType.ToString)
    '    Else
    '        Chart.Titles.Add("PartID ： " + yl.Part_ID.ToString)
    '    End If

    '    Chart.Titles(1).Font = New Font("Arial", 12, FontStyle.Bold)
    '    Chart.Titles(1).Color = Color.DarkBlue

    '    Chart.Palette = ChartColorPalette.Dundas
    '    Chart.BackColor = Color.White
    '    Chart.BackGradientEndColor = Color.Peru
    '    Chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss
    '    Chart.BorderStyle = ChartDashStyle.Solid
    '    Chart.BorderWidth = 3
    '    Chart.BorderColor = Color.DarkBlue

    '    Chart.ChartAreas.Add("Default")
    '    Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
    '    Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
    '    Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
    '    Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 14, FontStyle.Bold)
    '    Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
    '    Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

    '    Chart.UI.Toolbar.Enabled = False
    '    Chart.UI.ContextMenu.Enabled = True

    '    Dim series As Series
    '    Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, TimeColumn)
    '    weekGroupDT.DefaultView.Sort = (TimeColumn + " asc")
    '    weekGroupDT = weekGroupDT.DefaultView.ToTable
    '    Dim dtFilter As DataTable
    '    Dim dr As DataRow
    '    Dim foundRows() As DataRow
    '    Dim insideRows() As DataRow
    '    Dim failMode As String
    '    Dim DefectCode As String
    '    Dim failValue As Double
    '    Dim weekStr As String
    '    Dim failObj As FailObj
    '    Dim colorInx As Integer = 0
    '    Dim scriptStr As String = ""

    '    Dim pieChartWeek As String = ""
    '    For i As Integer = 0 To (weekGroupDT.Rows.Count - 1)
    '        pieChartWeek += (weekGroupDT.Rows(i)(TimeColumn)).ToString + ","
    '    Next
    '    If pieChartWeek.Length > 0 Then
    '        pieChartWeek = pieChartWeek.Substring(0, (pieChartWeek.Length - 1))
    '    End If
    '    ViewState("pieChartWeek") = pieChartWeek

    '    colorInx = (weekGroupDT.Rows.Count - 1)
    '    For toolIndex As Integer = 0 To (weekGroupDT.Rows.Count - 1)
    '        weekStr = (weekGroupDT.Rows(toolIndex)(TimeColumn)).ToString
    '        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '        '    foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_ratio_byXoutScrap desc")
    '        'Else
    '        foundRows = DtSet.Select(TimeColumn + "='" + weekStr + "'", "Fail_Ratio desc")
    '        'End If
    '        dtFilter = DtSet.Clone
    '        For x = 0 To (foundRows.Length - 1)
    '            dr = foundRows(x)
    '            dtFilter.LoadDataRow(dr.ItemArray, False)
    '        Next
    '        dtFilter.CaseSensitive = True

    '        series = Chart.Series.Add((weekStr))
    '        series.ChartArea = "Default"
    '        series.Type = SeriesChartType.Column
    '        series.Color = aryColor(colorInx)
    '        series.BorderColor = Color.White
    '        series.BorderWidth = 1

    '        Dim product_part As String = "PRODUCT"
    '        If rb_ProductPart.SelectedIndex = 1 Then
    '            product_part = "PART"
    '        End If
    '        'For i As Integer = 0 To (setupDT.Rows.Count - 1)
    '        'failMode = (DtSet.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "||")
    '        'insideRows = DtSet.Select(TimeColumn + "='" + weekStr + "' and Fail_Mode='" + failMode.Replace("||", "''") + "'")

    '        For i As Integer = 0 To (setupDT.Rows.Count - 1)
    '            failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "||")
    '            DefectCode = (setupDT.Rows(i)("DefectCode").ToString.Trim()).Replace("'", "||")
    '            insideRows = DtSet.Select(TimeColumn + "='" + weekStr + "' and DefectCode='" + DefectCode.Replace("||", "''") + "'")
    '            failValue = 0

    '            If insideRows.Length > 0 Then
    '                For j As Integer = 0 To (insideRows.Length - 1)
    '                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                    '    If Not IsDBNull(insideRows(j).Item("Fail_ratio_byXoutScrap")) Then
    '                    '        failValue += CType(insideRows(j).Item("Fail_ratio_byXoutScrap"), Double)
    '                    '    End If
    '                    'Else
    '                    If Not IsDBNull(insideRows(j).Item("Fail_Ratio")) Then
    '                        failValue += CType(insideRows(j).Item("Fail_Ratio"), Double)
    '                    End If
    '                    'End If
    '                Next
    '            End If

    '            If toolIndex = (weekGroupDT.Rows.Count - 1) Then
    '                failObj = New FailObj
    '                failObj.OriFail_Mode = failMode
    '                failObj.Fail_Mode = failMode
    '                failObj.Fail_Value = failValue
    '                DiffFAry.Add(failObj)
    '            ElseIf toolIndex = (weekGroupDT.Rows.Count - 2) Then
    '                'DiffBHash.Add(failMode, failValue)
    '            End If




    '            Dim strBumpingType As String = ""
    '            For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                If n = 0 Then
    '                    strBumpingType += listB_BumpingTypeShow.Items(n).Text
    '                Else
    '                    strBumpingType += "," & listB_BumpingTypeShow.Items(n).Text
    '                End If
    '            Next

    '            Dim sLotMerge As String = "True"

    '            If cb_Lot_Merge.Checked = False Then
    '                sLotMerge = "False"
    '            End If

    '            Dim sTimePeriod As String = "1"
    '            If rb_dayType.SelectedIndex = 0 Then
    '                sTimePeriod = "0"
    '            ElseIf rb_dayType.SelectedIndex = 1 Then
    '                sTimePeriod = "1"
    '            Else
    '                sTimePeriod = "2"
    '            End If

    '            Dim sGetPartID As String = Get_PartID().Replace("'", "")
    '            If rb_dayType.SelectedIndex = 0 Then
    '                ' Daily
    '                scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
    '                '=====================================================================================================================================================================================
    '                'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
    '                '=====================================================================================================================================================================================
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                Else
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                End If
    '            ElseIf rb_dayType.SelectedIndex = 1 Then
    '                ' Weekly
    '                scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
    '                '=====================================================================================================================================================================================
    '                'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}'), '{10,IsXoutScrap}', '{11,BumpingType}'"
    '                '=====================================================================================================================================================================================
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                Else
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                End If
    '                '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
    '            ElseIf rb_dayType.SelectedIndex = 2 Then
    '                ' Monthly
    '                scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}')"
    '                '=====================================================================================================================================================================================
    '                'javascript:openWindowWithPost('FailDetail_Monthly_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
    '                '=====================================================================================================================================================================================
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "True", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                Else
    '                    If rbl_lossItem.SelectedIndex = 0 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    Else
    '                        scriptStr = String.Format(scriptStr, (sGetPartID), failMode, weekStr, weekTemp, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "0", "False", strBumpingType, sLotMerge, sTimePeriod)
    '                    End If
    '                End If
    '                '("IVB L21", "4W ET short defect", "201329", "201329,201330,201331,201332", "CPU", "ALL")'
    '            End If

    '            Chart.Series((weekStr)).Points.AddXY(failMode, failValue)
    '            Chart.Series((weekStr)).Points(i).ToolTip = weekStr & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
    '            Chart.Series((weekStr)).Points(i).Href = scriptStr
    '            If cb_DRowData1.Checked = True Then
    '                Chart.Series((weekStr)).Points(i).Label = ((failValue.ToString) + "%")
    '            End If
    '        Next

    '        colorInx = (colorInx - 1)
    '    Next
    '    'Catch ex As Exception
    '    '    Dim sError As String = ex.ToString()
    '    'End Try
    'End Sub

    ' Draw Different Bar Chart
    Private Sub DrawDiffBarChart(ByRef Chart As Chart, ByRef DtSet As DataTable, ByRef setupDT As DataTable, ByRef DiffFAry As ArrayList, ByRef DiffBHash As Hashtable, ByVal TimeColumn As String)

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

        If rb_dayType.SelectedIndex = 0 Then
            Chart.Titles.Add("Fail Mode By Day")
        ElseIf rb_dayType.SelectedIndex = 1 Then
            Chart.Titles.Add("Fail Mode By Week")
        Else
            Chart.Titles.Add("Fail Mode By Month")
        End If

        Chart.Titles(0).Font = New Font("Arial", 12, FontStyle.Bold)
        Chart.Titles(0).Color = Color.DarkBlue

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        'Chart.ChartAreas("Default").AxisX.Title = "【 Fail Mode 】"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
       
        If setupDT.Rows.Count < 11 Then

            Chart.ChartAreas("Default").AxisX.LabelStyle.Font = New Font(FontFamily.GenericSansSerif, 18, FontStyle.Bold)
        Else
            Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 15, GraphicsUnit.Pixel)
        End If

        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        'Chart.ChartAreas("Default").AxisY.Interval = 20
        'Chart.ChartAreas("Default").AxisY.Minimum = -100
        'Chart.ChartAreas("Default").AxisY.Maximum = 100
        'Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

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
        obj11.Name = "Increase (Loss)"
        obj11.Style = LegendImageStyle.Marker
        obj11.MarkerColor = Color.Red
        Chart.Legends(0).CustomItems.Add(obj11)

        Dim obj2 As LegendItem = New LegendItem()
        obj2.MarkerSize = 10
        obj2.Name = "Decrease (Gain)"
        obj2.Style = LegendImageStyle.Marker
        obj2.MarkerColor = Color.DodgerBlue
        Chart.Legends(0).CustomItems.Add(obj2)

        Dim weekGroupDT As DataTable = UtilObj.fun_DataTable_SelectDistinct(DtSet, TimeColumn)
        weekGroupDT.DefaultView.Sort = (TimeColumn + " Desc")
        weekGroupDT = weekGroupDT.DefaultView.ToTable
        Dim failMode As String
        Dim failValue, beforeValue, finalValue As Double
        Dim failObj As FailObj
        Dim desStr As String = "Increase (Loss)"

        For i As Integer = 0 To (DiffFAry.Count - 1)

            failObj = CType(DiffFAry(i), FailObj)
            failMode = failObj.Fail_Mode
            failValue = failObj.Fail_Value
            If DiffBHash.Contains(failMode) Then
                beforeValue = CType(DiffBHash(failMode), Double)
            Else
                beforeValue = 0
            End If

            If (failValue >= 0) And (beforeValue >= 0) Then
                finalValue = Math.Round((beforeValue - failValue), 2)
            Else
                finalValue = 0
            End If
            Chart.Series("Diff").Points.AddXY(failObj.OriFail_Mode, finalValue)

            If finalValue <= 0 Then
                desStr = "Increase (Loss)"
                Chart.Series("Diff").Points(i).Color = Color.Red
            Else
                desStr = "Decrease (Gain)"
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
        Dim newfailMode As String
        Dim failValue As Double
        Dim colorInx As Integer = 0
        Dim scriptStr As String = ""
        Dim product_part As String = "PRODUCT"
        If rb_ProductPart.SelectedIndex = 1 Then
            product_part = "PART"
        End If

        colorInx = (plantAry.Length - 1)
        For toolIndex As Integer = 0 To (plantAry.Length - 1)
            'If cb_DRowData0.Checked = True Then '匹配報廢回歸
            '    foundRows = DtSet.Select("Plant='" + plantAry(toolIndex) + "'", "Fail_ratio_byXoutScrap desc")
            'Else
            foundRows = DtSet.Select("Plant='" + plantAry(toolIndex) + "'", "Fail_Ratio desc")
            'End If

            If foundRows.Length > 0 Then

                series = Chart.Series.Add(("Plant" + plantAry(toolIndex)))
                series.ChartArea = "Default"
                series.Type = SeriesChartType.Column
                series.Color = aryColor(colorInx)
                series.BorderColor = Color.White
                series.BorderWidth = 1

                For i As Integer = 0 To (setupDT.Rows.Count - 1)

                    newfailMode = (setupDT.Rows(i)("NewFailMode").ToString.Trim()).Replace("'", "''")
                    failMode = (setupDT.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "''")
                    failValue = 0

                    insideRows = DtSet.Select("Plant='" + plantAry(toolIndex) + "' and NewFailMode='" + newfailMode + "'")
                    If insideRows.Length > 0 Then
                        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        '    If Not IsDBNull(insideRows(0).Item("Fail_ratio_byXoutScrap")) Then
                        '        failValue = CType(insideRows(0).Item("Fail_ratio_byXoutScrap"), Double)
                        '    End If
                        'Else
                        If Not IsDBNull(insideRows(0).Item("Fail_Ratio")) Then
                            failValue = CType(insideRows(0).Item("Fail_Ratio"), Double)
                        End If
                        'End If
                    End If

                    scriptStr = "javascript:openWindowWithPost('YieldPieChart.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}'), '{11}', '{12}', '{13}')"
                    '=====================================================================================================================================================================================
                    'javascript:openWindowWithPost('FailDetail_Monthly_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                    '=====================================================================================================================================================================================
                    Dim sGetPartID As String = Get_PartID().Replace("'", "")
                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, WeekStr, WeekStr, (ddlProduct.SelectedValue), plantAry(toolIndex), ("all"), product_part, "", "10", "True", "")

                    Chart.Series(("Plant" + plantAry(toolIndex))).Points.AddXY(failMode, failValue)
                    Chart.Series(("Plant" + plantAry(toolIndex))).Points(i).ToolTip = "Plant" & plantAry(toolIndex) & vbCrLf & "FailMode=" & failMode & vbCrLf & "Value=" & Math.Round(failValue, 5).ToString
                    Chart.Series(("Plant" + plantAry(toolIndex))).Points(i).Href = scriptStr

                Next
                colorInx = (colorInx - 1)

            End If

        Next

    End Sub

    ' Show Weekly RowData 'MF_Stage, MF_Area, DefectCode_ID
    Private Sub showWeeklyRowData(ByRef sourceDT As DataTable, ByRef chipSetRawDT As DataTable)

        Dim plantAryList As ArrayList
        Dim allWeekDT, allPlantDT As DataTable
        Dim newDT As DataTable = sourceDT.Clone
        Dim nWeek As String = sourceDT.Rows(0)("yearWW")
        Dim dr As DataRow

        ' Step1. 取得最新週的所有 Fail_Mode Item
        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
        '    Dim foundRows As DataRow() = sourceDT.Select("yearWW='" + nWeek + "'", "Fail_ratio_byXoutScrap desc")
        '    For x = 0 To (foundRows.Length - 1)
        '        dr = foundRows(x)
        '        newDT.LoadDataRow(dr.ItemArray, False)
        '    Next
        'Else
        Dim foundRows As DataRow() = sourceDT.Select("yearWW='" + nWeek + "'", "Fail_Ratio desc")
        For x = 0 To (foundRows.Length - 1)
            dr = foundRows(x)
            newDT.LoadDataRow(dr.ItemArray, False)
        Next
        'End If
        newDT.CaseSensitive = True

        ' Step2. 取得所有週數
        allWeekDT = UtilObj.fun_DataTable_SelectDistinct(sourceDT, "yearWW")

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
        workTable.Columns.Add("Bumping Type", Type.GetType("System.String"))
        For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
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
            workRow(4) = newDT.Rows(i)("BumpingType")
            colIndex = 5

            ' === 加週數 ===
            For j As Integer = 0 To (allWeekDT.Rows.Count - 1)
                findRow = sourceDT.Select("NewFailMode='" + newFailModeStr + "' and yearWW='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()

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
            workRow(colIndex) = (rvalue).ToString()
            colIndex += 1

            ' === 加廠別 [ 如果是 Chip Set 有資料才有 ] ===
            For j As Integer = 0 To (plantAryList.Count - 1)

                findRow = chipSetRawDT.Select("NewFailMode='" + newFailModeStr + "' and plant='" + (plantAryList(j).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()
                colIndex += 1
            Next
            workTable.Rows.Add(workRow)

        Next

        If workTable.Rows.Count > 0 Then
            'Dim New_workTable As DataTable = New DataTable()
            'New_workTable.Columns.Add("Fail Mode", Type.GetType("System.String"))
            'New_workTable.Columns.Add("Stage", Type.GetType("System.String"))
            'New_workTable.Columns.Add("Area", Type.GetType("System.String")) ' DefectCode_ID
            'New_workTable.Columns.Add("ID", Type.GetType("System.String"))
            'For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            '    New_workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
            'Next
            'New_workTable.Columns.Add("Delta", Type.GetType("System.String"))

            'Dim dtDistinct As DataTable = workTable.DefaultView.ToTable(True, New String() {"Fail Mode"})
            'For i As Integer = 0 To (dtDistinct.Rows.Count - 1)
            '    For j As Integer = 0 To (workTable.Rows.Count - 1)

            '    Next
            '    New_workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
            'Next
            'Dim New_workRow As DataRow
            'Dim New_findRow As DataRow()

            'ViewState("RowData") = workTable
            'but_Excel.Enabled = True
        End If

        gv_rowdata.DataSource = workTable
        gv_rowdata.DataBind()
        UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)

    End Sub

    ' Show Daily RawData 
    Private Sub showDailyRowData(ByRef sourceDT As DataTable, ByRef chipSetRawDT As DataTable)

        Dim plantAryList As ArrayList
        Dim allWeekDT, allPlantDT As DataTable
        Dim newDT As DataTable = sourceDT.Clone
        Dim nWeek As String = sourceDT.Rows(0)("DataTime")
        Dim dr As DataRow

        ' Step1. 取得最新週的所有 Fail_Mode Item]]
        Dim foundRows As DataRow()

        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
        '    foundRows = sourceDT.Select("DataTime='" + nWeek + "'", "Fail_ratio_byXoutScrap desc")
        'Else
        foundRows = sourceDT.Select("DataTime='" + nWeek + "'", "Fail_Ratio desc")
        'End If

        For x = 0 To (foundRows.Length - 1)
            dr = foundRows(x)
            newDT.LoadDataRow(dr.ItemArray, False)
        Next
        newDT.CaseSensitive = True

        ' Step2. 取得所有週數
        allWeekDT = UtilObj.fun_DataTable_SelectDistinct(sourceDT, "DataTime")

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

        workTable.Columns.Add("Bumping Type", Type.GetType("System.String"))
        workTable.Columns.Add("MF Stage", Type.GetType("System.String"))
        workTable.Columns.Add("MF Area", Type.GetType("System.String"))
        workTable.Columns.Add("Defect Code", Type.GetType("System.String"))
        For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
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
            workRow(1) = newDT.Rows(i)("BumpingType")
            workRow(2) = newDT.Rows(i)("MF_Stage")
            workRow(3) = newDT.Rows(i)("MF_Area")
            workRow(4) = newDT.Rows(i)("DefectCode")
            colIndex = 5
            ' === 加週數 ===
            For j As Integer = 0 To (allWeekDT.Rows.Count - 1)
                findRow = sourceDT.Select("NewFailMode='" + newFailModeStr + "' and DataTime='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()

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
            workRow(colIndex) = (rvalue).ToString()
            colIndex += 1

            ' === 加廠別 [ 如果是 Chip Set 有資料才有 ] ===
            For j As Integer = 0 To (plantAryList.Count - 1)

                findRow = chipSetRawDT.Select("NewFailMode='" + newFailModeStr + "' and plant='" + (plantAryList(j).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()
                colIndex += 1
            Next
            workTable.Rows.Add(workRow)

        Next

        'If workTable.Rows.Count > 0 Then
        '    ViewState("RowData") = workTable
        '    but_Excel.Enabled = True
        'End If

        gv_rowdata.DataSource = workTable
        gv_rowdata.DataBind()

        'gv_rowdata.Columns(1).Visible = False
        UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)

    End Sub

    Private Sub showDailyRowData_FC(ByRef sourceDT As DataTable, ByRef chipSetRawDT As DataTable)

        Dim plantAryList As ArrayList
        Dim allWeekDT, allPlantDT As DataTable
        Dim newDT As DataTable = sourceDT.Clone
        Dim nWeek As String = sourceDT.Rows(0)("DataTime")
        Dim dr As DataRow

        ' Step1. 取得最新週的所有 Fail_Mode Item]]
        Dim foundRows As DataRow()

        'If cb_DRowData0.Checked = True Then '匹配報廢回歸
        '    foundRows = sourceDT.Select("DataTime='" + nWeek + "'", "Fail_ratio_byXoutScrap desc")
        'Else
        foundRows = sourceDT.Select("DataTime='" + nWeek + "'", "Fail_Ratio desc")
        'End If

        For x = 0 To (foundRows.Length - 1)
            dr = foundRows(x)
            newDT.LoadDataRow(dr.ItemArray, False)
        Next
        newDT.CaseSensitive = True

        ' Step2. 取得所有週數
        allWeekDT = UtilObj.fun_DataTable_SelectDistinct(sourceDT, "DataTime")

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

        'workTable.Columns.Add("Bumping Type", Type.GetType("System.String"))
        workTable.Columns.Add("MF Stage", Type.GetType("System.String"))
        workTable.Columns.Add("MF Area", Type.GetType("System.String"))
        workTable.Columns.Add("Defect Code", Type.GetType("System.String"))
        For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
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
        Dim defectcode As String = ""

        Dim colIndex = 0
        For i As Integer = 0 To (newDT.Rows.Count - 1)

            newFailModeStr = newDT.Rows(i)("NewFailMode").Replace("'", "''")
            If Not IsDBNull(newDT.Rows(i)("Fail_Mode")) Then
                failModeStr = newDT.Rows(i)("Fail_Mode").Replace("'", "''")
            End If
            If Not IsDBNull(newDT.Rows(i)("DefectCode")) Then
                defectcode = newDT.Rows(i)("DefectCode").Replace("'", "''")
            End If

            rvalue = 0
            fvalue = 0
            bvalue = 0

            workRow = workTable.NewRow
            workRow(0) = newDT.Rows(i)("Fail_Mode")
            'workRow(1) = newDT.Rows(i)("BumpingType")
            workRow(1) = newDT.Rows(i)("MF_Stage")
            workRow(2) = newDT.Rows(i)("MF_Area")
            If workRow(0) = "FM on SR/Pin/Ink" Or workRow(0) = "S/R/Ink UEM" Or workRow(0) = "Scratch on S/R/Bump/Ink" Then
                workRow(3) = ""
            Else
                workRow(3) = newDT.Rows(i)("DefectCode")
            End If





            colIndex = 4
            ' === 加週數 ===
            For j As Integer = 0 To (allWeekDT.Rows.Count - 1)
                findRow = sourceDT.Select("NewFailMode='" + newFailModeStr + "' and DataTime='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                'findRow = sourceDT.Select("DefectCode='" + defectcode + "' and DataTime='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 3)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()

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
            workRow(colIndex) = (rvalue).ToString()
            colIndex += 1

            ' === 加廠別 [ 如果是 Chip Set 有資料才有 ] ===
            For j As Integer = 0 To (plantAryList.Count - 1)
                findRow = chipSetRawDT.Select("NewFailMode='" + newFailModeStr + "' and plant='" + (plantAryList(j).ToString()) + "'")
                'findRow = chipSetRawDT.Select("DefectCode='" + defectcode + "' and plant='" + (plantAryList(j).ToString()) + "'")
                Try
                    'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    '    rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    'Else
                    rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    'End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()
                colIndex += 1
            Next
            workTable.Rows.Add(workRow)

        Next

        'If workTable.Rows.Count > 0 Then
        '    ViewState("RowData") = workTable
        '    but_Excel.Enabled = True
        'End If

        gv_rowdata.DataSource = workTable
        gv_rowdata.DataBind()

        'gv_rowdata.Columns(1).Visible = False
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
    Protected Sub but_Excel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_Excel.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim customStr As String = ""
        Dim plantStr As String = ""
        Dim partStr As String = ""
        Dim weekStr As String = ""
        Dim itemStr As String = ""
        Dim topStr As String = ""
        Dim myAdapter As SqlDataAdapter
        Dim topDT As DataTable = New DataTable
        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable


        ' --- Bumping Type ---
        Dim strBumpingType As String = ""
        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
            If n = 0 Then
                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
            Else
                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
            End If
        Next

        Dim sGetPartID As String = Get_PartID()
        ' --- DateTime ---

        Dim dateTimeTemp As String = ""
        If rbl_week.SelectedIndex = 1 Then
            For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                dateTimeTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
            Next
            If (dateTimeTemp <> "") Then
                dateTimeTemp = dateTimeTemp.Substring(0, (dateTimeTemp.Length - 1))
            End If


        Else
            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 1 Then
                dateTimeTemp = GetYearWW(Date.Now)
            End If

            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 0 Then
                dateTimeTemp = GetYearDay(Date.Now)
            End If

            If dateTimeTemp = "" And rb_dayType.SelectedIndex = 2 Then
                dateTimeTemp = GetYearMM(Date.Now)
            End If


        End If

        ' --- Yield Loss ID ---
        Dim Failmodeitem As String = ""

        Dim nTop As Integer = 10

        If rbl_lossItem.SelectedIndex = 0 Then
            topStr = "top(10)"
            nTop = 10
        ElseIf rbl_lossItem.SelectedIndex = 1 Then
            topStr = "top(20)"
            nTop = 20
        ElseIf rbl_lossItem.SelectedIndex = 2 Then
            topStr = "top(30)"
            nTop = 30
        ElseIf rbl_lossItem.SelectedIndex = 3 Then
            topStr = "top(40)"
            nTop = 40
        ElseIf rbl_lossItem.SelectedIndex = 4 Then
            topStr = "top(50)"
            nTop = 50
        Else
            nTop = 50
            For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                Failmodeitem += ((lb_LossShow.Items(i).Value).Replace("'", "''")) + ","
            Next
            If (Failmodeitem.Length > 0 AndAlso lb_LossShow.Items.Count > 0) Then
                Failmodeitem = Failmodeitem.Substring(0, (Failmodeitem.Length - 1))

            End If
        End If


        'Alfie--------------------------------------------------------------------------------------------------------------
        Dim yl As New YieldlossInfo
        'yl.BumpingType = Replace(strBumpingType, "'", "")
        yl.BumpingType = strBumpingType
        yl.Part_ID = Replace(sGetPartID, "'", "")
        yl.TimePeriod = rb_dayType.SelectedIndex
        yl.TimeRange = Replace(dateTimeTemp, "'", "")
        yl.xoutscrape = cb_DRowData0.Checked
        yl.nTop = nTop
        yl.Fail_Mode = Replace(Failmodeitem, "'", "")

        Dim sTemp As String = getTotalOriginal_SQL(yl)
        yl.TotalOriginal = Get_WB_TotalOriginal(sTemp)

        Dim workTable As DataTable
        Try

            conn.Open()
            ' --- Yield Loss ID ---
            Dim itemTemp As String = ""
            itemStr = ""



            Dim iTotal As String = "0"


            sqlStr = getTopWBSQL(yl)
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.SelectCommand.CommandTimeout = 3600
            myAdapter.Fill(topDT)

            If topDT.Rows.Count <> 0 Then



                sqlStr = getRawData_SQL(yl)
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.SelectCommand.CommandTimeout = 360
                workTable = New DataTable
                myAdapter.Fill(workTable)
                ' ViewState("RowData") = workTable
                ' but_Excel.Enabled = True
            Else
                lab_wait.Text = ""
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 
                    If rb_dayType.SelectedIndex = 0 Then
                        lab_wait.Text = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + " 無資料, 可使用天數自訂 !"
                    Else
                        lab_wait.Text = "最新一週無資料, 可使用週數自訂 !"
                    End If

                Else
                    ' Custom
                    lab_wait.Text = dateTimeTemp + " 無資料, 可使用天數自訂 !"
                End If
            End If
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        Dim dt As DataTable = workTable

        ExportToExcel(Page, dt, "YieldLoss")



        'Response.AddHeader("content-disposition", "attachment;filename=YieldLoss.csv")
        'Response.ContentType = "application/octet-stream"
        ''Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5")
        ''Response.ContentEncoding = Encoding.UTF8
        'Response.Charset = "UTF-8"
        'Response.ContentEncoding = System.Text.Encoding.UTF8
        'Response.HeaderEncoding = System.Text.Encoding.UTF8
        'Response.ContentType = "text/csv"

        ''Response.Write(sw);  
        ''Response.AppendHeader("content-disposition", "attachment; filename=" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8).Replace("+", "%20"));  
        ''context.Response.Flush();  
        ''context.Response.End();  

        'Dim tw As StringWriter = New System.IO.StringWriter()
        'Dim tmpStr As String = ""

        'Try
        '    ' --- 加欄位 ---
        '    For i As Integer = 0 To (dt.Columns.Count - 1)
        '        tmpStr += dt.Columns(i).ColumnName + ","
        '    Next
        '    tmpStr = tmpStr.Substring(0, (tmpStr.Length - 1))
        '    tw.WriteLine(tmpStr)
        '    ' --- 加資料 ---
        '    For i As Integer = 0 To (dt.Rows.Count - 1)
        '        tmpStr = ""
        '        For j As Integer = 0 To (dt.Columns.Count - 1)
        '            If j = 0 Then
        '                tmpStr += """" + dt.Rows(i)(j).ToString() + ""","
        '            Else
        '                tmpStr += dt.Rows(i)(j).ToString() + ","
        '            End If
        '        Next
        '        tmpStr = tmpStr.Substring(0, (tmpStr.Length - 1))
        '        tw.WriteLine(tmpStr)
        '    Next

        '    'Response.Write(tw)
        '    Dim byteArray As Byte() = System.Text.Encoding.Default.GetBytes(tw.ToString())
        '    Dim str As String = System.Text.Encoding.Default.GetString(byteArray)
        '    Response.BinaryWrite(byteArray)
        '    Response.End()

        'Catch ex As Exception

        'End Try

    End Sub
    Public Sub ExportToExcel(ByVal page As System.Web.UI.Page, ByVal dt As DataTable, ByRef FileName As String)
        If dt Is Nothing Then
            Return
        End If

        Response.ClearContent()
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        If FileName = "" Then
            FileName = "IPP"
        End If

        Response.AddHeader("content-disposition", "attachment;filename=" & Server.UrlEncode(FileName & ".xls"))
        Response.ContentType = "application/excel"
        'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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

    Protected Sub gv_rowdata_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv_rowdata.RowDataBound

        If e.Row.RowType = DataControlRowType.Header Then
            Dim dayAry(e.Row.Cells.Count - 1) As String
            For i As Integer = 0 To (e.Row.Cells.Count - 1)
                dayAry(i) = e.Row.Cells(i).Text
            Next
        End If
        Dim sLotMerge As String = "True"

        'If cb_Lot_Merge.Checked = False Then
        '    sLotMerge = "False"
        'End If
        If cb_Lot_Merge.Checked = False Then
            sLotMerge = "False"
            If Cb_SF.Checked = True Then
                sLotMerge += "_SF"
            End If
            If Cb_CR.Checked = False Then
                sLotMerge += "_CR"
            End If

        Else
            sLotMerge = "True"
            If Cb_SF.Checked = True Then
                sLotMerge += "_SF"
            End If
            If Cb_CR.Checked = False Then
                sLotMerge += "_CR"
            End If
        End If

        If ckFAI.Checked = True Then
            sLotMerge += "_FAI"
        End If


        If sLotMerge.IndexOf("CR") > 0 Then
            Dim ccc As Integer
            ccc += 1
        End If

        Dim sTimePeriod As String = "1"
        If rb_dayType.SelectedIndex = 0 Then
            sTimePeriod = "0"
        ElseIf rb_dayType.SelectedIndex = 1 Then
            sTimePeriod = "1"
        Else
            sTimePeriod = "2"
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then

            Dim scriptStr As String = ""
            Dim failMode As String = (e.Row.Cells(0).Text.Trim).Replace("&#39;", "000")
            Dim DefectCode As String = "" '(e.Row.Cells(3).Text.Trim).Replace("&#39;", "000")

            Dim pieChartWeek As String = ViewState("pieChartWeek").ToString ' 所有選擇的 Week
            Dim startPlant As Boolean = False
            Dim headerStr As String = ""
            Dim newWeekInt As Integer = 0
            Dim deltaDoublea As Double = 0
            Dim product_part As String = "PRODUCT"
            If rb_ProductPart.SelectedIndex = 1 Then
                product_part = "PART"
            End If

            Dim iStart As Integer = 0
            If rb_dayType.SelectedIndex = 0 Then
                iStart = 4
            ElseIf rb_dayType.SelectedIndex = 1 Then
                iStart = 4
            ElseIf rb_dayType.SelectedIndex = 2 Then
                iStart = 4
            End If

            If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
                iStart = 5
            End If

            Dim failTemp As String = failMode
            If Cb_Inline.Checked = True Then
                failMode = "Inline異常報廢"
            End If

            For i As Integer = iStart To (e.Row.Cells.Count - 1)

                deltaDoublea = 0
                headerStr = gv_rowdata.HeaderRow.Cells(i).Text

                If (headerStr.ToUpper) = "DELTA" Then
                    startPlant = True
                    newWeekInt = (i - 1)
                    deltaDoublea = CType(e.Row.Cells(i).Text, Double)
                    If deltaDoublea > 0 Then
                        e.Row.Cells(i).ForeColor = Color.Blue
                    ElseIf deltaDoublea < 0 Then
                        e.Row.Cells(i).ForeColor = Color.Red
                    End If
                    e.Row.Cells(i).Text = e.Row.Cells(i).Text + "%"
                Else
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += listB_BumpingTypeShow.Items(n).Text
                        Else
                            strBumpingType += "," & listB_BumpingTypeShow.Items(n).Text
                        End If
                    Next
                    Dim sGetPartID As String = Get_PartID().Replace("'", "")
                    If rb_dayType.SelectedIndex = 0 Then
                        If startPlant Then
                            ' 依廠別呈現資料
                            Dim newWeekStr As String = gv_rowdata.HeaderRow.Cells(newWeekInt).Text ' 要取得最新的 Week
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}"", ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                           
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            End If
                            'End If
                        Else
                            ' 依週數日期呈現資料
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}"", ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Monthly_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                           
                            If rbl_lossItem.SelectedIndex = 0 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            Else
                                scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                            End If
                            'End If
                        End If
                        e.Row.Cells(i).Text = scriptStr + (e.Row.Cells(i).Text) + "%</a>"
                    ElseIf rb_dayType.SelectedIndex = 1 Then
                        If startPlant Then
                            ' 依廠別呈現資料
                            Dim newWeekStr As String = gv_rowdata.HeaderRow.Cells(newWeekInt).Text ' 要取得最新的 Week
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}""), ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            Else
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            End If
                        Else
                            ' 依週數日期呈現資料
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}"", ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            Else
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            End If
                        End If
                        e.Row.Cells(i).Text = scriptStr + (e.Row.Cells(i).Text) + "%</a>"
                    ElseIf rb_dayType.SelectedIndex = 2 Then
                        If startPlant Then
                            ' 依廠別呈現資料
                            Dim newWeekStr As String = gv_rowdata.HeaderRow.Cells(newWeekInt).Text ' 要取得最新的 Week
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}"", ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}', '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            Else
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, newWeekStr, pieChartWeek, (ddlProduct.SelectedValue), headerStr, ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            End If
                        Else
                            ' 依週數日期呈現資料
                            scriptStr = "<a href='#' onclick='javascript:openWindowWithPost(""FailDetail_Test.aspx"", ""WEB"", ""{0}"",""{1}"",""{2}"",""{3}"", ""{4}"", ""{5}"", ""{6}"", ""{7}"", ""{8}"", ""{9}"", ""{10}"", ""{11}"", ""{12}"", ""{13}"", ""{14}"")'>"
                            '=====================================================================================================================================================================================
                            'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}'), '{10,IsXoutScrap}', '{11,BumpingType}')"
                            '=====================================================================================================================================================================================
                            If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "True", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            Else
                                If rbl_lossItem.SelectedIndex = 0 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "10", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "20", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "30", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "40", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "50", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                Else
                                    scriptStr = String.Format(scriptStr, (sGetPartID), failMode, headerStr, pieChartWeek, (ddlProduct.SelectedValue), "All", ("all"), product_part, "", "", "False", strBumpingType, sLotMerge, sTimePeriod, DefectCode)
                                End If
                            End If
                        End If
                        e.Row.Cells(i).Text = scriptStr + (e.Row.Cells(i).Text) + "%</a>"
                    End If

                End If

            Next

            'e.Row.Height = Unit.Pixel(30)
            'e.Row.Cells(0).Font.Size = FontUnit.XXSmall
            'e.Row.Cells(0).Font.Bold = True
            e.Row.Cells(0).ForeColor = Drawing.Color.Black

        End If

    End Sub

#Region " 上傳 Lot 分析"

    Protected Sub cb_uploadLot_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cb_uploadLot.CheckedChanged

        but_Execute.Enabled = False
        tr_upload.Visible = False
        tb_uploadLot.Visible = False
        tr_gvDisplay.Visible = False
        but_Excel.Visible = True

        If cb_uploadLot.Checked Then
            tr_upload.Visible = True
            but_Excel.Visible = False
        Else
            but_Execute.Enabled = True
            but_Excel.Visible = True
        End If

    End Sub

    ' 解析 Excel or TXT
    Protected Sub but_Uupload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_Uupload.Click

        Dim nFileName As String
        Dim file As StreamReader
        Dim line As String
        Dim dtr As DataRow
        Dim saveFullName As String = Page.MapPath(".") + "\\upload\\"

        If (uf_UfilePath.HasFile) Then

            Dim fileName As String = uf_UfilePath.FileName
            saveFullName += fileName
            uf_UfilePath.SaveAs(saveFullName)

            Dim fileN As FileInfo = New FileInfo(saveFullName)

            If (fileN.Extension).ToUpper = ".TXT" Or (fileN.Extension).ToUpper = ".CSV" Then

                Try
                    Dim lotAry As New ArrayList
                    file = New StreamReader(saveFullName, System.Text.Encoding.Default)
                    line = file.ReadLine()
                    While line <> Nothing
                        lotAry.Add(line.Trim)
                        line = file.ReadLine()
                    End While
                    file.Close()
                    handTXTFile(lotAry)
                Catch ex As Exception
                End Try

            Else
                ShowMessage("請上傳 TXT or CSV !!!")
            End If

        End If

    End Sub

    Private Sub handExcelFile(ByVal fName As String)

        Dim strCon As String = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + fName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;'"
        Dim sqlStr As String = ""
        Dim oledb_con As New OleDbConnection(strCon)
        Dim ole_apt As OleDbDataAdapter
        Dim ds As DataSet
        Dim dt As DataTable
        Dim rowCount As Integer = 0
        Dim excelSheetName As String = "Sheet1"
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdapter As SqlDataAdapter

        Try

            oledb_con.Open()
            ' Get Sheet Name 
            dt = New DataTable
            dt = oledb_con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            If dt Is Nothing Then
                Exit Try
            End If

            excelSheetName = dt.Rows(0)("TABLE_NAME")
            sqlStr = "SELECT * FROM [" + excelSheetName + "]"
            ole_apt = New OleDbDataAdapter(sqlStr, oledb_con)
            ds = New DataSet()
            ole_apt.Fill(ds)
            dt = ds.Tables(0)
            oledb_con.Close()
            Dim lot_list As String = ""

            For i As Integer = 0 To (dt.Rows.Count - 1)

                lot_list += "'" + (dt.Rows(i)(0)).ToString() + "',"

            Next
            lot_list = lot_list.Substring(0, (lot_list.Length - 1))
            lot_list = "(" + lot_list + ")"

            sqlStr = "select ww, Convert(char(19), DataTime, 120) as DataTime, "
            'If cb_DRowData0.Checked = True Then '匹配報廢回歸
            '    sqlStr += "Part_Id, production_id, Lot_Id, Customer_Id, Fail_Mode, Original_Input_QTY, Fail_Count_byXoutScrap, round(Fail_ratio_byXoutScrap,5) as Fail_ratio_byXoutScrap, 'ALL' as category "
            'Else
            sqlStr += "Part_Id, production_id, Lot_Id, Customer_Id, Fail_Mode, Original_Input_QTY, Fail_Count, round(Fail_ratio,5) as Fail_ratio, 'ALL' as category "
            'End If
            sqlStr += "from dbo.VW_BinCode_Daily_Lot "
            sqlStr += "where Lot_Id in {0} "
            sqlStr += "union "
            sqlStr += "select WW, Convert(char(19), DataTime, 120) as DataTime, "
            'If cb_DRowData0.Checked = True Then '匹配報廢回歸
            '    sqlStr += "Part_Id, production_id, Lot_Id, Customer_Id, Fail_Mode, Original_Input_QTY, Fail_Count_byXoutScrap, round(Fail_ratio_byXoutScrap,5) as Fail_ratio_byXoutScrap, category "
            'Else
            sqlStr += "Part_Id, production_id, Lot_Id, Customer_Id, Fail_Mode, Original_Input_QTY, Fail_Count, round(Fail_ratio,5) as Fail_ratio, category "
            'End If
            sqlStr += "from dbo.VW_BinCode_Detail_Daily_Lot "
            sqlStr += "where Lot_Id in {1} "
            sqlStr += "order by lot_id "
            sqlStr = String.Format(sqlStr, lot_list, lot_list)

            conn.Open()
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            dt = New DataTable()
            myAdapter.Fill(dt)
            conn.Close()

            If dt.Rows.Count > 0 Then
                tb_uploadLot.Visible = True
                Lot_GridView.DataSource = dt
                Lot_GridView.DataBind()
                UtilObj.Set_DataGridRow_OnMouseOver_Color(Lot_GridView, "#FFF68F", Lot_GridView.AlternatingRowStyle.BackColor)
            End If

        Catch ex As Exception

        Finally
            If oledb_con.State = ConnectionState.Open Then
                oledb_con.Close()
            End If
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

#Region "上傳 Lot 分析"

    Private Sub handTXTFile(ByRef lotAry As ArrayList)

        Dim sqlStr As String = ""
        Dim ds As DataSet
        Dim dt As DataTable
        Dim rowCount As Integer = 0
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim myAdapter As SqlDataAdapter
        Dim lot_list As String = ""
        Dim lotList As String = ""

        Try

            For i As Integer = 0 To (lotAry.Count - 1)
                lot_list += "'" + (lotAry(i)).ToString() + "',"
                lotList += (lotAry(i)).ToString() + ","
            Next
            lot_list = lot_list.Substring(0, (lot_list.Length - 1))
            lotList = lotList.Substring(0, (lotList.Length - 1))
            lot_list = "(" + lot_list + ")"

            ' --- Row Data SQL --- 
            sqlStr = "select distinct  "
            sqlStr += "ww, Convert(char(19), DataTime, 120) as DataTime, "
            'If cb_DRowData0.Checked = True Then '匹配報廢回歸
            '    sqlStr += "Part_Id, production_type, Lot_Id, Customer_Id, Fail_Mode, MF_Stage, DefectCode, BinCode, Original_Input_QTY, Fail_Count_byXoutScrap, round(Fail_ratio_byXoutScrap, 5) as Fail_ratio_byXoutScrap, FE_Plant, BE_Plant "
            'Else
            sqlStr += "Part_Id, production_type, Lot_Id, Customer_Id, Fail_Mode, MF_Stage, DefectCode, BinCode, Original_Input_QTY, Fail_Count, round(Fail_ratio, 5) as Fail_ratio, FE_Plant, BE_Plant "
            'End If
            sqlStr += "from dbo.VW_BinCode_Daily_Lot "
            sqlStr += "where 1=1 "
            sqlStr += "and Lot_Id in {0} "
            sqlStr += "order by lot_id, MF_Stage, Fail_Mode"
            sqlStr = String.Format(sqlStr, lot_list)

            conn.Open()
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            dt = New DataTable()
            myAdapter.Fill(dt)

            If dt.Rows.Count > 0 Then

                ' --- Lot Raw Data
                tb_uploadLot.Visible = True
                Lot_GridView.DataSource = dt
                Lot_GridView.DataBind()
                UtilObj.Set_DataGridRow_OnMouseOver_Color(Lot_GridView, "#FFF68F", Lot_GridView.AlternatingRowStyle.BackColor)
                ' --- Chart 

                sqlStr = "Select SUM(Original_Input_QTY) as total_count "
                sqlStr += "from ( Select   Original_Input_QTY "
                sqlStr += "from VW_BinCode_Daily_RawData "
                sqlStr += "where Lot_Id in {0} "
                sqlStr += "group by lot_id, Original_Input_QTY )a "
                sqlStr = String.Format(sqlStr, lot_list)
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                dt = New DataTable()
                myAdapter.Fill(dt)

                'If cb_DRowData0.Checked = True Then '匹配報廢回歸(XoutScrap)
                '    'Fail_Count_byXoutScrap
                '    sqlStr = "SELECT TOP(20) Fail_Mode, round((convert(float, SUM(Fail_Count_byXoutScrap))/" + (dt.Rows(0)("total_count")).ToString() + "), 4) * 100 as Ratio "
                'Else
                sqlStr = "SELECT TOP(20) Fail_Mode, round((convert(float, SUM(Fail_Count))/" + (dt.Rows(0)("total_count")).ToString() + "), 4) * 100 as Ratio "
                'End If
                sqlStr += "FROM VW_BinCode_Daily_Lot "
                sqlStr += "WHERE 1=1 "
                sqlStr += "AND Lot_Id IN {0} "
                If cb_Non8K.Checked = True Then
                    sqlStr += "AND Fail_Mode NOT LIKE '8K%' "
                End If

                If cb_NonIPQC.Checked = True Then
                    sqlStr += "AND Fail_Mode NOT LIKE 'IPQC%' "
                End If

                sqlStr += "GROUP BY Fail_Mode "
                sqlStr += "ORDER BY Ratio desc"
                sqlStr = String.Format(sqlStr, lot_list)
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                dt = New DataTable()
                myAdapter.Fill(dt)

                sqlStr = "SELECT b.yearWW, a.Production_Type, a.Category "
                sqlStr += "FROM VW_BinCode_Daily_Lot a, SystemDateMapping b "
                sqlStr += "WHERE b.DateTime = a.trtm "
                sqlStr += "AND Lot_Id IN {0} "
                sqlStr += "GROUP BY b.yearWW, a.Production_Type, a.Category "
                sqlStr += "ORDER BY b.yearWW DESC"
                sqlStr = String.Format(sqlStr, lot_list)
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                Dim newDT As DataTable = New DataTable()
                myAdapter.Fill(newDT)

                ' 畫 Chart
                upLoadLotChart(dt, newDT, lotList)

            End If
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub upLoadLotChart(ByRef dt As DataTable, ByRef newDT As DataTable, ByVal lot_list As String)

        Dim scriptStr As String = ""
        Dim productionType As String = ""
        Dim weekStr As String = ""
        Dim productStr As String = ""
        Dim pieChartWeek As String = ""

        Try
            productionType = (newDT.Rows(0)("Production_Type")).ToString()
            weekStr = (newDT.Rows(0)("yearWW")).ToString()
            productStr = (newDT.Rows(0)("Category")).ToString()
            pieChartWeek = ""
            Dim weekDT_Distinct As DataTable = newDT.DefaultView.ToTable(True, New String() {"yearWW"})
            For Each dr As DataRow In weekDT_Distinct.Rows
                pieChartWeek += (dr("yearWW") + ",")
            Next
            pieChartWeek = pieChartWeek.Substring(0, (pieChartWeek.Length - 1))
        Catch ex As Exception
            weekStr = ""
        End Try


        Dim Chart As New Dundas.Charting.WebControl.Chart()
        Chart = New Dundas.Charting.WebControl.Chart()
        Chart.ImageUrl = "temp/ULOTChart_#SEQ(1000,1)"
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

        Chart.ChartAreas.Add("Default")
        Chart.ChartAreas("Default").AxisY.LabelStyle.Format = "P2"
        Chart.ChartAreas("Default").AxisX.LabelStyle.Interval = 1
        Chart.ChartAreas("Default").AxisX.LabelStyle.FontAngle = -45 '文字對齊
        Chart.ChartAreas("Default").BorderStyle = ChartDashStyle.NotSet
        Chart.ChartAreas("Default").AxisY.LabelStyle.Font = New Font("Arial", 14, GraphicsUnit.Pixel)

        Chart.UI.Toolbar.Enabled = False
        Chart.UI.ContextMenu.Enabled = True

        Dim series As Series
        series = Chart.Series.Add("FailMode")
        series.ChartArea = "Default"
        series.Type = SeriesChartType.Column
        series.Color = Color.DodgerBlue
        series.ShowInLegend = False

        Dim failMode As String
        Dim failValue As Double
        Dim product_part As String = "PRODUCT"
        If rb_ProductPart.SelectedIndex = 1 Then
            product_part = "PART"
        End If

        For i As Integer = 0 To (dt.Rows.Count - 1)

            failMode = (dt.Rows(i)("Fail_Mode").ToString.Trim()).Replace("'", "''")
            failValue = 0
            If Not IsDBNull(dt.Rows(i)("Ratio")) Then
                failValue = CType(dt.Rows(i)("Ratio"), Double)
            End If

            Chart.Series("FailMode").Points.AddXY(failMode, failValue)
            Chart.Series("FailMode").Points(i).Label = ((failValue.ToString) + "%")
            If weekStr.Length > 0 Then
                Dim strBumpingType As String = ""
                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                    If n = 0 Then
                        strBumpingType += listB_BumpingTypeShow.Items(n).Text
                    Else
                        strBumpingType += "," & listB_BumpingTypeShow.Items(n).Text
                    End If
                Next

                scriptStr = "javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}')"
                '=====================================================================================================================================================================================
                'javascript:openWindowWithPost('FailDetail_Test.aspx', 'WEB', '{0,P}', '{1,F}', '{2,W}', '{3,WI}', '{4,Product}', '{5,Plant}', '{6,Customer}', '{7,TYPE}', '{8,LotList}', '{9,TopN}'), '{10,IsXoutScrap}', '{11,BumpingType}'"
                '=====================================================================================================================================================================================
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸(IsXoutScrap)
                '    If rbl_lossItem.SelectedIndex = 0 Then
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "10", "True", strBumpingType)
                '    ElseIf rbl_lossItem.SelectedIndex = 1 Then
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "20", "True", strBumpingType)
                '    ElseIf rbl_lossItem.SelectedIndex = 2 Then
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "30", "True", strBumpingType)
                '    ElseIf rbl_lossItem.SelectedIndex = 3 Then
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "40", "True", strBumpingType)
                '    ElseIf rbl_lossItem.SelectedIndex = 4 Then
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "50", "True", strBumpingType)
                '    Else
                '        scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", (ddlCustomer.SelectedValue.Trim()), product_part, lot_list, "", "True", strBumpingType)
                '    End If
                'Else
                Dim sTimePeriod As String = "1"
                If rb_dayType.SelectedIndex = 0 Then
                    sTimePeriod = "0"
                ElseIf rb_dayType.SelectedIndex = 1 Then
                    sTimePeriod = "1"
                Else
                    sTimePeriod = "2"
                End If

                Dim sLotMerge As String = "True"

                If cb_Lot_Merge.Checked = False Then
                    sLotMerge = "False"
                End If

                If rbl_lossItem.SelectedIndex = 0 Then
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "10", "False", strBumpingType, sLotMerge, sTimePeriod)
                ElseIf rbl_lossItem.SelectedIndex = 1 Then
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "20", "False", strBumpingType, sLotMerge, sTimePeriod)
                ElseIf rbl_lossItem.SelectedIndex = 2 Then
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "30", "False", strBumpingType, sLotMerge, sTimePeriod)
                ElseIf rbl_lossItem.SelectedIndex = 3 Then
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "40", "False", strBumpingType, sLotMerge, sTimePeriod)
                ElseIf rbl_lossItem.SelectedIndex = 4 Then
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "50", "False", strBumpingType, sLotMerge, sTimePeriod)
                Else
                    scriptStr = String.Format(scriptStr, productionType, failMode, weekStr, pieChartWeek, productStr, "All", ("all"), product_part, lot_list, "", "False", strBumpingType, sLotMerge, sTimePeriod)
                End If
                'End If
                Chart.Series("FailMode").Points(i).Href = scriptStr
            End If

        Next

        UploadLot_Chart.Controls.Add(Chart)

    End Sub

    'Private Sub monthly_failMode()

    '    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '    Dim sqlStr As String = ""
    '    Dim customStr As String = " "
    '    Dim plantStr As String = ""
    '    Dim partStr As String = " "
    '    Dim weekStr As String = " "
    '    Dim itemStr As String = " "
    '    Dim topStr As String = " "
    '    Dim myAdapter As SqlDataAdapter
    '    Dim topDT As DataTable = New DataTable
    '    Dim new_topDT As DataTable = New DataTable
    '    Dim rawDT, chipSetRawDT As DataTable

    '    If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 12 Then
    '        ShowMessage("選擇月數最多為 12 月")
    '        Exit Sub
    '    End If

    '    Try
    '        ' --- Customer ID ---
    '        ' customStr = "and customer_id='" + (ddlCustomer.SelectedValue.Trim()) + "' "
    '        ' --- Part ID ---
    '        'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
    '        Dim strBumpingType As String = ""
    '        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '            If n = 0 Then
    '                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '            Else
    '                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '            End If
    '        Next
    '        Dim BumpingType_Part As String = ""
    '        Dim sGetPartID As String = Get_PartID()
    '        If strBumpingType <> "" Then
    '            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '        End If
    '        If tr_BumpingType.Visible = False Then
    '            If rb_ProductPart.SelectedIndex = 0 Then
    '                If listB_BumpingTypeShow.Enabled = False Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Production_id={0} "
    '                    Else
    '                        partStr += "{0}"
    '                    End If
    '                Else
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Production_id={0} "
    '                    Else
    '                        partStr += "{0}"
    '                    End If
    '                End If
    '            ElseIf rb_ProductPart.SelectedIndex = 1 Then
    '                Dim oPartID = sGetPartID.Split(",")
    '                If oPartID.Length = 1 Then
    '                    If listB_BumpingTypeShow.Enabled = False Then
    '                        If sGetPartID <> "" Then
    '                            partStr += "and Part_Id = {0} "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    Else
    '                        If BumpingType_Part <> "" Then
    '                            partStr += "and Part_Id = {0} "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    End If
    '                ElseIf oPartID.Length > 1 Then
    '                    If listB_BumpingTypeShow.Enabled = False Then
    '                        If sGetPartID <> "" Then
    '                            partStr += "and Part_Id in({0}) "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    Else
    '                        If BumpingType_Part <> "" Then
    '                            partStr += "and Part_Id in({0}) "
    '                        Else
    '                            partStr += "{0}"
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Else
    '            If BumpingType_Part <> "" Then
    '                partStr += "and Part_Id in (" & BumpingType_Part & ") "
    '            Else
    '                Dim oPartID = sGetPartID.Split(",")
    '                If oPartID.Length = 1 Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Part_Id = " + sGetPartID
    '                    End If
    '                ElseIf oPartID.Length > 1 Then
    '                    If sGetPartID <> "" Then
    '                        partStr += "and Part_Id in(" + sGetPartID + ") "
    '                    End If
    '                End If
    '            End If
    '        End If
    '        If ddlProduct.SelectedValue <> "CPU" Then
    '            plantStr = "and Plant='All' "
    '        Else
    '            plantStr = ""
    '        End If

    '        Dim Month1 As String = DateTime.Now.ToString("yyyyMM")
    '        Dim Month2 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day).Day
    '                                                     ) + 1).ToString("yyyyMM")
    '        Dim Month3 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day).Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day
    '                                                     ) + 1).ToString("yyyyMM")
    '        Dim Month4 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day).Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day +
    '                                                      DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day - DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day).Day
    '                                                     ) + 1).ToString("yyyyMM")

    '        ' --- Month ID ---
    '        Dim weekTemp As String = ""
    '        If rbl_week.SelectedIndex = 1 Then
    '            For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
    '                'weekTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
    '                weekTemp += (lb_weekShow.Items(i).Value) + ","
    '            Next
    '            weekTemp = weekTemp.Substring(0, (weekTemp.Length - 1))
    '            weekStr = "and MM in (" + weekTemp + ") "
    '        Else
    '            ' Defaule 4 week
    '            'sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
    '            'weekTemp = "'" + Month1 + "','" + Month2 + "','" + Month3 + "','" + Month4 + "'"
    '            weekTemp = "and MM in (select top(4) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
    '        End If

    '        If rbl_lossItem.SelectedIndex = 0 Then
    '            topStr = "top(10)"
    '        ElseIf rbl_lossItem.SelectedIndex = 1 Then
    '            topStr = "top(20)"
    '        ElseIf rbl_lossItem.SelectedIndex = 2 Then
    '            topStr = "top(30)"
    '        ElseIf rbl_lossItem.SelectedIndex = 3 Then
    '            topStr = "top(40)"
    '        ElseIf rbl_lossItem.SelectedIndex = 4 Then
    '            topStr = "top(50)"
    '        End If

    '        conn.Open()
    '        ' --- Yield Loss ID ---
    '        Dim itemTemp As String = ""
    '        itemStr = ""

    '        If rbl_lossItem.SelectedIndex = 5 Then
    '            ' === Custom Item ===
    '            For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
    '                itemTemp += "'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "',"
    '            Next
    '            itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
    '            itemStr = "and fail_mode in (" + itemTemp + ") "

    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr = "select Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            Else
    '                sqlStr = "select Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            End If
    '            sqlStr = "from dbo.WB_BinCode_Summary_Monthly where 1=1 "

    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            sqlStr += itemStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 month
    '                'sqlStr += "and MM in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                'sqlStr += "and MM in (" + weekTemp + ") "
    '                sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
    '            Else
    '                ' Custom
    '                'sqlStr += "and MM in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                'sqlStr += weekStr
    '                sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly where 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
    '            End If
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                sqlStr += "order by Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
    '                sqlStr += "order by fail_ratio desc"
    '            End If
    '        Else
    '            ' === Top N === ' 如果選 Custom 就要呈現選擇的 item
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr = "select " + topStr + " Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            Else
    '                sqlStr = "select " + topStr + " Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
    '            End If
    '            sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "

    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            If cb_NonIPQC.Checked Then
    '                sqlStr += "and Fail_Mode <> 'IPQC defect' "
    '            End If
    '            If cb_Non8K.Checked Then
    '                sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
    '            End If
    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 week
    '                'sqlStr += "and MM in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
    '                'sqlStr += "and MM in (" + weekTemp + ") "
    '                sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
    '            Else
    '                ' Custom
    '                'sqlStr += "and MM in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
    '                'sqlStr += weekStr
    '                sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
    '            End If
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                sqlStr += "order by Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
    '                sqlStr += "order by fail_ratio desc"
    '            End If
    '        End If
    '        myAdapter = New SqlDataAdapter(sqlStr, conn)
    '        myAdapter.Fill(topDT)

    '        If (topDT.Rows.Count <> 0) Then

    '            ' === Raw Data ===
    '            lab_wait.Text = ""
    '            'sqlStr = "select a.MM, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
    '            'sqlStr += "from dbo.WB_BinCode_Summary_Monthly a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr = "select MM, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, MF_Area, DefectCode_ID, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode' "
    '            Else
    '                sqlStr = "select MM, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode' "
    '            End If
    '            sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
    '            ' --- WB 資料當時沒有 Customer_ID 2014/01/07 ---
    '            'sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
    '            ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '            'sqlStr += "and a.DefectCode_ID Not IN ('C9', '02') "
    '            sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '            If cb_NonIPQC.Checked Then
    '                sqlStr += "and Fail_Mode <> 'IPQC defect' "
    '            End If
    '            If cb_Non8K.Checked Then
    '                sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
    '            End If
    '            sqlStr += customStr
    '            sqlStr += partStr
    '            sqlStr += plantStr
    '            If rbl_week.SelectedIndex = 0 Then
    '                ' Defaule 4 week
    '                'sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
    '                'sqlStr += "and MM in ('" + weekTemp + "') "
    '                sqlStr += "and MM in (select TOP(4) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
    '            Else
    '                ' Custom
    '                'sqlStr += weekStr
    '                sqlStr += "and MM in (select TOP(4) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
    '            End If

    '            sqlStr += itemStr
    '            'sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
    '            'sqlStr += "order by b.yearWW desc, a.fail_ratio desc"
    '            If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                sqlStr += "group by MM, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, MF_Area, DefectCode_ID "
    '                sqlStr += "order by MM desc, Fail_ratio_byXoutScrap desc"
    '            Else
    '                sqlStr += "group by MM, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID "
    '                sqlStr += "order by MM desc, fail_ratio desc"
    '            End If
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            rawDT = New DataTable
    '            myAdapter.Fill(rawDT)

    '            ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
    '            chipSetRawDT = New DataTable
    '            If ddlProduct.SelectedValue = "CS" Then
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                    sqlStr = "select MM, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
    '                Else
    '                    sqlStr = "select MM, plant, Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
    '                End If
    '                sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
    '                ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
    '                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
    '                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
    '                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

    '                sqlStr += customStr
    '                sqlStr += partStr

    '                If rbl_week.SelectedIndex = 0 Then
    '                    'sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW ORDER BY yearWW DESC) "
    '                    'sqlStr += "and MM in ('" + weekTemp + "') "
    '                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
    '                Else
    '                    'sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW DESC) "
    '                    'sqlStr += weekStr
    '                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
    '                End If

    '                sqlStr += itemStr
    '                If cb_DRowData0.Checked = True Then '匹配報廢回歸
    '                    sqlStr += "group by MM, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
    '                    sqlStr += "order by plant desc, Fail_ratio_byXoutScrap desc"
    '                Else
    '                    sqlStr += "group by MM, plant, Fail_Mode, Fail_Ratio, MF_Stage "
    '                    sqlStr += "order by plant desc, fail_ratio desc"
    '                End If
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myAdapter.Fill(chipSetRawDT)

    '            End If

    '            conn.Close()
    '            BarChart_MM(rawDT, topDT, weekTemp)

    '            If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
    '                Chart_Panel.Controls.Add(New LiteralControl("<br>"))
    '                Dim Chart As New Dundas.Charting.WebControl.Chart()
    '                DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
    '                Chart_Panel.Controls.Add(Chart)
    '            End If

    '            tr_chartDisplay.Visible = True
    '            If cb_DRowData.Checked Then
    '                'showWeeklyRowData(rawDT, chipSetRawDT)
    '                showMonthlyRowData(rawDT, chipSetRawDT)
    '                tr_gvDisplay.Visible = True
    '            End If

    '        Else

    '            Dim inDT As DataTable = New DataTable()
    '            Dim sMonth = DateTime.Now.ToString("yyyyMM")
    '            sqlStr = "SELECT MM FROM WB_BinCode_Summary_Monthly WHERE DateTime='" + sMonth + "'"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myAdapter.Fill(inDT)
    '            lab_wait.Text = sMonth + " 無資料, 可使用月數自訂 !"

    '        End If

    '    Catch ex As Exception

    '    Finally
    '        If conn.State = ConnectionState.Open Then
    '            conn.Close()
    '        End If
    '    End Try

    'End Sub

    Private Sub monthly_failMode()

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
        Dim new_topDT As DataTable = New DataTable
        Dim rawDT, chipSetRawDT As DataTable

        If rbl_week.SelectedIndex = 1 And lb_weekShow.Items.Count > 12 Then
            ShowMessage("選擇月數最多為 12 月")
            Exit Sub
        End If

        Try
            ' --- Customer ID ---
            ' customStr = "and customer_id='" + (ddlCustomer.SelectedValue.Trim()) + "' "
            ' --- Part ID ---
            If tr_BumpingType.Visible = False Then
                'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
                Dim sGetPartID As String = Get_PartID()
                If (sGetPartID <> "") Then
                    partStr += "and part_id in (" + sGetPartID + ") "
                End If
            Else
                If listB_BumpingTypeShow.Items.Count = 0 Then
                    'partStr = "and part_id='" + (ddlPart.SelectedValue.Trim()) + "' "
                    Dim sGetPartID As String = Get_PartID()
                    If (sGetPartID <> "") Then
                        partStr += "and part_id in (" + sGetPartID + ") "
                    End If
                Else
                    Dim strBumpingType As String = ""
                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If n = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                        End If
                    Next

                    partStr = "and BumpingType_Id IN(" + strBumpingType + ") "
                End If
            End If
            If ddlProduct.SelectedValue <> "CPU" Then
                plantStr = "and Plant='All' "
            Else
                plantStr = ""
            End If

            Dim Month1 As String = DateTime.Now.ToString("yyyyMM")
            Dim Month2 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day).Day
                                                         ) + 1).ToString("yyyyMM")
            Dim Month3 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day).Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day
                                                         ) + 1).ToString("yyyyMM")
            Dim Month4 As String = DateTime.Now.AddDays(-(DateTime.Now.Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day).Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day +
                                                          DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day - DateTime.Now.AddDays(-DateTime.Now.Day - DateTime.Now.AddDays(-DateTime.Now.Day).Day).Day).Day
                                                         ) + 1).ToString("yyyyMM")

            ' --- Month ID ---
            Dim weekTemp As String = ""
            If rbl_week.SelectedIndex = 1 Then
                For i As Integer = 0 To (lb_weekShow.Items.Count - 1)
                    'weekTemp += "'" + (lb_weekShow.Items(i).Value) + "',"
                    weekTemp += (lb_weekShow.Items(i).Value) + ","
                Next
                weekTemp = weekTemp.Substring(0, (weekTemp.Length - 1))
                weekStr = "and MM in (" + weekTemp + ") "
            Else
                ' Defaule 4 week
                'sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
                'weekTemp = "'" + Month1 + "','" + Month2 + "','" + Month3 + "','" + Month4 + "'"
                weekTemp = "and MM in (select top(4) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
            End If

            If rbl_lossItem.SelectedIndex = 0 Then
                topStr = "top(10)"
            ElseIf rbl_lossItem.SelectedIndex = 1 Then
                topStr = "top(20)"
            ElseIf rbl_lossItem.SelectedIndex = 2 Then
                topStr = "top(30)"
            ElseIf rbl_lossItem.SelectedIndex = 3 Then
                topStr = "top(40)"
            ElseIf rbl_lossItem.SelectedIndex = 4 Then
                topStr = "top(50)"
            End If

            conn.Open()
            ' --- Yield Loss ID ---
            Dim itemTemp As String = ""
            itemStr = ""

            If rbl_lossItem.SelectedIndex = 5 Then
                ' === Custom Item ===
                For i As Integer = 0 To (lb_LossShow.Items.Count - 1)
                    itemTemp += "'" + ((lb_LossShow.Items(i).Value).Replace("'", "''")) + "',"
                Next
                itemTemp = itemTemp.Substring(0, (itemTemp.Length - 1))
                itemStr = "and fail_mode in (" + itemTemp + ") "

                'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                '    sqlStr = "select Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                'Else
                sqlStr = "select Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                'End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly_ByBT where 1=1 "
                    End If
                End If

                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                sqlStr += itemStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 month
                    'sqlStr += "and MM in (select yearWW from SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    'sqlStr += "and MM in (" + weekTemp + ") "
                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
                Else
                    ' Custom
                    'sqlStr += "and MM in (select Top(1) yearWW from SystemDateMapping where 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    'sqlStr += weekStr
                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly where 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
                End If
                'If cb_DRowData0.Checked = True Then '匹配報廢回歸
                '    sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                '    sqlStr += "order by Fail_ratio_byXoutScrap desc"
                'Else
                sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
                sqlStr += "order by fail_ratio desc"
                'End If
            Else
                ' === Top N === ' 如果選 Custom 就要呈現選擇的 item
                If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    sqlStr = "select " + topStr + " Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                Else
                    sqlStr = "select " + topStr + " Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+ '_' + MF_Stage) AS 'newFailMode' "
                End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly_ByBT where 1=1 "
                    End If
                End If

                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                End If
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    'sqlStr += "and MM in (SELECT yearWW FROM SystemDateMapping WHERE DateTime='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW)  "
                    'sqlStr += "and MM in (" + weekTemp + ") "
                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
                Else
                    ' Custom
                    'sqlStr += "and MM in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW desc) "
                    'sqlStr += weekStr
                    sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
                End If
                If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    sqlStr += "group by Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                    sqlStr += "order by Fail_ratio_byXoutScrap desc"
                Else
                    sqlStr += "group by Fail_Mode, Fail_Ratio, MF_Stage "
                    sqlStr += "order by fail_ratio desc"
                End If
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myAdapter.Fill(topDT)

            If (topDT.Rows.Count <> 0) Then

                ' === Raw Data ===
                lab_wait.Text = ""
                'sqlStr = "select a.MM, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, (a.Fail_Mode+'_'+a.MF_Stage) as 'NewFailMode', b.yearWW "
                'sqlStr += "from dbo.WB_BinCode_Summary_Monthly a, SystemDateMapping b where 1=1 and a.WW = b.yearWW "
                If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    sqlStr = "select MM, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, MF_Area, DefectCode_ID, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode' "
                Else
                    sqlStr = "select MM, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode' "
                End If
                If tr_BumpingType.Visible = False Then
                    sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                Else
                    If listB_BumpingTypeShow.Items.Count = 0 Then
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                    Else
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly_ByBT where 1=1 "
                    End If
                End If
                ' --- WB 資料當時沒有 Customer_ID 2014/01/07 ---
                'sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
                ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                'sqlStr += "and a.DefectCode_ID Not IN ('C9', '02') "
                sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                If cb_NonIPQC.Checked Then
                    sqlStr += "and Fail_Mode <> 'IPQC defect' "
                End If
                If cb_Non8K.Checked Then
                    sqlStr &= "and Fail_Mode NOT LIKE '8K%' "
                End If
                sqlStr += customStr
                sqlStr += partStr
                sqlStr += plantStr
                If rbl_week.SelectedIndex = 0 Then
                    ' Defaule 4 week
                    'sqlStr += "and b.yearWW in (select top(4) yearWW from SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' group by yearWW order by yearWW desc) "
                    'sqlStr += "and MM in ('" + weekTemp + "') "
                    sqlStr += "and MM in (select TOP(4) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
                Else
                    ' Custom
                    'sqlStr += weekStr
                    sqlStr += "and MM in (select TOP(4) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
                End If

                sqlStr += itemStr
                'sqlStr += "group by a.WW, a.Fail_Mode, a.Fail_Ratio, a.MF_Stage, a.MF_Area, a.DefectCode_ID, b.yearWW "
                'sqlStr += "order by b.yearWW desc, a.fail_ratio desc"
                If cb_DRowData0.Checked = True Then '匹配報廢回歸
                    sqlStr += "group by MM, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, MF_Area, DefectCode_ID "
                    sqlStr += "order by MM desc, Fail_ratio_byXoutScrap desc"
                Else
                    sqlStr += "group by MM, Fail_Mode, Fail_Ratio, MF_Stage, MF_Area, DefectCode_ID "
                    sqlStr += "order by MM desc, fail_ratio desc"
                End If
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                rawDT = New DataTable
                myAdapter.Fill(rawDT)

                ' === Chip Set By Plant === 最新一週的 ChipSet 分廠別的資料 
                chipSetRawDT = New DataTable
                If ddlProduct.SelectedValue = "CS" Then
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        sqlStr = "select MM, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
                    Else
                        sqlStr = "select MM, plant, Fail_Mode, Fail_Ratio, MF_Stage, (Fail_Mode+'_'+MF_Stage) as 'NewFailMode'  "
                    End If
                    If tr_BumpingType.Visible = False Then
                        sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                    Else
                        If listB_BumpingTypeShow.Items.Count = 0 Then
                            sqlStr += "from dbo.WB_BinCode_Summary_Monthly where 1=1 "
                        Else
                            sqlStr += "from dbo.WB_BinCode_Summary_Monthly_ByBT where 1=1 "
                        End If
                    End If
                    ' sqlStr += "and customer_id='" + ddlCustomer.SelectedValue.Trim() + "' "
                    ' 2013/03/13 IPQC code C9及02 改不報廢,  Yield Loss改不顯示 [Mail]
                    'sqlStr += "and DefectCode_ID Not IN ('C9', '02') "
                    sqlStr += "AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN 'C9' ELSE 'N/A' END) AND DefectCode_ID != (CASE WHEN MF_Stage = 'IPQC' THEN '02' ELSE 'N/A' END) "

                    sqlStr += customStr
                    sqlStr += partStr

                    If rbl_week.SelectedIndex = 0 Then
                        'sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE DateTime<='" + DateTime.Now.ToString("yyyy-MM-dd") + "' GROUP BY yearWW ORDER BY yearWW DESC) "
                        'sqlStr += "and MM in ('" + weekTemp + "') "
                        sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly GROUP BY MM ORDER BY MM desc) "
                    Else
                        'sqlStr += "and ww in (SELECT Top(1) yearWW FROM SystemDateMapping WHERE 1=1 " + weekStr + " GROUP BY yearWW ORDER BY yearWW DESC) "
                        'sqlStr += weekStr
                        sqlStr += "and MM in (select TOP(1) MM from WB_BinCode_Summary_Monthly WHERE 1=1 " + weekStr + " GROUP BY MM ORDER BY MM desc) "
                    End If

                    sqlStr += itemStr
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        sqlStr += "group by MM, plant, Fail_Mode, Fail_ratio_byXoutScrap, MF_Stage "
                        sqlStr += "order by plant desc, Fail_ratio_byXoutScrap desc"
                    Else
                        sqlStr += "group by MM, plant, Fail_Mode, Fail_Ratio, MF_Stage "
                        sqlStr += "order by plant desc, fail_ratio desc"
                    End If
                    myAdapter = New SqlDataAdapter(sqlStr, conn)
                    myAdapter.Fill(chipSetRawDT)

                End If

                conn.Close()
                BarChart_MM(rawDT, topDT, weekTemp)

                If ddlProduct.SelectedValue = "CS" And chipSetRawDT.Rows.Count > 0 Then
                    Chart_Panel.Controls.Add(New LiteralControl("<br>"))
                    Dim Chart As New Dundas.Charting.WebControl.Chart()
                    DrawChipSetPlantBarChart(Chart, chipSetRawDT, topDT)
                    Chart_Panel.Controls.Add(Chart)
                End If

                tr_chartDisplay.Visible = True
                If cb_DRowData.Checked Then
                    'showWeeklyRowData(rawDT, chipSetRawDT)
                    showMonthlyRowData(rawDT, chipSetRawDT)


                    tr_gvDisplay.Visible = True
                End If

                'ViewState("RowData") = workTable
                'but_Excel.Enabled = True

            Else

                Dim inDT As DataTable = New DataTable()
                Dim sMonth = DateTime.Now.ToString("yyyyMM")
                sqlStr = "SELECT MM FROM WB_BinCode_Summary_Monthly WHERE DateTime='" + sMonth + "'"
                myAdapter = New SqlDataAdapter(sqlStr, conn)
                myAdapter.Fill(inDT)
                lab_wait.Text = sMonth + " 無資料, 可使用月數自訂 !"

            End If

        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Private Sub showMonthlyRowData(ByRef sourceDT As DataTable, ByRef chipSetRawDT As DataTable)

        Dim plantAryList As ArrayList
        Dim allWeekDT, allPlantDT As DataTable
        Dim newDT As DataTable = sourceDT.Clone
        Dim nWeek As String = sourceDT.Rows(0)("MM")
        Dim dr As DataRow

        ' Step1. 取得最新月的所有 Fail_Mode Item
        If cb_DRowData0.Checked = True Then '匹配報廢回歸
            Dim foundRows As DataRow() = sourceDT.Select("MM='" + nWeek + "'", "Fail_ratio_byXoutScrap desc")
            For x = 0 To (foundRows.Length - 1)
                dr = foundRows(x)
                newDT.LoadDataRow(dr.ItemArray, False)
            Next
        Else
            Dim foundRows As DataRow() = sourceDT.Select("MM='" + nWeek + "'", "Fail_Ratio desc")
            For x = 0 To (foundRows.Length - 1)
                dr = foundRows(x)
                newDT.LoadDataRow(dr.ItemArray, False)
            Next
        End If
        newDT.CaseSensitive = True

        ' Step2. 取得所有月數
        allWeekDT = UtilObj.fun_DataTable_SelectDistinct(sourceDT, "MM")

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
        workTable.Columns.Add("Bumping Type", Type.GetType("System.String"))
        For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
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
            workRow(4) = newDT.Rows(i)("BumpingType")
            colIndex = 5

            ' === 加週數 ===
            For j As Integer = 0 To (allWeekDT.Rows.Count - 1)
                findRow = sourceDT.Select("NewFailMode='" + newFailModeStr + "' and MM='" + (allWeekDT.Rows(j)(0).ToString()) + "'")
                Try
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    Else
                        rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()

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
            workRow(colIndex) = (rvalue).ToString()
            colIndex += 1

            ' === 加廠別 [ 如果是 Chip Set 有資料才有 ] ===
            For j As Integer = 0 To (plantAryList.Count - 1)

                findRow = chipSetRawDT.Select("NewFailMode='" + newFailModeStr + "' and plant='" + (plantAryList(j).ToString()) + "'")
                Try
                    If cb_DRowData0.Checked = True Then '匹配報廢回歸
                        rvalue = CType(findRow(0)("Fail_ratio_byXoutScrap"), Double)
                    Else
                        rvalue = CType(findRow(0)("Fail_Ratio"), Double)
                    End If
                    rvalue = Math.Round(rvalue, 2)
                Catch ex As Exception
                    rvalue = 0
                End Try
                workRow(colIndex) = (rvalue).ToString()
                colIndex += 1
            Next
            workTable.Rows.Add(workRow)

        Next

        If workTable.Rows.Count > 0 Then
            'Dim New_workTable As DataTable = New DataTable()
            'New_workTable.Columns.Add("Fail Mode", Type.GetType("System.String"))
            'New_workTable.Columns.Add("Stage", Type.GetType("System.String"))
            'New_workTable.Columns.Add("Area", Type.GetType("System.String")) ' DefectCode_ID
            'New_workTable.Columns.Add("ID", Type.GetType("System.String"))
            'For i As Integer = 0 To (allWeekDT.Rows.Count - 1)
            '    New_workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
            'Next
            'New_workTable.Columns.Add("Delta", Type.GetType("System.String"))

            'Dim dtDistinct As DataTable = workTable.DefaultView.ToTable(True, New String() {"Fail Mode"})
            'For i As Integer = 0 To (dtDistinct.Rows.Count - 1)
            '    For j As Integer = 0 To (workTable.Rows.Count - 1)

            '    Next
            '    New_workTable.Columns.Add((allWeekDT.Rows(i)(0)).ToString(), Type.GetType("System.String"))
            'Next
            'Dim New_workRow As DataRow
            'Dim New_findRow As DataRow()

            'ViewState("RowData") = workTable
            'but_Excel.Enabled = True
        End If

        gv_rowdata.DataSource = workTable
        gv_rowdata.DataBind()
        UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)

    End Sub
#End Region

#End Region

    Private Sub moveList(ByRef sourceList As ListBox, ByRef destList As ListBox)

        Dim sourceAry As New ArrayList()
        Dim DestAry As New ArrayList()

        For i As Integer = 0 To sourceList.Items.Count - 1
            If sourceList.Items(i).Selected Then
                If DestAry.Contains(sourceList.Items(i).Value) = False Then
                    DestAry.Add(sourceList.Items(i).Value)
                End If

            Else
                If sourceAry.Contains(sourceList.Items(i).Value) = False Then
                    sourceAry.Add(sourceList.Items(i).Value)
                End If


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
    Public Class DecComparer
        Implements IComparer
        Dim myComapar As CaseInsensitiveComparer = New CaseInsensitiveComparer()
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
            Return myComapar.Compare(y, x)
        End Function
    End Class

    Private Sub ListBoxSort(ByVal lbx As ListBox)
        '利用sortedlist 類為listbox排序 
        'Dim slist As New SortedList()
        Dim slist As SortedList = New SortedList(New DecComparer())
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

    Private Sub ListBoxSort(ByVal lbx As ListBox, ByVal bASC As Boolean)
        '利用sortedlist 類為listbox排序 
        Dim slist As SortedList
        If bASC = True Then
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
    Protected Sub but_BumpingTypeRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_BumpingTypeRight.Click
        moveList(listB_BumpingTypeSource, listB_BumpingTypeShow)

        Get_Part_Process()

        ListBoxSort(listB_BumpingTypeSource)
        ListBoxSort(listB_BumpingTypeShow)
    End Sub

    'Protected Sub but_BumpingTypeRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_BumpingTypeRight.Click
    '    moveList(listB_BumpingTypeSource, listB_BumpingTypeShow)

    '    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '    Dim sqlStr As String = ""
    '    Dim myDT As DataTable
    '    Dim myAdapter As SqlDataAdapter

    '    If listB_BumpingTypeSource.Items.Count = 0 Then
    '        'ddlCustomer.Enabled = True
    '        'ddlPart.Enabled = True
    '        'rb_ProductPart.Enabled = True
    '        'rb_ProductPart.Items(0).Selected = True

    '        'ddlCustomer.Enabled = False
    '        'rb_ProductPart.Enabled = False
    '        'ddlPart.Enabled = False

    '        If ddlProduct.Text = "PPS" Then
    '            ' -- Part ID --
    '            rb_ProductPart.Items(0).Enabled = False
    '            rb_ProductPart.SelectedIndex = 1
    '            sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            If ddlCustomer.Text <> "All" Then
    '                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            End If
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            If listB_BumpingTypeSource.Items.Count > 0 Then
    '                ' --- Bumping Type ---
    '                Dim strBumpingType As String = ""
    '                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                    If n = 0 Then
    '                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    Else
    '                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    End If
    '                Next

    '                Dim BumpingType_Part As String = ""
    '                If strBumpingType <> "" Then
    '                    BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                    sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                    If BumpingType_Part <> "" Then
    '                        sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                    End If
    '                End If
    '            End If
    '            sqlStr += "group by part_id order by part_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '            Next
    '        Else
    '            ' -- Production ID --
    '            'rb_ProductPart.SelectedIndex = 0
    '            rb_ProductPart.Items(0).Enabled = True
    '            sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            If ddlCustomer.Text <> "All" Then
    '                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            End If
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            If listB_BumpingTypeSource.Items.Count > 0 Then
    '                ' --- Bumping Type ---
    '                Dim strBumpingType As String = ""
    '                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                    If n = 0 Then
    '                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    Else
    '                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    End If
    '                Next

    '                Dim BumpingType_Part As String = ""
    '                If strBumpingType <> "" Then
    '                    BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                    sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                    If BumpingType_Part <> "" Then
    '                        sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                    End If
    '                End If
    '            End If
    '            sqlStr += "group by production_id order by production_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '            Next
    '        End If
    '    Else
    '        If ddlProduct.Text = "PPS" Then
    '            If listB_BumpingTypeShow.Items.Count = 0 Then
    '                'ddlCustomer.Enabled = True
    '                'rb_ProductPart.Enabled = True
    '                'ddlPart.Enabled = True

    '                If ddlProduct.Text = "PPS" Then
    '                    ' -- Part ID --
    '                    rb_ProductPart.Items(0).Enabled = False
    '                    rb_ProductPart.SelectedIndex = 1
    '                    sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by part_id order by part_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                    Next
    '                Else
    '                    ' -- Production ID --
    '                    'rb_ProductPart.SelectedIndex = 0
    '                    rb_ProductPart.Items(0).Enabled = True
    '                    sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by production_id order by production_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                    Next
    '                End If
    '            Else
    '                'ddlCustomer.Enabled = False
    '                'rb_ProductPart.Enabled = False
    '                'ddlPart.Enabled = False

    '                If ddlProduct.Text = "PPS" Then
    '                    ' -- Part ID --
    '                    rb_ProductPart.Items(0).Enabled = False
    '                    rb_ProductPart.SelectedIndex = 1
    '                    sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by part_id order by part_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)

    '                    If listB_OLProcessShow.Items.Count = 0 And listB_BackendShow.Items.Count = 0 Then
    '                        listB_PartSource.Items.Clear()
    '                    End If

    '                    'listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        Dim aa As New ListItem
    '                        aa.Text = myDT.Rows(i)("part_id").ToString()
    '                        aa.Value = myDT.Rows(i)("part_id").ToString()
    '                        If listB_PartSource.Items.Contains(aa) = False Then
    '                            listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                        End If
    '                    Next
    '                Else
    '                    ' -- Production ID --
    '                    'rb_ProductPart.SelectedIndex = 0
    '                    rb_ProductPart.Items(0).Enabled = True
    '                    sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by production_id order by production_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                    Next
    '                End If
    '            End If
    '        Else
    '            ddlCustomer.Enabled = True
    '            'rb_ProductPart.Enabled = True
    '            'ddlPart.Enabled = True

    '            If ddlProduct.Text = "PPS" Then
    '                ' -- Part ID --
    '                rb_ProductPart.Items(0).Enabled = False
    '                rb_ProductPart.SelectedIndex = 1
    '                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                sqlStr += "and fail_function=1 "
    '                If ddlCustomer.Text <> "All" Then
    '                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                End If
    '                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                If listB_BumpingTypeSource.Items.Count > 0 Then
    '                    ' --- Bumping Type ---
    '                    Dim strBumpingType As String = ""
    '                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                        If n = 0 Then
    '                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        Else
    '                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        End If
    '                    Next

    '                    Dim BumpingType_Part As String = ""
    '                    If strBumpingType <> "" Then
    '                        BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                        sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                        If BumpingType_Part <> "" Then
    '                            sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                        End If
    '                    End If
    '                End If
    '                sqlStr += "group by part_id order by part_id"
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myDT = New DataTable
    '                myAdapter.Fill(myDT)
    '                'UtilObj.FillController(myDT, ddlPart, 0)
    '                listB_PartSource.Items.Clear()
    '                listB_PartShow.Items.Clear()
    '                For i As Integer = 0 To myDT.Rows.Count - 1
    '                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                Next
    '            Else
    '                ' -- Production ID --
    '                'rb_ProductPart.SelectedIndex = 0
    '                rb_ProductPart.Items(0).Enabled = True
    '                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                sqlStr += "and fail_function=1 "
    '                If ddlCustomer.Text <> "All" Then
    '                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                End If
    '                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                If listB_BumpingTypeSource.Items.Count > 0 Then
    '                    ' --- Bumping Type ---
    '                    Dim strBumpingType As String = ""
    '                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                        If n = 0 Then
    '                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        Else
    '                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        End If
    '                    Next

    '                    Dim BumpingType_Part As String = ""
    '                    If strBumpingType <> "" Then
    '                        BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                        sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                        If BumpingType_Part <> "" Then
    '                            sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                        End If
    '                    End If
    '                End If
    '                sqlStr += "group by production_id order by production_id"
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myDT = New DataTable
    '                myAdapter.Fill(myDT)
    '                'UtilObj.FillController(myDT, ddlPart, 0)
    '                listB_PartSource.Items.Clear()
    '                listB_PartShow.Items.Clear()
    '                For i As Integer = 0 To myDT.Rows.Count - 1
    '                    listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                Next
    '            End If
    '            rb_ProductPart.Items(0).Selected = True
    '        End If
    '        If ddlCustomer.Items.Count > 0 Then
    '            ddlCustomer.SelectedIndex = 0
    '        End If
    '        'If ddlPart.Items.Count > 0 Then
    '        '    ddlPart.SelectedIndex = 0
    '        'End If
    '    End If

    '    ListBoxSort(listB_BumpingTypeSource)
    '    ListBoxSort(listB_BumpingTypeShow)
    'End Sub

    Protected Sub but_BumpingTypeLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_BumpingTypeLeft.Click
        moveList(listB_BumpingTypeShow, listB_BumpingTypeSource)

        Get_Part_Process()

        ListBoxSort(listB_BumpingTypeSource)
        ListBoxSort(listB_BumpingTypeShow)
    End Sub

    'Protected Sub but_BumpingTypeLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_BumpingTypeLeft.Click
    '    moveList(listB_BumpingTypeShow, listB_BumpingTypeSource)

    '    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
    '    Dim sqlStr As String = ""
    '    Dim myDT As DataTable
    '    Dim myAdapter As SqlDataAdapter

    '    If listB_BumpingTypeSource.Items.Count = 0 Then
    '        'ddlCustomer.Enabled = True
    '        'ddlPart.Enabled = True
    '        'rb_ProductPart.Enabled = True
    '        'rb_ProductPart.Items(0).Selected = True

    '        'ddlCustomer.Enabled = False
    '        'rb_ProductPart.Enabled = False
    '        'ddlPart.Enabled = False

    '        If ddlProduct.Text = "PPS" Then
    '            ' -- Part ID --
    '            rb_ProductPart.Items(0).Enabled = False
    '            rb_ProductPart.SelectedIndex = 1
    '            sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            If ddlCustomer.Text <> "All" Then
    '                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            End If
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            If listB_BumpingTypeSource.Items.Count > 0 Then
    '                ' --- Bumping Type ---
    '                Dim strBumpingType As String = ""
    '                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                    If n = 0 Then
    '                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    Else
    '                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    End If
    '                Next

    '                Dim BumpingType_Part As String = ""
    '                If strBumpingType <> "" Then
    '                    BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                    sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                    If BumpingType_Part <> "" Then
    '                        sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                    End If
    '                End If
    '            End If
    '            sqlStr += "group by part_id order by part_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '            Next
    '        Else
    '            ' -- Production ID --
    '            'rb_ProductPart.SelectedIndex = 0
    '            rb_ProductPart.Items(0).Enabled = True
    '            sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '            sqlStr += "and fail_function=1 "
    '            If ddlCustomer.Text <> "All" Then
    '                sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '            End If
    '            sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '            If listB_BumpingTypeSource.Items.Count > 0 Then
    '                ' --- Bumping Type ---
    '                Dim strBumpingType As String = ""
    '                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                    If n = 0 Then
    '                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    Else
    '                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                    End If
    '                Next

    '                Dim BumpingType_Part As String = ""
    '                If strBumpingType <> "" Then
    '                    BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                    sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                    If BumpingType_Part <> "" Then
    '                        sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                    End If
    '                End If
    '            End If
    '            sqlStr += "group by production_id order by production_id"
    '            myAdapter = New SqlDataAdapter(sqlStr, conn)
    '            myDT = New DataTable
    '            myAdapter.Fill(myDT)
    '            'UtilObj.FillController(myDT, ddlPart, 0)
    '            listB_PartSource.Items.Clear()
    '            listB_PartShow.Items.Clear()
    '            For i As Integer = 0 To myDT.Rows.Count - 1
    '                listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '            Next
    '        End If
    '    Else
    '        If ddlProduct.Text = "PPS" Then
    '            If listB_BumpingTypeShow.Items.Count = 0 Then
    '                ddlCustomer.Enabled = True
    '                'rb_ProductPart.Enabled = True
    '                'ddlPart.Enabled = True

    '                If ddlProduct.Text = "PPS" Then
    '                    ' -- Part ID --
    '                    rb_ProductPart.Items(0).Enabled = False
    '                    rb_ProductPart.SelectedIndex = 1
    '                    sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by part_id order by part_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                    Next
    '                Else
    '                    ' -- Production ID --
    '                    'rb_ProductPart.SelectedIndex = 0
    '                    rb_ProductPart.Items(0).Enabled = True
    '                    sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by production_id order by production_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                    Next
    '                End If
    '            Else
    '                'ddlCustomer.Enabled = False
    '                'rb_ProductPart.Enabled = False
    '                'ddlPart.Enabled = False

    '                If ddlProduct.Text = "PPS" Then
    '                    ' -- Part ID --
    '                    rb_ProductPart.Items(0).Enabled = False
    '                    rb_ProductPart.SelectedIndex = 1
    '                    sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by part_id order by part_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                    Next
    '                Else
    '                    ' -- Production ID --
    '                    'rb_ProductPart.SelectedIndex = 0
    '                    rb_ProductPart.Items(0).Enabled = True
    '                    sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                    sqlStr += "and fail_function=1 "
    '                    If ddlCustomer.Text <> "All" Then
    '                        sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                    End If
    '                    sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                    If listB_BumpingTypeSource.Items.Count > 0 Then
    '                        ' --- Bumping Type ---
    '                        Dim strBumpingType As String = ""
    '                        For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                            If n = 0 Then
    '                                strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            Else
    '                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                            End If
    '                        Next

    '                        Dim BumpingType_Part As String = ""
    '                        If strBumpingType <> "" Then
    '                            BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                            sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                            If BumpingType_Part <> "" Then
    '                                sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                            End If
    '                        End If
    '                    End If
    '                    sqlStr += "group by production_id order by production_id"
    '                    myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                    myDT = New DataTable
    '                    myAdapter.Fill(myDT)
    '                    'UtilObj.FillController(myDT, ddlPart, 0)
    '                    listB_PartSource.Items.Clear()
    '                    listB_PartShow.Items.Clear()
    '                    For i As Integer = 0 To myDT.Rows.Count - 1
    '                        listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                    Next
    '                End If
    '            End If
    '        Else
    '            ddlCustomer.Enabled = True
    '            'rb_ProductPart.Enabled = True
    '            'ddlPart.Enabled = True

    '            If ddlProduct.Text = "PPS" Then
    '                ' -- Part ID --
    '                rb_ProductPart.Items(0).Enabled = False
    '                rb_ProductPart.SelectedIndex = 1
    '                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                sqlStr += "and fail_function=1 "
    '                If ddlCustomer.Text <> "All" Then
    '                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                End If
    '                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                If listB_BumpingTypeSource.Items.Count > 0 Then
    '                    ' --- Bumping Type ---
    '                    Dim strBumpingType As String = ""
    '                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                        If n = 0 Then
    '                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        Else
    '                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        End If
    '                    Next

    '                    Dim BumpingType_Part As String = ""
    '                    If strBumpingType <> "" Then
    '                        BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                        sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                        If BumpingType_Part <> "" Then
    '                            sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                        End If
    '                    End If
    '                End If
    '                sqlStr += "group by part_id order by part_id"
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myDT = New DataTable
    '                myAdapter.Fill(myDT)
    '                'UtilObj.FillController(myDT, ddlPart, 0)
    '                listB_PartSource.Items.Clear()
    '                listB_PartShow.Items.Clear()
    '                For i As Integer = 0 To myDT.Rows.Count - 1
    '                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
    '                Next
    '            Else
    '                ' -- Production ID --
    '                'rb_ProductPart.SelectedIndex = 0
    '                rb_ProductPart.Items(0).Enabled = True
    '                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
    '                sqlStr += "and fail_function=1 "
    '                If ddlCustomer.Text <> "All" Then
    '                    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
    '                End If
    '                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
    '                If listB_BumpingTypeSource.Items.Count > 0 Then
    '                    ' --- Bumping Type ---
    '                    Dim strBumpingType As String = ""
    '                    For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
    '                        If n = 0 Then
    '                            strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        Else
    '                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
    '                        End If
    '                    Next

    '                    Dim BumpingType_Part As String = ""
    '                    If strBumpingType <> "" Then
    '                        BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
    '                        sqlStr += "and bumping_type in (" + strBumpingType + ") "
    '                        If BumpingType_Part <> "" Then
    '                            sqlStr += "and part_id in (" + BumpingType_Part + ") "
    '                        End If
    '                    End If
    '                End If
    '                sqlStr += "group by production_id order by production_id"
    '                myAdapter = New SqlDataAdapter(sqlStr, conn)
    '                myDT = New DataTable
    '                myAdapter.Fill(myDT)
    '                'UtilObj.FillController(myDT, ddlPart, 0)
    '                listB_PartSource.Items.Clear()
    '                listB_PartShow.Items.Clear()
    '                For i As Integer = 0 To myDT.Rows.Count - 1
    '                    listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
    '                Next
    '            End If
    '            rb_ProductPart.Items(0).Selected = True
    '        End If
    '        If ddlCustomer.Items.Count > 0 Then
    '            ddlCustomer.SelectedIndex = 0
    '        End If
    '        'If ddlPart.Items.Count > 0 Then
    '        '    ddlPart.SelectedIndex = 0
    '        'End If
    '    End If

    '    ListBoxSort(listB_BumpingTypeSource)
    '    ListBoxSort(listB_BumpingTypeShow)
    'End Sub


    ' >>
    Protected Sub but_PartRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_PartRight.Click
        moveList(listB_PartSource, listB_PartShow)

        ListBoxSort(listB_PartSource, True)
        ListBoxSort(listB_PartShow, True)
    End Sub

    ' <<
    Protected Sub but_PartLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles but_PartLeft.Click
        moveList(listB_PartShow, listB_PartSource)

        ListBoxSort(listB_PartSource, True)
        ListBoxSort(listB_PartShow, True)
    End Sub

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sCustomer_ID As String = ""
        If ddlProduct.SelectedValue = "PPS" Or ddlProduct.SelectedValue = "PCB" Then
            Get_Part_Process(TextBox1.Text)
        Else

            Get_Part_Process_CPU_CS(TextBox1.Text)
        End If

        Exit Sub

        ' --- 檢查有無 Group Part, 沒有就呈現 Part ID ---
        'sqlStr = "select production_id from " + confTable + " where 1=1 "
        For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
            If n = 0 Then
                sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
            Else
                sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
            End If
        Next
        sqlStr = "select "
        If rb_ProductPart.SelectedIndex = 0 Then
            sqlStr += "production_id "
        ElseIf rb_ProductPart.SelectedIndex = 1 Then
            sqlStr += "Part_id "
        End If
        sqlStr += "from " + confTable + " where 1=1 "

        If sCustomer_ID <> "" Then
            sqlStr += "and customer_id in (" + sCustomer_ID + ") "
        End If

        'If ddlCustomer.Text <> "All" Then
        '    sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
        'End If
        sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "
        If listB_BumpingTypeShow.Items.Count > 0 Then
            Dim strBumpingType As String = ""
            For x As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                If x = 0 Then
                    strBumpingType += "'" & listB_BumpingTypeShow.Items(x).Text & "'"
                Else
                    strBumpingType += ",'" & listB_BumpingTypeShow.Items(x).Text & "'"
                End If
            Next
            sqlStr += "and Bumping_Type in (" + strBumpingType + ") "
        End If
        If TextBox1.Text <> "" Then
            If rb_ProductPart.SelectedIndex = 0 Then
                sqlStr += "and production_id like '%" + TextBox1.Text + "%' "
            ElseIf rb_ProductPart.SelectedIndex = 1 Then
                sqlStr += "and Part_id like '%" + TextBox1.Text + "%' "
            End If
        End If
        sqlStr += "and yield_function=1 "
        'sqlStr += "and production_id <> '' "
        'sqlStr += "group by production_id"
        sqlStr += "group by "
        If rb_ProductPart.SelectedIndex = 0 Then
            sqlStr += "production_id"
        ElseIf rb_ProductPart.SelectedIndex = 1 Then
            sqlStr += "Part_id"
        End If

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
                    'If ddlCustomer.Text <> "All" Then
                    '    sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
                    'End If
                    If sCustomer_ID <> "" Then
                        sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                    End If

                    sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "
                    If listB_BumpingTypeShow.Items.Count > 0 Then
                        Dim strBumpingType As String = ""
                        For x As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                            If x = 0 Then
                                strBumpingType += "'" & listB_BumpingTypeShow.Items(x).Text & "'"
                            Else
                                strBumpingType += ",'" & listB_BumpingTypeShow.Items(x).Text & "'"
                            End If
                        Next
                        sqlStr += "and Bumping_Type in (" + strBumpingType + ") "
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
                    'UtilObj.FillController(myDT, ddlPart, 0, "Part_id", "Memo")
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                    Next
                Else
                    'rb_ProductPart.SelectedIndex = 0
                    rb_ProductPart.Items(0).Enabled = True
                    ' 有就秀 Group
                    'rb_ProductPart.SelectedIndex = 0
                    'UtilObj.FillController(myDT, ddlPart, 0)
                    listB_PartSource.Items.Clear()
                    listB_PartShow.Items.Clear()
                    For i As Integer = 0 To myDT.Rows.Count - 1
                        
                        listB_PartSource.Items.Add(myDT.Rows(i)(0).ToString())
                    Next
                End If
            Else
                ' 轉向 Part ID
                rb_ProductPart.SelectedIndex = 1
                sqlStr = "select Part_id, Memo from " + confTable + " where 1=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If


                sqlStr += "and category ='" + ddlProduct.SelectedValue + "' "
                If listB_BumpingTypeShow.Items.Count > 0 Then
                    Dim strBumpingType As String = ""
                    For x As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                        If x = 0 Then
                            strBumpingType += "'" & listB_BumpingTypeShow.Items(x).Text & "'"
                        Else
                            strBumpingType += ",'" & listB_BumpingTypeShow.Items(x).Text & "'"
                        End If
                    Next
                    sqlStr += "and Bumping_Type in (" + strBumpingType + ") "
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
            'If ddlCustomer.Text <> "All" Then
            '    sqlStr += "and customer_id ='" + ddlCustomer.SelectedValue + "' "
            'End If
            If sCustomer_ID <> "" Then
                sqlStr += "and customer_id in (" + sCustomer_ID + ") "
            End If

            If listB_BumpingTypeShow.Items.Count > 0 Then
                ' --- Bumping Type ---
                Dim strBumpingType As String = ""
                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                    If n = 0 Then
                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    Else
                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    End If
                Next

                If strBumpingType <> "" Then
                    sqlStr += "and Bumping_Type IN(" + strBumpingType + ") "
                End If
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
    Private Sub Get_Part_Process_CPU_CS(Optional ByVal sfilter As String = "")
       Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim categoryStr As String = ""
        Dim sCustomer_ID As String = ""
        Try

            conn.Open()
            For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
                If n = 0 Then
                    sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
                Else
                    sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
                End If
            Next

            ' -- Production_ID --
            If rb_ProductPart.SelectedIndex = 0 Then
                sqlStr = "select production_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If


                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "and production_id like '%" + sfilter + "%' "
                sqlStr += "group by production_id order by production_id"
            Else
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If

                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If
                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "
                sqlStr += "and part_id like '%" + sfilter + "%' "

                sqlStr += "group by part_id order by part_id"
            End If
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            'UtilObj.FillController(myDT, ddlPart, 0)
            listB_PartSource.Items.Clear()
            listB_PartShow.Items.Clear()
            For i As Integer = 0 To myDT.Rows.Count - 1
                If rb_ProductPart.SelectedIndex = 0 Then
                    listB_PartSource.Items.Add(myDT.Rows(i)("production_id").ToString())
                Else
                    listB_PartSource.Items.Add(myDT.Rows(i)("Part_id").ToString())
                End If
            Next

            ' -- Yield Loss Item --
            sqlStr = "select fail_mode from dbo.VW_BinCode_Summary where 1=1 "
            If ddlProduct.Text = "WB" Then
                sqlStr += "and customer_id='All' "
            Else
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If
            End If
            'sqlStr += "and part_id='" + ddlPart.SelectedValue + "' "
            Dim sGetPartID As String = Get_PartID()
            If sGetPartID <> "" Then
                sqlStr += "and part_id in(" + sGetPartID + ") "
            End If
            sqlStr += "group by fail_mode order by fail_mode"
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            UtilObj.FillLitsBoxController(myDT, lb_LossSource, 1)

            conn.Close()
        Catch ex As Exception
            Dim sError As String = ex.ToString()
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub


    Private Sub Get_Part_Process(Optional ByVal sfilter As String = "")
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlhead As String = ""
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sCustomer_ID As String = ""

        'Alfie
        If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then

            For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
                If n = 0 Then
                    sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
                Else
                    sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
                End If
            Next

            rb_ProductPart.Items(0).Enabled = False
            rb_ProductPart.SelectedIndex = 1
            sqlhead = "select * from ("




            sqlStr += "select part_id from Customer_Prodction_Mapping_BU_Rename a,[MES].[dbo].[ProductInfo] b where 1=1 "
            'sqlStr += " and a.Part_Id=b.Part_No and fail_function=1 "
            sqlStr += " and a.Part_Id=b.Part_No  "

            'If sCustomer_ID <> "" Then
            '    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
            'End If

            If ddlProduct.Text = "PPS" Then
                If sCustomer_ID <> "" Then
                    sqlStr += "and Assfct in (" + sCustomer_ID + ") "
                End If
            Else
                If sCustomer_ID <> "" Then
                    sqlStr += "and Dircu in (" + sCustomer_ID + ") "
                End If
            End If



            sqlStr += "and a.category='" + ddlProduct.SelectedValue + "' "

            If listB_BumpingTypeShow.Items.Count > 0 Then
                ' --- Bumping Type ---
                Dim strBumpingType As String = ""
                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                    If n = 0 Then
                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    Else
                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    End If
                Next

                Dim BumpingType_Part As String = ""
                If strBumpingType <> "" Then
                    If ddlProduct.Text = "PCB" Then
                        sqlStr += "and Notes in (" + strBumpingType + ") "
                    Else
                        sqlStr += "and bumping_type in (" + strBumpingType + ") "
                    End If

                End If
            End If

            If listB_OLProcessShow.Items.Count > 0 Then

                Dim strOLProcess As String = ""
                For n As Integer = 0 To (listB_OLProcessShow.Items.Count - 1)
                    If n = 0 Then
                        strOLProcess += "'" & listB_OLProcessShow.Items(n).Text & "'"
                    Else
                        strOLProcess += ",'" & listB_OLProcessShow.Items(n).Text & "'"
                    End If
                Next

                Dim OLProcess_Part As String = ""
                If strOLProcess <> "" Then
                    sqlStr += "and OL_Process in (" + strOLProcess + ") "
                End If

            End If

            If listB_BackendShow.Items.Count > 0 Then


                Dim strBackend As String = ""
                For n As Integer = 0 To (listB_BackendShow.Items.Count - 1)
                    If n = 0 Then
                        strBackend += "'" & listB_BackendShow.Items(n).Text & "'"
                    Else
                        strBackend += ",'" & listB_BackendShow.Items(n).Text & "'"
                    End If
                Next
                If strBackend <> "" Then
                    sqlStr += "and Backend in (" + strBackend + ") "
                End If


            End If


            If sqlStr <> "" Then
                sqlhead += sqlStr
                sqlhead += ")c "

                If sfilter <> "" Then
                    sqlhead += " where part_id like '%" + sfilter + "%'"

                End If
                sqlhead += "order by part_id "
                myAdapter = New SqlDataAdapter(sqlhead, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)

                listB_PartSource.Items.Clear()
                'listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                Next

            Else
                sqlhead = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 and fail_function=1 and category='PPS'"



                If sfilter <> "" Then
                    sqlhead += " and part_id like '%" + sfilter + "%' "

                End If

                'If sCustomer_ID <> "" Then
                '    sqlhead += " and customer_id in (" + sCustomer_ID + ") "
                'End If

                If ddlProduct.Text = "PPS" Then
                    If sCustomer_ID <> "" Then
                        sqlhead += " and assfct in (" + sCustomer_ID + ") "
                    End If
                Else
                    If sCustomer_ID <> "" Then
                        sqlhead += " and Dircu in (" + sCustomer_ID + ") "
                    End If
                End If



                myAdapter = New SqlDataAdapter(sqlhead, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)

                listB_PartSource.Items.Clear()
                'listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                Next
            End If

        End If

    End Sub

    Private Sub Get_Part_Process1(Optional ByVal sfilter As String = "")
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlhead As String = ""
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter
        Dim sCustomer_ID As String = ""

        'Alfie
        If ddlProduct.Text = "PPS" Or ddlProduct.Text = "PCB" Then

            For n As Integer = 0 To (listB_CustomerTarget.Items.Count - 1)
                If n = 0 Then
                    sCustomer_ID += "'" & listB_CustomerTarget.Items(n).Text & "'"
                Else
                    sCustomer_ID += ",'" & listB_CustomerTarget.Items(n).Text & "'"
                End If
            Next





            rb_ProductPart.Items(0).Enabled = False
            rb_ProductPart.SelectedIndex = 1
            sqlhead = "select * from ("

            If listB_BumpingTypeShow.Items.Count > 0 Then
                sqlStr = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 "
                sqlStr += "and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If

                sqlStr += "and category='" + ddlProduct.SelectedValue + "' "

                ' --- Bumping Type ---
                Dim strBumpingType As String = ""
                For n As Integer = 0 To (listB_BumpingTypeShow.Items.Count - 1)
                    If n = 0 Then
                        strBumpingType += "'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    Else
                        strBumpingType += ",'" & listB_BumpingTypeShow.Items(n).Text & "'"
                    End If
                Next

                Dim BumpingType_Part As String = ""
                If strBumpingType <> "" Then
                    'BumpingType_Part = Get_BumpingType_PartID(strBumpingType)
                    sqlStr += "and bumping_type in (" + strBumpingType + ") "
                    'If BumpingType_Part <> "" Then
                    '    sqlStr += "and part_id in (" + BumpingType_Part + ") "
                    'End If
                End If

                'If ddlCustomer.SelectedItem.Text <> "" And ddlCustomer.SelectedItem.Text.ToLower <> "all" Then
                '    sqlStr += "and customer_id in ('" + ddlCustomer.SelectedItem.Text + "') "
                'End If

            End If


            If listB_OLProcessShow.Items.Count > 0 Then

                If sqlStr <> "" Then
                    sqlStr += "union "
                End If

                sqlStr += "select part_id from Customer_Prodction_Mapping_BU_Rename a,[MES].[dbo].[ProductInfo] b where 1=1 "
                sqlStr += " and a.Part_Id=b.Part_No and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If

                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If

                sqlStr += "and a.category='" + ddlProduct.SelectedValue + "' "



                Dim strOLProcess As String = ""
                For n As Integer = 0 To (listB_OLProcessShow.Items.Count - 1)
                    If n = 0 Then
                        strOLProcess += "'" & listB_OLProcessShow.Items(n).Text & "'"
                    Else
                        strOLProcess += ",'" & listB_OLProcessShow.Items(n).Text & "'"
                    End If
                Next

                Dim OLProcess_Part As String = ""
                If strOLProcess <> "" Then
                    sqlStr += "and OL_Process in (" + strOLProcess + ") "
                End If

                'If ddlCustomer.SelectedItem.Text <> "" And ddlCustomer.SelectedItem.Text.ToLower <> "all" Then
                '    sqlStr += "and customer_id in ('" + ddlCustomer.SelectedItem.Text + "') "
                'End If
            End If



            If listB_BackendShow.Items.Count > 0 Then

                If sqlStr <> "" Then
                    sqlStr += "union"
                End If

                sqlStr += " select part_id from Customer_Prodction_Mapping_BU_Rename a,[MES].[dbo].[ProductInfo] b where 1=1 "
                sqlStr += " and a.Part_Id=b.Part_No and fail_function=1 "
                'If ddlCustomer.Text <> "All" Then
                '    sqlStr += "and a.customer_id='" + ddlCustomer.SelectedValue + "' "
                'End If
                If sCustomer_ID <> "" Then
                    sqlStr += "and customer_id in (" + sCustomer_ID + ") "
                End If

                sqlStr += "and a.category='" + ddlProduct.SelectedValue + "' "


                Dim strBackend As String = ""
                For n As Integer = 0 To (listB_BackendShow.Items.Count - 1)
                    If n = 0 Then
                        strBackend += "'" & listB_BackendShow.Items(n).Text & "'"
                    Else
                        strBackend += ",'" & listB_BackendShow.Items(n).Text & "'"
                    End If
                Next
                If strBackend <> "" Then
                    sqlStr += "and Backend in (" + strBackend + ") "
                End If

                'If ddlCustomer.SelectedItem.Text <> "" And ddlCustomer.SelectedItem.Text.ToLower <> "all" Then
                '    sqlStr += "and customer_id in ('" + ddlCustomer.SelectedItem.Text + "') "
                'End If

            End If


            If sqlStr <> "" Then
                sqlhead += sqlStr
                sqlhead += ")c "

                If sfilter <> "" Then
                    sqlhead += " where part_id like '%" + sfilter + "%'"

                End If
                sqlhead += "order by part_id "
                myAdapter = New SqlDataAdapter(sqlhead, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)

                listB_PartSource.Items.Clear()
                'listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                Next

            Else
                sqlhead = "select part_id from Customer_Prodction_Mapping_BU_Rename where 1=1 and fail_function=1 and category='PPS'"



                If sfilter <> "" Then
                    sqlhead += " and part_id like '%" + sfilter + "%' "

                End If

                If sCustomer_ID <> "" Then
                    sqlhead += " and customer_id in (" + sCustomer_ID + ") "
                End If

                myAdapter = New SqlDataAdapter(sqlhead, conn)
                myDT = New DataTable
                myAdapter.Fill(myDT)

                listB_PartSource.Items.Clear()
                'listB_PartShow.Items.Clear()
                For i As Integer = 0 To myDT.Rows.Count - 1
                    listB_PartSource.Items.Add(myDT.Rows(i)("part_id").ToString())
                Next
            End If


        End If
    End Sub
  
    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles but_OLProcessRight.Click
        moveList(listB_OLProcessSource, listB_OLProcessShow)

        Get_Part_Process()
       

        ListBoxSort(listB_OLProcessSource)
        ListBoxSort(listB_OLProcessShow)
    End Sub

    Protected Sub Button4_Click(sender As Object, e As System.EventArgs) Handles Button4.Click
        moveList(listB_BackendSource, listB_BackendShow)

        Get_Part_Process()

        ListBoxSort(listB_BackendSource)
        ListBoxSort(listB_BackendShow)

        'Alfie
        Exit Sub
    End Sub

    Protected Sub Button3_Click(sender As Object, e As System.EventArgs) Handles Button3.Click
        moveList(listB_OLProcessShow, listB_OLProcessSource)

        Get_Part_Process()


        ListBoxSort(listB_OLProcessSource)
        ListBoxSort(listB_OLProcessShow)
    End Sub

    Protected Sub Button5_Click(sender As Object, e As System.EventArgs) Handles Button5.Click
        'moveList(listB_BackendSource, listB_BackendShow)
        moveList(listB_BackendShow, listB_BackendSource)
        Get_Part_Process()

        ListBoxSort(listB_BackendSource)
        ListBoxSort(listB_BackendShow)
    End Sub

    Protected Sub Button2_Click1(sender As Object, e As System.EventArgs) Handles Button2.Click
        moveList(listB_CustomerSource, listB_CustomerTarget)

        Get_Part_Process()
        ListBoxSort(listB_CustomerSource, True)
        ListBoxSort(listB_CustomerTarget, True)
    End Sub

    Protected Sub Button6_Click(sender As Object, e As System.EventArgs) Handles Button6.Click
        'moveList(listB_BackendSource, listB_BackendShow)
        moveList(listB_CustomerTarget, listB_CustomerSource)
        Get_Part_Process()

        ListBoxSort(listB_CustomerSource, True)
        ListBoxSort(listB_CustomerTarget, True)
    End Sub
End Class
