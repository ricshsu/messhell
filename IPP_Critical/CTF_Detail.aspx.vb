Imports System.Data
Imports System.Data.SqlClient
Imports NYPCB.SPC

Partial Class CTF_Detail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Me.but_reCaculate.Attributes.Add("onclick", "javascript:document.getElementById('lab_wait').innerText='Please wait ......';" & _
                                                    "javascript:document.getElementById('but_reCaculate').disabled=true;" & _
                                                    Me.Page.GetPostBackEventReference(but_reCaculate))
        If Not Me.IsPostBack Then

            If Request("PART") <> Nothing Then
                Me.tb_part.Text = Request("PART")
                Me.tb_measItem.Text = Request("ITEM")
                Me.tx_OUSL.Text = Request("U")
                Me.tx_OCL.Text = Request("C")
                Me.tx_OLSL.Text = Request("L")
            End If

        End If

    End Sub

    Protected Sub RBL_SPCTYPE_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles RBL_SPCTYPE.SelectedIndexChanged

        If RBL_SPCTYPE.SelectedIndex = 0 Then
            Me.tr_usl.Visible = True
            Me.tr_lsl.Visible = True
        ElseIf RBL_SPCTYPE.SelectedIndex = 1 Then
            Me.tr_usl.Visible = True
            Me.tr_lsl.Visible = False
        Else
            Me.tr_usl.Visible = False
            Me.tr_lsl.Visible = True
        End If

    End Sub

    ' 重算
    Protected Sub but_reCaculate_Click(sender As Object, e As System.EventArgs) Handles but_reCaculate.Click

        ' Step1. 更新表格 CTF_Monitor_Performance_Meas_Item_SPEC
        ' Step2. 取得資料 CTF_Monitor_Performance_RawData 
        ' Step3. 更新表格 CTF_Monitor_Performance_Lot_Summary 

        Dim sqlStr As String = ""
        Dim updateStr As String = ""
        Dim insertStr As String = ""
        Dim myDT As DataTable
        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim comm As SqlCommand = Nothing
        
        If RBL_SPCTYPE.SelectedIndex = 0 Then ' 雙邊
            updateStr = "update set SPC_TYPE='{0}', USL={1}, LSL={2} "
            insertStr = "insert (SPC_TYPE, USL, LSL) values('{0}', {1}, {2});"
            updateStr = String.Format(updateStr, "2", tb_USL.Text.Trim(), tb_LSL.Text.Trim())
            insertStr = String.Format(insertStr, "2", tb_USL.Text.Trim(), tb_LSL.Text.Trim())
        ElseIf RBL_SPCTYPE.SelectedIndex = 1 Then ' 單邊[上]
            updateStr = "update set SPC_TYPE='{0}', USL={1} "
            insertStr = "insert (SPC_TYPE, USL) values('{0}', {1});"
            updateStr = String.Format(updateStr, "U", tb_USL.Text.Trim())
            insertStr = String.Format(insertStr, "U", tb_USL.Text.Trim())
        Else ' 單邊[下]
            updateStr = "update set SPC_TYPE='{0}', USL={1} "
            insertStr = "insert (SPC_TYPE, LSL) values('{0}', {1});"
            updateStr = String.Format(updateStr, "L", tb_LSL.Text.Trim())
            insertStr = String.Format(insertStr, "L", tb_LSL.Text.Trim())
        End If

        sqlStr += "MERGE INTO CTF_Monitor_Performance_Meas_Item_SPEC as t_fv "
        sqlStr += "USING (select '{0}' PART_ID, '{1}' MEAS_ITEM) as s_fv "
        sqlStr += "ON t_fv.PART_ID = s_fv.PART_ID and t_fv.MEAS_ITEM = s_fv.MEAS_ITEM "
        sqlStr += "WHEN MATCHED THEN "
        sqlStr += updateStr
        sqlStr += "WHEN NOT MATCHED THEN "
        sqlStr += insertStr
        sqlStr = String.Format(sqlStr, (tb_part.Text.Trim()), (tb_measItem.Text.Trim()))
        '----- Step1 START ---
        Try

            conn.Open()
            comm = conn.CreateCommand()
            comm.CommandText = sqlStr
            comm.ExecuteNonQuery()
            conn.Close()

        Catch ex As Exception
            Exit Sub
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        '----- Step2 & Step3 START ---
        Try
            CalStatistics((Me.tb_part.Text), (Me.tb_measItem.Text))
            exeScript()
        Catch ex As Exception
            Exit Sub
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' Statistics
    Private Sub CalStatistics(ByVal part_id As String, ByVal meas_item As String)

        Dim connRead As SqlConnection = Nothing
        Dim sAdpt As SqlDataAdapter = Nothing
        Dim dtRaw As DataTable
        Dim sqlStr As String = ""

        Try

            sqlStr += "select Lot_Id, Machine_Id, Data_Value "
            sqlStr += "from dbo.CTF_Monitor_Performance_RawData "
            sqlStr += "where 1 = 1 "
            sqlStr += "and Part_Id='{0}' "
            sqlStr += "and Meas_Item='{1}' "
            sqlStr += "group by Lot_Id, Machine_Id, Data_Value "
            sqlStr += "order by Lot_Id, Machine_Id "
            sqlStr = String.Format(sqlStr, part_id, meas_item)

            connRead = New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
            connRead.Open()
            sAdpt = New SqlDataAdapter(sqlStr, connRead)
            dtRaw = New DataTable
            sAdpt.Fill(dtRaw)
            connRead.Close()

            Dim arrSeries() As Decimal
            Dim noLineSpc As Boolean = False
            Dim FMACHINE As String = ""
            Dim BMACHINE As String = ""
            Dim FLOT As String = ""
            Dim BLOT As String = ""
            Dim rawAry As ArrayList = New ArrayList

            For i As Integer = 0 To (dtRaw.Rows.Count - 1)

                If i = 0 Then
                    FLOT = dtRaw.Rows(i)("LOT_ID")
                    FMACHINE = dtRaw.Rows(i)("Machine_Id")
                End If
                BLOT = dtRaw.Rows(i)("LOT_ID")
                BMACHINE = dtRaw.Rows(i)("Machine_Id")

                If FLOT <> BLOT Then

                    ' 計算統計值
                    ReDim arrSeries(rawAry.Count - 1)
                    For j As Integer = 0 To (rawAry.Count - 1)
                        arrSeries(j) = CType(rawAry(j), Decimal)
                    Next
                    InsertStatistics(arrSeries, FLOT, FMACHINE)

                    FLOT = BLOT
                    FMACHINE = BMACHINE
                    rawAry = New ArrayList

                End If

                If Not IsDBNull(dtRaw.Rows(i)("Data_Value")) Then
                    rawAry.Add(dtRaw.Rows(i)("Data_Value"))
                End If

            Next

            ' 計算統計值 - 最後一筆
            ReDim arrSeries(rawAry.Count - 1)
            For j As Integer = 0 To (rawAry.Count - 1)
                arrSeries(j) = CType(rawAry(j), Decimal)
            Next
            InsertStatistics(arrSeries, BLOT, BMACHINE)

        Catch ex As Exception

        Finally
            If connRead.State = ConnectionState.Open Then
                connRead.Close()
            End If
        End Try

    End Sub

    ' Insert Statistics Data to DataBase
    Private Sub InsertStatistics(ByRef decimalAry() As Decimal, ByVal lot_id As String, ByVal machine_id As String)

        Dim objFormula As SPCFormula = Nothing
        Dim objStatisc As SPCStatisc = Nothing
        Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim comm As SqlCommand = Nothing
        Dim sqlStr As String = ""
        Dim usl As Decimal
        Dim lsl As Decimal
        Dim decimalAryStr As String = ""

        For i As Integer = 0 To (decimalAry.Length - 1)
            decimalAryStr += Math.Round(decimalAry(i), 5).ToString + "|"
        Next
        decimalAryStr = decimalAryStr.Substring(0, decimalAryStr.Length - 1)

        If RBL_SPCTYPE.SelectedIndex = 0 Then ' 雙邊
            usl = CType(Me.tb_USL.Text, Decimal)
            lsl = CType(Me.tb_LSL.Text, Decimal)
            objFormula = New SPCFormula(decimalAry, usl, lsl, Nothing)
            objStatisc = objFormula.Calculate()
        ElseIf RBL_SPCTYPE.SelectedIndex = 1 Then ' 單邊[上]
            usl = CType(Me.tb_USL.Text, Decimal)
            objFormula = New SPCFormula(decimalAry, usl, Nothing, Nothing)
            objStatisc = objFormula.Calculate()
        Else ' 單邊[下]
            lsl = CType(Me.tb_LSL.Text, Decimal)
            objFormula = New SPCFormula(decimalAry, Nothing, lsl, Nothing)
            objStatisc = objFormula.Calculate()
        End If

        sqlStr = ""
        sqlStr += "update CTF_Monitor_Performance_Lot_Summary "
        sqlStr += "set Mean_Value={0}, Std_Value={1}, Min_Value={2}, Max_Value={3}, CP={4}, CPK={5}, CSL={6} "
        sqlStr += "where 1=1 "
        sqlStr += "and part_id='{7}' "
        sqlStr += "and meas_item='{8}' "
        sqlStr += "and lot_id='{9}' "
        sqlStr += "and machine_id='{10}' "
        sqlStr += "and rowdata='{11}' "

        sqlStr = String.Format(sqlStr,
                               Math.Round(objStatisc.xMean, 6).ToString(),
                               Math.Round(objStatisc.Sigma, 6).ToString(),
                               Math.Round(objStatisc.Minimum, 6).ToString(),
                               Math.Round(objStatisc.Maximum, 6).ToString(),
                               (objStatisc.Cp).ToString(),
                               (objStatisc.Cpk).ToString(),
                               (objStatisc.Target).ToString(),
                               (Me.tb_part.Text.Trim()),
                               (Me.tb_measItem.Text.Trim()),
                               lot_id,
                               machine_id,
                               decimalAryStr)

        Try

            conn.Open()
            comm = conn.CreateCommand()
            comm.CommandText = sqlStr
            comm.ExecuteNonQuery()
            conn.Close()

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    ' 導回原頁面
    Public Sub exeScript()

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("window.opener.document.getElementById('but_Execute').click();")
        sb.Append("this.close();")
        sb.Append("</script>")
        Dim myCSManager As ClientScriptManager = Page.ClientScript
        myCSManager.RegisterStartupScript(Me.GetType(), "SetStatusScript", sb.ToString())

    End Sub

End Class
