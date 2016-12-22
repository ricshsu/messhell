Imports System.Data.SqlClient
Imports System.Data

Partial Class CTF_RowData
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not Me.IsPostBack Then

            Try

                ViewState("PART") = Request("PART") 'PART 
                ViewState("LOT") = Request("LOT") 'LOT
                ViewState("Machine") = Request("M") 'Machine
                ViewState("ITEM") = Request("ITEM") 'ITEM
                ViewState("ETime") = Request("ET") 'End Time
                pageInit()

            Catch ex As Exception

            End Try

        End If

    End Sub

    Private Sub pageInit()

        Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("iSVRConnectionString").ToString)
        Dim sqlStr As String = ""
        Dim myDT As DataTable
        Dim myAdapter As SqlDataAdapter

        sqlStr += "select Part_Id, lot_id, Machine_id, Meas_item, "
        sqlStr += "round(Max_Value, 5) as Max_Value, "
        sqlStr += "round(Min_Value, 5) as Min_Value, "
        sqlStr += "round(Mean_value, 5) as Mean_value, "
        sqlStr += "round(Std_value, 5) as Std_value, "
        sqlStr += "round(Cp, 5) as CP, "
        sqlStr += "round(Cpk, 5) as CPK, "
        sqlStr += "RowData "
        sqlStr += "from CTF_Monitor_Performance_Lot_Summary "
        sqlStr += "where 1=1 "
        sqlStr += "and part_id='" + ViewState("PART") + "' "
        sqlStr += "and lot_id='" + ViewState("LOT") + "' "
        sqlStr += "and Lot_Meas_End_DataTime='" + ViewState("ETime") + "' "
        sqlStr += "order by Meas_item "

        Try

            conn.Open()
            myAdapter = New SqlDataAdapter(sqlStr, conn)
            myDT = New DataTable
            myAdapter.Fill(myDT)
            conn.Close()

            gv_rowdata.DataSource = myDT
            gv_rowdata.DataBind()
            UtilObj.Set_DataGridRow_OnMouseOver_Color(gv_rowdata, "#FFF68F", gv_rowdata.AlternatingRowStyle.BackColor)

        Catch ex As Exception

        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try


    End Sub

End Class
