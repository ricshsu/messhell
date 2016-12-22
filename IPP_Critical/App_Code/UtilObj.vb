Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics

Public Class UtilObj


    Public Shared Function FillController(ByVal DT As DataTable, ByRef ddl As DropDownList, ByVal nAll As Integer, ByVal valueName As String, ByVal textName As String) As Boolean

        Dim status As Integer = False
        Dim listItem As ListItem
        Try

            ddl.Items.Clear()
            If nAll = 1 Then
                ddl.Items.Add("All")
            End If

            For i As Integer = 0 To DT.Rows.Count - 1
                If Not DBNull.Value.Equals(DT.Rows(i)(0)) Then
                    listItem = New ListItem
                    listItem.Value = DT.Rows(i)(valueName).ToString()
                    listItem.Text = DT.Rows(i)(textName).ToString()
                    ddl.Items.Add(listItem)
                End If
            Next

            If ddl.Items.Count > 0 Then
                ddl.Items(0).Selected = True
            End If
            status = True
        Catch ex As Exception
            status = False
        End Try

        Return status

    End Function

    Public Shared Function FillController(ByVal DT As DataTable, ByRef ddl As DropDownList, ByVal nAll As Integer) As Boolean

        Dim status As Integer = False
        Try

            ddl.Items.Clear()
            If nAll = 1 Then
                ddl.Items.Add("All")
            End If

            For i As Integer = 0 To DT.Rows.Count - 1
                If Not DBNull.Value.Equals(DT.Rows(i)(0)) Then
                    ddl.Items.Add(DT.Rows(i)(0))
                End If
            Next

            If ddl.Items.Count > 0 Then
                ddl.Items(0).Selected = True
            End If
            status = True
        Catch ex As Exception
            status = False
        End Try

        Return status

    End Function

    Public Shared Function FillController(ByRef DT As DataTable, ByRef ddl As DropDownList, ByVal nAll As Integer, ByVal colName As String) As Boolean

        Dim status As Integer = False
        Try

            ddl.Items.Clear()
            If nAll = 1 Then
                ddl.Items.Add("All")
            End If

            For i As Integer = 0 To DT.Rows.Count - 1
                If Not DBNull.Value.Equals(DT.Rows(i)(colName)) Then
                    ddl.Items.Add(DT.Rows(i)(colName))
                End If
            Next

            If ddl.Items.Count > 0 Then
                ddl.Items(0).Selected = True
            End If
            status = True
        Catch ex As Exception
            status = False
        End Try

        Return status

    End Function

    Public Shared Function FillLitsBoxController(ByVal DT As DataTable, ByRef ddlSource As ListBox, ByVal nAll As Integer) As Boolean

        Dim status As Integer = False
        Try
            ddlSource.Items.Clear()
            For i As Integer = 0 To DT.Rows.Count - 1
                If Not DBNull.Value.Equals(DT.Rows(i)(0)) Then
                    ddlSource.Items.Add(DT.Rows(i)(0))
                End If
            Next
            status = True
        Catch ex As Exception
            status = False
        End Try

        Return status

    End Function

    Public Shared Function Set_DataGridRow_OnMouseOver_Color(ByRef DataGrid As GridView, ByVal MouseOverColor As String, ByVal MouseOutColor As System.Drawing.Color)
        Dim i As Integer
        For i = 0 To (DataGrid.Rows.Count - 1)
            DataGrid.Rows(i).Attributes("onmouseover") = "this.style.backgroundColor='" & MouseOverColor & "';"
            If i Mod 2 = 0 Then
                DataGrid.Rows(i).Attributes("onmouseout") = "this.style.backgroundColor='';"
            Else
                DataGrid.Rows(i).Attributes("onmouseout") = "this.style.backgroundColor='#" & GetColorCode(MouseOutColor) & "';"
            End If
        Next
    End Function

    Shared Function GetColorCode(ByVal Icolor As System.Drawing.Color) As String
        If Icolor.Equals(Color.Transparent) Then
            GetColorCode = ""
            Exit Function
        End If
        If Icolor.Equals(Color.Empty) Then
            GetColorCode = ""
            Exit Function
        End If
        GetColorCode = Right("00" & Hex(Icolor.R), 2) & Right("00" & Hex(Icolor.G), 2) & Right("00" & Hex(Icolor.B), 2)
    End Function

    Public Shared Function fun_DataTable_SelectDistinct(ByVal SourceTable As DataTable, ByVal ParamArray FieldNames() As String) As DataTable

        Dim lastValues() As Object
        Dim newTable As DataTable
        If FieldNames Is Nothing OrElse FieldNames.Length = 0 Then
            Throw New ArgumentNullException("FieldNames")
        End If
        lastValues = New Object(FieldNames.Length - 1) {}
        newTable = New DataTable

        For Each field As String In FieldNames
            newTable.Columns.Add(field, SourceTable.Columns(field).DataType)
        Next

        For Each Row As DataRow In SourceTable.Select("", String.Join(", ", FieldNames))
            If Not fieldValuesAreEqual(lastValues, Row, FieldNames) Then
                newTable.Rows.Add(createRowClone(Row, newTable.NewRow(), FieldNames))
                setLastValues(lastValues, Row, FieldNames)
            End If
        Next
        Return newTable

    End Function

    Private Shared Function fieldValuesAreEqual(ByVal lastValues() As Object, ByVal currentRow As DataRow, ByVal fieldNames() As String) As Boolean
        Dim areEqual As Boolean = True

        For i As Integer = 0 To fieldNames.Length - 1
            If lastValues(i) Is Nothing OrElse Not lastValues(i).Equals(currentRow(fieldNames(i))) Then
                areEqual = False
                Exit For
            End If
        Next

        Return areEqual
    End Function

    Private Shared Function createRowClone(ByVal sourceRow As DataRow, ByVal newRow As DataRow, ByVal fieldNames() As String) As DataRow
        For Each field As String In fieldNames
            newRow(field) = sourceRow(field)
        Next

        Return newRow
    End Function

    Private Shared Sub setLastValues(ByVal lastValues() As Object, ByVal sourceRow As DataRow, ByVal fieldNames() As String)
        For i As Integer = 0 To fieldNames.Length - 1
            lastValues(i) = sourceRow(fieldNames(i))
        Next
    End Sub

    Private Sub KillProcess(ByVal ProcName As String) '關閉工作管理員的 excel 
        Dim thisProc As System.Diagnostics.Process
        Dim allRelationalProcs() As Process = System.Diagnostics.Process.GetProcessesByName(ProcName)
        For Each thisProc In allRelationalProcs
            If Not thisProc.CloseMainWindow() Then
                thisProc.Kill()
            End If
        Next
    End Sub

    Public Shared Sub exeScript(ByRef myCSManager As ClientScriptManager)

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("window.opener.document.getElementById('but_Execute').click();")
        sb.Append("this.close();")
        sb.Append("</script>")
        myCSManager.RegisterStartupScript(myCSManager.GetType(), "SetStatusScript", sb.ToString())

    End Sub

End Class
