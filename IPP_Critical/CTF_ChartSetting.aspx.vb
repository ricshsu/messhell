
Partial Class CTF_ChartSetting
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not Me.IsPostBack Then



        End If

    End Sub

    Protected Sub but_reCaculate_Click(sender As Object, e As System.EventArgs) Handles but_reCaculate.Click

        Dim chartType As String = "0"
        Dim chartValue As String = "0"
        Dim isGroup As String = "0"

        chartType = rb_chartType.SelectedValue
        chartValue = rdb_chartValue.SelectedValue
        isGroup = "1"

        Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder()
        sb.Append("<script language='javascript'>")
        sb.Append("window.opener.document.getElementById('txb_chart_value').value='" + chartType + "-" + chartValue + "-" + isGroup + "';")
        sb.Append("this.close();")
        sb.Append("</script>")
        Dim myCSManager As ClientScriptManager = Page.ClientScript
        myCSManager.RegisterStartupScript(Me.GetType(), "SetStatusScript", sb.ToString())

    End Sub

End Class
