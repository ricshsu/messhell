<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DailyReportDetail.aspx.vb" Inherits="DailyReportDetail" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Daily Yield Report Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<script type="text/javascript" language='javascript'>
function openWindowWithPost(url, name, vala, valb, valc, vald) 
{
    
    var newWindow = window.open(url,  name, 'height=700,width=850,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=no,status=no');
    var html = "";
    html += "<html><head></head><body><form id='formid' method='post' action='" + url + "'>";
    html += "<input type='hidden' name='P' value='" + vala + "'/>";
    html += "<input type='hidden' name='F' value='" + valb + "'/>";
    html += "<input type='hidden' name='W' value='" + valc + "'/>";
    html += "<input type='hidden' name='WI' value='" + vald + "'/>";
    html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
    html += "script></body></html>";
    newWindow.document.write(html);

}
function LinkPoint(dataID) 
{
    var linkPoint = document.getElementById(('gv_lotYield_' + dataID));
    linkPoint.style.background = "#FF6A6A";
    document.location = ('#' + dataID);
}
</script>
</head>
<body>
<form id="form1" runat="server">
<table>

<tr>
<td>
<table>
<asp:Panel ID="ThendPanel" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
</table>
</td>
</tr>

<tr>
<td style='width=1100px'>
<asp:GridView ID="gv_lotYield" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3">
<AlternatingRowStyle BackColor="#DBEEFF" />
<EditRowStyle BackColor="#999999" />
<Columns>
<asp:TemplateField>
<HeaderTemplate>No.</HeaderTemplate>
<ItemTemplate><%# Container.DataItemIndex + 1 %></ItemTemplate>
    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
</asp:TemplateField>
</Columns>
<FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
<HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
<PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
<RowStyle BackColor="#ffffff" ForeColor="Black" Width="80"/>
<SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
<SortedAscendingCellStyle BackColor="#E9E7E2" />
<SortedAscendingHeaderStyle BackColor="#506C8C" />
<SortedDescendingCellStyle BackColor="#FFFDF8" />
<SortedDescendingHeaderStyle BackColor="#6F8DAE" />
<SortedAscendingCellStyle BackColor="#E9E7E2"></SortedAscendingCellStyle>
<SortedAscendingHeaderStyle BackColor="#506C8C"></SortedAscendingHeaderStyle>
<SortedDescendingCellStyle BackColor="#FFFDF8"></SortedDescendingCellStyle>
<SortedDescendingHeaderStyle BackColor="#6F8DAE"></SortedDescendingHeaderStyle>
</asp:GridView>


</td>
</tr>

</table>
</form>
</body>
</html>
