<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CTF_ChartDetail.aspx.vb" Inherits="CTF_ChartDetail" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Critical Item Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<script src="js/jsdomenu.js" type="text/javascript"></script>
<script src="js/jsdomenu.inc.js" type="text/javascript"></script>
</head>
<body style="MARGIN-TOP: 10px; MARGIN-LEFT: 10px" onload="initjsDOMenu();" MS_POSITIONING="FlowLayout">
<form id="form1" runat="server">
<table border="1">
 
<!-- Chart Panel -->
<tr>
<td style="height: 1px;" >
<asp:Panel ID="ChartPanel" runat="server" BorderColor="#99FF33" BackColor="White" EnableViewState="False">
</asp:Panel>
</td>
</tr>

<!-- Grid View Start -->
<tr>
<td colspan=4>
<asp:GridView ID="Lot_GridView" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3">
<AlternatingRowStyle BackColor="#DBEEFF" />
<EditRowStyle BackColor="#999999" />
<FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
<HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
<PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
<RowStyle BackColor="#ffffff" ForeColor="Black"/>
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
