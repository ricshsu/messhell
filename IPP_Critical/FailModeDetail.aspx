<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FailModeDetail.aspx.cs" Inherits="FailModeDetail" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Fail Mode Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<script language=javascript>
function LinkPoint(dataID) 
{
    var linkPoint = document.getElementById(('GV_LotRowData_' + dataID));
    linkPoint.style.background = "#FF6A6A";
    document.location = ('#' + dataID);
}
</script>
</head>
<body>
<form id="form1" runat="server">
<table border="1" width="1080px">

<asp:Label ID="lab_custom" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_product" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_productType" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_partID" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_dateF" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_dateT" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_modeType" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_PA" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_PB" runat="server" Text="" Visible="False"></asp:Label>
<asp:Label ID="lab_PC" runat="server" Text="" Visible="False"></asp:Label>

<!-- Title -->
<tr>
<td colspan='6' class="Table_One_Title" valign='middle' align="center" style='font-size:large;font-weight: bold'>
Fail&nbsp; Mode&nbsp; Detail
</td>
</tr>

<!-- Product OR Part-->
<tr id="tr_function_condition" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>
    Mode Matching 
</td>
<td colspan=5 style='width:880px'>
<asp:RadioButtonList ID="rb_mode_selected" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" Width="300px" onselectedindexchanged="rb_mode_selected_SelectedIndexChanged" >
<asp:ListItem Value="0">Fail Mode</asp:ListItem>
<asp:ListItem Value="1">Defect Code</asp:ListItem>
<asp:ListItem Value="2">PT AOI</asp:ListItem>
</asp:RadioButtonList>
</td>
</tr>

<!-- PT AOI -->
<tr id="tr_ptaoi" runat="server" visible="false">
    <td align=left style='background:#E9E7E2; ' class="style1">PT AOI
    </td>
    <td colspan=5 class="style2">      
    <table>
    <tr>
    <td colspan='2'><asp:Label ID="lab_generalType" runat="server">Layer : </asp:Label></td>
    <td>
    <asp:DropDownList runat="server" ID="ddl_ptaoi_layer" Width="150px" 
            AutoPostBack="True" 
            onselectedindexchanged="ddl_ptaoi_layer_SelectedIndexChanged"></asp:DropDownList>
    </td>
    </tr>
    <tr>

    <td>
    <asp:ListBox ID="listB_ptaoi_source" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    <td style='width:20px; height:20px'>
    <asp:Button ID="but_ptaoi_right" runat="server" Text=">>" CssClass="BT_2" onclick="but_right_Click" />
    <asp:Button ID="but_ptaoi_left"  runat="server" Text="<<" CssClass="BT_2" onclick="but_left_Click"  />
    </td>
    
    <td>      
    <asp:ListBox ID="listB_ptaoi_display" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    <td align=left valign=bottom>
    <asp:Button ID="but_ptaoi_decision" runat="server" Text="OK" Height="28px" 
            Width="40px" CssClass="BT_2" Font-Bold="True" Font-Size="Medium" 
            onclick="but_ptaoi_decision_Click" />
    </td>
    </tr>
    </table>
    </td>
</tr>

<!-- Stage -->
<tr id='tr_Stage' runat='server' visible='false' style='height:80px'>
<td align=left style='background:#E9E7E2; width:200px; height:80px;'>Stage (Category)</td>
<td valign='bottom' colspan='5' style='width:880px; height:80px;'> 
<table style='height:80px'>
<tr style='height:80px'>
<td valign='bottom' style='width:150px; height:80px;'>       
<asp:ListBox ID="lb_StageSource" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>
<td style='width:20px; height:80px;'>
<asp:Button ID="but_stageTo" runat="server" Text=">>" CssClass=BT_2 onclick="but_right_Click" />
<asp:Button ID="but_stageBack" runat="server" Text="<<" CssClass=BT_2 onclick="but_left_Click"/>
</td>
<td valign='bottom' style='width:150px; height:80px;'>      
<asp:ListBox ID="lb_StageShow" runat="server" Height="80px" Width="150px" 
        SelectionMode="Multiple"></asp:ListBox>
</td>
<td valign='bottom' style='width:20px;'>
<asp:Button ID="but_stageOK" runat="server" Text="OK" Height="25px" Width="35px" CssClass="BT_2" Font-Bold="True" Font-Size="Medium" onclick="but_stageOK_Click" />
</td>
</tr> 
</table> 
</td>
</tr>

<!-- Fail Mode -->
<tr id='tr_failMode' runat='server' visible='false' style='height:80px'>
<td align=left style='background:#E9E7E2; width:200px; height:80px;'>Fail Mode</td>
<td valign='bottom' colspan='5' style='width:880px; height:80px;'> 
<table style='height:80px'>
<tr style='height:80px'>
<td valign='bottom' style='width:150px; height:80px;'>       
<asp:ListBox ID="lb_failModeSource" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>
<td style='width:20px; height:80px;'>
<asp:Button ID="but_failTo" runat="server" Text=">>" CssClass=BT_2 onclick="but_right_Click" />
<asp:Button ID="but_failBack" runat="server" Text="<<" CssClass=BT_2 onclick="but_left_Click" />
</td>
    
<td valign='bottom' style='width:150px; height:80px;'>      
<asp:ListBox ID="lb_failModeShow" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>
<td valign='bottom' style='width:20px;'>
<asp:Button ID="but_failModeOK" runat="server" Text="OK" Height="25px" Width="35px" 
        CssClass="BT_2" Font-Bold="True" Font-Size="Medium" 
        onclick="but_failModeOK_Click" />
</td>
</tr> 
</table> 
</td>
</tr>

<!-- Defect Code -->
<tr id='tr_defectCode' runat='server' visible='false' style='height:80px'>
<td align=left style='background:#E9E7E2; width:200px; height:80px;'>Defect Code</td>
<td valign='bottom' colspan='5' style='width:880px; height:80px;'> 
<table style='height:80px'>
<tr style='height:80px'>
<td valign='bottom' style='width:150px; height:80px;'>       
<asp:ListBox ID="lb_dcodeSource" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>
<td style='width:20px; height:80px;'>
<asp:Button ID="but_dcodeTo" runat="server" Text=">>" CssClass=BT_2 onclick="but_right_Click" />
<asp:Button ID="but_dcodeBack" runat="server" Text="<<" CssClass=BT_2 onclick="but_left_Click" />
</td>
    
<td valign='bottom' style='width:150px; height:80px;'>      
<asp:ListBox ID="lb_dcodeShow" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>

<td valign='bottom' style='width:20px;'>
<asp:Button ID="but_DefectCodeOK" runat="server" Text="OK" Height="25px" Width="35px" 
        CssClass="BT_2" Font-Bold="True" Font-Size="Medium" 
        onclick="but_DefectCodeOK_Click" />
</td>

</tr> 
</table> 
</td>
</tr>

<!-- Query -->
<tr id="tr_execute" runat="server" visible="false">
<td colspan='6' align='right' width="100%">
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label> 
    &nbsp;&nbsp;&nbsp;
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Inquery" Height=30px 
        Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium" 
        onclick="but_Execute_Click"/>
    </td>
</tr>

<!-- Chart Panel -->
<tr id="tr_chartPanel" runat="server" visible="false" enableviewstate="true">
<td colspan='6'>
<table>
<tr>
<td class='Table_Two_Title' valign='middle' align='center' style='width:700px;font-size:middle;font-weight:bold'>
<asp:Label ID="lab_chartTitle" runat="server" Text=""></asp:Label>
</td>
<td style='width:400px'>
</td>
</tr>
<tr><td colspan=2 valign='middle' align="left" style='font-size:x-large;font-weight: bold'>
<asp:Chart ID="chartObj" runat="server" EnableViewState="True"></asp:Chart>
</td></tr>
</td>
</tr>
</table>
</td>
</tr>
</table>

<table>

<!-- Result -->
<tr id="tr_result" runat="server" visible="false">
<td colspan='6' align='left' width="100%">
<asp:GridView ID="GV_LotRowData" runat="server" BackColor="White" 
        BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
        onrowdatabound="GV_LotRowData_RowDataBound">
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
