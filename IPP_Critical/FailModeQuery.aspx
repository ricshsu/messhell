<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FailModeQuery.aspx.cs" Inherits="FailModeQuery" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>Fail Mode Query</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />

<script language="javascript" type="text/javascript">
function openWindowWithPost(customerID, product, productType, partID, dateF, dateT, modeType, PA, PB, PC) 
{
    var newWindow = window.open("FailModeDetail.aspx", "FailDetail", 'height=700, width=850, top=100, left=200,toolbar=yes,menubar=yes,scrollbars=yes,resizable=yes,location=yes,status=yes');
    var html = "";
    html += "<html><head></head><body><form id='formid' method='post' action='FailModeDetail.aspx'>";
    html += "<input type='hidden' name='customerID' value='" + customerID + "'/>";
    html += "<input type='hidden' name='product' value='" + product + "'/>";
    html += "<input type='hidden' name='productType' value='" + productType + "'/>";
    html += "<input type='hidden' name='partID' value='" + partID + "'/>";
    html += "<input type='hidden' name='dateF' value='" + dateF + "'/>";
    html += "<input type='hidden' name='dateT' value='" + dateT + "'/>";
    html += "<input type='hidden' name='modeType' value='" + modeType + "'/>";
    html += "<input type='hidden' name='PA' value='" + PA + "'/>";
    html += "<input type='hidden' name='PB' value='" + PB + "'/>";
    html += "<input type='hidden' name='PC' value='" + PC + "'/>";
    html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
    html += "script></body></html>";
    newWindow.document.write(html);
}
</script>


<script type="text/javascript" language='javascript'>
    
</script>

<style type="text/css">
        .style1
        {
            width: 200px;
            height: 80px;
        }
        .style2
        {
            height: 80px;
        }
    </style>

</head>
<body>
<form id="form1" runat="server">
<table border="1" width="1080px">
<!-- Title -->
<tr>
<td colspan='6' class="Table_One_Title" valign='middle' align="center" style='font-size:large;font-weight: bold'>
Fail&nbsp; Mode&nbsp; Query
</td>
</tr>
<!-- Product OR Part-->
<tr id="tr_selectMode" runat='server' visible='false'>
<td align='left' style='background:#E9E7E2; width:200px'>
Lot Select Mode
</td>
<td colspan=5 style='width:880px'>
<asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" onselectedindexchanged="RadioButtonList1_SelectedIndexChanged">
<asp:ListItem Value="0" Selected="True">Condition Select</asp:ListItem>
<asp:ListItem Value="1">Upload Lot</asp:ListItem>
</asp:RadioButtonList>
</td>
</tr>
<!-- Customer -->
<tr id="tr_customer" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>Customer</td>
<td colspan=5 style='width:880px'>
<asp:DropDownList ID="ddlCustomer" runat="server" Width="150px" AutoPostBack="True" 
        Height="25px" onselectedindexchanged="ddlCustomer_SelectedIndexChanged"></asp:DropDownList>
</td>
</tr>
<!-- CPU OR ChipSet -->
<tr id="tr_cpucs" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>CPU or ChipSet</td>
<td colspan=5 style='width:880px'>
<asp:DropDownList ID="ddlCategory" runat="server" Width="150px" Height="25px" 
        AutoPostBack="True" onselectedindexchanged="ddlCategory_SelectedIndexChanged"></asp:DropDownList>
</td>
</tr>
<!-- Product OR Part-->
<tr id="tr_partSelect" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>
Product Type OR Part ID
</td>
<td colspan=5 style='width:880px'>
<asp:RadioButtonList ID="rbl_BySource" runat="server" RepeatDirection="Horizontal" 
        AutoPostBack="True" Width="180px" 
        onselectedindexchanged="rbl_BySource_SelectedIndexChanged">
<asp:ListItem Value="0" Selected="True">Product Type</asp:ListItem>
<asp:ListItem Value="1">Part ID</asp:ListItem>
</asp:RadioButtonList>
</td>
</tr>
<!-- Product OR Part Data -->
<tr id="tr_partData" runat='server' visible='true'>
<td align=left style='background:#E9E7E2; width:200px'>
    <asp:Label ID="lab_ProductType_PartID" runat="server" Text="Product Type"></asp:Label>
    </td>
<td colspan=5 style='width:880px'>
<asp:DropDownList ID="ddlPart" runat="server" Width="150px" Height="25px" 
        onselectedindexchanged="ddlPart_SelectedIndexChanged" AutoPostBack="True"></asp:DropDownList>
</td>
</tr>
<!-- Upload Lot -->
<tr id="tr_upload" runat=server visible=false>
<td align='left' style='background:#E9E7E2; width:200px'>Upload Lot</td>
<td colspan=5 style='width:880px'>
<asp:FileUpload ID="uf_UfilePath" runat="server" Width=255px ></asp:FileUpload>
&nbsp;
<asp:Button ID="but_Uupload" runat="server" Text="Upload" onclick="but_Uupload_Click" />
</td>
</tr>
<!-- Lot List -->
<tr id="tr_lotlist" runat='server' visible='false'>
<td align='left' style='background:#E9E7E2; width:200px'>Upload Lot List</td>
<td colspan=5 style='width:880px'>
<asp:ListBox ID="lb_lotList" runat="server" Height="100px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
</td>
</tr>

<tr id="tr_time" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>
Time Query Range
</td>
<td colspan=5 style='width:880px'>
<!-- By Date -->
<asp:RadioButton ID="rb_byDate" runat="server" text="By Date" AutoPostBack="True" 
        Checked="True" oncheckedchanged="rb_byDate_CheckedChanged" 
        Visible="False" />
<asp:TextBox ID="txtDateFrom" runat="server" Columns="10" MaxLength="10" 
        Width="110px" Enabled=true ontextchanged="txtDateFrom_TextChanged"></asp:TextBox>
<obout:Calendar ID="Calendar1" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
TextArrowRight=">>" TextBoxId="txtDateFrom"></obout:Calendar>
&nbsp;~&nbsp;
<asp:TextBox ID="txtDateTo" runat="server" Columns="10" MaxLength="10" Width="110px" Enabled=true></asp:TextBox>
<obout:Calendar ID="Calendar2" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
TextArrowRight=">>" TextBoxId="txtDateTo">
</obout:Calendar>  
<!-- By Week -->
&nbsp;&nbsp;&nbsp;
<asp:RadioButton ID="rb_byWeek" runat="server" text="By Week" AutoPostBack="True" 
        oncheckedchanged="rb_byWeek_CheckedChanged" Visible=false />&nbsp;&nbsp;&nbsp;
<asp:DropDownList ID="ddlWeekStart" runat="server"  Visible=false></asp:DropDownList>
<asp:DropDownList ID="ddlWeekEnd" runat="server" Enabled=false Visible=false></asp:DropDownList>
</td>
</tr>

<!-- Product OR Part-->
<tr id="tr_function_condition" runat='server' visible='true'>
<td align='left' style='background:#E9E7E2; width:200px'>
Mode Condition
</td>
<td colspan=5 style='width:880px'>
<asp:RadioButtonList ID="rb_mode_selected" runat="server" 
        RepeatDirection="Horizontal" AutoPostBack="True" Width="300px" 
        onselectedindexchanged="rb_mode_selected_SelectedIndexChanged" >
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
    <asp:Button ID="but_ptaoi_right" runat="server" Text=">>" CssClass="BT_2" 
            onclick="but_ptaoi_right_Click" />
    <asp:Button ID="but_ptaoi_left"  runat="server" Text="<<" CssClass="BT_2" 
            onclick="but_ptaoi_left_Click"  />
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
<asp:Button ID="but_stageTo" runat="server" Text=">>" CssClass=BT_2 
        onclick="but_stageTo_Click" />
<asp:Button ID="but_stageBack" runat="server" Text="<<" CssClass=BT_2 
        onclick="but_stageBack_Click"/>
</td>
<td valign='bottom' style='width:150px; height:80px;'>      
<asp:ListBox ID="lb_StageShow" runat="server" Height="80px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
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
<asp:Button ID="but_failTo" runat="server" Text=">>" CssClass=BT_2 
        onclick="but_failTo_Click" />
<asp:Button ID="but_failBack" runat="server" Text="<<" CssClass=BT_2 
        onclick="but_failBack_Click" />
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
<asp:Button ID="but_dcodeTo" runat="server" Text=">>" CssClass=BT_2 
        onclick="but_dcodeTo_Click" />
<asp:Button ID="but_dcodeBack" runat="server" Text="<<" CssClass=BT_2 
        onclick="but_dcodeBack_Click" />
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
<tr id="tr_chartPanel" runat="server" visible="false">
<td colspan='6'>
<table>
<asp:Panel ID="ChartPanel" runat="server" Visible="true">
</asp:Panel>
</table>
</td>
</tr>

<!-- Result -->
<tr id="tr_result" runat="server" visible="false">
<td colspan='6' align='left' width="100%">
<asp:GridView ID="GV_LotRowData" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3">
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
