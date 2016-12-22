<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DailyReport.aspx.vb" Inherits="DailyReport" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Daily Report</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<script type="text/javascript" language='javascript'>
    function openWindowWithPost(url, name, customerID, productType, vala, valb, valc, vald, dataType) {

        var newWindow = window.open(url, name, 'height=700, width=850, top=100, left=200,toolbar=yes,menubar=yes,scrollbars=yes,resizable=yes,location=yes,status=yes');
        var html = "";
        html += "<html><head></head><body><form id='formid' method='post' action='" + url + "'>";
        html += "<input type='hidden' name='customerID' value='" + customerID + "'/>";
        html += "<input type='hidden' name='productType' value='" + productType + "'/>";
        html += "<input type='hidden' name='partID' value='" + vala + "'/>";
        html += "<input type='hidden' name='YieldType' value='" + valb + "'/>";
        html += "<input type='hidden' name='RangeTime' value='" + valc + "'/>";
        html += "<input type='hidden' name='RangeType' value='" + vald + "'/>";
        html += "<input type='hidden' name='DataType' value='" + dataType + "'/>";
        html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
        html += "script></body></html>";
        newWindow.document.write(html);
    }
</script>
    </head>

<body>
<form id="form1" runat="server">
<!-- Condition Select -->
<table border="1" width="1080px"> 
<!-- Title -->
<tr>
<td colspan='2' class="Table_One_Title" valign=middle align="center" style='font-size:x-large;font-weight: bold; height:20px'>
Daily&nbsp;&nbsp;&nbsp;Yield&nbsp;&nbsp;&nbsp;Report
</td>
</tr>

<!-- Product Type -->
<tr>
<td nowrap='nowrap' align='left' style='background:#E9E7E2; width:200px; height:20px;'>
Product Type
</td>
<td>
<table>
<tr>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<asp:DropDownList ID="ddl_ProductType" runat="server" Width="170px" Height="25px" 
        AutoPostBack="True"></asp:DropDownList>
</td>
</tr>
</table>
</td>
</tr>


<!-- 廠商 Customer-->
<tr>
<td nowrap='nowrap' align='left' style='background:#E9E7E2;width:200px; height:20px;'>
Customer
</td>
<td style='width:800px; height:20px;'>
<table>
<tr>
<td>
<asp:DropDownList ID="ddl_CustomerID" runat="server" Width="170px" Height="25px" 
        AutoPostBack="True"></asp:DropDownList>
</td>

<td nowrap='nowrap'>
    &nbsp;</td>
</tr>
</table>
</td>
</tr>


<!-- Yield Mode -->
<tr>
<td nowrap='nowrap' align='left' style='background:#E9E7E2; width:200px; height:20px;'>
Yield Mode
</td>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<table>
<tr>
<td>
<asp:DropDownList ID="ddl_YieldMode" runat="server" Width="170px" Height="25px" AutoPostBack="True"></asp:DropDownList>
</td>
</tr>
</table>
</td>
</tr>

<!-- Product No / Part No -->
<tr>
<td nowrap='nowrap' align='left' style='background:#E9E7E2; width:200px; height:20px;'>
Product No / Part No
</td>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<table>
<tr>
<td>
<asp:RadioButtonList ID="rbl_BySource" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" Width="180px">
<asp:ListItem Value="0" Selected="True">Product No</asp:ListItem>
<asp:ListItem Value="1">Part No</asp:ListItem>
</asp:RadioButtonList>
</td>
<td>
<asp:DropDownList ID="ddlPart" runat="server" Width="200px" Height="20px"></asp:DropDownList>
</td>
</tr>
</table>
</td>
</tr>

<!-- Time Range -->
<tr>
<td nowrap='nowrap' align='left' style='background:#E9E7E2; width:200px; height:20px;'>
Time Range
</td>
<td>
<table>
<tr>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<asp:RadioButtonList ID="rb_DataTimeCustor" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" Width="180px">
<asp:ListItem Value="0" Selected="True">Default</asp:ListItem>
<asp:ListItem Value="1">Custom</asp:ListItem>
</asp:RadioButtonList>
</td>
</tr>
</table>
</td>
</tr>

<!-- Date Range -->
<tr id='tr_dateRange' runat='server' visible='false'>
<td align='left' style='background:#E9E7E2; width:200px; height:20px;'>
Specified Time Range
</td>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<table>
<tr>
<td colspan=3>
<asp:RadioButtonList ID="rbl_lossItem" runat="server" RepeatDirection="Horizontal" Width="318px" AutoPostBack="True">
<asp:ListItem Selected="True">Daily</asp:ListItem>
<asp:ListItem>Weekly</asp:ListItem>
<asp:ListItem>Monthly</asp:ListItem>
</asp:RadioButtonList>
</td>
</tr>
<tr>
<td style='width:120px'>       
<asp:ListBox ID="listB_timeSource" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple"></asp:ListBox>
</td>
<td style='width:20px'>
<asp:Button ID="but_dateRight" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
        Width="30px" />
<br />
<asp:Button ID="but_dateLeft" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
        Width="30px"/>
</td>
<td>      
<asp:ListBox ID="listB_timeShow" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple"></asp:ListBox>
</td>
</tr>  
</table>
</td>
</tr>

<!-- Report Week -->
<tr id='tr_week' runat='server' visible='true'>
<td nowrap='nowrap' align='left' style='background:#E9E7E2; width:200px; height:20px;'>Yield Item</td>
<td nowrap='nowrap' style='width:800px; height:20px;'>
<table>
<tr>
<td style='width:120px'>       
<asp:ListBox ID="lb_weekSource" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple">
        <asp:ListItem>InLine Yield</asp:ListItem>
        <asp:ListItem>O/S Yield</asp:ListItem>
        <asp:ListItem>Bump Yield</asp:ListItem>
        <asp:ListItem>FVI Yield</asp:ListItem>
</asp:ListBox>
</td>
<td style='width:30px'>
    <asp:Button ID="but_weekTo" runat="server" Text=">>" CssClass=BT_2 
        Height="25px" Width="30px" />
    <br />
    <asp:Button ID="but_weekBack" runat="server" Text="<<" CssClass=BT_2 
        Height="25px" Width="30px"/>
</td>
<td>      
<asp:ListBox ID="lb_weekShow" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple"></asp:ListBox>
</td>
</tr>  
</table>
</td>
</tr>

</table>  

<!-- Query -->
<!-- START --> 
<table border="1" width="1080px">
    <tr>
    <td align="right" style='width:1080px'>
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>
    <asp:CheckBox ID="cb_customerDay" runat="server" Text="Customize Time" AutoPostBack="True" Visible="false" />
    <asp:TextBox ID="txtDateFrom" runat="server" Columns="10" MaxLength="10" Width="65px" Enabled=false Visible="false"></asp:TextBox>
    <asp:DropDownList ID="ddlHourFrom" runat="server" Width="50" Enabled='false' Visible="false" ></asp:DropDownList>
    <obout:Calendar ID="Calendar1" runat="server" Enabled=false Columns="1" DateFormat="yyyyMMdd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateFrom" Visible="false" ></obout:Calendar>
    &nbsp;&nbsp;&nbsp;<asp:TextBox ID="txtDateTo" runat="server" Columns="10" MaxLength="10" Width="65px" Enabled=false Visible="false"></asp:TextBox>
    <asp:DropDownList ID="ddlHourTo" runat="server" Width="50" Enabled='false' Visible="false"></asp:DropDownList>   
    <obout:Calendar ID="Calendar2" runat="server" Enabled=false Columns="1" DateFormat="yyyyMMdd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateTo" Visible="false" >
    </obout:Calendar>  
    &nbsp;&nbsp;&nbsp;    
    <asp:CheckBox ID="cb_ShowToday" runat="server" Text="Show Today" AutoPostBack="True" />
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Inquery" Height=30px Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    </td>
    </tr>
    </table>
<!-- E N D-->

<!-- Chart Start -->
<table>
<tr id="tr_chartDisplay" runat='server' visible='false'>
<td class="style12">
<table>
<tr>
<td>
<asp:Panel ID="Chart_Panel" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
</td>
</tr>
</table>
</td>
</tr>
<!-- Chart E n d -->

<!-- GridView Start -->
<tr id="tr_gvDisplay" runat=server visible=false>
<td class="style12">
<asp:GridView ID="gv_rowdata" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3">
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
<!-- GridView E n d -->
</table>

</form>
</body>
</html>


