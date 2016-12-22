<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Critical_KPP.aspx.vb" Inherits="IPP_Critical" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Critical Item</title>
    <link href="css/Table.css" rel="Stylesheet" />
    <link href="css/Button.css" rel="Stylesheet" />
    <link href="css/PUPPY.css" rel="Stylesheet" />
    <link href="css/TabContainer.css" rel="Stylesheet" />
    <link href="css/DW.css" rel="Stylesheet" />
    <script language=javascript>
        function openWin(dfrom, dto, datasource, partid, yimpact, kmodule, critical, edaitem, isHL) 
        {
          var linkStr = "Critical_KPPDetails.aspx?DF=" + dfrom + "&DT=" + dto + "&DS=" + datasource + "&DP=" + partid + "&YI=" + yimpact + "&KM=" + kmodule + "&CI=" + critical + "&EI=" + edaitem + "&HL=" + isHL ;
          window.open(linkStr, "newwindow", "height=800, width=1440, top=100, left=200, toolbar =no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no");
        }
    </script>
</head>
<body>
<form id="form1" runat="server">

<table border="1" width="1100px">
    
    <tr>
    <td colspan=4 class="Table_One_Title" valign=middle align="center" style='font-size:x-large;font-weight: bold'>
        Critical&nbsp; KPP&nbsp; Monitor
        (By Process Time)</td>
    </tr>

    <tr>

    <td align=left style='width:200px;background:#E9E7E2;'>
    Data Source
    </td>
    
    <td style='width:300px;'>
    <asp:DropDownList ID="ddlDataSource" runat="server" Width="120px" 
            AutoPostBack="True"></asp:DropDownList>
    </td>
    
    <td align=left style='width:200px;background:#E9E7E2;'>
    Part No. (料號)
    </td>
    
    <td style='width:300px;'>    
    <asp:DropDownList ID="ddlPartNo" runat="server" Width="120px"></asp:DropDownList>
    </td>

    </tr>

    <tr>

    <td align=left style='width:200px;background:#E9E7E2;'>
    Yield Impact
    </td>
    
    <td style='width:300px;'>    
    <asp:DropDownList ID="ddlYImpact" runat="server" Width="120px"></asp:DropDownList>
    </td>

    <td align=left style='width:200px;background:#E9E7E2;'>
    Key Module
    </td>
    
    <td style='width:300px;'>    
    <asp:DropDownList ID="ddlKModule" runat="server" Width="120px"></asp:DropDownList>
    </td>

    </tr>

    <tr>

    <td align=left style='width:200px;background:#E9E7E2;'>
    Critical Item
    </td>
    
    <td style='width:300px;'>    
    <asp:DropDownList ID="ddlCItem" runat="server" Width="120px"></asp:DropDownList>
    </td>

    <td align=left style='width:200px;background:#E9E7E2;'>
    Time Query
    </td>

    <td style='width:300px;'>
    <asp:TextBox ID="txtDateFrom" runat="server" Columns="10" MaxLength="10" Width="110px"></asp:TextBox>
    <obout:Calendar ID="Calendar1" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateFrom"></obout:Calendar>
    &nbsp;~&nbsp;
    <asp:TextBox ID="txtDateTo" runat="server" Columns="10" MaxLength="10" Width="110px"></asp:TextBox>
    <obout:Calendar ID="Calendar2" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateTo">
    </obout:Calendar>     
    </td>

    </tr>
  
    <tr>

    <td colspan=4 align=right class="style16">
        &nbsp;&nbsp;
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#0066FF" 
            Font-Size="Small"></asp:Label> 
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;<asp:CheckBox ID="cb_DailyEvent" runat="server" Text="DailyEvent (以區間結束日期為 Event Day)" 
            AutoPostBack="True" Font-Bold="True" Font-Size="Small" 
            ForeColor="#990000" />
        &nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Inquiry" Height=30px Width="86px" 
            CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    </td>

    </tr>    

<!-- Grid View Start -->
<tr id="tr_chartPanel" runat="server" visible="false">
<td style="height: 1px;" colspan=4>
<table>
<asp:Panel ID="Panel1" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
</table>
</td>
</tr>

<!-- Grid View E n d -->
</table>

<table style="width: 1172px">
<tr align="right">
<td>
<asp:Label runat="server" ID="lab_userid" style='font-size:small;'></asp:Label>
</td>
</tr>
</table>

<asp:Label ID="lab_Date" runat="server" Text="Label" Visible="False"></asp:Label>
</form>
</body>
</html>
