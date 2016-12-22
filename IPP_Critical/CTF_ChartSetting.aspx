<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CTF_ChartSetting.aspx.vb" Inherits="CTF_ChartSetting" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Chart Setting</title>
    <link href="~/css/Button.css" rel="Stylesheet" />
    <link href="~/css/PUPPY.css" rel="Stylesheet" />
    <link href="~/css/TabContainer.css" rel="Stylesheet" />
    <link href="~/css/DW.css" rel="Stylesheet" />
    <style type="text/css">
        .style1
        {
            width: 128px;
        }
        .login
        {
            height: 179px;
            width: 380px;
        }
        .style2
        {
            height: 30px;
        }
        .style3
        {
            width: 128px;
            height: 25px;
        }
        .style4
        {
            height: 25px;
        }
    </style>
</head>
<body>
    
    <form id="form1" runat="server">
    <center>
    <div class="accountInfo">
    <fieldset class="login"><legend>Chart Setting</legend>
    <BR />
    <table border="1" style="height: 107px">                
    
    <tr>
    <td class="style1" align=left style='background:#E9E7E2'>
    Chart Type
    </td>
    
    <td class="style10" align=left>
        <asp:RadioButtonList ID="rb_chartType" runat="server" RepeatDirection="Horizontal">
            <asp:ListItem Selected="True" Value="0">Trend</asp:ListItem>
            <asp:ListItem Value="1">Bar</asp:ListItem>
        </asp:RadioButtonList>
    </td>
    </tr>


    <tr>
    <td class="style1" align=left style='background:#E9E7E2'>
        Chart Value 
    </td>
    <td align=left>
    <asp:RadioButtonList ID="rdb_chartValue" runat="server" RepeatDirection="Horizontal">
            <asp:ListItem Selected="True" Value="0">Mean</asp:ListItem>
            <asp:ListItem Value="1">Std</asp:ListItem>
            <asp:ListItem Value="2">Both</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>

    <tr>
    <td colspan=2 class="style2" align=right>
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>&nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_reCaculate" runat="server" Text="確定" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>          
    </td>
    </tr>

    </table>
  
    </fieldset>
    </div>
    </center>
    </form>
    
</body>
</html>
