<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CTF_Detail.aspx.vb" Inherits="CTF_Detail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>SPC Define</title>
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
            height: 320px;
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
    <fieldset class="login"><legend>規格定義</legend>
    <BR />
    <table border="1" style="height: 107px">                
    <tr>
    <td class="style1" align=left style='background:#E9E7E2'>
        Part ID :
    </td>
    
    <td class="style10">
    <asp:TextBox ID="tb_part" runat="server" Enabled="False"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td class="style3" align=left style='background:#E9E7E2'>
        Meas_Item :
    </td>
    <td class="style4">
    <asp:TextBox ID="tb_measItem" runat="server" Enabled="False"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td class="style3" align=left style='background:#E9E7E2'>
    Original USL :
    </td>
    <td class="style4">
    <asp:TextBox ID="tx_OUSL" runat="server" Enabled="False"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td class="style3" align=left style='background:#E9E7E2'>
    Original CL :
    </td>
    <td class="style4">
    <asp:TextBox ID="tx_OCL" runat="server" Enabled="False"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td class="style3" align=left style='background:#E9E7E2'>
    Original LSL :
    </td>
    <td class="style4">
    <asp:TextBox ID="tx_OLSL" runat="server" Enabled="False"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td colspan=2>
        <asp:RadioButtonList ID="RBL_SPCTYPE" runat="server" 
            RepeatDirection="Horizontal" AutoPostBack="True">
            <asp:ListItem Selected="True" Value="2">雙   邊</asp:ListItem>
            <asp:ListItem Value="1_U">單邊-上界</asp:ListItem>
            <asp:ListItem Value="1_D">單邊-下界</asp:ListItem>
        </asp:RadioButtonList>
    </td>
    </tr>

    <!-- 上界 -->
    <tr id=tr_usl runat=server>
    <td class="style3" align=left style='background:#E9E7E2'>
    USL :
    </td>
    <td class="style4">
    <asp:TextBox ID="tb_USL" runat="server"></asp:TextBox>
    </td>
    </tr>

    <!-- 下界 -->
    <tr id=tr_lsl runat=server>
    <td class="style1" align=left style='background:#E9E7E2'>
    LSL :
    </td>
    <td class="style10">
    <asp:TextBox ID="tb_LSL" runat="server"></asp:TextBox>
    </td>
    </tr>


    <tr>
    <td colspan=2 class="style2" align=right>
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>&nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_reCaculate" runat="server" Text="更新" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>          
    </td>
    </tr>

    </table>
  
    </fieldset>
    </div>
    </center>
    </form>
    
</body>
</html>
