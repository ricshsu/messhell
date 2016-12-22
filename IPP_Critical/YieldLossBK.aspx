<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YieldLossBK.aspx.vb" Inherits="YieldLoss" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Yield Loss</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<script type="text/javascript" language='javascript'>
  function SearchList() {
      
      var tb = document.getElementById('<%= txb_ylInput.ClientID %>');
      var l = document.getElementById('<%= lb_LossSource.ClientID %>');

      if (tb.value == "") 
      {
          l.selectedIndex = -1;
      }
      else 
      {
          for (var i = 0; i < l.options.length; i++) 
          {
              if (l.options[i].value.toLowerCase().match(tb.value.toLowerCase())) 
              {
                  l.selectedIndex = i;
                  l.options[i].selected = true;
                  return false;
              }
              else 
              {
                  l.selectedIndex = -1;
              }
          }
      }

  }

  function ListSwapNode(tNowIndex, tIndex, tList) {
     
      var tSelectedIndex = tNowIndex;
      var tListOptions = tList.options;

      //先判斷要做swap的item存不存在
      if (tSelectedIndex + tIndex >= 0 && tListOptions[tSelectedIndex + tIndex] && tListOptions[tSelectedIndex]) 
      {
        //進行swapNode
        //tListOptions[tSelectedIndex + tIndex].swapNode(tListOptions[tSelectedIndex]); For IE
        swapNodes(tListOptions[tSelectedIndex + tIndex], tListOptions[tSelectedIndex]); //FOR Chrome, FireFox
      }
  }

  //進行listitem的排序
  function Sort(pAction) {
      
      var tSortField = document.getElementById('<%= lb_LossShow.ClientID %>');
      var tSelectedIndex = tSortField.selectedIndex;

      //如果沒有選擇任何一個item
      if (tSelectedIndex == -1) {
        alert("請先選擇欲調整的顯示欄位!!");
      }
      else {
          for (var i = 0; i < tSortField.options.length; i++) {
              if (tSortField.options[i].selected == true)
              {
                  if (pAction == 'Up') {
                      //上移
                      ListSwapNode(i, -1, tSortField);
                  }
                  else {
                      //下移
                      ListSwapNode(i, +1, tSortField);
                  }
                  break;
              }
          }
      }
  }

  function swapNodes(item1, item2) {
      var itemtmp = item1.cloneNode(1);
      var parent = item1.parentNode;
      item2 = parent.replaceChild(itemtmp, item2);
      parent.replaceChild(item2, item1);
      parent.replaceChild(item1, itemtmp);
      itemtmp = null;
  }

  function openWindowWithPost(url, name, vala, valb, valc, vald, vale, valf) 
  {
    
    var newWindow = window.open(url,  name, 'height=700,width=850,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
    var html = "<html><head></head><body><form id='formid' method='post' action='" + url + "'>";
    html += "<input type='hidden' name='P' value='" + vala + "'/>";
    html += "<input type='hidden' name='F' value='" + valb + "'/>";
    html += "<input type='hidden' name='W' value='" + valc + "'/>";
    html += "<input type='hidden' name='WI' value='" + vald + "'/>";
    html += "<input type='hidden' name='Product' value='" + vale + "'/>";
    html += "<input type='hidden' name='Plant' value='" + valf + "'/>";
    html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
    html += "script></body></html>";
    newWindow.document.write(html);

  }

</script>
<style type="text/css">
#listbox
{
 overflow: auto;
}
#resulTb
{
 width: 289px;
}
.style11
{
 width: 150px;
}
    .style12
    {
        height: 1px;
    }
</style>
</head>

<body>
<form id="form1" runat="server">

<table border="1" width="1080px">
    
    <!-- Title -->
    <tr>
    <td colspan=2 class="Table_One_Title" valign=middle align="center" style='font-size:x-large;font-weight: bold'>
        Yield&nbsp; Loss&nbsp; Report
    </td>
    </tr>

    <!-- Vendor -->
    <tr>
    <td class="style11" align='left' style='background:#E9E7E2'>
    Vendor (廠商)
    </td>
    <td>
    <asp:DropDownList ID="ddlCustomer" runat="server" Width="150px" AutoPostBack="True" Height="25px"></asp:DropDownList>
    </td>
    </tr>

    <tr>
    <td class="style11" align='left' style='background:#E9E7E2'>
    Product (產品)
    </td>
    <td>
    <asp:DropDownList ID="ddlProduct" runat="server" Width="150px" Height="25px" AutoPostBack="True"></asp:DropDownList>
    </td>
    </tr>

    <!-- Part -->
    <tr>
    <td class="style11" align=left style='background:#E9E7E2'>
    Part (料號)
    </td>
    <td>
    <asp:DropDownList ID="ddlPart" runat="server" Width="150px" Height="25px"></asp:DropDownList>
    
    </td>
    </tr>

    <!-- Report Week -->
    <tr>
    <td class="style11" align=left style='background:#E9E7E2'>
    Report Week
    </td>
    <td>
    <asp:RadioButtonList ID="rbl_week" runat="server" RepeatDirection="Horizontal" Width="200px" AutoPostBack="True">
    <asp:ListItem Selected="True">Default</asp:ListItem>
    <asp:ListItem>Custom</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>

    <!-- Custom Week -->
    <tr id=tr_week runat=server visible=false>
    <td class="style11" align=left style='background:#E9E7E2'>Custom Week</td>
    <td>
    <table>
    <tr>
    <td style='width:150px'>       
    <asp:ListBox ID="lb_weekSource" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="but_weekTo" runat="server" Text=">>" CssClass=BT_2 />
    <br />
    <asp:Button ID="but_weekBack" runat="server" Text="<<" CssClass=BT_2/>
    </td>
    <td>      
    <asp:ListBox ID="lb_weekShow" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>
    
    </tr> 
    </table> 
    </td>
    </tr>

    <!-- Yield Loss Item -->
    <tr>
    <td class="style11" align=left style='background:#E9E7E2'>
    Yield Loss Item
    </td>
    <td>
    <asp:RadioButtonList ID="rbl_lossItem" runat="server" 
            RepeatDirection="Horizontal" Width="200px" AutoPostBack="True">
            <asp:ListItem Selected="True">Top10</asp:ListItem>
            <asp:ListItem>Top20</asp:ListItem>
            <asp:ListItem>Custom</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>
    <tr id='tr_lossItem' runat=server visible=false>
    <td class="style11" align=left style='background:#E9E7E2'>Custom Item</td>
    <td>
    <table>

    <tr>
    <td colspan=3>
    <asp:TextBox ID="txb_ylInput" runat="server" Width="145px" onkeyup="return SearchList();" AutoCompleteType="Disabled"></asp:TextBox>
    </td>
    </tr>
    
    <tr>
    <td style='width:150px'>      
    <asp:ListBox ID="lb_LossSource" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    <td style='width:20px'>
    <asp:Button ID="but_lossItemTo" runat="server" Text=">>" CssClass=BT_2/>
    <br />
    <asp:Button ID="but_lossItemBack" runat="server" Text="<<" CssClass=BT_2/>
    </td>
    
    <td>      
    <asp:ListBox ID="lb_LossShow" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    </tr>

    </table>
    </td>
    </tr>
    
    <!-- Query -->
    <tr>
    <td colspan=2 align=right>
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label> 
    &nbsp;&nbsp;&nbsp;
    <asp:CheckBox ID="cb_DRowData" runat="server" Text="Display RowData" Checked="True" Visible="False" />
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Excel" runat="server" Text="Export Excel" Height=30px Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium" Visible="False"/>
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Inquery" Height=30px Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    </td>
    </tr> 
       
    </table>

<table>
<!-- Chart Start -->
<tr id="tr_chartDisplay" runat=server visible=false>
<td class="style12">
<asp:Panel ID="Chart_Panel" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
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