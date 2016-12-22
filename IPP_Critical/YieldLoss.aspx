<%@ Page Language="VB" AutoEventWireup="false" MaintainScrollPositionOnPostback="true" CodeFile="YieldLoss.aspx.vb" Inherits="YieldLoss_Test" %>
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

  function openWindowWithPost(url, name, vala, valb, valc, vald, vale, valf, valCustomer, TYPE, LotList, TopN, IsXoutScrap, BumpingType, LotMerge, TimePeriod)   
  {  
    var newWindow = window.open(url,  name, 'height=700,width=1040,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
    var html = "<html><head></head><body><form id='form1' method='post' action='" + url + "'>";
    html += "<input type='hidden' name='P' value='" + vala + "'/>";
    html += "<input type='hidden' name='F' value='" + valb + "'/>";
    html += "<input type='hidden' name='W' value='" + valc + "'/>";
    html += "<input type='hidden' name='WI' value='" + vald + "'/>";
    html += "<input type='hidden' name='Product' value='" + vale + "'/>";
    html += "<input type='hidden' name='Plant' value='" + valf + "'/>";
    html += "<input type='hidden' name='Customer' value='" + valCustomer + "'/>";
    html += "<input type='hidden' name='TYPE' value='" + TYPE + "'/>";
    html += "<input type='hidden' name='LotList' value='" + LotList + "'/>";
    html += "<input type='hidden' name='TopN' value='" + TopN + "'/>";
    html += "<input type='hidden' name='IsXoutScrap' value='" + IsXoutScrap + "'/>";
    html += "<input type='hidden' name='BumpingType' value='" + BumpingType + "'/>";
    html += "<input type='hidden' name='LotMerge' value='" + LotMerge + "'/>";
    html += "<input type='hidden' name='TimePeriod' value='" + TimePeriod + "'/>";
    html += "</form><script type='text/javascript'>document.getElementById(\"form1\").submit();</";
    html += "script></body></html>";
    newWindow.document.write(html);
  }

  function openWindowWithPostDaily(url, name, vala, valb, valc, vald, vale, valf, valg, lotList, TopN, IsXoutScrap, BumpingType, LotMerge, TimePeriod) 
  {
      var newWindow = window.open(url, name, 'height=700,width=1200,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
      var html = "<html><head></head><body><form id='form1' method='post' action='" + url + "'>";
      html += "<input type='hidden' name='C' value='" + vala + "'/>";
      html += "<input type='hidden' name='CA' value='" + valb + "'/>";
      html += "<input type='hidden' name='P' value='" + valc + "'/>";
      html += "<input type='hidden' name='F' value='" + vald + "'/>";
      html += "<input type='hidden' name='D' value='" + vale + "'/>";
      html += "<input type='hidden' name='PLANT' value='" + valf + "'/>";
      html += "<input type='hidden' name='TYPE' value='" + valg + "'/>";
      html += "<input type='hidden' name='LotList' value='" + lotList + "'/>";
      html += "<input type='hidden' name='TopN' value='" + TopN + "'/>";
      html += "<input type='hidden' name='IsXoutScrap' value='" + IsXoutScrap + "'/>";
      html += "<input type='hidden' name='BumpingType' value='" + BumpingType + "'/>";
      html += "<input type='hidden' name='LotMerge' value='" + LotMerge + "'/>";
      html += "<input type='hidden' name='TimePeriod' value='" + TimePeriod + "'/>";
      html += "</form><script type='text/javascript'>document.getElementById(\"form1\").submit();</";
      html += "script></body></html>";
      newWindow.document.write(html);
  }

  function openWindowWithPostMonthly(url, name, vala, valb, valc, vald, vale, valf, valCustomer, TYPE, lotList, TopN, IsXoutScrap, BumpingType, LotMerge, TimePeriod) {
      var newWindow = window.open(url, name, 'height=700,width=1040,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
      var html = "<html><head></head><body><form id='form1' method='post' action='" + url + "'>";
      html += "<input type='hidden' name='P' value='" + vala + "'/>";
      html += "<input type='hidden' name='F' value='" + valb + "'/>";
      html += "<input type='hidden' name='W' value='" + valc + "'/>";
      html += "<input type='hidden' name='WI' value='" + vald + "'/>";
      html += "<input type='hidden' name='Product' value='" + vale + "'/>";
      html += "<input type='hidden' name='Plant' value='" + valf + "'/>";
      html += "<input type='hidden' name='Customer' value='" + valCustomer + "'/>";
      html += "<input type='hidden' name='TYPE' value='" + TYPE + "'/>";
      html += "<input type='hidden' name='LotList' value='" + lotList + "'/>";
      html += "<input type='hidden' name='TopN' value='" + TopN + "'/>";
      html += "<input type='hidden' name='IsXoutScrap' value='" + IsXoutScrap + "'/>";
      html += "<input type='hidden' name='BumpingType' value='" + BumpingType + "'/>";
      html += "<input type='hidden' name='LotMerge' value='" + LotMerge + "'/>";
      html += "<input type='hidden' name='TimePeriod' value='" + TimePeriod + "'/>";
      html += "</form><script type='text/javascript'>document.getElementById(\"form1\").submit();</";
      html += "script></body></html>";
      newWindow.document.write(html);
  }
</script>
    <style type="text/css">
        .style1
        {
            width: 200px;
            height: 38px;
        }
        .style2
        {
            width: 880px;
            height: 38px;
        }
    </style>
</head>
<body>
<form id="form1" runat="server">

<table border="1" width="1080px">
    
    <!-- Title -->
    <tr>
    <td colspan='2' class="Table_One_Title" valign='middle' align="center" style='font-size:x-large;font-weight: bold'>
        Yield&nbsp; Loss&nbsp; Report
    </td>
    </tr>

    <!-- Custormer -->
    <tr>
    <td align='left' style='background:#E9E7E2; width:200px'>Product Type&nbsp;</td>
    <td style='width:880px'>
    <asp:DropDownList ID="ddlProduct" runat="server" Width="200px" Height="25px" AutoPostBack="True"></asp:DropDownList>
    </td>
    </tr>

    <!-- WB BumpingType-->
    <tr ID="tr_BumpingType" runat="server">
    <td nowrap='nowrap' align='left' style='background:#E9E7E2;width:200px; height:20px;'>
        Bumping Type
    </td>
    <td style='width:800px; height:20px;'>
    <table>
    <tr>
        <%--<td>
    <asp:DropDownList ID="ddl_BumpingType" runat="server" Width="170px" Height="25px" 
            AutoPostBack="True"></asp:DropDownList>
    </td>--%>
    <td style='width:120px'>       
    <asp:ListBox ID="listB_BumpingTypeSource" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="but_BumpingTypeRight" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
            Width="30px" />
    <br />
    <asp:Button ID="but_BumpingTypeLeft" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
            Width="30px"/>
    </td>
    <td>      
    <asp:ListBox ID="listB_BumpingTypeShow" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td nowrap='nowrap'>
        &nbsp;</td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- OL Process-->
    <tr ID="tr_OLProcess" runat="server">
    <td nowrap='nowrap' align='left' style='background:#E9E7E2;width:200px; height:20px;'>
        OL_Process
    </td>
    <td style='width:800px; height:20px;'>
    <table>
    <tr>
        <%--<td>
    <asp:DropDownList ID="ddl_BumpingType" runat="server" Width="170px" Height="25px" 
            AutoPostBack="True"></asp:DropDownList>
    </td>--%>
    <td style='width:120px'>       
    <asp:ListBox ID="listB_OLProcessSource" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="but_OLProcessRight" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
            Width="30px" />
    <br />
    <asp:Button ID="Button3" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
            Width="30px"/>
    </td>
    <td>      
    <asp:ListBox ID="listB_OLProcessShow" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td nowrap='nowrap'>
        &nbsp;</td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Backend-->
    <tr ID="tr_Backend" runat="server">
    <td nowrap='nowrap' align='left' style='background:#E9E7E2;width:200px; height:20px;'>
        Backend
    </td>
    <td style='width:800px; height:20px;'>
    <table>
    <tr>
        <%--<td>
    <asp:DropDownList ID="ddl_BumpingType" runat="server" Width="170px" Height="25px" 
            AutoPostBack="True"></asp:DropDownList>
    </td>--%>
    <td style='width:120px'>       
    <asp:ListBox ID="listB_BackendSource" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="Button4" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
            Width="30px" />
    <br />
    <asp:Button ID="Button5" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
            Width="30px"/>
    </td>
    <td>      
    <asp:ListBox ID="listB_BackendShow" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td nowrap='nowrap'>
        &nbsp;</td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Customer_id-->
    <tr ID="tr1" runat="server">
    <td nowrap='nowrap' align='left' style='background:#E9E7E2;width:200px; height:20px;'>
        Customer_id
    </td>
    <td style='width:800px; height:20px;'>
    <table>
    <tr>
        <%--<td>
    <asp:DropDownList ID="ddl_BumpingType" runat="server" Width="170px" Height="25px" 
            AutoPostBack="True"></asp:DropDownList>
    </td>--%>
    <td style='width:120px'>       
    <asp:ListBox ID="listB_CustomerSource" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="Button2" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
            Width="30px" />
    <br />
    <asp:Button ID="Button6" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
            Width="30px"/>
    </td>
    <td>      
    <asp:ListBox ID="listB_CustomerTarget" runat="server" Height="90px" Width="172px" 
            SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td nowrap='nowrap'>
        &nbsp;</td>
    </tr>
    </table>
    </td>
    </tr>
    <!-- Vendor -->
  <%--  <tr>
    <td align='left' style='background:#E9E7E2; width:200px'>Customer1</td>
    <td style='width:880px'>
        &nbsp;</td>
    </tr>
--%>
    

    <!-- Part -->
    <tr>
    <td align=left style='background:#E9E7E2; ' class="style1">
        Product No / Part No</td>
    <td class="style2">
    <table>
    <tr>
    <td>
    <asp:RadioButtonList ID="rb_ProductPart" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" Width="180px">
<asp:ListItem Value="0" Selected="True">Product No</asp:ListItem>
<asp:ListItem Value="1">Part No</asp:ListItem>
</asp:RadioButtonList>
    <asp:Label ID="Label2" runat="server" Text="Keyword"></asp:Label>
    <asp:TextBox ID="TextBox1" runat="server" Width="101px"></asp:TextBox>
    <asp:Button ID="Button1" runat="server" Text="Filter" />
</td>
        <%--    <td>
    <asp:DropDownList ID="ddlPart" runat="server" Width="200px" Height="25px" AutoPostBack="True"></asp:DropDownList>
    </td>--%>
    <td style='width:120px'>       
<asp:ListBox ID="listB_PartSource" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple"></asp:ListBox>
</td>
<td style='width:20px'>
<asp:Button ID="but_PartRight" runat="server" Text=">>" CssClass=BT_2 Height="25px" 
        Width="30px" />
<br />
<asp:Button ID="but_PartLeft" runat="server" Text="<<" CssClass=BT_2 Height="25px" 
        Width="30px"/>
</td>
<td>      
<asp:ListBox ID="listB_PartShow" runat="server" Height="90px" Width="172px" 
        SelectionMode="Multiple"></asp:ListBox>
    <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
</td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Daily OR Weekly -->
    <tr>
    <td align='left' style='background:#E9E7E2; width:200px'>Time Period</td>
    <td style='width:880px'>
    <table>
    <tr>
    <td>
    <asp:RadioButtonList ID="rb_dayType" runat="server" RepeatDirection="Horizontal" Width="226px" AutoPostBack="True">
    <asp:ListItem Selected="True">Daily</asp:ListItem>
    <asp:ListItem>Weekly</asp:ListItem>
        <asp:ListItem>Monthly</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Default Day or Week Type -->
    <tr>
    <td align='left' style='background:#E9E7E2; width:200px'>
        Time Range
     </td>
    <td style='width:880px'>
    <table>
    <tr>
    <td>
    <asp:RadioButtonList ID="rbl_week" runat="server" RepeatDirection="Horizontal" Width="200px" AutoPostBack="True">
    <asp:ListItem Selected="True">Default</asp:ListItem>
    <asp:ListItem>Custom</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Custom Day or Week Type or Month Type -->
    <tr id='tr_week' runat='server' visible='false'>
    <td align=left style='background:#E9E7E2; width:200px'>Specified Time Range</td>
    <td style='width:880px'>
    <table>
    <tr>
    <td style='width:150px'>       
    <asp:ListBox ID="lb_weekSource" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>
    <td style='width:20px'>
    <asp:Button ID="but_weekTo" runat="server" Text=">>" CssClass=BT_2 Height="30px" Width="25px" />
    <br />
    <asp:Button ID="but_weekBack" runat="server" Text="<<" CssClass=BT_2 Height="30px" Width="25px"/>
    </td>
    <td>      
    <asp:ListBox ID="lb_weekShow" runat="server" Height="118px" Width="150px" SelectionMode="Multiple"></asp:ListBox>
    </td>
    </tr> 
    </table> 
    </td>
    </tr>

    <!-- Default Yield Loss Item -->
    <tr>
    <td align=left style='background:#E9E7E2; width:200px'>
        Yield Loss Item Rank
    </td>
    <td style='width:880px'>
    <table>
    <tr>
    <td>
    <asp:RadioButtonList ID="rbl_lossItem" runat="server" RepeatDirection="Horizontal" Width="310px" AutoPostBack="True">
    <asp:ListItem Selected="True">Top10</asp:ListItem>
    <asp:ListItem>Top20</asp:ListItem>
        <asp:ListItem>Top30</asp:ListItem>
        <asp:ListItem>Top40</asp:ListItem>
        <asp:ListItem>Top100</asp:ListItem>
    <asp:ListItem>Custom</asp:ListItem>
    </asp:RadioButtonList>
    </td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Custom Yield Loss Item -->
    <tr id='tr_lossItem' runat=server visible=false>
    <td align=left style='background:#E9E7E2; width:200px'>Specified Yield Loss Item 
        Rank</td>
    <td style='width:880px'>
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
    
    <tr id="tr_upload" runat=server visible=false>
    <td align='left' style='background:#E9E7E2; width:200px'>
    Upload Lot
    </td>
    <td style='width:880px'>
    <asp:FileUpload ID="uf_UfilePath" runat="server" ></asp:FileUpload>
    &nbsp;
    <asp:Button ID="but_Uupload" runat="server" Text="Upload" />
    </td>
    </tr>

    <!-- Query -->
    <tr>
    <td colspan='2' align='right' width="100%">
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label> 
    &nbsp;<asp:CheckBox ID="ckFAI" runat="server" Text="試製" Checked="True" 
            Visible="False" />
                <asp:CheckBox ID="cb_Lot_Merge" runat="server" Text="合批" 
            Visible="False" />
        &nbsp;<asp:CheckBox ID="cb_DRowData1" runat="server" Text="長條圖數值顯示" />
        <asp:CheckBox ID="cb_DRowData0" runat="server" Text="匹配報廢回歸" />
    <asp:CheckBox ID="cb_DRowData" runat="server" Text="Display RowData" Checked="True" Visible="False" />
    <asp:CheckBox ID="cb_Non8K" runat="server" Text="Non 8K" Checked="True" />
    <asp:CheckBox ID="cb_NonIPQC" runat="server" Text="Non IPQC" />
    <asp:CheckBox ID="Cb_SF" runat="server" Text="SF" />
    <asp:CheckBox ID="Cb_CR" runat="server" Text="客退" />
    <asp:CheckBox ID="Cb_Inline" runat="server" Text="Inline" />
    <asp:CheckBox ID="cb_uploadLot" runat="server" Text="Upload Lot" Checked="false" 
            Visible="true" AutoPostBack="True" />
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Excel" runat="server" Text="Export Excel" Height=30px 
            Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    &nbsp;&nbsp;&nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Query" Height=30px 
            Width="110px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
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
<AlternatingRowStyle BackColor="White" />
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

<!-- LOT GridView Start -->
<table id='tb_uploadLot' runat=server visible=false>
<!-- Lot Chart -->
<tr>
<td class="style12">
<asp:Panel ID="UploadLot_Chart" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
</td>
</tr>
<!-- Lot RowData -->
<tr>
<td class="style12">
<asp:GridView ID="Lot_GridView" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3">
<AlternatingRowStyle BackColor="White" />
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
<!-- LOT GridView E n d -->

</form>
</body>
</html>