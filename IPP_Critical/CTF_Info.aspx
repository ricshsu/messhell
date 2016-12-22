<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CTF_Info.aspx.vb" Inherits="CTF_Info" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>CTF Tool Monitor Report</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<style type="text/css">
        #resulTb
        {
            width: 289px;
        }
        .style12
        {
            width: 100px;
            height: 25px;
        }
        .style13
        {
            width: 125px;
            height: 25px;
        }
    </style>
<script type="text/javascript" language="javascript">
    
    function openWin(part_id, mItem, tool, valueType) {
        var linkStr = "CTF_ChartDetail.aspx?PART=" + part_id + "&ITEM=" + mItem + "&TOOL=" + tool + "&TYPE=" + valueType;
        window.open(linkStr, "newwindow", "fullscreen=yes, scrollbars=yes, resizable=yes");
    }

    function openTransfer() {
        window.open("CTF_Transfer.aspx", "newwindow", "height=520,width=650,top=150,left=200, scrollbars=yes, resizable=yes");
    }

    function openWindowWithPost(part_id, mItem, tool, valueType) {

        var newWindow = window.open("CTF_ChartDetail.aspx", "ChartDetail", 'height=700,width=850,top=100,left=200,toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no');
        var html = "<html><head></head><body><form id='formid' method='post' action='CTF_ChartDetail.aspx'>";
        html += "<input type='hidden' name='PART' value='" + part_id + "'/>";
        html += "<input type='hidden' name='ITEM' value='" + mItem + "'/>";
        html += "<input type='hidden' name='TOOL' value='" + tool + "'/>";
        html += "<input type='hidden' name='TYPE' value='" + valueType + "'/>";
        html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
        html += "script></body></html>";
        newWindow.document.write(html);

    }


    function SearchList() {

        var tb = document.getElementById('<%= txb_lotinput.ClientID %>');
        var l = document.getElementById('<%= lb_lotSource.ClientID %>');

        if (tb.value == "") {
            l.selectedIndex = -1;
        }
        else {
            for (var i = 0; i < l.options.length; i++) {
                if (l.options[i].value.toLowerCase().match(tb.value.toLowerCase())) {
                    l.selectedIndex = i;
                    l.options[i].selected = true;
                    return false;
                }
                else {
                    l.selectedIndex = -1;
                }
            }
        }

    }

    function ListSwapNode(tNowIndex, tIndex, tList) {

        var tSelectedIndex = tNowIndex;
        var tListOptions = tList.options;

        //先判斷要做swap的item存不存在
        if (tSelectedIndex + tIndex >= 0 && tListOptions[tSelectedIndex + tIndex] && tListOptions[tSelectedIndex]) {
            //進行swapNode
            //tListOptions[tSelectedIndex + tIndex].swapNode(tListOptions[tSelectedIndex]); For IE
            swapNodes(tListOptions[tSelectedIndex + tIndex], tListOptions[tSelectedIndex]); //FOR Chrome, FireFox
        }
    }

    function Inquery() 
    {
        var l_source = document.getElementById('<%= lb_lotSource.ClientID %>');
        var l_show = document.getElementById('<%= lb_lotShow.ClientID %>');

        if (l_show.options.length == 0) 
        {
          for (var i = 0;i<(l_source.options.length-1);i++)
          {
            var opt = document.createElement("option");
            l_show.options.add(opt);
            opt.text = l_source.options[i].text;
            opt.value = l_source.options[i].value;
            l_source.remove(i);
          }

        for (var i = (l_source.options.length - 1); i >= 0; i--) 
          {
              l_source.options[i] = null;
          }
          l_source.selectedIndex = -1;
        }
    }

</script>
</head>
<body MS_POSITIONING="FlowLayout">
<form id="form1" runat="server">

<table border="1" width="900px">
   
    <tr>
    <td colspan=6 class="Table_One_Title" valign=middle align="center" style='font-size:large;font-weight: bold'>
    CTF&nbsp;Tool&nbsp;Monitor Data
    </td>
    </tr>

    <tr>
    <td align=left style='background:#E9E7E2;width:150px'>
        料號
    </td>
    <td style='width:150px'>
    <asp:DropDownList ID="ddlPart" runat="server" Width="148px" AutoPostBack="True"></asp:DropDownList>
    </td>

    <td align=left style='background:#E9E7E2;width:150px'>量測項目</td>
    <td style='width:150px'>    
    <asp:DropDownList ID="ddlMeasItem" runat="server" Width="148px" AutoPostBack="True"></asp:DropDownList>
    </td>

    <td align=left style='background:#E9E7E2;width:150px'>機台</td>
    <td style='width:150px'>    
    <asp:DropDownList ID="ddlMachineID" runat="server" Width="148px" AutoPostBack="True"></asp:DropDownList>
    </td>
    </tr>

    <tr>
    
    <td align=left style='background:#E9E7E2;width:150px'>LOT</td>
    
    <td colspan=5>      
    <table>
    <tr>
    <td colspan=3>
    <asp:TextBox ID="txb_lotinput" runat="server" Width="145px" onkeyup="return SearchList();" AutoCompleteType="Disabled"></asp:TextBox>
    </td>
    </tr>

    <tr>

    <td>
    <asp:ListBox ID="lb_lotSource" runat="server" Height="120px" Width="145px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    <td style='width:20px'>
    <asp:Button ID="but_lotTo" runat="server" Text=">>" CssClass=BT_2 />
    <br/>
    <asp:Button ID="but_lotBack" runat="server" Text="<<" CssClass=BT_2/>
    </td>
    
    <td>      
    <asp:ListBox ID="lb_lotShow" runat="server" Height="120px" Width="145px" SelectionMode="Multiple"></asp:ListBox>
    </td>

    </tr>

    </table>
    </td>

    </tr>

    <tr>
    <td colspan=6 align=right>
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>
    <asp:CheckBox ID="cb_showchart" runat="server" Text="Show Chart" />
    &nbsp;&nbsp;
    <asp:Button ID="but_parse" runat="server" Text="資料轉置" Height="30px" Width="86px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    &nbsp;&nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="執行查詢" Height=30px Width="86px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
    </td>
    </tr>

    <tr id='tr_chart' runat=server visible=false>
    <td colspan=6 align=left>
    <table>
    <asp:Panel ID="Chart_Panel" runat="server"></asp:Panel>
    </table>
    </td>
    </tr>

</table>   

<table border="1" width="900px">
<!-- Grid View Start -->
<tr id='tr_rowData' runat=server visible=false>
<td style="height:1px;">
<asp:GridView ID="gv_data" runat="server" AutoGenerateColumns="False" BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3" Width="1040px">
<AlternatingRowStyle BackColor="#F7F6F3" ForeColor="#333333"/>
<Columns>
            <asp:BoundField DataField="Part_ID" HeaderText="Part"/>
            <asp:BoundField DataField="Lot_ID" HeaderText="Lot" />
            <asp:BoundField DataField="Machine_ID" HeaderText="Machine" />
            <asp:BoundField DataField="Meas_Item" HeaderText="Meas_Item" />
            <asp:BoundField DataField="CSL" HeaderText="CSL" />
            <asp:BoundField DataField="Max_value" HeaderText="Max" />
            <asp:BoundField DataField="Min_Value" HeaderText="Min">
            </asp:BoundField>
            <asp:BoundField DataField="Mean_Value" HeaderText="Mean" />
            <asp:BoundField DataField="Std_Value" HeaderText="Std" />
            <asp:BoundField HeaderText="CP" 
                DataField="CP" />
            <asp:BoundField DataField="CPK" HeaderText="CPK" />
            <asp:BoundField DataField="Lot_Meas_Start_DataTime" HeaderText="StartTime" />
            <asp:BoundField DataField="Lot_Meas_End_DataTime" HeaderText="EndTime" />
            <asp:BoundField HeaderText="SPC" />
        </Columns>
<EditRowStyle BackColor="#999999" />
<FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
<HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" HorizontalAlign="Center" />
<PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
<RowStyle ForeColor="#333333" />
<SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
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
<!-- Grid View E n d -->
</table>

<asp:HiddenField ID="txb_chart_value" runat="server" />
</form>
</body>
</html>

