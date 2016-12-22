<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Critical_Lot.aspx.vb" Inherits="Critical_Lot" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Shipping Judgement KPP Monitor</title>
<!-- Bootstrap -->

    <link href="css/bootstrap.min.css" rel="stylesheet"/>
	<link href="css/daterangepicker-bs3.css" rel="stylesheet"/>
    <link href="css/bootstrap-theme.min.css" rel="stylesheet" type="text/css" />

    <script type="text/javascript" src="js/jsFunction.js"></script>
    <script type="text/javascript" src="js/jquery-1.11.3.min.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
    <script type="text/javascript" src="js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/daterangepicker.js"></script>

    <script type="text/javascript">
    function openWin(PCategory, dfrom, dto, partid, critical, valueType, isHL, HLLot) {
        var linkStr = "Critical_LotDetails.aspx?CTYPE=" + PCategory + "&DF=" + dfrom + "&DT=" + dto + "&PART=" + partid + "&CI=" + critical + "&VT=" + valueType + "&HL=" + isHL + "&HLLOT=" + HLLot; ;
        window.open(linkStr, "newwindow", "height=800, width=1150, top=100, left=200, toolbar=no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no");
    }
    </script>
    <style type="text/css">
    #progressBackgroundLayer
    {
            position: fixed;
            top: 0px;
            bottom: 0px;
            left: 0px;
            right: 0px;
            overflow: hidden;
            padding: 0;
            margin: 0;
            background-color: #000;
            filter: alpha(opacity=30);
            opacity: 0.5;
            z-index: 1000;
    }
    #processMessageLayer
    {
            position: fixed;
            text-align: center;
            width: 15%;
            border: none;
            padding: 5px;
            background-color: #fff;
            vertical-align: middle;
            z-index: 1001;
            top: 40%;
            left: 40%;
    }
    .style1
    {
      height: 74px;
    }
    .style2
    {
      width: 20px;
      height: 74px;
    }
    
    #daterange{
        width:200px;
    }
    .daterangepicker.show-calendar .calendar {
        width: 700px;
    }
    
    /* 調整Table樣式 */
    .grayBG{
        width:120px;
		background: #eee;
		font-weight: bold;
	}
    </style>

    <!--[if IE]>
    <style>
    .daterangepicker.dropdown-menu{
        width: 180px;
    }

    .daterangepicker.dropdown-menu.show-calendar{
        width: 730px ;
    }
    
    .daterangepicker .ranges .input-mini{
        width: 50px;
    }
    </style>
    <![endif]-->
</head>
<body>
<form id="form1" runat="server">
<asp:ScriptManager ID="ScriptManager" runat="server">
</asp:ScriptManager>

<!-- Condition Start -->
<table style="width:1100px" class="table table-condensed">
    
    <!-- Title -->
    <tr>
    <td colspan="4" class="label-primary" valign="middle" align="center" style='font-size:x-large;color:white;font-weight: bold'>
    Shipping Judgement KPP Monitor
    </td>
    </tr>
    <!-- Data Source -->
    <tr> 
    <td class="grayBG">
    BU Category
    </td>

    <td colspan="3">
    <asp:DropDownList ID="ddl_Category" runat="server" Width="200px" AutoPostBack="True">
        <asp:ListItem Value="critical_lot">CPU</asp:ListItem>
        <asp:ListItem>CS</asp:ListItem>
        <asp:ListItem>PPS</asp:ListItem>
        </asp:DropDownList>
    </td>
    </tr>
    <!-- Critical Item -->
    
<tr>

    <td class="grayBG">
    Data Source
    </td>
    
    <td>    
    <asp:DropDownList ID="c" runat="server" Width="200px" AutoPostBack="True"></asp:DropDownList>
    </td>

    <td class="grayBG">
    Critical Item
    </td>
    
    <td>    
    <asp:DropDownList ID="ddlItem" runat="server" Width="200px" AutoPostBack="True"></asp:DropDownList>
    </td>

</tr>

    <!-- Part -->
    <tr>
    <td class="grayBG">Part No (料號)&nbsp;</td>
    <td colspan=3>      
    <table>
    <tr>
    <td colspan=3>
    <asp:TextBox ID="txb_lotinput" runat="server" Width="190px" AutoCompleteType="Disabled"></asp:TextBox>
    </td>
    </tr>

    <tr>
    <td class="style1">
    <asp:ListBox ID="lb_lotSource" runat="server" Height="70px" Width="192px" SelectionMode="Multiple">
    </asp:ListBox>
    </td>

    <td class="style2">
    <asp:Button ID="but_lotTo" runat="server" Text=">>" CssClass="btn btn-xs btn-warning" />
    <br />
    <asp:Button ID="but_lotBack" runat="server" Text="<<" CssClass="btn btn-xs btn-warning"/>
    </td>
    
    <td class="style1">      
    <asp:ListBox ID="lb_lotShow" runat="server" Height="70px" Width="190px" SelectionMode="Multiple"></asp:ListBox>
    </td>
    </tr>
    </table>
    </td>
    </tr>

    <!-- Time -->
    <tr>
    <td class="grayBG">
    Time Query
    </td>

    <td colspan='3' style='width:1000px'>
    <asp:TextBox ID="daterange" runat="server" ></asp:TextBox>
    <asp:HiddenField ID="txtDateFrom" runat="server"></asp:HiddenField>
    <asp:HiddenField ID="txtDateTo" runat="server"></asp:HiddenField>

    <script type="text/javascript">
        $(function () {
            if ($("#txtDateFrom").val() != "") {
                $("#daterange").val($("#txtDateFrom").val() + " - " + $("#txtDateTo").val());
            }
            //$(daterange).val($(txtDateFrom).val() + " - " + $(txtDateTo).val());

            $('#daterange').daterangepicker({
                format: 'YYYY-MM-DD',
                showDropdowns: true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Yesterday': [moment().subtract(1, 'days'), moment().subtract(1, 'days')],
                    'Last 7 Days': [moment().subtract(6, 'days'), moment()],
                    'Last 30 Days': [moment().subtract(29, 'days'), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract(1, 'month').startOf('month'), moment().subtract(1, 'month').endOf('month')]
                }
            },
		    function (start, end, label) {
		        $("#txtDateFrom").val(start.format('YYYY-MM-DD'));
		        $("#txtDateTo").val(end.format('YYYY-MM-DD'));
		        $("#daterange").val($("#txtDateFrom").val() + " - " + $("#txtDateTo").val());
		        //alert("A new date range was chosen: " + start.format('YYYY-MM-DD') + ' to ' + end.format('YYYY-MM-DD'));
		    });

        });
    </script>
    </td>
    
    </tr>
    
    <!-- Excute -->
    <tr>
    <td colspan=4 align=right class="style16">
    <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#0066FF" Font-Size="Small"></asp:Label>
    <asp:CheckBox ID="cb_DailyEvent" runat="server" Text="DailyEvent (以區間結束日期為 Event Day)" AutoPostBack="True" Font-Bold="True" Font-Size="Small" ForeColor="#990000" />
    &nbsp;
    <asp:Button ID="but_Execute" runat="server" Text="Query" CssClass="btn btn-primary" />
    
    </td>
    </tr>
       
</table>
<!-- Condition End   -->

<!-- Chart Area Start -->
<table>
<tr id="tr_chartPanel" runat="server" visible="true">
<td style="height:1px;width:850px;">
<table>
<asp:Panel ID="ChartPanel" runat="server" BorderColor="#99FF33" BackColor="White">

</asp:Panel>
</table>
</td>
</tr>
</table>
<!-- Chart Area E n d -->

<table style="width: 1000px">
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
