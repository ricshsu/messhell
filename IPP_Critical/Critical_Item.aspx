<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Critical_Item.aspx.vb" Inherits="IPP_Critical" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Critical Item</title>

    <link href="css/bootstrap.min.css" rel="stylesheet"/>
	<link href="css/daterangepicker-bs3.css" rel="stylesheet"/>
    <link href="css/bootstrap-theme.min.css" rel="stylesheet" type="text/css" />

    <script type="text/javascript" src="js/jsFunction.js"></script>
    <script type="text/javascript" src="js/jquery-1.11.3.min.js"></script>
	<script type="text/javascript" src="js/moment.min.js"></script>
    <script type="text/javascript" src="js/bootstrap.min.js"></script>
	<script type="text/javascript" src="js/daterangepicker.js"></script>

    <script language=javascript>
    function openWin(Category, dfrom, dto, partid, main_id, sub_id, isHL, showProcess, HLLot) 
    {
       var linkStr = "Critical_ItemDetails.aspx?CTYPE=" + Category + "&DF=" + dfrom + "&DT=" + dto + "&DP=" + partid + "&MAIN_ID=" + main_id + "&SUB_ID=" + sub_id + "&HL=" + isHL + "&SP=" + showProcess + "&HLLOT=" + HLLot;
       window.open(linkStr, "newwindow", "height=800, width=1150, top=100, left=200, toolbar =no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no");
    }
    </script>
    <style>
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
</head>

<body>
<form id="form1" runat="server">
<asp:ScriptManager ID="ScriptManager" runat="server">
</asp:ScriptManager>


<table style="width:1100px" class="table table-condensed">
    
    <tr>
        <td colspan="4" class="label-primary" valign="middle" align="center" style='font-size:x-large;color:white;font-weight: bold'>
        Critical KPP Monitor (By Measure Time)
        </td>
    </tr>

    <tr>
        <td class="grayBG">BU Category</td>
        <td><asp:DropDownList ID="ddlDataSource" runat="server" Width="200px" AutoPostBack="True"></asp:DropDownList></td>
    
        <td class="grayBG">Part No. (料號)</td>
        <td><asp:DropDownList ID="ddlPartNo" runat="server" Width="200px"></asp:DropDownList></td>
    </tr>

    <tr>
        <td class="grayBG">Yield Impact</td>
        <td><asp:DropDownList ID="ddlYImpact" runat="server" Width="200px"></asp:DropDownList></td>

        <td class="grayBG">Key Module</td>
        <td><asp:DropDownList ID="ddlKModule" runat="server" Width="200px"></asp:DropDownList></td>
    </tr>

    <tr>
    <td class="grayBG">Critical Item</td>
    <td colspan="4"><asp:DropDownList ID="ddlCItem" runat="server" Width="200px"></asp:DropDownList></td>
    </tr>
    <tr>
    <td class="grayBG">Time Query</td>

    <td colspan="4">
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
  
    <tr>
        <td colspan="4" align="right">
        <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#0066FF" Font-Size="Small"></asp:Label>　　
        <asp:CheckBox ID="cb_showlabel" runat="server" Text="Show label on the left" AutoPostBack="True" Font-Bold="True" Font-Size="Small" ForeColor="#990000" />　　
        <asp:CheckBox ID="cb_DailyEvent" runat="server" Text="DailyEvent (以區間結束日期為 Event Day)"  AutoPostBack="True" Font-Bold="True" Font-Size="Small"  ForeColor="#990000" />　　
        <asp:Button ID="but_Execute" runat="server" Text="Query" CssClass="btn btn-primary"/>
        </td>
    </tr>    
</table>

<table style="width: 1100px">
<tr id="tr_chartPanel" runat="server" visible="false">
<td style="height: 1px;" colspan=4>
<asp:Panel ID="Panel1" runat="server" BorderColor="#99FF33" BackColor="White">
</asp:Panel>
</td>
</tr>
</table>

<asp:Label runat="server" ID="lab_userid" style='font-size:small;'></asp:Label>
<asp:Label ID="lab_Date" runat="server" Text="Label" Visible="False"></asp:Label>




</form>
</body>

</html>
