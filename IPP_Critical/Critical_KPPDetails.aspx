<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Critical_KPPDetails.aspx.cs" Inherits="Critical_KPPDetails" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<%@ Register assembly="Ext.Net" namespace="Ext.Net" tagprefix="ext" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Critical Item Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<link href="/resources/css/examples.css" rel="stylesheet" />
<script src="js/jsdomenu.js" type="text/javascript"></script>
<script src="js/jsdomenu.inc.js" type="text/javascript"></script>
<script language=javascript>

    function LinkPoint(dataID) 
    {
        var linkPoint = document.getElementById(('GridView1_' + dataID));
        linkPoint.style.background = "#FF6A6A";
        document.location = ('#' + dataID);
    }

    function openCommand(url, name, partID, lotID, paramID, userID, e) {

        var xPosition = e.clientX;
        var yPosition = e.clientY;
        var fertrue = "height=630,width=520,top=" + xPosition + ",left=" + yPosition + ",toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no";
        var newWindow = window.open(url, name, fertrue);
        var html = "<html><body><form id='formid' method='post' action='" + url + "'>";
        html += "<input type='hidden' name='partID' value='" + partID + "'/>";
        html += "<input type='hidden' name='lotID' value='" + lotID + "'/>";
        html += "<input type='hidden' name='item' value='" + paramID + "'/>";
        html += "<input type='hidden' name='userID' value='" + userID + "'/>";
        html += "<input type='hidden' name='fType' value='Critical_KPP'/>";
        html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
        html += "script></body></html>";
        newWindow.document.write(html);

    }

    function setPosition(e) {
        var WindowObj = document.getElementById('Window1');
        WindowObj.x = e.clientX;
        WindowObj.y = e.clientY;
    }

    function pageReFresh() {
        document.getElementById('Button1').click();
    }

</script>
</head>
<body style="MARGIN-TOP: 10px; MARGIN-LEFT: 10px" onload="initjsDOMenu();" MS_POSITIONING="FlowLayout">
<form id="form1" runat="server">
<ext:ResourceManager ID="ResourceManager1" runat="server" />
<!-- Ext.NET -->
<ext:Window ID="Window1" runat="server" Closable="false" Resizable="false" Height="150" Width="350" Icon="Lock" Title="Login For Command" Draggable="false" Modal="true"
BodyPadding="5" Layout="FormLayout" X=500 Y=600 >
<Items>
<ext:TextField 
ID="txtUsername" 
runat="server"                     
FieldLabel="Username" 
AllowBlank="false"
BlankText="Your username is required."
Text="" />
<ext:TextField 
ID="txtPassword" 
runat="server" 
InputType="Password" 
FieldLabel="Password" 
AllowBlank="false" 
BlankText="Your password is required."
Text="" />
</Items>
<Buttons>
<ext:Button ID="btnLogin" runat="server" Text="Login" Icon="Accept">
                    <Listeners>
                        <Click Handler="
                            if (!#{txtUsername}.validate() || !#{txtPassword}.validate()) {
                                Ext.Msg.alert('Error','The Username and Password fields are both required');
                                // return false to prevent the btnLogin_Click Ajax Click event from firing.
                                return false; 
                            }" />
                    </Listeners>
                    <DirectEvents>
                        <Click OnEvent="btnLogin_Click">
                            <EventMask ShowMask="true" Msg="Verifying..." MinDelay="500" />
                        </Click>
                    </DirectEvents>
                </ext:Button>
                <ext:Button ID="btnCancel" runat="server" Text="Cancel" Icon="Decline">
                    <Listeners>
                        <Click Handler="#{Window1}.hide();" />
                    </Listeners>
</ext:Button>
</Buttons>
</ext:Window>
<!-- Ext.NET -->
<table>
<!-- Correlation -->
<tr id='tr_correlation' runat='server' visible='false' >
<td colspan=4 align="left" style='font-size:x-large;font-weight: bold;'>
Correlation 
<asp:CheckBox ID="chb_correlation" runat="server" AutoPostBack="True" />
</td>
</tr>
<!-- Correlation Item -->
<tr id='tr_yieldImpact' runat='server' visible='false'>
    <td class="style1" align=left style='background:#E9E7E2'>
    Yield Impact
    </td>
    <td class="style15">    
    <asp:DropDownList ID="ddlYImpact" runat="server" Width="120px"></asp:DropDownList>
    </td>
    <td class="style11" align=left style='background:#E9E7E2'>
    Key Module
    </td>
    <td class="style13">    
    <asp:DropDownList ID="ddlKModule" runat="server" Width="120px"></asp:DropDownList>
    </td>
</tr>
<tr id='tr_critical' runat='server' visible='false'>
    <td class="style11" align=left style='background:#E9E7E2'>
    Critical Item
    </td>
    <td class="style14">    
    <asp:DropDownList ID="ddlCItem" runat="server" Width="120px"></asp:DropDownList>
    </td> 
    <td class="style11" align=left style='background:#E9E7E2'>
    Time Query
    </td>
    <td class="style10">
    <asp:TextBox ID="txtDateFrom" runat="server" Columns="10" MaxLength="10" Width="65px"></asp:TextBox>
    <obout:Calendar ID="Calendar1" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateFrom"></obout:Calendar> &nbsp;~&nbsp; <asp:TextBox ID="txtDateTo" runat="server" Columns="10" MaxLength="10" Width="65px"></asp:TextBox>
    <obout:Calendar ID="Calendar2" runat="server" Columns="1" DateFormat="yyyy-MM-dd" DatePickerImagePath="images/calendar.gif"
    DatePickerMode="True" FirstDayOfWeek="6" ScriptPath="Calendar/calendarscript"
    StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<" 
    TextArrowRight=">>" TextBoxId="txtDateTo">
    </obout:Calendar>     
    </td>
</tr>
<tr id='tr_inquery' runat='server' visible='false'>
<td class="style11" align=left style='background:#E9E7E2'>
Layer
</td>
<td class="style14">    
<asp:DropDownList ID="ddl_layer" runat="server" Width="120px"></asp:DropDownList>
</td> 
<td colspan=2 align=right class="style16">
<asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>&nbsp;&nbsp;&nbsp;
<asp:Button ID="but_Execute" runat="server" Text="Inquery" Height=30px Width="86px" 
        CssClass="BT_1" Font-Bold="True" Font-Size="Medium" />
</td>
</tr> 
<!-- Chart Panel -->
<tr>
<td colspan=4 style="height: 1px;" >
<asp:Panel ID="Panel1" runat="server" BorderColor="#99FF33" BackColor="White" EnableViewState="False">
</asp:Panel>
</td>
</tr>
</table>


<table>
<!-- Grid View Start -->
<tr>
<td>
<asp:Button ID="Button1" runat="server" Text="ReFresh" Height=25px Width="86px" CssClass="BT_1" Font-Bold="True" Font-Size="Small" onclick="Button1_Click"/>
</td>
<td align=right>
<ext:Button ID="Button3" runat="server" Text="登入下註解" Icon="LockOpen"><Listeners><Click Handler="#{Window1}.show();" /></Listeners></ext:Button>
</td>
</tr>

<tr>
<td colspan=2>
<asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" 
       BackColor="White" BorderColor="Black" BorderStyle="None" BorderWidth="1px"  
        CellPadding="3" Width="1250px" onrowdatabound="GridView1_RowDataBound">
        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
        <Columns>
            <asp:TemplateField>
            <HeaderTemplate>No.</HeaderTemplate>
            <ItemTemplate><%# Container.DataItemIndex + 1 %></ItemTemplate>
            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:TemplateField>
            <asp:BoundField DataField="Part" HeaderText="PartID" />
            <asp:BoundField DataField="Lot" HeaderText="LotID" />
            <asp:BoundField DataField="Parametric_Measurement" HeaderText="Parametric" />
            <asp:BoundField DataField="MPID" HeaderText="Operation" />
            <asp:BoundField DataField="EQPID" HeaderText="Machine No" />
            <asp:BoundField DataField="MachineName" HeaderText="Machine Name" />
            <asp:BoundField DataField="Layer" HeaderText="Layer" />
            <asp:BoundField DataField="trtm" HeaderText="Process Time" />
            <asp:BoundField DataField="MeasureTime" HeaderText="Measure Time" />
            <asp:BoundField DataField="InStationDate" HeaderText="Operation IN" />
            <asp:BoundField DataField="InMachineDate" HeaderText="Machine IN" />
            <asp:BoundField DataField="OutMachineDate" HeaderText="Machine Out" />
            <asp:BoundField DataField="OutStationDate" HeaderText="Operation Out" />
            <asp:BoundField DataField="meanval" HeaderText="Mean" />
            <asp:BoundField DataField="std" HeaderText="Std" />
            <asp:BoundField DataField="maxval" HeaderText="Max Value" />
            <asp:BoundField DataField="minval" HeaderText="Min Value" />
            <asp:BoundField DataField="CP" HeaderText="CP" />
            <asp:BoundField DataField="CPK" HeaderText="CPK" />
            <asp:TemplateField HeaderText="Comment" ShowHeader="False">
                <ItemTemplate>
                    <asp:Button ID="but_exidComment" runat="server" Text="Edit" />
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField DataField="comment" HeaderText="commentStr" />
        </Columns>
        <EditRowStyle BackColor="#999999" />
        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" HorizontalAlign="Center" />
        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" HorizontalAlign="Center" />
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
</form>
</body>
</html>
