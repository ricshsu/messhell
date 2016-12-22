<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Critical_LotDetails.aspx.cs" Inherits="Critical_LotDetails" %>
<%@ Register TagPrefix="obout" Namespace="OboutInc.Calendar2" Assembly="obout_Calendar2_Net" %>
<%@ Register assembly="Ext.Net" namespace="Ext.Net" tagprefix="ext" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<title>Critical Item Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<link href="/resources/css/examples.css" rel="stylesheet" />
<script language=javascript>
    
    function LinkPoint(dataID) {
        
        var linkPoint = document.getElementById(('Lot_GridView_' + dataID));
        linkPoint.style.background = "#FF6A6A";
        document.location = ('#' + dataID);

        /*
        var linkPoint = document.getElementById(('Lot_GridView_' + dataID));
        var preIDObj = document.getElementById('preID');
        var preColorObj = document.getElementById('preColor');
        if (preIDObj.value == "") {
        preIDObj.value = dataID;
        preColorObj.value = linkPoint.style.backgroundColor;
        } else {
        document.getElementById('Lot_GridView_' + (preIDObj.value)).style.backgroundColor = preColorObj.value;
        preIDObj.value = dataID;
        preColorObj.value = linkPoint.style.backgroundColor;
        alert(preColorObj.value);
        }
        linkPoint.style.background = "#FF6A6A";
        document.location = ('#' + dataID);
        */
    }

    function openCommand(url, name, partID, lotID, paramID, userID, e) 
    {
        
        var xPosition = e.clientX;
        var yPosition = e.clientY;
        var fertrue = "height=630,width=520,top=" + xPosition + ",left=" + yPosition+ ",toolbar=no,menubar=no,scrollbars=yes,resizable=yes,location=no,status=no";
        var newWindow = window.open(url, name, fertrue);
        var html = "<html><body><form id='formid' method='post' action='" + url + "'>";
        html += "<input type='hidden' name='partID' value='" + partID + "'/>";
        html += "<input type='hidden' name='lotID' value='" + lotID + "'/>";
        html += "<input type='hidden' name='item' value='" + paramID + "'/>";
        html += "<input type='hidden' name='userID' value='" + userID + "'/>";
        html += "<input type='hidden' name='fType' value='Critical_Lot'/>";
        html += "</form><script type='text/javascript'>document.getElementById(\"formid\").submit();</";
        html += "script></body></html>";
        newWindow.document.write(html);

    }

    function setPosition(e) {
        var WindowObj = document.getElementById('Window1');
        WindowObj.x = e.clientX;
        WindowObj.y = e.clientY;
    }

    function pageReFresh() 
    {
        document.getElementById('Button1').click();
    }
</script>
</head>
<body style="MARGIN-TOP: 10px; MARGIN-LEFT: 10px" MS_POSITIONING="FlowLayout">
<form id="form1" runat="server">
<ext:ResourceManager ID="ResourceManager1" runat="server" />
<!-- Ext.NET -->
<ext:Window ID="Window1" runat="server" Closable="false" Resizable="false" Height="150" Width="350" Icon="Lock" Title="Login For Command" Draggable="false" Modal="true"
BodyPadding="5" Layout="FormLayout" X=500 Y=1300 >
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
<table border="1">
<!-- Correlation -->
<tr id='tr_correlation' runat='server' visible='false' >
<td colspan=4 align="left" style='font-size:x-large;font-weight: bold'>
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
<asp:Button ID="but_Execute" runat="server" Text="Inquery" Height=30px Width="86px" CssClass="BT_1" Font-Bold="True" Font-Size="Medium"/>
</td>
</tr> 
</table>

<!-- Chart Panel -->
<table border="0">
<tr>
<td style="height: 1px;" >
<table>
<asp:Panel ID="ChartPanel" runat="server" BorderColor="#99FF33" BackColor="White" EnableViewState="False">
</asp:Panel>
</table>
</td>
</tr>
</table>

<!-- Grid View -->
<table>
<tr>
<td>
<asp:Button ID="Button1" runat="server" Text="ReFresh" Height=25px Width="86px" CssClass="BT_1" Font-Bold="True" Font-Size="Small" onclick="Button1_Click"/>
</td>
<td align=right>
<ext:Button ID="Button3" runat="server" Text="登入下註解" Icon="LockOpen"><Listeners><Click Handler="#{Window1}.show();" /></Listeners></ext:Button>
</td>
</tr>

<tr>
<td colspan='2' style='width:1100px'>
<asp:GridView ID="Lot_GridView" runat="server" BackColor="White" 
        BorderColor="Black" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
        AutoGenerateColumns="False" onrowdatabound="Lot_GridView_RowDataBound1">
    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
    <Columns>
        <asp:TemplateField>
            <HeaderTemplate>No.</HeaderTemplate>
            <ItemTemplate><%# Container.DataItemIndex + 1 %></ItemTemplate>
        </asp:TemplateField>
        <asp:BoundField DataField="Part" HeaderText="Part" >
        <HeaderStyle Width="75px" />
        </asp:BoundField>
        <asp:BoundField DataField="Lot" HeaderText="Lot" >
        <HeaderStyle Width="75px" />
        </asp:BoundField>
        <asp:BoundField DataField="Parametric_Measurement" HeaderText="Parameter" >
        <HeaderStyle Width="100px" />
        </asp:BoundField>
        <asp:BoundField DataField="Plant" HeaderText="Plant" >
        <HeaderStyle Width="75px" />
        </asp:BoundField>
        <asp:BoundField DataField="MIS_OP" HeaderText="Operation" />
        <asp:BoundField DataField="mchno" HeaderText="Machine" >
        <HeaderStyle Width="75px" />
        </asp:BoundField>
        <asp:BoundField DataField="trtm" HeaderText="DataTime" >
        <HeaderStyle Width="200px" />
        </asp:BoundField>
        <asp:BoundField DataField="meanval" HeaderText="Mean" >
        <HeaderStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="std" HeaderText="Std" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="maxval" HeaderText="Max" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="minval" HeaderText="Min" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="CP" HeaderText="CP" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="CPK" HeaderText="CPK" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="USL" HeaderText="USL" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:BoundField DataField="LSL" HeaderText="LSL" >
        <ItemStyle Width="50px" />
        </asp:BoundField>
        <asp:TemplateField HeaderText="Comment" ShowHeader="False">
            <ItemTemplate>
                <asp:Button ID="Button2" runat="server" Text="Edit" />
            </ItemTemplate>
        </asp:TemplateField>
        <asp:BoundField DataField="comment" HeaderText="CommentStr" />
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
</table>
</form>
    <p>
        <input id="preID" type="hidden" value=""/>
        <input id="preColor" type="hidden" value=""/>
    </p>
</body>
</html>
