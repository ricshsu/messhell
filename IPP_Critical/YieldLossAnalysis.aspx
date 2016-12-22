<%@ Page Language="VB" AutoEventWireup="false" CodeFile="YieldLossAnalysis.aspx.vb"
    Inherits="YieldLossAnalysis" %>

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
        .style5
        {
            width: 200px;
            height: 34px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <table border="1" width="1080px">
        <!-- Title -->
        <tr>
            <td colspan='2' class="Table_One_Title" valign='middle' align="center" style='font-size: x-large;
                font-weight: bold'>
                Yield&nbsp;Loss Commonality&nbsp;Analysis
            </td>
        </tr>
        <!-- Target -->
        <tr id="tr_upload" runat="server" visible="true">
            <td align='left' style='background: #E9E7E2; width: 200px'>
                Target Lot
            </td>
            <td style='width: 880px'>
                <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal"
                    Width="228px" AutoPostBack="True">
                    <asp:ListItem Selected="True">By Time Range</asp:ListItem>
                    <asp:ListItem>By Upload</asp:ListItem>
                </asp:RadioButtonList>
                <asp:FileUpload ID="uf_UfilePath" runat="server" Width="314px" Enabled="False"></asp:FileUpload>
                <asp:Button ID="but_Uupload" runat="server" Text="Upload" Enabled="False" 
                    Width="52px" /> 
                     <asp:Label ID="Label7" runat="server" Font-Bold="True" ForeColor="Red" 
                                Text="(紅色表示upload lot)"></asp:Label>
                    <asp:TextBox ID="TextBox2" runat="server" Width="881px" Enabled="False"></asp:TextBox>
            </td>
        </tr>
        <!-- Product -->
        <tr id="ProductType"  runat='server' visible='true'>
            <td align='left' style='background: #E9E7E2; width: 200px'>
                Product Type&nbsp;
            </td>
            <td style='width: 880px'>
                <asp:DropDownList ID="ddlProduct" runat="server" Width="200px" Height="25px" AutoPostBack="True">
                    <asp:ListItem Selected="True">PPS</asp:ListItem>
                    <asp:ListItem>CPU</asp:ListItem>
                    <asp:ListItem>CS</asp:ListItem>
                    <asp:ListItem>PCB</asp:ListItem>
                </asp:DropDownList>
            </td>
        </tr>
        <!-- Vendor -->
        <tr id="CustomerType"  runat='server' visible='true'>
            <td align='left' style='background: #E9E7E2; width: 200px'>
                Customer
            </td>
            <td style='width: 880px'>
                <asp:DropDownList ID="ddlCustomer" runat="server" Width="200px" AutoPostBack="True"
                    Height="25px">
                </asp:DropDownList>
            </td>
        </tr>
        <!-- Part -->
        <tr id="CategoryType" runat='server' visible='true'>
            <td align="left" style='background: #E9E7E2;' class="style1">
                Product No / Part No
            </td>
            <td class="style2">
                <table>
                    <tr>
                        <td>
                            <asp:RadioButtonList ID="rb_ProductPart" runat="server" RepeatDirection="Horizontal"
                                AutoPostBack="True" Width="292px">
                                <asp:ListItem Value="0" Enabled="False">Product No</asp:ListItem>
                                <asp:ListItem>BumpingType</asp:ListItem>
                                <asp:ListItem Value="1" Selected="True">Part No</asp:ListItem>
                            </asp:RadioButtonList>
                            <asp:Label ID="Label2" runat="server" Text="Keyword"></asp:Label>
                            <asp:TextBox ID="TextBox1" runat="server" Width="101px"></asp:TextBox>
                            <asp:Button ID="btn_Filter_Part" runat="server" Text="OK" BackColor="#0033CC" 
                                BorderColor="#0033CC" Font-Bold="True" ForeColor="White" />
                        </td>
                        <%--    <td>
    <asp:DropDownList ID="ddlPart" runat="server" Width="200px" Height="25px" AutoPostBack="True"></asp:DropDownList>
    </td>--%>
                        <td style='width: 120px'>
                            <asp:ListBox ID="listB_PartSource" runat="server" Height="90px" Width="172px" 
                                SelectionMode="Multiple">
                            </asp:ListBox>
                        </td>
                        <td style='width: 20px'>
                            <asp:Button ID="but_PartRight" runat="server" Text=">>" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                            <br />
                            <asp:Button ID="but_PartLeft" runat="server" Text="<<" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                        </td>
                        <td>
                            <asp:ListBox ID="listB_PartShow" runat="server" Height="90px" Width="172px"
                            SelectionMode="Multiple">
                            </asp:ListBox>
                            <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red" 
                                Text="(*備註：開放複選十個料號)"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <!-- Station No -->
        <tr>
            <td align='left' style='background: #E9E7E2; width: 200px; height: 10px'>
                Station No&nbsp;
            </td>
            <td colspan="5" style='width: 900px; white-space: nowrap'>
                <asp:RadioButtonList ID="rbl_Station" runat="server" 
                    RepeatDirection="Horizontal" Width="232px"
                    AutoPostBack="True" Height="20px">
                    <asp:ListItem Selected="True">By FVI (FI0/X81)</asp:ListItem>
                    <asp:ListItem>By Process </asp:ListItem>
                </asp:RadioButtonList>
                 <asp:Label ID="Label5" runat="server" Text="Keyword"></asp:Label>
                            <asp:TextBox ID="TextBox3" runat="server" Width="101px" 
                    Enabled="False"></asp:TextBox>
                            <asp:Button ID="Button1" runat="server" Text="OK" BackColor="#0033CC" 
                                BorderColor="#0033CC" Font-Bold="True" ForeColor="White" 
                    Enabled="False" />
                &nbsp;&nbsp;&nbsp;&nbsp;
                <asp:DropDownList ID="ddlStation" runat="server" Width="200px" Height="23px" 
                    AutoPostBack="True" Enabled="False">
                </asp:DropDownList>
                <asp:CheckBox ID="chkParallel" runat="server" Checked="True" Enabled="False" 
                    Text="Parallel" />
                <asp:CheckBox ID="chkRebuilt" runat="server" Checked="True" Enabled="False" 
                    Text="Rebuilt" />
            </td>
        </tr>
        <tr id="tr_time" runat='server' visible='true'>
            <td align='left' style='background: #E9E7E2; width: 200px' id="TimeRange">
                Time Query Range
            </td>
            <td colspan="5" style='width: 880px'>
                <!-- By Date -->
                <asp:TextBox ID="txtDateFrom" runat="server" Columns="10" MaxLength="10" Width="110px"
                    Enabled="true"></asp:TextBox>
                <obout:Calendar ID="Calendar1" runat="server" Columns="1" DateFormat="yyyy-MM-dd"
                    DatePickerImagePath="images/calendar.gif" DatePickerMode="True" FirstDayOfWeek="6"
                    ScriptPath="Calendar/calendarscript" StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<"
                    TextArrowRight=">>" TextBoxId="txtDateFrom">
                </obout:Calendar>
                &nbsp;~&nbsp;
                <asp:TextBox ID="txtDateTo" runat="server" Columns="10" MaxLength="10" Width="110px"
                    Enabled="true"></asp:TextBox>
                <obout:Calendar ID="Calendar2" runat="server" Columns="1" DateFormat="yyyy-MM-dd"
                    DatePickerImagePath="images/calendar.gif" DatePickerMode="True" FirstDayOfWeek="6"
                    ScriptPath="Calendar/calendarscript" StyleFolder="Calendar/styles/blocky" TextArrowLeft="<<"
                    TextArrowRight=">>" TextBoxId="txtDateTo">
                </obout:Calendar>
                <!-- By Week -->
                &nbsp;&nbsp;&nbsp;
                <asp:RadioButton ID="rb_byWeek" runat="server" Text="By Week" AutoPostBack="True"
                    Visible="false" />&nbsp;&nbsp;&nbsp;
                <asp:DropDownList ID="ddlWeekStart" runat="server" Visible="false">
                </asp:DropDownList>
                <asp:DropDownList ID="ddlWeekEnd" runat="server" Enabled="false" Visible="false">
                </asp:DropDownList>
            </td>
        </tr>

        <!-- Custom Yield Loss Item -->
                <!-- Part -->
        <tr>
            <td align="left" style='background: #E9E7E2;' class="style1">
                 Yield Loss Item
            </td>
            <td class="style2">
                <table>
                    <tr>
                        <td>
                            <asp:RadioButtonList ID="RadioButtonList2" runat="server" RepeatDirection="Horizontal"
                                AutoPostBack="True" Width="233px" Height="16px">
                                <asp:ListItem Value="0" Selected="True">Fail Mode</asp:ListItem>
                                <asp:ListItem Value="2">MIS Defect Mode</asp:ListItem>
                            </asp:RadioButtonList>
                            <asp:Label ID="Label3" runat="server" Text="Keyword"></asp:Label>
                            <asp:TextBox ID="txt_YieldlossFilter" runat="server" Width="101px"></asp:TextBox>
                            <asp:Button ID="Button2" runat="server" Text="OK" BackColor="#0033CC" 
                                BorderColor="#0033CC" Font-Bold="True" ForeColor="White" />
                        </td>
                        <%--    <td>
    <asp:DropDownList ID="ddlPart" runat="server" Width="200px" Height="25px" AutoPostBack="True"></asp:DropDownList>
    </td>--%>
                        <td style='width: 120px'>
                            <asp:ListBox ID="lst_Yieldloss_Source" runat="server" Height="90px" 
                                Width="172px" SelectionMode="Multiple">
                            </asp:ListBox>
                        </td>
                        <td style='width: 20px'>
                            <asp:Button ID="Button3" runat="server" Text=">>" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                            <br />
                            <asp:Button ID="Button4" runat="server" Text="<<" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                        </td>
                        <td>
                            <asp:ListBox ID="lst_Yieldloss_Target" runat="server" Height="90px" 
                                Width="172px">
                            </asp:ListBox>
                              <asp:Label ID="Label4" runat="server" Font-Bold="True" ForeColor="Red" 
                                Text="(*備註：目前只開放單一報廢選項)"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
          <tr id="Advanced" runat='server' visible='false'>
            <td align='left' style='background: #E9E7E2; ' class="style5">
                Advanced&nbsp; Extension</td>
                  <td class="style2">
                <table>
                    <tr>
                        <td>
                         Backward :&nbsp;
                            <asp:DropDownList ID="ddlADStart" runat="server" Width="100px" Height="22px" 
                    AutoPostBack="True">
                                <asp:ListItem>1</asp:ListItem>
                                <asp:ListItem>2</asp:ListItem>
                                <asp:ListItem>3</asp:ListItem>
                                <asp:ListItem>4</asp:ListItem>
                                <asp:ListItem>5</asp:ListItem>
                </asp:DropDownList>
                         Forward&nbsp;&nbsp; :&nbsp; <asp:DropDownList ID="ddlADEnd" runat="server" 
                                Width="100px" Height="23px" 
                    AutoPostBack="True">
                                <asp:ListItem>1</asp:ListItem>
                                <asp:ListItem>2</asp:ListItem>
                                <asp:ListItem>3</asp:ListItem>
                                <asp:ListItem>4</asp:ListItem>
                                <asp:ListItem>5</asp:ListItem>
                </asp:DropDownList>
                         &nbsp;</td>
                        <td>
                            <asp:RadioButtonList ID="RadioButtonList3" runat="server" RepeatDirection="Horizontal"
                                AutoPostBack="True" Width="292px">
                                <asp:ListItem Value="0" Enabled="False">Product No</asp:ListItem>
                                <asp:ListItem Selected="True">BumpingType</asp:ListItem>
                                <asp:ListItem Value="1" Enabled="False">Part No</asp:ListItem>
                            </asp:RadioButtonList>
                            <asp:Label ID="Label6" runat="server" Text="Keyword"></asp:Label>
                            <asp:TextBox ID="TextBox4" runat="server" Width="101px"></asp:TextBox>
                            <asp:Button ID="Button5" runat="server" Text="OK" BackColor="#0033CC" 
                                BorderColor="#0033CC" Font-Bold="True" ForeColor="White" />
                        </td>

                        <td style='width: 120px'>
                            <asp:ListBox ID="lst_Ad_Source" runat="server" Height="90px" Width="172px">
                            </asp:ListBox>
                        </td>
                        <td style='width: 20px'>
                            <asp:Button ID="btn_Ad_Right" runat="server" Text=">>" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                            <br />
                            <asp:Button ID="Button7" runat="server" Text="<<" Height="25px"
                                Width="30px" BackColor="#0033CC" BorderColor="#0033CC" Font-Bold="True" 
                                ForeColor="White" />
                        </td>
                        <td>
                            <asp:ListBox ID="lst_Ad_Target" runat="server" Height="90px" Width="172px">
                            </asp:ListBox>
                        </td>
                    </tr>
                </table>
            </td>
              <%-- <td colspan="5" style='white-space: nowrap' class="style4">
                 <asp:RadioButtonList ID="RadioButtonList3" runat="server" 
                    RepeatDirection="Horizontal" Width="212px"
                    AutoPostBack="True" Height="20px" Visible="False">
                    <asp:ListItem Selected="True">By FVI (FI0)</asp:ListItem>
                    <asp:ListItem>By Process </asp:ListItem>
                </asp:RadioButtonList>
                 BumpingType:&nbsp;
                <asp:DropDownList ID="ddlExtenBumpingType" runat="server" Width="200px" 
                     Height="23px">
                </asp:DropDownList>
            &nbsp;&nbsp; Week:
                <asp:DropDownList ID="ddlExtenWeek" runat="server" Width="91px" Height="25px">
                    <asp:ListItem Selected="True">1</asp:ListItem>
                    <asp:ListItem>2</asp:ListItem>
                    <asp:ListItem>3</asp:ListItem>
                    <asp:ListItem>4</asp:ListItem>
                    <asp:ListItem>5</asp:ListItem>
                </asp:DropDownList>
            </td>--%>
        </tr>
        <!-- Query -->
        <tr>
            <td colspan='2' align='right' width="100%">
                <asp:Label ID="lab_wait" runat="server" Font-Bold="True" ForeColor="#CC0000"></asp:Label>
                &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="ckFAI" runat="server" Text="試製" />
                &nbsp;&nbsp;
                <asp:CheckBox ID="ckMachine" runat="server" Text="機台commonality" 
            Checked="False" AutoPostBack="True" /> 
                <asp:TextBox ID="txtMachine" runat="server" Height="18px" Visible="False" 
                    Width="151px"></asp:TextBox>
                &nbsp;&nbsp;
                 <asp:CheckBox ID="ckAdvanced" runat="server" Text="Advanced Extension" 
            Checked="False" AutoPostBack="True" Visible="False" /> &nbsp;&nbsp;
                <asp:CheckBox ID="cb_Lot_Merge" runat="server" Text="合批" /> &nbsp;&nbsp;
            <asp:CheckBox ID="cb_SF" runat="server" Text="SF" /> &nbsp;&nbsp;
            <asp:CheckBox ID="cb_CR" runat="server" Text="客退" /> &nbsp;&nbsp;
              <asp:CheckBox ID="Cb_Oversea" runat="server" Text="海外托工" AutoPostBack="True" /> &nbsp;&nbsp;
                <asp:Button ID="but_Export" runat="server" Text="Export" Height="30px" Width="110px"
                    CssClass="BT_1" Font-Bold="True" Font-Size="Medium" />
                <asp:Button ID="but_Execute" runat="server" Text="Query" Height="30px" Width="110px"
                    CssClass="BT_1" Font-Bold="True" Font-Size="Medium" />
                     
            </td>
        </tr>
    </table>
    <table>
        <!-- Chart Start -->
        <tr id="tr_chartDisplay" runat="server" visible="false">
            <td class="style12">
                <asp:Panel ID="Chart_Panel" runat="server" BorderColor="#99FF33" BackColor="White">
                </asp:Panel>
            </td>
        </tr>
        <!-- Chart E n d -->
        <!-- GridView Start -->
        <tr id="tr_gvDisplay" runat="server" visible="false">
            <td class="style12">
                <asp:GridView ID="gv_rowdata" runat="server" BackColor="White" BorderColor="Black"
                    BorderStyle="None" BorderWidth="1px" CellPadding="3" AllowPaging="True" 
                    EnableSortingAndPagingCallbacks="True" PageSize="100">
                    <AlternatingRowStyle BackColor="#DBEEFF" />
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#ffffff" ForeColor="Black" />
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
    <table id='tb_uploadLot' runat="server" visible="false">
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
                <asp:GridView ID="Lot_GridView" runat="server" BackColor="White" BorderColor="Black"
                    BorderStyle="None" BorderWidth="1px" CellPadding="3">
                    <AlternatingRowStyle BackColor="#DBEEFF" />
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#ffffff" ForeColor="Black" />
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
