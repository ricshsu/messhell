<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FailDetail_Test.aspx.cs"
    Inherits="FailDetail_Test" %>

<%@ Register Assembly="Ext.Net" Namespace="Ext.Net" TagPrefix="ext" %>
<html>
<head id="Head1" runat="server">
    <title>FailMode Detail</title>
    <link href="/resources/css/examples.css" rel="stylesheet" />
    <link href="css/Table.css" rel="Stylesheet" />
    <link href="css/Button.css" rel="Stylesheet" />
    <link href="css/PUPPY.css" rel="Stylesheet" />
    <link href="css/TabContainer.css" rel="Stylesheet" />
    <link href="css/DW.css" rel="Stylesheet" />
    <script type="text/javascript" language='javascript'>
        function LinkPoint(dataID) {
            var linkPoint = document.getElementById(('GV_LotRowData_' + dataID));
            linkPoint.style.background = "#FF6A6A";
            document.location = ('#' + dataID);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <ext:ResourceManager ID="ResourceManager1" runat="server" />
    <table border='1' cellpadding='1' cellspacing='1'>
        <tr>
            <td>
                <asp:Panel ID="titlePanel" runat="server" BorderColor="#99FF33" BackColor="White">
                </asp:Panel>
            </td>
        </tr>
        <tr>
            <td>
                <ext:TabStrip ID="TabStrip3" runat="server" Plain="true" Width="1000px">
                    <Items>
                        <ext:Tab Text="Lot Detail" ActionItemID="elm1" />
                        <ext:Tab Text="Fail Detail(Pie Chart)" ActionItemID="elm2" />
                        <ext:Tab Text="Bump Detail" ActionItemID="elm3" />
                        <ext:Tab Text="Fail Detail(Pareto)" ActionItemID="elm4" />
                        <ext:Tab Text="Inline(Raw Data)" ActionItemID="elm5" />
                        <ext:Tab Text="Fail By Plant(Pie Chart)" ActionItemID="elm6" />
                        <ext:Tab Text="Total By Plant(Pie Chart)" ActionItemID="elm7" />
                    </Items>
                </ext:TabStrip>
                <!-- Thrend Chart -->
                <div id="elm1" style="padding: 5px; border: 2px;">
                    <table>
                        <tr valign='top'>
                            <td>
                                <asp:Panel ID="ThendPanel" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#0000CC" align='center' style="font-size: medium;">
                                <asp:Label ID="lab_lotRowData" runat="server" ForeColor="White"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:GridView ID="GV_LotRowData" runat="server" BackColor="White" BorderColor="Black"
                                    BorderStyle="None" BorderWidth="1px" CellPadding="3" OnRowDataBound="GV_LotRowData_RowDataBound">
                                    <AlternatingRowStyle BackColor="#DBEEFF" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                No.</HeaderTemplate>
                                            <ItemTemplate>
                                                <%# Container.DataItemIndex + 1 %></ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                    </Columns>
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
                </div>
                <!-- Pie Chart -->
                <div id="elm2" style="padding: 5px;">
                    <table>
                        <tr>
                            <td align="left">
                                <asp:Panel ID="PiePanel" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td align="left" style='width=1000px'>
                                <asp:GridView ID="gv_pie" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None"
                                    BorderWidth="1px" CellPadding="3" OnRowDataBound="gv_pie_RowDataBound1">
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
                </div>
                <!-- Pareto -->
                <div id="elm3" style="padding: 5px; border: 2px;">
                    <table>
                        <!-- Pareto Chart -->
                        <tr>
                            <td>
                                <asp:Panel ID="DetailParetoPanel" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <!-- Title -->
                        <tr>
                            <td bgcolor="#0000CC" align='center' style="font-size: medium;">
                                <asp:Label ID="lab_DetailTitle" runat="server" ForeColor="White"></asp:Label>
                            </td>
                        </tr>
                        <!-- Pareto Chart Detail -->
                        <tr>
                            <td>
                                <asp:GridView ID="gr_lotview" runat="server" BackColor="White" BorderColor="Black"
                                    BorderStyle="None" BorderWidth="1px" CellPadding="3" OnRowDataBound="gr_lotview_RowDataBound1">
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
                </div>
                <!-- New Pareto Chart -->
                <div id="elm4" style="padding: 5px; border: 2px;">
                    <table>
                        <tr valign='top'>
                            <td>
                                <asp:Panel ID="NewParetoPanel" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#0000CC" align='center' style="font-size: medium;">
                                <asp:Label ID="lab_NewLotRowData" runat="server" ForeColor="White"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:GridView ID="GV_NewLotRowData" runat="server" BackColor="White" BorderColor="Black"
                                    BorderStyle="None" BorderWidth="1px" CellPadding="3" OnRowDataBound="GV_NewLotRowData_RowDataBound">
                                    <AlternatingRowStyle BackColor="#DBEEFF" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                No.</HeaderTemplate>
                                            <ItemTemplate>
                                                <%# Container.DataItemIndex + 1 %></ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                    </Columns>
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
                </div>
                <!-- inline raw data -->
                <div id="elm5" style="padding: 5px; border: 2px;">
                    <table>                        
                        <tr>
                            <td bgcolor="#0000CC" align='center' style="font-size: medium;">
                                <asp:Label ID="Label1" runat="server" ForeColor="White"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:GridView ID="GridView1" runat="server" BackColor="White" BorderColor="Black"
                                    BorderStyle="None" BorderWidth="1px" CellPadding="3" >
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
                </div>
                <!-- Pie Chart1 -->
                <div id="elm6" style="padding: 5px;">
                    <table>
                        <tr>
                            <td align="left">
                                <asp:Panel ID="PiePanel2" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td align="left" style='width=1000px'>
                                <asp:GridView ID="GridView2" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None"
                                    BorderWidth="1px" CellPadding="3" OnRowDataBound="gv_pie_RowDataBound1">
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
                </div>

                <!-- total -->
                <div id="elm7" style="padding: 5px;">
                    <table>
                        <tr>
                            <td align="left">
                                <asp:Panel ID="PiePanel3" runat="server" BorderColor="#99FF33" BackColor="White">
                                </asp:Panel>
                            </td>
                        </tr>
                        <tr>
                            <td align="left" style='width=1000px'>
                                <asp:GridView ID="GridView3" runat="server" BackColor="White" BorderColor="Black" BorderStyle="None"
                                    BorderWidth="1px" CellPadding="3" OnRowDataBound="gv_pie_RowDataBound1">
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
                </div>
                <asp:Label ID="Label2" runat="server" Text="Label"></asp:Label>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
