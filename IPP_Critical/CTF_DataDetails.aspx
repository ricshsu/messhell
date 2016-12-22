<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CTF_DataDetails.aspx.cs" Inherits="CTF_DataDetails" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Xml.Xsl" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Linq" %>
<%@ Register Assembly="Ext.Net" Namespace="Ext.Net" TagPrefix="ext" %>

<html>
<head id="Head1" runat="server">
<title>CTF Detail</title>
<link href="css/Table.css" rel="Stylesheet" />
<link href="css/Button.css" rel="Stylesheet" />
<link href="css/PUPPY.css" rel="Stylesheet" />
<link href="css/TabContainer.css" rel="Stylesheet" />
<link href="css/DW.css" rel="Stylesheet" />
<link href="/resources/css/examples.css" rel="stylesheet" />
<script>
        var template = '<span style="color:{0};">{1}</span>';
        var change = function (value) {
            return Ext.String.format(template, (value > 0) ? "green" : "red", value);
        };

        var pctChange = function (value) {
            return Ext.String.format(template, (value > 0) ? "green" : "red", value + "%");
        };

        var exportData = function (format) {
            App.FormatType.setValue(format);
            var store = App.GridPanel1.store;

            store.submitData(null, { isUpload: true });
        };

        
</script>
<style type="text/css">
.x-grid-custom .x-grid3-row-table TD 
{
  line-height:200px;
  height:100px;
}
</style>
</head>
<body>
<center>
<form id="form1" runat="server">
<table>
<tr>
<td>
<ext:ResourceManager ID="ResourceManager1" runat="server" />
<ext:Store ID="Store1" runat="server" OnReadData="Store1_RefreshData" OnSubmitData="Store1_Submit" PageSize="10">
<Model>
                <ext:Model ID="Model1" runat="server">
                    <Fields>
                        <ext:ModelField Name="Part_Id" />
                        <ext:ModelField Name="lot_id"  />
                        <ext:ModelField Name="Machine_id" />
                        <ext:ModelField Name="Meas_item" />
                        <ext:ModelField Name="Max_Value" Type="Float" />
                        <ext:ModelField Name="Min_Value" Type="Float" />
                        <ext:ModelField Name="Mean_value" Type="Float" />
                        <ext:ModelField Name="Std_value" Type="Float" />
                        <ext:ModelField Name="CP" Type="Float" />
                        <ext:ModelField Name="CPK" Type="Float" />
                        <ext:ModelField Name="RowData" />
                    </Fields>
                </ext:Model>
</Model>
</ext:Store>
        
<ext:Hidden ID="FormatType" runat="server" />

<ext:GridPanel id="GridPanel1" runat="server" StoreID="Store1" Title="CTF Detail" Width="1100" Height="450" Cls="x-grid-custom" columnLines="true" RowLines="true">

<ColumnModel ID="ColumnModel1" runat="server" >
<Columns>
                    
                    <ext:Column ID="Column1" runat="server" Text="PART" Width="60" DataIndex="Part_Id">
                        <Editor>
                            <ext:TextField ID="TextField1" runat="server" />
                        </Editor>
                    </ext:Column>

                    <ext:Column ID="Column2" runat="server" Text="LOT" Width="110" DataIndex="lot_id">
                        <Editor>
                            <ext:TextField ID="TextField2" runat="server" />
                        </Editor>
                    </ext:Column>

                    <ext:Column ID="Column3" runat="server" Text="Tool" Width="60" DataIndex="Machine_id">
                        <Editor>
                            <ext:TextField ID="TextField3" runat="server" />
                        </Editor>
                    </ext:Column>

                    <ext:Column ID="Column4" runat="server" Text="MeasItem" Width="120" DataIndex="Meas_item">
                        <Editor>
                            <ext:TextField ID="TextField4" runat="server" />
                        </Editor>
                    </ext:Column>

                    <ext:Column ID="Column5" runat="server" Text="Max" Width="40" DataIndex="Max_Value">
                    </ext:Column>

                    <ext:Column ID="Column6" runat="server" Text="Min" Width="40" DataIndex="Min_Value">
                    </ext:Column>

                    <ext:Column ID="Column7" runat="server" Text="Mean" Width="40" DataIndex="Mean_value">
                    </ext:Column>

                    <ext:Column ID="Column8" runat="server" Text="Std" Width="40" DataIndex="Std_value">
                    </ext:Column>

                    <ext:Column ID="Column9" runat="server" Text="CP" Width="40"  DataIndex="CP">
                    </ext:Column>

                    <ext:Column ID="Column10" runat="server" Text="CPK" Width="40"  DataIndex="CPK">
                    </ext:Column>

                    <ext:Column ID="Column11" runat="server" Text="RowData" Width="510" AutoScroll="true" DataIndex="RowData" >
                        <Editor>
                            <ext:TextField ID="TextField5" runat="server"/>
                        </Editor>
                    </ext:Column>

</Columns>
</ColumnModel>
            
<SelectionModel>
<ext:RowSelectionModel ID="RowSelectionModel1" runat="server" Mode="Multi" />
</SelectionModel>
    
<TopBar>
                <ext:Toolbar ID="Toolbar1" runat="server">
                    <Items>
                        <ext:ToolbarFill ID="ToolbarFill1" runat="server" />
                        <ext:Button ID="Button2" runat="server" Text="To Excel" Icon="PageExcel">
                            <Listeners>
                                <Click Handler="exportData('xls');" />
                            </Listeners>
                        </ext:Button>
                    </Items>
                </ext:Toolbar>
            </TopBar>
            
<BottomBar>
<ext:PagingToolbar ID="PagingToolbar1" runat="server" StoreID="Store1" />
</BottomBar>

</ext:GridPanel>

    

</td>
</tr>
</table>    
</form>
</center>
</body>
</html>
