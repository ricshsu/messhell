<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Critical_Command.aspx.cs" Inherits="Critical_Command" %>
<%@ Import Namespace="System.Collections.Generic"%>
<%@ Register Assembly="Ext.Net" Namespace="Ext.Net" TagPrefix="ext" %>
<html>
<head id="Head1" runat="server">
    <title>Critical_Command</title>
    <link href="/resources/css/examples.css" rel="stylesheet" />    
    <script>
        function backOpener() {
            this.close();
        }
    </script>
    <script>
        
        var updateRecord = function (form) {

            if (form.getForm()._record == null) {
                return;
            }

            if (!form.getForm().isValid()) {
                Ext.net.Notification.show({
                    iconCls: "icon-exclamation",
                    html: "Form is invalid",
                    title: "Error"
                });
                return false;
            }

            form.getForm().updateRecord();
        };

        var addRecord = function (form, grid) {
            if (!form.getForm().isValid()) {
                Ext.net.Notification.show({
                    iconCls: "icon-exclamation",
                    html: "Form is invalid",
                    title: "Error"
                });

                return false;
            }
            grid.store.insert(0, new Person(form.getForm().getValues()));
            //form.getForm().reset();
        };
    </script>
</head>
<body>
<center>
    <form id="Form1" runat="server">
        <table>
        <tr style='height:10px'>
        <td></td>
        </tr>
        </table>
        <ext:ResourceManager ID="ResourceManager1" runat="server" />
        <ext:Store ID="Store1" runat="server">
            <Model>
                <ext:Model ID="Model1" runat="server" IDProperty="Id" Name="Person" ClientIdProperty="PhantomId">
                   
                    <Fields>
                        <ext:ModelField Name="Id" Type="Int" UseNull="true" />
                        <ext:ModelField Name="part_id" UseNull="true" />
                        <ext:ModelField Name="lot" />
                        <ext:ModelField Name="meas_item" />
                        <ext:ModelField Name="plant" />
                        <ext:ModelField Name="tool" />
                        <ext:ModelField Name="trtm" />
                        <ext:ModelField Name="user" />
                        <ext:ModelField Name="updatetime" />
                        <ext:ModelField Name="comment" />
                        <ext:ModelField Name="mean" />
                        <ext:ModelField Name="std" />
                        <ext:ModelField Name="max" />
                        <ext:ModelField Name="min" />
                        <ext:ModelField Name="cp" />
                        <ext:ModelField Name="cpk" />
                    </Fields>

                    <Validations>
                        <ext:LengthValidation Field="meas_item" Min="1" />                    
                    </Validations>

                </ext:Model>
            </Model>            
        </ext:Store>
        

        <!-- Form Panel-->
        <ext:FormPanel 
            ID="UserForm" 
            runat="server"
            Frame="true"
            LabelAlign="Right"
            Title=""
            Icon="User"
            Width="750">
            <Items>
                
                <ext:TextField ID="tf_part" runat="server"
                    FieldLabel="Part(料號)"
                    Name="part_id"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
                
                <ext:TextField ID="tf_lot" runat="server"
                    FieldLabel="Lot"
                    Name="lot"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />

                <ext:TextField ID="tf_item" runat="server"
                    FieldLabel="Parameter"
                    Name="meas_item"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
                <ext:TextField ID="tf_plant" runat="server"
                    FieldLabel="Plant(廠別)"
                    Name="plant"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
                <ext:TextField ID="tf_machine" runat="server"
                    FieldLabel="Tool"
                    Name="tool"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />

                <ext:TextField ID="tf_trtm" runat="server"
                    FieldLabel="時間"
                    Name="trtm"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />

                <ext:TextField ID="tf_user" runat="server"
                    FieldLabel="更新人員"
                    Name="user"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    Hidden=true
                    />

                 <ext:TextField ID="tf_uTime" runat="server"
                    FieldLabel="更新時間"
                    Name="updatetime"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    Hidden=true
                    />

                <ext:TextField ID="tf_command" runat="server"
                    FieldLabel="Comment"
                    Name="comment"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
            </Items>            
            
            <Buttons>
                
                <ext:Button ID="Button2" 
                    runat="server"
                    Text="增加"
                    Icon="Add">
                    <Listeners>
                        <Click Handler="addRecord(#{UserForm}, #{GridPanel1});" />
                    </Listeners>
                </ext:Button>
                
            </Buttons>

        </ext:FormPanel>

        <!-- GridPanel -->
        <ext:GridPanel 
            ID="GridPanel1" 
            runat="server"
            Icon="Table"
            Frame="true"
            Title="Spec Setup"
            Height="350"
            Width="750"
            StoreID="Store1"
            StyleSpec="margin-top: 10px">
            <ColumnModel ID="ctl394">
                
                <Columns>
                    <ext:Column ID="Column1" runat="server" Text="UpdateTime" Width="120" DataIndex="updatetime"/>
                    <ext:Column ID="Column2" runat="server" Text="LOT" Width="100" DataIndex="lot"/>
                    <ext:Column ID="Column3" runat="server" Text="Machine" Width="100" DataIndex="tool"/>
                    <ext:Column ID="Column4" runat="server" Text="Comment" Width="310" DataIndex="comment" />
                    <ext:Column ID="Column5" runat="server" Text="User" Width="100" DataIndex="user" />

                    <ext:Column ID="Column6" runat="server" Text="ID"  Width="20" DataIndex="Id" Hidden="true"/>
                    <ext:Column ID="Column7" runat="server" Text="PartID" Width="40" DataIndex="part_id" Hidden="true"/>
                    <ext:Column ID="Column8" runat="server" Text="Param" Width="100" DataIndex="meas_item" Hidden="true"/>
                    <ext:Column ID="Column9" runat="server" Text="Plant" Width="40" DataIndex="plant" Hidden="true"/>       
                    <ext:Column ID="Column10" runat="server" Text="Trtm" Width="100" DataIndex="trtm" Hidden="true"/>
                    <ext:Column ID="Column11" runat="server" Text="Mean" Width="80" DataIndex="mean" Hidden="true" />
                    <ext:Column ID="Column12" runat="server" Text="Std" Width="100" DataIndex="std" Hidden="true" />
                    <ext:Column ID="Column13" runat="server" Text="Max" Width="320" DataIndex="max" Hidden="true" />
                    <ext:Column ID="Column14" runat="server" Text="Min" Width="80" DataIndex="min" Hidden="true" />
                    <ext:Column ID="Column15" runat="server" Text="CP" Width="100" DataIndex="cp" Hidden="true" />
                    <ext:Column ID="Column16" runat="server" Text="CPK" Width="320" DataIndex="cpk" Hidden="true" />
                </Columns>

            </ColumnModel>

            <Buttons>
                <ext:Button ID="Button6" runat="server" Text="確定" Icon="Disk">
                    <DirectEvents>
                        <Click OnEvent="SaveClick">
                            <ExtraParams>
                                <ext:Parameter Name="data" Value="#{Store1}.getChangedData({skipIdForNewRecords : false})" Mode="Raw" Encode="true" />
                            </ExtraParams>
                        </Click>
                    </DirectEvents>
                </ext:Button>
            </Buttons>

            <Plugins>
                <ext:CellEditing ID="CellEditing1" runat="server" />
            </Plugins>
        </ext:GridPanel>

    </form>
</center>
    <asp:Label ID="lab_part" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lab_lot" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lab_item" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lab_user" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="lab_fType" runat="server" Visible="False"></asp:Label>
</body>
</html>
