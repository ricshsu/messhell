<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CTF_SpecSetup.aspx.cs" Inherits="CTF_SpecSetup" %>
<%@ Import Namespace="System.Collections.Generic"%>
<%@ Register Assembly="Ext.Net" Namespace="Ext.Net" TagPrefix="ext" %>
<html>
<head id="Head1" runat="server">
    <title>Grid with AutoSave - Ext.NET Examples</title>
    <link href="/resources/css/examples.css" rel="stylesheet" />    
    <script>
        function backOpener() {
            window.opener.document.getElementById('but_Execute').click();
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
            form.getForm().reset();
        };
    </script>
</head>
<body>
<center>
    <form id="Form1" runat="server">
        <ext:ResourceManager ID="ResourceManager1" runat="server" />
        <ext:Store ID="Store1" runat="server">
            <Model>
                <ext:Model ID="Model1" runat="server" IDProperty="Id" Name="Person" ClientIdProperty="PhantomId">
                   
                    <Fields>
                        <ext:ModelField Name="Id" Type="Int" UseNull="true" />
                        <ext:ModelField Name="part_id" />
                        <ext:ModelField Name="meas_item" />
                        <ext:ModelField Name="USL" UseNull="true"/>
                        <ext:ModelField Name="CL"  UseNull="true"/>
                        <ext:ModelField Name="LSL" UseNull="true"/>
                        <ext:ModelField Name="SPEC_TYPE" />
                    </Fields>

                    <Validations>
                        <ext:LengthValidation Field="meas_item" Min="1" />                    
                    </Validations>

                </ext:Model>
            </Model>            
        </ext:Store>
          
        <ext:GridPanel 
            ID="GridPanel1" 
            runat="server"
            Icon="Table"
            Frame="true"
            Title="Spec Setup"
            Height="400"
            Width="500"
            StoreID="Store1"
            StyleSpec="margin-top: 10px">
            <ColumnModel ID="ctl394">
                
                <Columns>
                    
                    <ext:Column ID="Column1" runat="server" Text="ID" Width="40" DataIndex="Id" Hidden=true/>
                    <ext:Column ID="Column2" runat="server" Text="PartID" Flex="1" Width="40" DataIndex="part_id" />
                    <ext:Column ID="Column3" runat="server" Text="Meas_Item" Flex="1" DataIndex="meas_item" />
                    <ext:Column ID="Column4" runat="server" Text="USL" Flex="1" DataIndex="USL" />
                    <ext:Column ID="Column5" runat="server" Text="CL" Flex="1" DataIndex="CL" />
                    <ext:Column ID="Column6" runat="server" Text="LSL" Flex="1" DataIndex="LSL" />
                    <ext:Column ID="Column7" runat="server" Text="SPEC_TYPE" Flex="1" DataIndex="SPEC_TYPE" />
                    
                    <ext:CommandColumn ID="CommandColumn1" runat="server" Width="70">
                        <Commands>
                            <ext:GridCommand Text="取消" ToolTip-Text="Reject row changes" CommandName="reject" Icon="ArrowUndo" >
                            <ToolTip Text="Reject row changes"></ToolTip>
                            </ext:GridCommand>
                        </Commands>
                        <PrepareToolbar Handler="toolbar.items.get(0).setVisible(record.dirty);" />
                        <Listeners>
                            <Command Handler="record.reject();" />
                        </Listeners>
                    </ext:CommandColumn>

                </Columns>

            </ColumnModel>
           
            <TopBar>
                <ext:Toolbar ID="Toolbar1" runat="server">
                    <Items>                        
                        <ext:Button ID="Button5" runat="server" Text="刪除" Icon="Exclamation">
                            <Listeners>
                                <Click Handler="var selection = #{GridPanel1}.getView().getSelectionModel().getSelection()[0];
                                                if (selection) {
                                                    #{GridPanel1}.store.remove(selection);
                                                    #{UserForm}.getForm().reset();
                                                }" />
                            </Listeners>
                        </ext:Button>
                    </Items>
                </ext:Toolbar>
            </TopBar>
            
            <SelectionModel>
                <ext:RowSelectionModel ID="RowSelectionModel1" runat="server" Mode="Single">
                    <Listeners>                        
                        <Select Handler="#{UserForm}.getForm().loadRecord(record);" />
                    </Listeners>
                </ext:RowSelectionModel>
            </SelectionModel>

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

        <!-- Form Panel-->
        <ext:FormPanel 
            ID="UserForm" 
            runat="server"
            Frame="true"
            LabelAlign="Right"
            Title="Spec Setup -- All fields are required"
            Width="500">
            <Items>

               <ext:TextField ID="tf_part" runat="server"
                    FieldLabel="料號"
                    Name="part_id"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
                
                <ext:TextField ID="tf_item" runat="server"
                    FieldLabel="量測項目"
                    Name="meas_item"
                    AllowBlank="false"
                    AnchorHorizontal="100%"
                    />
                
                <ext:TextField ID="tf_usl" runat="server"
                    FieldLabel="上界"
                    Name="USL"
                    AllowBlank="true"
                    AnchorHorizontal="100%"
                    />

                <ext:TextField ID="tf_cl" runat="server"
                    FieldLabel="中心線"
                    Name="CL"
                    AllowBlank="true"
                    AnchorHorizontal="100%"
                    />
                
                <ext:TextField ID="tf_lsl" runat="server"
                    FieldLabel="下界"
                    Name="LSL"
                    AllowBlank="true"
                    AnchorHorizontal="100%"
                    />

                <ext:ComboBox ID="cb_side" runat="server" AllowBlank="false" Editable="false" Name="SPEC_TYPE" FieldLabel="雙邊 / 單邊界限" AnchorHorizontal="100%">
                <Items>
                      <ext:ListItem Text="雙邊" Value="T" />
                      <ext:ListItem Text="單邊(上)" Value="U" />
                      <ext:ListItem Text="單邊(下)" Value="L" />
                </Items>
                </ext:ComboBox>

            </Items>            
            
            <Buttons>

                <ext:Button ID="Button1" 
                    runat="server"
                    Text="更新"
                    Icon="BasketEdit">
                    <Listeners>
                        <Click Handler="updateRecord(#{UserForm});" />
                    </Listeners>
                </ext:Button>
                
                <ext:Button ID="Button2" 
                    runat="server"
                    Text="增加"
                    Icon="Add">
                    <Listeners>
                        <Click Handler="addRecord(#{UserForm}, #{GridPanel1});" />
                    </Listeners>
                </ext:Button>
                
                <ext:Button ID="Button3" 
                    runat="server"
                    Text="Reset">
                    <Listeners>
                        <Click Handler="#{UserForm}.getForm().reset();" />
                    </Listeners>
                </ext:Button>

            </Buttons>

        </ext:FormPanel>

        <table>
        <tr>
        <td>
        
            <asp:Button ID="TransferAll" runat="server" onclick="TransferAll_Click" 
                Text="All" />
        
        </td>
        </tr>
        </table>
    </form>
</center>
    <asp:Label ID="lab_part" runat="server" Visible="False" ViewStateMode="Enabled"></asp:Label>
</body>
</html>
