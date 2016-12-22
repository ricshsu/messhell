<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CTF_Transfer.aspx.cs" Inherits="CTF_Transfer" %>
<%@ Import Namespace="System.Threading" %>
<%@ Register Assembly="Ext.Net" Namespace="Ext.Net" TagPrefix="ext" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CTF 資料轉置介面</title>
<link href="~/css/Button.css" rel="Stylesheet" />
    <link href="~/css/PUPPY.css" rel="Stylesheet" />
    <link href="~/css/TabContainer.css" rel="Stylesheet" />
    <link href="~/css/DW.css" rel="Stylesheet" />
    <link href="/resources/css/examples.css" rel="stylesheet" />
    <style>
        .status {
            color:#555;
        }
        .x-progress.left-align .x-progress-text {
            text-align:left;
        }
        .x-progress.custom {
            height:19px;
            border:1px solid #686868;
            padding:0 2px;
        }
        .ext-strict .x-progress.custom {
            height:17px;
        }       
        .custom .x-progress-bar {
            height:17px;
            border: none;
            background:transparent url(custom-bar.gif) repeat-x !important;
            border-top:1px solid #BEBEBE;
        }

        .ext-strict .custom .x-progress-bar {
            height: 15px;
        }
        .login
        {
            width: 610px;
             height: 510px;
        }
        .accountInfo
        {
            width: 600px;
            height: 450px;
        }
        .style1
        {
            height: 50px;
        }
    </style>
     
</head>
<body>
    <form id="form1" runat="server">
    <center>
    <div class="accountInfo">
    <fieldset class="login" style='border:3px;'><legend>資料轉置</legend>
    <BR />
    <table style="height: 440px">                
    <tr>
    <td align=left class="style1" align=center valign=top>
    <ext:ResourceManager ID="ResourceManager1" runat="server" />
    <!-- -->
    <ext:Button ID="ShowProgress1" runat="server" Text="開始" OnDirectClick="StartLongAction" />
    <br />
    <ext:ProgressBar ID="Progress1" runat="server" Width="600" Hidden="true" />
    <ext:TaskManager ID="TaskManager1" runat="server">
    <Tasks>
    <ext:Task TaskID="longactionprogress" Interval="100" AutoRun="false" 
     OnStart="#{ShowProgress1}.setDisabled(true);"
     OnStop="#{ShowProgress1}.setDisabled(false);">
     <DirectEvents>
       <Update OnEvent="RefreshProgress" />
     </DirectEvents>                    
    </ext:Task>
    </Tasks>
    </ext:TaskManager>
    <!-- -->
    </td>
    </tr>

    <tr>
    <td align=left align=center valign=top>
    <ext:TextArea ID="ta_result" runat="server" Height="400px" Width="600px" EnableViewState="true"></ext:TextArea>
    </td>
    </tr>

    <tr>
    <td>
    <asp:Button ID="but_Transfer" runat="server" Text="Transfer" 
            onclick="but_Transfer_Click" />
    </td>
    </tr>

    </table>
    </fieldset>
    </div>
    </center>
    </form>
</body>
</html>
