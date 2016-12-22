<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<%@ Register assembly="DundasWebChart" namespace="Dundas.Charting.WebControl" tagprefix="DCWC" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Button ID="Button1" runat="server" Text="Button" />
    
    </div>
    <DCWC:Chart ID="Chart1" runat="server">
        <Series>
            <DCWC:Series Name="Series1">
            </DCWC:Series>
            <DCWC:Series Name="Series2">
            </DCWC:Series>
        </Series>
        <ChartAreas>
            <DCWC:ChartArea Name="Default">
                <AxisX Reverse="True">
                    <LabelStyle Format="F0" />
                </AxisX>
                <AxisX2 Reverse="True">
                </AxisX2>
            </DCWC:ChartArea>
        </ChartAreas>
        <Legends>
            <DCWC:Legend Name="Default">
            </DCWC:Legend>
        </Legends>
        <Titles>
            <DCWC:Title Name="Title1" Text="jjlkjlkhj">
            </DCWC:Title>
        </Titles>
    </DCWC:Chart>
    </form>
</body>
</html>
