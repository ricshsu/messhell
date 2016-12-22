Imports Microsoft.VisualBasic
Imports Dundas.Charting.WebControl
Imports System.Drawing

Public Class ChartClass

    Public Shared Sub setCommonAttribute(ByRef ch As Dundas.Charting.WinControl.Chart)

        Dim AreaName As String = "AreaName"
        With ch

            .ChartAreas.Add(AreaName)
            'Tool bar
            .UI.Toolbar.Enabled = True
            .UI.ContextMenu.Enabled = True

            'Width & Height

            '.Width = TabPage.Height

            '.Height = TabPage.Width

            ' X線
            .ChartAreas(AreaName).AxisX.LabelStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
            '.ChartAreas(AreaName).AxisX.l()
            .ChartAreas(AreaName).AxisX.StartFromZero = False
            .ChartAreas(AreaName).AxisX.MajorGrid.Enabled = False
            .ChartAreas(AreaName).AxisX.MajorGrid.LineStyle = Dundas.Charting.WinControl.ChartDashStyle.Dash
            .ChartAreas(AreaName).AxisX.ScrollBar.ButtonColor = Color.LightGray

            ' Y線
            .ChartAreas(AreaName).AxisY.LabelStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
            .ChartAreas(AreaName).AxisY.MajorGrid.Enabled = True
            .ChartAreas(AreaName).AxisY.MajorGrid.LineStyle = Dundas.Charting.WinControl.ChartDashStyle.Dash
            .ChartAreas(AreaName).AxisY.ScrollBar.ButtonColor = Color.LightGray
            .ChartAreas(AreaName).CursorX.UserEnabled = True
            .ChartAreas(AreaName).CursorX.UserSelection = True
            .ChartAreas(AreaName).CursorY.UserEnabled = True
            .ChartAreas(AreaName).CursorY.UserSelection = True

            'Border
            .BorderLineColor = System.Drawing.Color.Silver
            .BorderLineStyle = ChartDashStyle.Solid

            'Legend
            .Legend.Alignment = System.Drawing.StringAlignment.Near
            .Legend.Docking = Dundas.Charting.WinControl.LegendDocking.Bottom
            .Legend.LegendStyle = Dundas.Charting.WinControl.LegendStyle.Table
            .Legend.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)

        End With

    End Sub

End Class

