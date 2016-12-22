Imports Dundas.Charting.WebControl
Partial Class _Default
    Inherits System.Web.UI.Page

    '...
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim instance As AxisDataView
        Dim type As ScrollType
        'instance.Scroll(type)

        ' Add a series.
        Chart1.Series.Add("Series1")
        ' Set the chart types of the series.
        'Chart1.Series("Series1").Type = SeriesChartType.Bar
        ' Set the chart type of the series. Note that by default a series is plotted
        ' in the "Default" chart area, and if this chart area does not exist then the series will
        ' use the first available chart area.
        Chart1.Series("Series1").ChartArea = "Default"
        ' Add data to Series1. Note that we are only setting the Y values,
        ' so X values will automatically be 0, and data will be plotted using the
        ' index of data points in the data point collection.
        'Chart1.ImageType = ChartImageType.Emf

        Chart1.ChartAreas("Default").CursorX.UserEnabled = True
        Chart1.ChartAreas("Default").CursorX.UserSelection = True

        Chart1.ChartAreas("Default").CursorY.UserEnabled = True
        Chart1.ChartAreas("Default").CursorY.UserSelection = True

        Chart1.ChartAreas("Default").CursorX.SelectionColor = Drawing.Color.PaleGoldenrod

        Chart1.ChartAreas("Default").AxisY.ScrollBar.ChartArea.AxisY.View.Zoomable = True

        ' Set the timeout parameter.
        'Chart1.ChartScrollTimeout = 1000

        If Not Me.IsPostBack Then
            For i = 0 To 400
                Chart1.Series("Series1").Points.AddY(i + 1)
            Next
        End If


        'Chart1.ChartAreas("Default").AxisX.View.SizeType = DateTimeIntervalType.Auto
        'Chart1.ChartAreas("Default").AxisX.View.Size = 500
        'Chart1.ChartAreas("Default").AxisX.View.MinSize = 200
        'Chart1.ChartAreas("Default").AxisX.View.MinSizeType = DateTimeIntervalType.Auto
        'Chart1.ChartAreas("Default").AxisX.View.Position = 12
        Chart1.ChartAreas("Default").AxisX.ScrollBar.Enabled = True

        Chart1.ChartAreas("Default").AxisX.View.Size = 9
        Chart1.ChartAreas("Default").AxisX.View.MinSize = 9
        Chart1.ChartAreas("Default").AxisX.View.SmallScrollMinSize = 10
        Chart1.ChartAreas("Default").AxisX.View.SmallScrollSize = 10
        Chart1.ChartAreas("Default").AxisX.ScrollBar.ChartArea.AxisX.Interval = 1
        'Chart1.ChartAreas("Default").AxisX.View.Zoom(0.0, 9.0)
        Chart1.ChartAreas("Default").AxisX.View.Zoomable = True
        Chart1.ChartAreas("Default").AxisY.View.Zoomable = True

        'Chart1.ChartAreas("Default").AxisX.View.IsZoomed

        'Chart1.ChartAreas("Default").AxisX.Margin = True

        'Chart1.ChartAreas("Default").AxisX.View.SmallScrollSize = 2
        'Chart1.ChartAreas("Default").AxisX.StartFromZero = True
        'Chart1.ChartAreas("Default").AxisX.Maximum = 1000
        'Chart1.ChartAreas("Default").CursorX.Interval = 0

        'Chart1.ChartAreas("Default").AxisX.View.Zoomable = True
        Dim bb As Boolean = Chart1.ChartAreas("Default").AxisX.View.IsZoomed


        Chart1.ChartAreas("Default").AlignWithChartArea = "Default"
        Chart1.ChartAreas("Default").AlignOrientation = AreaAlignOrientation.Vertical
        Chart1.ChartAreas("Default").AlignType = AreaAlignType.All



    End Sub


End Class
