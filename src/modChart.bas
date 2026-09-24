Attribute VB_Name = "modChart"

Option Explicit

Private Const CLR_CLOUD As Long = 15127237     ' RGB(197, 210, 230)
Private Const CLR_FRONTIER As Long = 6567967   ' RGB(31, 56, 100)
Private Const CLR_ASSET As Long = 5855577      ' RGB(89, 89, 89)
Private Const CLR_TANG As Long = 1137349       ' RGB(197, 90, 17)
Private Const CLR_CML As Long = 8355711        ' RGB(127, 127, 127)
Private Const CLR_GRID As Long = 14277081      ' RGB(217, 217, 217)

' Scatter chart on the Optimizer sheet (H8:Q31). Row arguments point to the
' asset table, the frontier table and the random-portfolio block in S:T.
Public Sub DrawFrontierChart(ByVal ws As Worksheet, _
        ByVal assetFirst As Long, ByVal assetLast As Long, _
        ByVal frFirst As Long, ByVal frLast As Long, _
        ByVal cloudFirst As Long, ByVal cloudLast As Long, _
        ByVal xMax As Double, ByVal yMin As Double, ByVal yMax As Double)
    Dim oc As ChartObject, ch As Chart, s As Series, p As Long, cd As Long
    cd = COL_CHARTDATA

    For Each oc In ws.ChartObjects: oc.Delete: Next oc
    Set oc = ws.ChartObjects.Add(Left:=ws.Range("H8").Left, Top:=ws.Range("H8").Top + 4, _
                                 Width:=ws.Range("H8:Q8").Width, Height:=ws.Range("H8:H31").Height)
    oc.Placement = xlFreeFloating
    Set ch = oc.Chart
    ch.ChartType = xlXYScatter
    ch.HasTitle = False

    ' 1) Random portfolio cloud
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Random portfolios"
    s.ChartType = xlXYScatter
    s.XValues = ws.Range(ws.Cells(cloudFirst, cd), ws.Cells(cloudLast, cd))
    s.Values = ws.Range(ws.Cells(cloudFirst, cd + 1), ws.Cells(cloudLast, cd + 1))
    s.MarkerStyle = xlMarkerStyleCircle
    s.MarkerSize = 3
    s.MarkerBackgroundColor = CLR_CLOUD
    s.MarkerForegroundColor = CLR_CLOUD

    ' 2) Efficient frontier (Volatility in D, Return in C)
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Efficient frontier"
    s.ChartType = xlXYScatterLinesNoMarkers
    s.XValues = ws.Range(ws.Cells(frFirst, 4), ws.Cells(frLast, 4))
    s.Values = ws.Range(ws.Cells(frFirst, 3), ws.Cells(frLast, 3))
    s.Format.Line.Visible = msoTrue
    s.Format.Line.ForeColor.RGB = CLR_FRONTIER
    s.Format.Line.Weight = 2.25

    ' 3) Capital market line
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Capital market line"
    s.ChartType = xlXYScatterLinesNoMarkers
    s.XValues = ws.Range(ws.Cells(10, cd), ws.Cells(11, cd))
    s.Values = ws.Range(ws.Cells(10, cd + 1), ws.Cells(11, cd + 1))
    s.Format.Line.Visible = msoTrue
    s.Format.Line.ForeColor.RGB = CLR_CML
    s.Format.Line.Weight = 1.25
    s.Format.Line.DashStyle = msoLineDash

    ' 4) Individual assets, labelled with the ticker
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Assets"
    s.ChartType = xlXYScatter
    s.XValues = ws.Range(ws.Cells(assetFirst, 4), ws.Cells(assetLast, 4))
    s.Values = ws.Range(ws.Cells(assetFirst, 3), ws.Cells(assetLast, 3))
    s.MarkerStyle = xlMarkerStyleSquare
    s.MarkerSize = 6
    s.MarkerBackgroundColor = CLR_ASSET
    s.MarkerForegroundColor = CLR_ASSET
    s.HasDataLabels = True
    For p = 1 To assetLast - assetFirst + 1
        s.Points(p).DataLabel.Text = CStr(ws.Cells(assetFirst + p - 1, 2).Value)
        s.Points(p).DataLabel.Position = xlLabelPositionRight
    Next p

    ' 5) GMVP
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "GMVP"
    s.ChartType = xlXYScatter
    s.XValues = ws.Range(ws.Cells(ROW_GMVP, 4), ws.Cells(ROW_GMVP, 4))
    s.Values = ws.Range(ws.Cells(ROW_GMVP, 3), ws.Cells(ROW_GMVP, 3))
    s.MarkerStyle = xlMarkerStyleDiamond
    s.MarkerSize = 9
    s.MarkerBackgroundColor = CLR_BLUE
    s.MarkerForegroundColor = CLR_BLUE

    ' 6) Tangency
    Set s = ch.SeriesCollection.NewSeries
    s.Name = "Tangency"
    s.ChartType = xlXYScatter
    s.XValues = ws.Range(ws.Cells(ROW_TANG, 4), ws.Cells(ROW_TANG, 4))
    s.Values = ws.Range(ws.Cells(ROW_TANG, 3), ws.Cells(ROW_TANG, 3))
    s.MarkerStyle = xlMarkerStyleTriangle
    s.MarkerSize = 9
    s.MarkerBackgroundColor = CLR_TANG
    s.MarkerForegroundColor = CLR_TANG

    ' Axes, legend, fonts
    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Volatility (annual)"
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = False
        .MinimumScale = 0
        .MaximumScale = xMax
        If xMax > 0.5 Then .MajorUnit = 0.1 Else .MajorUnit = 0.05
    End With
    With ch.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Expected return (annual)"
        .TickLabels.NumberFormat = "0%"
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = CLR_GRID
        .MinimumScale = yMin
        .MaximumScale = yMax
        .MajorUnit = 0.1
    End With
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.ChartArea.Format.Line.Visible = msoFalse
    ch.ChartArea.Font.Name = "Arial"
    ch.ChartArea.Font.Size = 9
    ch.Axes(xlCategory).AxisTitle.Font.Bold = False
    ch.Axes(xlValue).AxisTitle.Font.Bold = False
End Sub
