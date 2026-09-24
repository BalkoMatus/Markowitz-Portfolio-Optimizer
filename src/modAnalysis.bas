Attribute VB_Name = "modAnalysis"

Option Explicit

' ---- sheet names and layout (used by modDataImport, modChart and the form) ----
Public Const DATA_SHEET As String = "Data"
Public Const OUT_SHEET As String = "Optimizer"
Public Const MODEL_NAME As String = "Markowitz Portfolio Optimizer"

Public Const DATA_HDR_ROW As Long = 7      ' Data sheet: header row (Date, tickers)
Public Const DATA_COL As Long = 2          ' Data sheet: dates in column B, prices from column C

Public Const ROW_GMVP As Long = 19         ' Optimizer sheet: fixed result rows
Public Const ROW_TANG As Long = 20
Public Const ROW_ASSETS As Long = 24       ' first asset row
Public Const COL_CHARTDATA As Long = 19    ' column S, source data for the chart

' ---- colours (RGB values written out, Const cannot call RGB) ----
Public Const CLR_NAVY As Long = 4131840     ' RGB(0, 12, 63)
Public Const CLR_BLUE As Long = 13791538    ' RGB(50, 113, 210)
Public Const CLR_LIGHT As Long = 16773863   ' RGB(231, 242, 255)
Public Const CLR_INPUT As Long = 16711680   ' RGB(0, 0, 255)
Public Const CLR_GREY As Long = 7763574     ' RGB(118, 118, 118)
Public Const CLR_WHITE As Long = 16777215

Public Const FMT_PCT As String = "0.0%_);(0.0%)"
Public Const FMT_RATIO As String = "0.00_);(0.00)"
Public Const FMT_INT As String = "#,##0_)"
Public Const FMT_DATE As String = "yyyy-mm-dd_)"

' Read the Data sheet -> tickers(), prices() (rows=time, cols=assets).
Public Sub ReadPriceData(ByRef tickers() As String, ByRef prices() As Double)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long, n As Long, t As Long
    lastRow = ws.Cells(ws.Rows.Count, DATA_COL).End(xlUp).Row
    lastCol = ws.Cells(DATA_HDR_ROW, ws.Columns.Count).End(xlToLeft).Column
    n = lastCol - DATA_COL                ' first column is Date
    t = lastRow - DATA_HDR_ROW            ' rows below the header
    If n < 2 Or t < 3 Then Err.Raise vbObjectError + 30, , "Data sheet needs >= 2 assets and >= 3 price rows."

    ReDim tickers(0 To n - 1)
    For j = 1 To n: tickers(j - 1) = CStr(ws.Cells(DATA_HDR_ROW, DATA_COL + j).Value): Next j

    ReDim prices(1 To t, 1 To n)
    Dim v As Variant
    For i = 1 To t
        For j = 1 To n
            v = ws.Cells(DATA_HDR_ROW + i, DATA_COL + j).Value
            If Not IsNumeric(v) Or IsEmpty(v) Then Err.Raise vbObjectError + 31, , _
                "Non-numeric price at row " & (DATA_HDR_ROW + i) & ", col " & (DATA_COL + j) & "."
            prices(i, j) = CDbl(v)
        Next j
    Next i
End Sub

' Main entry: compute GMVP, tangency, frontier; write results + chart.
Public Sub RunOptimization(ByVal rf As Double, ByVal periodsPerYear As Double, _
                           ByVal nFrontier As Long, ByVal nCloud As Long)
    Dim tickers() As String, prices() As Double
    ReadPriceData tickers, prices

    Dim r() As Double, mu() As Double, Sigma() As Double
    r = PricesToReturns(prices, True)
    mu = MeanVector(r, periodsPerYear)
    Sigma = CovMatrix(r, periodsPerYear)

    Dim n As Long: n = UBound(mu, 1)
    Dim gmvp() As Double, tang() As Double
    gmvp = GMVPWeights(Sigma)
    tang = TangencyWeights(mu, Sigma, rf)
    Dim fc As FrontierConst: fc = FrontierConstants(mu, Sigma)

    Dim gRet As Double, gVol As Double, tRet As Double, tVol As Double
    gRet = PortReturn(gmvp, mu): gVol = PortStdDev(gmvp, Sigma)
    tRet = PortReturn(tang, mu): tVol = PortStdDev(tang, Sigma)

    Dim ws As Worksheet: Set ws = GetOrCreateSheet(OUT_SHEET)
    Application.ScreenUpdating = False
    ws.Cells.Clear
    DrawOptimizerFrame ws

    ' ===== Inputs =====
    ws.Range("C8").Value = rf
    ws.Range("C9").Value = periodsPerYear
    ws.Range("C10").Value = nFrontier
    ws.Range("C11").Value = nCloud
    ws.Range("C12").Value = n
    ws.Range("C13").Value = UBound(r, 1)
    WriteSampleDates ws.Range("C14"), ws.Range("C15")

    ' ===== Portfolios =====
    ws.Cells(ROW_GMVP, 3).Value = gRet
    ws.Cells(ROW_GMVP, 4).Value = gVol
    ws.Cells(ROW_GMVP, 5).Value = PortSharpe(gmvp, mu, Sigma, rf)
    ws.Cells(ROW_TANG, 3).Value = tRet
    ws.Cells(ROW_TANG, 4).Value = tVol
    ws.Cells(ROW_TANG, 5).Value = PortSharpe(tang, mu, Sigma, rf)

    ' ===== Assets and weights =====
    Dim j As Long, rw As Long
    For j = 1 To n
        rw = ROW_ASSETS + j - 1
        ws.Cells(rw, 2).Value = tickers(j - 1)
        ws.Cells(rw, 3).Value = mu(j, 1)
        ws.Cells(rw, 4).Value = Sqr(Sigma(j, j))
        ws.Cells(rw, 5).Value = gmvp(j, 1)
        ws.Cells(rw, 6).Value = tang(j, 1)
    Next j
    ws.Range(ws.Cells(ROW_ASSETS, 3), ws.Cells(ROW_ASSETS + n - 1, 6)).NumberFormat = FMT_PCT

    Dim totRow As Long: totRow = ROW_ASSETS + n
    ws.Cells(totRow, 2).Value = "Total"
    ws.Cells(totRow, 5).Formula = "=SUM(" & ws.Range(ws.Cells(ROW_ASSETS, 5), ws.Cells(totRow - 1, 5)).Address(False, False) & ")"
    ws.Cells(totRow, 6).Formula = "=SUM(" & ws.Range(ws.Cells(ROW_ASSETS, 6), ws.Cells(totRow - 1, 6)).Address(False, False) & ")"
    With ws.Range(ws.Cells(totRow, 2), ws.Cells(totRow, 6))
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    ws.Range(ws.Cells(totRow, 5), ws.Cells(totRow, 6)).NumberFormat = FMT_PCT

    ' ===== Frontier points =====
    Dim secRow As Long: secRow = totRow + 2
    DrawSection ws, secRow, 2, 6, "Frontier points"
    DrawColumnHeaders ws, secRow + 1, 2, Array("Point", "Return", "Volatility"), True

    Dim gmvpRet As Double, lo As Double, hi As Double, stepSize As Double, tr As Double
    Dim w() As Double, k As Long, frFirst As Long, frLast As Long, maxRet As Double
    gmvpRet = fc.b / fc.a
    maxRet = mu(1, 1)
    For j = 2 To n
        If mu(j, 1) > maxRet Then maxRet = mu(j, 1)
    Next j
    ' The range depends on the assets, not on the tangency portfolio, so the curve
    ' stays put when the risk-free rate changes. It is extended only if the tangency
    ' portfolio lies above it.
    lo = gmvpRet
    hi = gmvpRet + 1.5 * (maxRet - gmvpRet)
    If tRet > hi Then hi = tRet + 0.25 * (tRet - gmvpRet)
    If hi <= lo Then hi = lo + Abs(lo) + 0.1
    stepSize = (hi - lo) / (nFrontier - 1)
    frFirst = secRow + 2
    frLast = frFirst + nFrontier - 1
    For k = 0 To nFrontier - 1
        tr = lo + k * stepSize
        w = FrontierWeights(mu, Sigma, tr)
        ws.Cells(frFirst + k, 2).Value = PointLabel(k)
        ws.Cells(frFirst + k, 3).Value = tr
        ws.Cells(frFirst + k, 4).Value = PortStdDev(w, Sigma)
    Next k
    ws.Range(ws.Cells(frFirst, 2), ws.Cells(frLast, 2)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(frFirst, 3), ws.Cells(frLast, 4)).NumberFormat = FMT_PCT

    ' ===== Chart frame =====
    ' Axis limits come from the frontier range and the assets, so the frame does not
    ' move with the risk-free rate. Long-only random portfolios stay inside it.
    Dim xMax As Double, yMax As Double, yMin As Double
    xMax = ws.Cells(frLast, 4).Value
    yMax = hi
    yMin = 0#
    For j = 1 To n
        If Sqr(Sigma(j, j)) > xMax Then xMax = Sqr(Sigma(j, j))
        If mu(j, 1) > yMax Then yMax = mu(j, 1)
        If mu(j, 1) < yMin Then yMin = mu(j, 1)
    Next j
    xMax = RoundUpTo(xMax * 1.05, 0.05)
    yMax = RoundUpTo(yMax * 1.05, 0.1)
    If yMin < 0 Then yMin = -RoundUpTo(-yMin * 1.05, 0.1)

    ' ===== Chart data (columns S:T) =====
    ' Capital market line from (0, rf) through the tangency portfolio, cut at the chart frame.
    Dim c As Long: c = COL_CHARTDATA
    Dim cmlSlope As Double, cmlX As Double
    If tVol > 0 Then cmlSlope = (tRet - rf) / tVol
    cmlX = xMax
    If cmlSlope > 0 Then
        If rf + cmlSlope * cmlX > yMax Then cmlX = (yMax - rf) / cmlSlope
    End If
    ws.Cells(10, c).Value = 0#
    ws.Cells(10, c + 1).Value = rf
    ws.Cells(11, c).Value = cmlX
    ws.Cells(11, c + 1).Value = rf + cmlSlope * cmlX

    Dim vols() As Double, rets() As Double
    RandomCloud mu, Sigma, nCloud, vols, rets
    Dim cloud() As Variant: ReDim cloud(1 To nCloud, 1 To 2)
    For k = 1 To nCloud: cloud(k, 1) = vols(k): cloud(k, 2) = rets(k): Next k
    ws.Range(ws.Cells(15, c), ws.Cells(14 + nCloud, c + 1)).Value = cloud
    With ws.Range(ws.Cells(10, c), ws.Cells(14 + nCloud, c + 1))
        .NumberFormat = FMT_PCT
        .Font.Color = CLR_GREY
    End With

    DrawFrontierChart ws, ROW_ASSETS, totRow - 1, frFirst, frLast, 15, 14 + nCloud, xMax, yMin, yMax

    Application.ScreenUpdating = True
    MsgBox "Optimization complete. See the '" & OUT_SHEET & "' sheet.", vbInformation
End Sub

' Static part of the Optimizer sheet: widths, header, section titles, labels, formats.
Private Sub DrawOptimizerFrame(ByVal ws As Worksheet)
    ws.Columns(1).ColumnWidth = 2.7
    ws.Columns(2).ColumnWidth = 26
    ws.Range("C:F").ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 3
    ws.Range("H:Q").ColumnWidth = 9
    ws.Columns(18).ColumnWidth = 3
    ws.Range("S:T").ColumnWidth = 12

    DrawSheetHeader ws, "Optimization Results", COL_CHARTDATA + 1
    With ws.Range("B5")
        .Value = "All figures annualized unless stated. Weights can be negative (short positions)."
        .Font.Italic = True
        .Font.Size = 9
    End With

    ' Inputs
    Dim labels As Variant, k As Long
    labels = Array("Risk-free rate", "Periods per year", "Frontier points", "Random portfolios", _
                   "Assets", "Return observations", "First price date", "Last price date")
    DrawSection ws, 7, 2, 6, "Inputs"
    For k = 0 To 7: ws.Cells(8 + k, 2).Value = labels(k): Next k
    ws.Range("C8").NumberFormat = FMT_PCT
    ws.Range("C9:C13").NumberFormat = FMT_INT
    ws.Range("C14:C15").NumberFormat = FMT_DATE
    ws.Range("C14:C15").HorizontalAlignment = xlRight
    With ws.Range("B16")
        .Value = "Inputs used in the last run. To change them, click Open Optimizer on the Cover sheet."
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = CLR_GREY
    End With

    ' Portfolios
    DrawSection ws, 17, 2, 6, "Portfolios"
    DrawColumnHeaders ws, 18, 2, Array("", "Return", "Volatility", "Sharpe ratio"), True
    ws.Cells(ROW_GMVP, 2).Value = "Global minimum variance"
    ws.Cells(ROW_TANG, 2).Value = "Tangency (max Sharpe)"
    ws.Range(ws.Cells(ROW_GMVP, 3), ws.Cells(ROW_TANG, 4)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(ROW_GMVP, 5), ws.Cells(ROW_TANG, 5)).NumberFormat = FMT_RATIO

    ' Assets and weights
    DrawSection ws, 22, 2, 6, "Assets and weights"
    DrawColumnHeaders ws, 23, 2, Array("Ticker", "Return", "Volatility", "GMVP", "Tangency"), True

    ' Chart and chart data
    DrawSection ws, 7, 8, 17, "Efficient frontier"
    DrawSection ws, 7, COL_CHARTDATA, COL_CHARTDATA + 1, "Chart data"
    ws.Cells(8, COL_CHARTDATA).Value = "Capital market line"
    DrawColumnHeaders ws, 9, COL_CHARTDATA, Array("Volatility", "Return"), False
    ws.Cells(13, COL_CHARTDATA).Value = "Random portfolios"
    DrawColumnHeaders ws, 14, COL_CHARTDATA, Array("Volatility", "Return"), False
    ws.Range(ws.Cells(8, COL_CHARTDATA), ws.Cells(14, COL_CHARTDATA + 1)).Font.Color = CLR_GREY
End Sub

' "1 (GMVP)" for the first point, then plain numbers.
Private Function PointLabel(ByVal k As Long) As String
    PointLabel = CStr(k + 1)
    If k = 0 Then PointLabel = PointLabel & " (GMVP)"
End Function

' Round x up to the next multiple of stepSize (for axis limits).
Private Function RoundUpTo(ByVal x As Double, ByVal stepSize As Double) As Double
    RoundUpTo = -Int(-x / stepSize + 1E-09) * stepSize
End Function

' First and last date on the Data sheet.
Private Sub WriteSampleDates(ByVal firstCell As Range, ByVal lastCell As Range)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, DATA_COL).End(xlUp).Row
    firstCell.Value = ws.Cells(DATA_HDR_ROW + 1, DATA_COL).Value
    lastCell.Value = ws.Cells(lastRow, DATA_COL).Value
End Sub

' ---- shared formatting helpers ----

' Navy band with the model name (row 1) and the sheet title bar (row 3).
Public Sub DrawSheetHeader(ByVal ws As Worksheet, ByVal title As String, ByVal lastCol As Long)
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 9
    ws.Rows(3).RowHeight = 21
    ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).Interior.Color = CLR_NAVY
    With ws.Cells(1, 2)
        .Value = MODEL_NAME
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = CLR_WHITE
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(3, 2), ws.Cells(3, lastCol)).Interior.Color = CLR_LIGHT
    With ws.Cells(3, 2)
        .Value = title
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = CLR_BLUE
        .VerticalAlignment = xlCenter
    End With
End Sub

' Bold section title with a blue rule underneath.
Public Sub DrawSection(ByVal ws As Worksheet, ByVal rw As Long, ByVal c1 As Long, _
                       ByVal c2 As Long, ByVal caption As String)
    ws.Cells(rw, c1).Value = caption
    ws.Cells(rw, c1).Font.Bold = True
    With ws.Range(ws.Cells(rw, c1), ws.Cells(rw, c2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = CLR_BLUE
    End With
End Sub

' Bold column headers with a thin rule; numbers right-aligned.
Public Sub DrawColumnHeaders(ByVal ws As Worksheet, ByVal rw As Long, ByVal c1 As Long, _
                             ByVal headers As Variant, ByVal firstLeft As Boolean)
    Dim k As Long, nCols As Long
    nCols = UBound(headers) - LBound(headers) + 1
    For k = 0 To nCols - 1
        ws.Cells(rw, c1 + k).Value = headers(LBound(headers) + k)
    Next k
    With ws.Range(ws.Cells(rw, c1), ws.Cells(rw, c1 + nCols - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    If firstLeft Then ws.Cells(rw, c1).HorizontalAlignment = xlLeft
End Sub

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

Public Sub TestRunOptimization()
    RunOptimization 0.03, 252#, 25, 2000
End Sub

Public Sub ShowOptimizer()
    frmOptimizer.Show vbModeless
End Sub
