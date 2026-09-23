Attribute VB_Name = "modDataImport"
Option Explicit

' DATA_SHEET, DATA_HDR_ROW and DATA_COL are declared in modAnalysis.

' Download daily closing prices for the given tickers from Stooq.
' lookbackDays = how much history to pull (365 = ~1 year).
Public Sub ImportFromStooq(ByRef tickers() As String, ByVal lookbackDays As Long)
    Dim n As Long, k As Long, maps() As Object, csv As String
    Dim d1 As Date, d2 As Date, s1 As String, s2 As String

    n = UBound(tickers) - LBound(tickers) + 1
    ReDim maps(1 To n)
    d2 = Date: d1 = d2 - lookbackDays
    s1 = Format(d1, "yyyymmdd"): s2 = Format(d2, "yyyymmdd")

    For k = 1 To n
        csv = HttpGet("https://stooq.com/q/d/l/?s=" & _
              LCase(Trim(tickers(LBound(tickers) + k - 1))) & _
              "&d1=" & s1 & "&d2=" & s2 & "&i=d")
        If InStr(1, csv, "Date,Open", vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 20, , "Stooq returned no usable data for '" & _
                tickers(LBound(tickers) + k - 1) & _
                "'. US stocks need a .us suffix (e.g. aapl.us). Check ticker / internet."
        End If
        Set maps(k) = ParseStooqCsv(csv)
    Next k

    WriteAlignedData tickers, maps, "Daily closing prices downloaded from Stooq."
    MsgBox "Import complete. " & n & " tickers written to the '" & DATA_SHEET & "' sheet.", vbInformation
End Sub

' Fill the Data sheet with realistic synthetic prices so the tool always
' has something to run on (one-factor model: r = alpha + beta * market + noise).
Public Sub LoadExampleData()
    Dim tickers() As String
    tickers = Split("AAPL,MSFT,JPM,XOM,KO", ",")
    Dim nDays As Long: nDays = 252
    Dim n As Long: n = UBound(tickers) - LBound(tickers) + 1

    Dim p0() As Double, beta() As Double, sige() As Double, alpha() As Double
    ReDim p0(1 To n): ReDim beta(1 To n): ReDim sige(1 To n): ReDim alpha(1 To n)
    p0(1) = 190: beta(1) = 1.2: sige(1) = 0.012: alpha(1) = 0.0006   ' AAPL
    p0(2) = 410: beta(2) = 1.1: sige(2) = 0.011: alpha(2) = 0.0005   ' MSFT
    p0(3) = 195: beta(3) = 1.3: sige(3) = 0.013: alpha(3) = 0.0003   ' JPM
    p0(4) = 115: beta(4) = 0.7: sige(4) = 0.014: alpha(4) = 0.0002   ' XOM
    p0(5) = 60#: beta(5) = 0.5: sige(5) = 0.008: alpha(5) = 0.0001   ' KO

    Dim muM As Double, sigM As Double
    muM = 0.0004: sigM = 0.009                ' daily market drift / vol

    Dim prices() As Double, j As Long
    ReDim prices(1 To nDays + 1, 1 To n)
    For j = 1 To n: prices(1, j) = p0(j): Next j

    Randomize
    Dim t As Long, mRet As Double, r As Double
    For t = 2 To nDays + 1
        mRet = muM + sigM * NormRand()
        For j = 1 To n
            r = alpha(j) + beta(j) * mRet + sige(j) * NormRand()
            prices(t, j) = prices(t - 1, j) * Exp(r)
        Next j
    Next t

    Dim ws As Worksheet: Set ws = GetOrCreateSheet(DATA_SHEET)
    ws.Cells.Clear
    ws.Cells(DATA_HDR_ROW, DATA_COL).Value = "Date"
    For j = 1 To n: ws.Cells(DATA_HDR_ROW, DATA_COL + j).Value = tickers(LBound(tickers) + j - 1): Next j

    Dim dt As Date, rIdx As Long
    dt = Date - Int((nDays + 1) * 7 / 5) - 2: rIdx = DATA_HDR_ROW + 1   ' ~253 weekdays ending near today
    For t = 1 To nDays + 1
        Do While Weekday(dt, vbMonday) > 5: dt = dt + 1: Loop   ' skip weekends
        ws.Cells(rIdx, DATA_COL).Value = dt
        For j = 1 To n: ws.Cells(rIdx, DATA_COL + j).Value = prices(t, j): Next j
        dt = dt + 1: rIdx = rIdx + 1
    Next t

    FormatDataSheet ws, n, nDays + 1, _
        "Simulated prices from a one-factor model (example data). Use Import CSV to load real prices."
    MsgBox "Example data loaded: " & n & " assets, " & nDays & " daily returns.", vbInformation
End Sub


Private Function HttpGet(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    If http.Status <> 200 Then _
        Err.Raise vbObjectError + 21, , "Download failed (HTTP " & http.Status & "). Check internet."
    HttpGet = http.responseText
End Function

' Parse Stooq/Yahoo daily CSV -> dictionary "date" -> closing price.
' Locale-safe: uses Val() (always '.' = decimal) instead of IsNumeric,
' which respects regional comma-decimal settings and would reject prices.
Private Function ParseStooqCsv(ByVal csv As String) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lines() As String, parts() As String, i As Long, ln As String
    csv = Replace(Replace(csv, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(csv, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 Then
            parts = Split(ln, ",")
            If UBound(parts) >= 4 Then
                ' a data row has a date like 2024-01-15 in column 0;
                ' the header ("Date") has no dash, so it is skipped.
                If InStr(parts(0), "-") > 0 Then
                    dict(parts(0)) = Val(parts(4))   ' Close, period-decimal safe
                End If
            End If
        End If
    Next i
    Set ParseStooqCsv = dict
End Function

' Intersect dates across tickers, sort ascending, write to Data sheet.
Private Sub WriteAlignedData(ByRef tickers() As String, ByRef maps() As Object, ByVal note As String)
    Dim n As Long, k As Long: n = UBound(maps)
    Dim common As Object: Set common = CreateObject("Scripting.Dictionary")
    Dim keyD As Variant, inAll As Boolean
    For Each keyD In maps(1).Keys
        inAll = True
        For k = 2 To n
            If Not maps(k).Exists(keyD) Then inAll = False: Exit For
        Next k
        If inAll Then common(keyD) = 1
    Next keyD
    If common.Count < 2 Then Err.Raise vbObjectError + 22, , "Not enough overlapping dates across tickers."

    Dim dates() As String, m As Long, a As Long, b As Long, tmp As String
    m = common.Count: ReDim dates(1 To m): a = 1
    For Each keyD In common.Keys: dates(a) = CStr(keyD): a = a + 1: Next keyD
    For a = 1 To m - 1
        For b = a + 1 To m
            If dates(b) < dates(a) Then tmp = dates(a): dates(a) = dates(b): dates(b) = tmp
        Next b
    Next a

    Dim ws As Worksheet: Set ws = GetOrCreateSheet(DATA_SHEET)
    ws.Cells.Clear
    ws.Cells(DATA_HDR_ROW, DATA_COL).Value = "Date"
    For k = 1 To n: ws.Cells(DATA_HDR_ROW, DATA_COL + k).Value = tickers(LBound(tickers) + k - 1): Next k

    Dim out() As Variant: ReDim out(1 To m, 1 To n + 1)
    For a = 1 To m
        out(a, 1) = IsoToDate(dates(a))
        For k = 1 To n: out(a, k + 1) = maps(k)(dates(a)): Next k
    Next a
    ws.Range(ws.Cells(DATA_HDR_ROW + 1, DATA_COL), ws.Cells(DATA_HDR_ROW + m, DATA_COL + n)).Value = out
    FormatDataSheet ws, n, m, note
End Sub

' "2024-01-15" -> Date (locale-safe). Anything else is left as text.
Private Function IsoToDate(ByVal s As String) As Variant
    If Len(s) >= 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        IsoToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    Else
        IsoToDate = s
    End If
End Function

' Header band, title, source note and number formats for the price table.
Private Sub FormatDataSheet(ByVal ws As Worksheet, ByVal nAssets As Long, _
                            ByVal nRows As Long, ByVal note As String)
    Dim lastCol As Long, lastRow As Long
    lastCol = DATA_COL + nAssets
    If lastCol < DATA_COL + 7 Then lastCol = DATA_COL + 7
    lastRow = DATA_HDR_ROW + nRows

    ws.Columns(1).ColumnWidth = 2.7
    ws.Columns(DATA_COL).ColumnWidth = 12
    ws.Range(ws.Columns(DATA_COL + 1), ws.Columns(lastCol)).ColumnWidth = 11
    DrawSheetHeader ws, "Price Data", lastCol
    With ws.Cells(5, DATA_COL)
        .Value = note
        .Font.Italic = True
        .Font.Size = 9
    End With

    With ws.Range(ws.Cells(DATA_HDR_ROW, DATA_COL), ws.Cells(DATA_HDR_ROW, DATA_COL + nAssets))
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = CLR_BLUE
    End With
    ws.Cells(DATA_HDR_ROW, DATA_COL).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_HDR_ROW + 1, DATA_COL), ws.Cells(lastRow, DATA_COL)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(DATA_HDR_ROW + 1, DATA_COL), ws.Cells(lastRow, DATA_COL)).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_HDR_ROW + 1, DATA_COL + 1), ws.Cells(lastRow, DATA_COL + nAssets)).NumberFormat = "#,##0.00_)"
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


Private Function NormRand() As Double
    Dim u1 As Double, u2 As Double
    Do: u1 = Rnd: Loop While u1 <= 0#
    u2 = Rnd
    NormRand = Sqr(-2# * Log(u1)) * Cos(6.28318530717959 * u2)
End Function


Public Sub TestExampleData()
    LoadExampleData
End Sub

Public Sub TestStooqImport()
    Dim tk() As String
    tk = Split("aapl.us,msft.us,jpm.us,xom.us,ko.us", ",")
    ImportFromStooq tk, 365
End Sub

Public Sub DebugStooq()
    Dim s As String
    On Error Resume Next
    s = HttpGet("https://stooq.com/q/d/l/?s=aapl.us&i=d")
    On Error GoTo 0
    Debug.Print "----- first 600 chars of Stooq response -----"
    Debug.Print Left$(s, 600)
    Debug.Print "----- end -----"
End Sub

Public Sub ImportFromCsvFiles()
    Dim fnames As Variant
    fnames = Application.GetOpenFilename( _
        "CSV files (*.csv),*.csv", , "Select one CSV per ticker", , True)
    If VarType(fnames) = vbBoolean Then Exit Sub      ' cancelled

    Dim n As Long, k As Long, fpath As String, csv As String
    n = UBound(fnames) - LBound(fnames) + 1
    Dim tickers() As String, maps() As Object
    ReDim tickers(0 To n - 1)
    ReDim maps(1 To n)

    For k = 1 To n
        fpath = CStr(fnames(LBound(fnames) + k - 1))
        csv = ReadTextFile(fpath)
        tickers(k - 1) = TickerFromPath(fpath)
        Set maps(k) = ParseStooqCsv(csv)
        If maps(k).Count = 0 Then _
            Err.Raise vbObjectError + 23, , "No price rows found in: " & fpath & _
            "  (expected a CSV with Date and Close columns)."
    Next k

    WriteAlignedData tickers, maps, "Daily closing prices imported from CSV files."
    MsgBox "Imported " & n & " CSV file(s) into the '" & DATA_SHEET & "' sheet.", vbInformation
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim ff As Integer, buf As String
    ff = FreeFile
    Open path For Binary Access Read As #ff
    buf = Space$(LOF(ff))
    Get #ff, 1, buf
    Close #ff
    ReadTextFile = buf
End Function

Private Function TickerFromPath(ByVal path As String) As String
    Dim nm As String, p As Long
    p = InStrRev(path, "\")
    nm = Mid$(path, p + 1)
    p = InStrRev(nm, ".")
    If p > 0 Then nm = Left$(nm, p - 1)
    nm = UCase$(nm)
    nm = Replace(nm, "_US_D", "")     ' Stooq names files like aapl_us_d.csv
    nm = Replace(nm, ".US", "")
    TickerFromPath = nm
End Function
