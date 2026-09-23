VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptimizer 
   Caption         =   "Markowitz Portfolio Optimizer"
   ClientHeight    =   6120
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8360.001
   StartUpPosition =   1  'CenterOwner
   TypeInfoVer     =   45
End
Attribute VB_Name = "frmOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    btnExample_Click          ' start pre-filled and ready to run
End Sub

Private Sub btnExample_Click()
    txtRf.Value = "3"
    txtPeriods.Value = "252"
    txtFrontier.Value = "25"
    txtCloud.Value = "2000"
    On Error GoTo eh
    If DataRowCount() < 3 Then
        LoadExampleData
        txtResults.Value = "Example parameters filled and built-in demo data loaded." & vbCrLf & _
                           "Click Calculate to run the optimization."
    Else
        txtResults.Value = "Example parameters filled. Using the data already on the Data sheet." & vbCrLf & _
                           "Click Calculate to run the optimization."
    End If
    Exit Sub
eh:
    MsgBox "Could not load example data: " & Err.Description, vbExclamation
End Sub

Private Sub btnImport_Click()
    On Error GoTo eh
    ImportFromCsvFiles
    txtResults.Value = "CSV data imported to the Data sheet (" & DataRowCount() & " rows)." & vbCrLf & _
                       "Click Calculate to run."
    Exit Sub
eh:
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnCalculate_Click()
    If Not NumOK(txtRf.Value) Then MsgBox "Risk-free rate must be a number, e.g. 3 for 3%.", vbExclamation: Exit Sub
    If Not NumOK(txtPeriods.Value) Then MsgBox "Periods per year must be a number, e.g. 252.", vbExclamation: Exit Sub
    If Not NumOK(txtFrontier.Value) Then MsgBox "Frontier points must be a whole number, e.g. 25.", vbExclamation: Exit Sub
    If Not NumOK(txtCloud.Value) Then MsgBox "Random portfolios must be a whole number, e.g. 2000.", vbExclamation: Exit Sub

    Dim rf As Double, ppy As Double, nFr As Long, nCl As Long
    rf = ParseNum(txtRf.Value) / 100#       ' user types percent -> decimal
    ppy = ParseNum(txtPeriods.Value)
    nFr = CLng(ParseNum(txtFrontier.Value))
    nCl = CLng(ParseNum(txtCloud.Value))

    If ppy <= 0 Then MsgBox "Periods per year must be positive.", vbExclamation: Exit Sub
    If nFr < 3 Then MsgBox "Use at least 3 frontier points.", vbExclamation: Exit Sub
    If nCl < 1 Then MsgBox "Random portfolios must be at least 1.", vbExclamation: Exit Sub
    If DataRowCount() < 3 Then MsgBox "No price data found. Use Import CSV or Example first.", vbExclamation: Exit Sub

    On Error GoTo eh
    Application.Cursor = xlWait
    RunOptimization rf, ppy, nFr, nCl
    Application.Cursor = xlDefault
    txtResults.Value = BuildSummary()
    Exit Sub
eh:
    Application.Cursor = xlDefault
    MsgBox "Calculation failed: " & Err.Description, vbExclamation
End Sub

Private Sub btnClear_Click()
    txtRf.Value = ""
    txtPeriods.Value = ""
    txtFrontier.Value = ""
    txtCloud.Value = ""
    txtResults.Value = ""
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' ---- helpers ----

Private Function NumOK(ByVal s As String) As Boolean
    s = Trim(s)
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", ".")
    Dim i As Long, c As String, dotSeen As Boolean
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c = "." Then
            If dotSeen Then Exit Function
            dotSeen = True
        ElseIf c = "-" Then
            If i <> 1 Then Exit Function
        ElseIf c < "0" Or c > "9" Then
            Exit Function
        End If
    Next i
    NumOK = True
End Function

Private Function ParseNum(ByVal s As String) As Double
    ParseNum = Val(Replace(Trim(s), ",", "."))
End Function

Private Function DataRowCount() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    DataRowCount = ws.Cells(ws.Rows.Count, DATA_COL).End(xlUp).Row - DATA_HDR_ROW
    If DataRowCount < 0 Then DataRowCount = 0
End Function

Private Function BuildSummary() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(OUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then BuildSummary = "Results written to the Optimizer sheet.": Exit Function

    Dim g As Long, t As Long, s As String
    g = ROW_GMVP: t = ROW_TANG
    s = "Results (annualized):" & vbCrLf & vbCrLf
    s = s & "GMVP" & vbCrLf
    s = s & "  Return:     " & Format(ws.Cells(g, 3).Value, "0.00%") & vbCrLf
    s = s & "  Volatility: " & Format(ws.Cells(g, 4).Value, "0.00%") & vbCrLf
    s = s & "  Sharpe:     " & Format(ws.Cells(g, 5).Value, "0.000") & vbCrLf & vbCrLf
    s = s & "Tangency (max Sharpe)" & vbCrLf
    s = s & "  Return:     " & Format(ws.Cells(t, 3).Value, "0.00%") & vbCrLf
    s = s & "  Volatility: " & Format(ws.Cells(t, 4).Value, "0.00%") & vbCrLf
    s = s & "  Sharpe:     " & Format(ws.Cells(t, 5).Value, "0.000") & vbCrLf & vbCrLf
    s = s & "See the Optimizer sheet for full weights and the chart."
    BuildSummary = s
End Function
