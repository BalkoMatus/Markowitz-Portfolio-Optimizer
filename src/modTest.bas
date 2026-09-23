Attribute VB_Name = "modTest"
Option Explicit

Public Sub TestEngine()
    Dim p() As Double, d As Variant, i As Long, j As Long
    d = Array( _
        Array(100#, 50#, 200#), _
        Array(101#, 49.5, 202#), _
        Array(102.5, 50.2, 199#), _
        Array(101.8, 51#, 203#), _
        Array(103#, 50.8, 205#), _
        Array(104.2, 51.5, 204#))
    ReDim p(1 To 6, 1 To 3)
    For i = 1 To 6
        For j = 1 To 3: p(i, j) = d(i - 1)(j - 1): Next j
    Next i

    Dim r() As Double, mu() As Double, Sigma() As Double
    r = PricesToReturns(p, True)
    mu = MeanVector(r, 252#)
    Sigma = CovMatrix(r, 252#)

    Dim gmvp() As Double, tang() As Double
    gmvp = GMVPWeights(Sigma)
    tang = TangencyWeights(mu, Sigma, 0.02)

    Debug.Print "--- Annualized expected returns ---"
    For j = 1 To 3: Debug.Print "  Asset"; j; "="; Format(mu(j, 1), "0.0000"): Next j
    Debug.Print "--- GMVP ---"
    For j = 1 To 3: Debug.Print "  w"; j; "="; Format(gmvp(j, 1), "0.0000"): Next j
    Debug.Print "  ret="; Format(PortReturn(gmvp, mu), "0.0000"); "  std="; Format(PortStdDev(gmvp, Sigma), "0.0000")
    Debug.Print "--- Tangency (rf=2%) ---"
    For j = 1 To 3: Debug.Print "  w"; j; "="; Format(tang(j, 1), "0.0000"): Next j
    Debug.Print "  ret="; Format(PortReturn(tang, mu), "0.0000"); _
                "  std="; Format(PortStdDev(tang, Sigma), "0.0000"); _
                "  Sharpe="; Format(PortSharpe(tang, mu, Sigma, 0.02), "0.0000")
End Sub




