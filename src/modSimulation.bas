Attribute VB_Name = "modSimulation"

Option Explicit

' Generate nPts random long-only portfolios (weights >= 0, sum = 1,
' uniform over the simplex) and return their (volatility, return).
Public Sub RandomCloud(ByRef mu() As Double, ByRef Sigma() As Double, _
                       ByVal nPts As Long, ByRef vols() As Double, ByRef rets() As Double)
    Dim nAssets As Long: nAssets = UBound(mu, 1)
    ReDim vols(1 To nPts)
    ReDim rets(1 To nPts)
    Dim w() As Double: ReDim w(1 To nAssets, 1 To 1)
    Dim p As Long, j As Long, s As Double, u As Double
    Randomize
    For p = 1 To nPts
        s = 0#
        For j = 1 To nAssets
            u = Rnd
            If u < 1E-09 Then u = 1E-09
            w(j, 1) = -Log(u)              ' exponential -> uniform Dirichlet
            s = s + w(j, 1)
        Next j
        For j = 1 To nAssets: w(j, 1) = w(j, 1) / s: Next j
        rets(p) = PortReturn(w, mu)
        vols(p) = PortStdDev(w, Sigma)
    Next p
End Sub

