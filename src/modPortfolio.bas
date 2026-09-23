Attribute VB_Name = "modPortfolio"
Option Explicit

Public Type FrontierConst
    a As Double   ' 1' Sinv 1
    b As Double   ' 1' Sinv mu
    c As Double   ' mu' Sinv mu
    d As Double   ' A*C - B^2
End Type

' Price matrix P (rows=time, cols=assets) -> returns ((rows-1) x assets)
Public Function PricesToReturns(ByRef p() As Double, ByVal useLog As Boolean) As Double()
    Dim t As Long, n As Long, i As Long, j As Long, r() As Double
    t = UBound(p, 1): n = UBound(p, 2)
    If t < 2 Then Err.Raise vbObjectError + 10, , "PricesToReturns: need at least 2 price rows."
    ReDim r(1 To t - 1, 1 To n)
    For i = 2 To t
        For j = 1 To n
            If p(i - 1, j) <= 0 Or p(i, j) <= 0 Then _
                Err.Raise vbObjectError + 11, , "PricesToReturns: non-positive price found."
            If useLog Then
                r(i - 1, j) = Log(p(i, j) / p(i - 1, j))
            Else
                r(i - 1, j) = p(i, j) / p(i - 1, j) - 1#
            End If
        Next j
    Next i
    PricesToReturns = r
End Function

' Per-asset mean return, annualized by periodsPerYear -> mu (n x 1)
Public Function MeanVector(ByRef r() As Double, ByVal periodsPerYear As Double) As Double()
    Dim t As Long, n As Long, i As Long, j As Long, s As Double, mu() As Double
    t = UBound(r, 1): n = UBound(r, 2)
    ReDim mu(1 To n, 1 To 1)
    For j = 1 To n
        s = 0#
        For i = 1 To t: s = s + r(i, j): Next i
        mu(j, 1) = (s / t) * periodsPerYear
    Next j
    MeanVector = mu
End Function

' Sample covariance matrix (n x n), annualized by periodsPerYear
Public Function CovMatrix(ByRef r() As Double, ByVal periodsPerYear As Double) As Double()
    Dim t As Long, n As Long, i As Long, a As Long, b As Long
    Dim mean() As Double, s As Double, cov() As Double
    t = UBound(r, 1): n = UBound(r, 2)
    If t < 2 Then Err.Raise vbObjectError + 12, , "CovMatrix: need at least 2 return rows."
    ReDim mean(1 To n)
    For a = 1 To n
        s = 0#
        For i = 1 To t: s = s + r(i, a): Next i
        mean(a) = s / t
    Next a
    ReDim cov(1 To n, 1 To n)
    For a = 1 To n
        For b = a To n
            s = 0#
            For i = 1 To t
                s = s + (r(i, a) - mean(a)) * (r(i, b) - mean(b))
            Next i
            s = (s / (t - 1)) * periodsPerYear   ' sample covariance, annualized
            cov(a, b) = s
            cov(b, a) = s
        Next b
    Next a
    CovMatrix = cov
End Function

' Ones column vector (n x 1)
Private Function Ones(ByVal n As Long) As Double()
    Dim v() As Double, i As Long
    ReDim v(1 To n, 1 To 1)
    For i = 1 To n: v(i, 1) = 1#: Next i
    Ones = v
End Function

' A, B, C, D frontier scalars from mu and Sigma
Public Function FrontierConstants(ByRef mu() As Double, ByRef Sigma() As Double) As FrontierConst
    Dim n As Long, SigInv() As Double, ones1() As Double
    Dim SinvOne() As Double, SinvMu() As Double, fc As FrontierConst
    n = UBound(mu, 1)
    SigInv = MatInverse(Sigma)
    ones1 = Ones(n)
    SinvOne = MatMult(SigInv, ones1)
    SinvMu = MatMult(SigInv, mu)
    fc.a = VecDot(ones1, SinvOne)
    fc.b = VecDot(ones1, SinvMu)
    fc.c = VecDot(mu, SinvMu)
    fc.d = fc.a * fc.c - fc.b * fc.b
    FrontierConstants = fc
End Function

' Global Minimum Variance Portfolio weights: Sinv 1 / A
Public Function GMVPWeights(ByRef Sigma() As Double) As Double()
    Dim n As Long, i As Long, SigInv() As Double, ones1() As Double
    Dim SinvOne() As Double, a As Double, w() As Double
    n = UBound(Sigma, 1)
    SigInv = MatInverse(Sigma)
    ones1 = Ones(n)
    SinvOne = MatMult(SigInv, ones1)
    a = VecDot(ones1, SinvOne)
    ReDim w(1 To n, 1 To 1)
    For i = 1 To n: w(i, 1) = SinvOne(i, 1) / a: Next i
    GMVPWeights = w
End Function

' Tangency (max-Sharpe) weights: Sinv(mu - rf) / (1' Sinv(mu - rf))
Public Function TangencyWeights(ByRef mu() As Double, ByRef Sigma() As Double, ByVal rf As Double) As Double()
    Dim n As Long, i As Long, SigInv() As Double, excess() As Double
    Dim SinvE() As Double, ones1() As Double, denom As Double, w() As Double
    n = UBound(mu, 1)
    SigInv = MatInverse(Sigma)
    ReDim excess(1 To n, 1 To 1)
    For i = 1 To n: excess(i, 1) = mu(i, 1) - rf: Next i
    SinvE = MatMult(SigInv, excess)
    ones1 = Ones(n)
    denom = VecDot(ones1, SinvE)
    If Abs(denom) < 1E-12 Then _
        Err.Raise vbObjectError + 13, , "TangencyWeights: undefined (denominator ~ 0)."
    ReDim w(1 To n, 1 To 1)
    For i = 1 To n: w(i, 1) = SinvE(i, 1) / denom: Next i
    TangencyWeights = w
End Function

' Minimum-variance weights for a TARGET return: w = g + h*targetRet
Public Function FrontierWeights(ByRef mu() As Double, ByRef Sigma() As Double, ByVal targetRet As Double) As Double()
    Dim n As Long, i As Long, SigInv() As Double, ones1() As Double
    Dim SinvOne() As Double, SinvMu() As Double
    Dim a As Double, b As Double, c As Double, d As Double, g As Double, h As Double, w() As Double
    n = UBound(mu, 1)
    SigInv = MatInverse(Sigma)
    ones1 = Ones(n)
    SinvOne = MatMult(SigInv, ones1)
    SinvMu = MatMult(SigInv, mu)
    a = VecDot(ones1, SinvOne)
    b = VecDot(ones1, SinvMu)
    c = VecDot(mu, SinvMu)
    d = a * c - b * b
    ReDim w(1 To n, 1 To 1)
    For i = 1 To n
        g = (c * SinvOne(i, 1) - b * SinvMu(i, 1)) / d
        h = (a * SinvMu(i, 1) - b * SinvOne(i, 1)) / d
        w(i, 1) = g + h * targetRet
    Next i
    FrontierWeights = w
End Function

' --- Portfolio statistics given weights ---
Public Function PortReturn(ByRef w() As Double, ByRef mu() As Double) As Double
    PortReturn = VecDot(w, mu)
End Function

Public Function PortVariance(ByRef w() As Double, ByRef Sigma() As Double) As Double
    Dim Sw() As Double
    Sw = MatMult(Sigma, w)
    PortVariance = VecDot(w, Sw)
End Function

Public Function PortStdDev(ByRef w() As Double, ByRef Sigma() As Double) As Double
    Dim v As Double
    v = PortVariance(w, Sigma)
    If v < 0 Then v = 0
    PortStdDev = Sqr(v)
End Function

Public Function PortSharpe(ByRef w() As Double, ByRef mu() As Double, ByRef Sigma() As Double, ByVal rf As Double) As Double
    Dim sd As Double
    sd = PortStdDev(w, Sigma)
    If sd <= 0 Then PortSharpe = 0: Exit Function
    PortSharpe = (PortReturn(w, mu) - rf) / sd
End Function

