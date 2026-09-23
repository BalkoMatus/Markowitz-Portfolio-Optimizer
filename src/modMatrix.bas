Attribute VB_Name = "modMatrix"
Option Explicit

Public Function MatMult(ByRef a() As Double, ByRef b() As Double) As Double()
    Dim p As Long, q As Long, q2 As Long, r As Long
    Dim i As Long, j As Long, k As Long, s As Double
    Dim c() As Double

    p = UBound(a, 1): q = UBound(a, 2)
    q2 = UBound(b, 1): r = UBound(b, 2)
    If q <> q2 Then Err.Raise vbObjectError + 1, , "MatMult: inner dimensions do not match."

    ReDim c(1 To p, 1 To r)
    For i = 1 To p
        For j = 1 To r
            s = 0#
            For k = 1 To q
                s = s + a(i, k) * b(k, j)
            Next k
            c(i, j) = s
        Next j
    Next i
    MatMult = c
End Function

Public Function MatTranspose(ByRef a() As Double) As Double()
    Dim p As Long, q As Long, i As Long, j As Long, t() As Double
    p = UBound(a, 1): q = UBound(a, 2)
    ReDim t(1 To q, 1 To p)
    For i = 1 To p
        For j = 1 To q
            t(j, i) = a(i, j)
        Next j
    Next i
    MatTranspose = t
End Function

Public Function MatInverse(ByRef a() As Double) As Double()
    Dim n As Long, i As Long, j As Long, k As Long, p As Long
    Dim factor As Double, pivot As Double, big As Double, tmp As Double
    Dim m() As Double, Inv() As Double

    n = UBound(a, 1)
    If UBound(a, 2) <> n Then Err.Raise vbObjectError + 2, , "MatInverse: matrix is not square."

    ReDim m(1 To n, 1 To n)
    ReDim Inv(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            m(i, j) = a(i, j)
            Inv(i, j) = IIf(i = j, 1#, 0#)
        Next j
    Next i

    For k = 1 To n
        big = Abs(m(k, k)): p = k
        For i = k + 1 To n
            If Abs(m(i, k)) > big Then big = Abs(m(i, k)): p = i
        Next i
        If big < 1E-12 Then Err.Raise vbObjectError + 3, , "MatInverse: matrix is singular (cannot invert)."
        If p <> k Then
            For j = 1 To n
                tmp = m(k, j): m(k, j) = m(p, j): m(p, j) = tmp
                tmp = Inv(k, j): Inv(k, j) = Inv(p, j): Inv(p, j) = tmp
            Next j
        End If
        pivot = m(k, k)
        For j = 1 To n
            m(k, j) = m(k, j) / pivot
            Inv(k, j) = Inv(k, j) / pivot
        Next j
        For i = 1 To n
            If i <> k Then
                factor = m(i, k)
                If factor <> 0# Then
                    For j = 1 To n
                        m(i, j) = m(i, j) - factor * m(k, j)
                        Inv(i, j) = Inv(i, j) - factor * Inv(k, j)
                    Next j
                End If
            End If
        Next i
    Next k

    MatInverse = Inv
End Function


Public Function VecDot(ByRef a() As Double, ByRef b() As Double) As Double
    Dim n As Long, i As Long, s As Double
    n = UBound(a, 1)
    s = 0#
    For i = 1 To n
        s = s + a(i, 1) * b(i, 1)
    Next i
    VecDot = s
End Function
