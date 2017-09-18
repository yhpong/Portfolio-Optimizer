Attribute VB_Name = "mPortOpt"
Option Explicit
'Requires: modMath


'"Portfolio Optimization in R", M. Andercut
'https://arxiv.org/pdf/1307.0450.pdf
'Find unconstrained efficient frontiers using Lagrange multiplier
'Input: x_r(), vector holding the returns of N stocks
'Input: x_covar(), NxN covariance matrix
'Output: mv(1 to r_bin, 1 to 2): first column is return, second column is variance
'Output: ws(1 to N, 1 to r_bin): weight of each stock
Sub EF_Lagrange(x_r() As Double, x_covar() As Double, mv() As Double, ws() As Double, Optional r_bin As Long = 20, Optional r_min As Variant, Optional r_max As Variant, Optional r_tgt As Variant, Optional var_out As Variant, Optional wr_tgt As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double, a11 As Double, a22 As Double, a12 As Double, d As Double
Dim A() As Double, f() As Double, g() As Double
    n = UBound(x_r, 1)
    A = modMath.Matrix_Inverse(x_covar)
    k = modMath.Identity_Chk(modMath.M_Dot(A, x_covar), tmp_x)
    If k = 0 Then
        Debug.Print "EF_Lagrange: covariance matrix is not invertible."
        Erase A
        Exit Sub
    End If
    a11 = 0: a12 = 0: a22 = 0
    For i = 1 To n
        For j = 1 To n
            a11 = a11 + A(i, j)
            a12 = a12 + x_r(i) * A(i, j)
            a22 = a22 + x_r(i) * A(i, j) * x_r(j)
        Next j
    Next i
    d = a11 * a22 - a12 * a12
    ReDim f(1 To n)
    ReDim g(1 To n)
    For i = 1 To n
        tmp_x = 0
        tmp_y = 0
        For j = 1 To n
            tmp_x = tmp_x + A(i, j) * x_r(j)
            tmp_y = tmp_y + A(i, j)
        Next j
        f(i) = (a22 * tmp_y - a12 * tmp_x) / d
        g(i) = (-a12 * tmp_y + a11 * tmp_x) / d
    Next i
    
    If IsMissing(r_min) = True Then r_min = a12 / a11
    If IsMissing(r_max) = True Then
        r_max = x_r(1)
        For i = 2 To n
            If x_r(i) > r_max Then r_max = x_r(i)
        Next i
    End If
    
    ReDim mv(1 To r_bin, 1 To 2)
    ReDim ws(1 To n, 1 To r_bin)
    For k = 1 To r_bin
        mv(k, 1) = r_min + (k - 1) * (r_max - r_min) / (r_bin - 1)
        mv(k, 2) = (a11 / d) * ((mv(k, 1) - a12 / a11) ^ 2) + 1# / a11
        For i = 1 To n
            ws(i, k) = f(i) + mv(k, 1) * g(i)
        Next i
    Next k
    
    If IsMissing(r_tgt) = False Then
        If IsMissing(var_out) = False Then
            var_out = (a11 / d) * ((r_tgt - a12 / a11) ^ 2) + 1# / a11
        End If
        If IsMissing(wr_tgt) = False Then
            ReDim wr_tgt(1 To n)
            For i = 1 To n
                wr_tgt(i) = f(i) + r_tgt * g(i)
            Next i
        End If
    End If
    
    Erase f, g, A
End Sub


'Return portfolio w() that gives return r_tgt and statisfies 0 < w_i <w_max
'Input: x_r(), N x 1 vector of expected return of each stock
'       x_covar(), N x N covariance matrix
'       r_tgt, target return
'       w_max, if left blank then no contraint is imposed on maximum weight of a single stock
'Output: r_out, return of optimized portolio, shoul dbe the same as r_tgt
'        var_out, varaince of optimized portfolio
'        w(), N x 1 vector holding the weight of each stock
Sub EF_InteriorPt_single(x_r() As Double, x_covar() As Double, r_tgt As Double, r_out As Double, var_out As Double, w() As Double, _
        Optional w_max As Variant, Optional w_min As Variant, Optional iter_max As Long = 1000, Optional tol As Double = 0.00000000001)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim q() As Double, QQ() As Double, A() As Double, B() As Double
    
    n = UBound(x_r, 1)
    
    ReDim q(1 To n)
    ReDim QQ(1 To n, 1 To n)
    ReDim A(1 To 2, 1 To n)
    ReDim B(1 To 2)
    ReDim w(1 To n)
    For i = 1 To n
        For j = 1 To n
            QQ(i, j) = 2 * x_covar(i, j)
        Next j
    Next i
    B(1) = r_tgt
    B(2) = 1
    For i = 1 To n
        A(1, i) = x_r(i)    'w^T.x_r = r_tgt
        A(2, i) = 1         'sum(w)=1
        w(i) = 1 / n        'initial guess is equal weight
    Next i
    
    If IsMissing(w_max) = True And IsMissing(w_min) = True Then
        Call mQPSolve.IPM(w, QQ, q, A, B, , , iter_max, tol)
    ElseIf IsMissing(w_max) = False And IsMissing(w_min) = True Then
        Call mQPSolve.IPM(w, QQ, q, A, B, w_max, , iter_max, tol)
    ElseIf IsMissing(w_max) = True And IsMissing(w_min) = False Then
        Call mQPSolve.IPM(w, QQ, q, A, B, , w_min, iter_max, tol)
    ElseIf IsMissing(w_max) = False And IsMissing(w_min) = False Then
        Call mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, iter_max, tol)
    End If

    ReDim Preserve w(1 To n)
    r_out = 0
    var_out = 0
    For i = 1 To n
        r_out = r_out + w(i) * x_r(i)
        For j = 1 To n
            var_out = var_out + w(i) * x_covar(i, j) * w(j)
        Next j
    Next i
    
    Erase q, QQ, A, B
    Application.StatusBar = False
End Sub


'Return portfolios along the efficient frontier that statisfy 0 < w_i <w_max
'Input: x_r(), N x 1 vector of expected return of each stock
'       x_covar(), N x N covariance matrix
'       w_max, if left blank then no contraint is imposed on maximum weight of a single stock
'       r_bin, number of portfolios to retrieved, default is 20
'       r_min, minimum target return, if left blank, then minimum variance portofiolo is returned
'       r_max, maximum target return, if left blank then it's chosen
'           to be the 3rd quartile of x_r().
'Output: mv(), r_bin x 2 array that holds the return and varaince of each portfolio
'        ws(), N x r_bin array that holds the stock weight in each portfolio
Sub EF_InteriorPt(x_r() As Double, x_covar() As Double, mv() As Double, ws() As Double, _
        Optional w_max As Variant, Optional w_min As Variant, Optional iter_max As Long = 1000, Optional tol As Double = 0.00000000001, Optional r_bin As Long = 20, Optional r_min As Variant, Optional r_max As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, k_start As Long
Dim r_tgt As Double, tmp_x As Double, tmp_y As Double, qrng As Double
Dim w() As Double, q() As Double, QQ() As Double, A() As Double, B() As Double

    n = UBound(x_r, 1)
    ReDim mv(1 To r_bin, 1 To 2)
    ReDim ws(1 To n, 1 To r_bin)
    
    ReDim q(1 To n)
    ReDim QQ(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            QQ(i, j) = 2 * x_covar(i, j)
        Next j
    Next i
    
    k_start = 1
    w = modMath.fQuartile(x_r)
    qrng = w(3) - w(1)
    If IsMissing(r_max) = True Then r_max = w(3) + 0.25 * (w(4) - w(3))
    
    'Solve for minimum variance portfolio
    If IsMissing(r_min) = True Then
        ReDim A(1 To 1, 1 To n)
        ReDim B(1 To 1)
        ReDim w(1 To n)
        B(1) = 1
        For i = 1 To n
            A(1, i) = 1
            w(i) = 1# / n
        Next i
    
        If IsMissing(w_max) = True And IsMissing(w_min) = True Then
            Call mQPSolve.IPM(w, QQ, q, A, B, , , iter_max, tol)
        ElseIf IsMissing(w_max) = False And IsMissing(w_min) = True Then
            Call mQPSolve.IPM(w, QQ, q, A, B, w_max, , iter_max, tol)
        ElseIf IsMissing(w_max) = True And IsMissing(w_min) = False Then
            Call mQPSolve.IPM(w, QQ, q, A, B, , w_min, iter_max, tol)
        ElseIf IsMissing(w_max) = False And IsMissing(w_min) = False Then
            Call mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, iter_max, tol)
        End If
        
        For i = 1 To n
            ws(i, 1) = w(i)
            mv(1, 1) = mv(1, 1) + w(i) * x_r(i)
            For j = 1 To n
                mv(1, 2) = mv(1, 2) + w(i) * x_covar(i, j) * w(j)
            Next j
        Next i
        r_min = mv(1, 1)
        k_start = 2
        If r_max <= r_min Then r_max = r_min + qrng
    Else
        ReDim w(1 To n)
        For i = 1 To n
            w(i) = 1# / n
        Next i
    End If
    
    ReDim A(1 To 2, 1 To n)
    ReDim B(1 To 2)
    B(2) = 1
    For i = 1 To n
        A(1, i) = x_r(i)
        A(2, i) = 1
    Next i
    
    For k = k_start To r_bin
        DoEvents
        Application.StatusBar = "EF_InteriorPt_Max: " & k & "/" & r_bin
        
        r_tgt = r_min + (k - 1) * (r_max - r_min) / (r_bin - 1)
        B(1) = r_tgt
    
    '    For i = 1 To n     'Use previous portoflio instead of re-initializing
    '        w(i) = 1 / n
    '        w(n + i) = w_max - w(i)
    '    Next i
        
        If IsMissing(w_max) = True And IsMissing(w_min) = True Then
            Call mQPSolve.IPM(w, QQ, q, A, B, , , iter_max, tol)
        ElseIf IsMissing(w_max) = False And IsMissing(w_min) = True Then
            Call mQPSolve.IPM(w, QQ, q, A, B, w_max, , iter_max, tol)
        ElseIf IsMissing(w_max) = True And IsMissing(w_min) = False Then
            Call mQPSolve.IPM(w, QQ, q, A, B, , w_min, iter_max, tol)
        ElseIf IsMissing(w_max) = False And IsMissing(w_min) = False Then
            Call mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, iter_max, tol)
        End If
        
        For i = 1 To n
            ws(i, k) = w(i)
            mv(k, 1) = mv(k, 1) + w(i) * x_r(i)
            For j = 1 To n
                mv(k, 2) = mv(k, 2) + w(i) * x_covar(i, j) * w(j)
            Next j
        Next i
    
    Next k
    
    Erase q, QQ, A, B, w
    Application.StatusBar = False
End Sub
