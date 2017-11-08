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
Dim A() As Double, f() As Double, g() As Double, tmp_vec() As Double
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
        tmp_vec = modMath.fQuartile(x_r)
        r_max = tmp_vec(3)
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
Sub EF_InteriorPt_single(x_r() As Double, x_covar() As Double, r_tgt As Double, _
        r_out As Double, var_out As Double, w() As Double, _
        Optional w_max As Variant = Null, Optional w_min As Variant = Null, _
        Optional x_sector As Variant = Null, Optional sector_list As Variant = Null, _
        Optional sector_w_max As Variant = Null, Optional sector_w_min As Variant = Null, _
        Optional x_ctry As Variant = Null, Optional ctry_list As Variant = Null, _
        Optional ctry_w_max As Variant = Null, Optional ctry_w_min As Variant = Null, _
        Optional iter_max As Long = 1000, Optional tol As Double = 0.0000000001, _
        Optional w_init As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim q() As Double, QQ() As Double, A() As Double, B() As Double
Dim C As Variant, c_max As Variant, c_min As Variant
Dim tmpBool As Boolean

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
    Next i
    'Constraints on sector/country wgts
    If IsNull(x_sector) = False Or IsNull(x_ctry) = False Then
        Call Construct_Constraints(C, c_max, c_min, _
            x_sector, sector_list, sector_w_max, sector_w_min, _
            x_ctry, ctry_list, ctry_w_max, ctry_w_min)
    Else
        C = Null
        c_max = Null
        c_min = Null
    End If
    'Initial guess of weight vector
    If IsMissing(w_init) Then
        Call Init_wgt(n, w, w_max, w_min, C, c_max, c_min)
    Else
        w = w_init
        Call Trim_Wgt(w, w_max, w_min)
    End If
    
    'Run optimizer
    tmpBool = mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, C, c_max, c_min, iter_max, tol)

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
Sub EF_InteriorPt(x_r As Variant, x_covar As Variant, mv() As Double, ws() As Double, _
        Optional w_max As Variant = Null, Optional w_min As Variant = Null, _
        Optional x_sector As Variant = Null, Optional sector_list As Variant = Null, _
        Optional sector_w_max As Variant = Null, Optional sector_w_min As Variant = Null, _
        Optional x_ctry As Variant = Null, Optional ctry_list As Variant = Null, _
        Optional ctry_w_max As Variant = Null, Optional ctry_w_min As Variant = Null, _
        Optional iter_max As Long = 1000, Optional tol As Double = 0.0000000001, _
        Optional r_bin As Long = 20, Optional r_min As Variant = Null, Optional r_max As Variant = Null, _
        Optional w_init As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, k_start As Long
Dim r_tgt As Double, tmp_x As Double, tmp_y As Double, qrng As Double
Dim w() As Double, q() As Double, QQ() As Double, A() As Double, B() As Double
Dim C As Variant, c_max As Variant, c_min As Variant
Dim tmpBool As Boolean

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
    
    'Constraints on sector/country wgts
    If IsNull(x_sector) = False Or IsNull(x_ctry) = False Then
        Call Construct_Constraints(C, c_max, c_min, _
            x_sector, sector_list, sector_w_max, sector_w_min, _
            x_ctry, ctry_list, ctry_w_max, ctry_w_min)
    Else
        C = Null
        c_max = Null
        c_min = Null
    End If
    
    'Use 3rd-quartile return if r_max is not supplied
    k_start = 1
    w = modMath.fQuartile(x_r)
    qrng = w(3) - w(1)
    If IsNull(r_max) Then r_max = w(3) '+ 0.25 * (w(3) - w(2))
    
    'Initial guess of weights vector
    If IsMissing(w_init) Then
        Call Init_wgt(n, w, w_max, w_min, C, c_max, c_min)
    Else
        w = w_init
        Call Trim_Wgt(w, w_max, w_min)
    End If
    
    'Solve for minimum variance portfolio if r_min is not supplied
    If IsNull(r_min) Then
        DoEvents
        Application.StatusBar = "EF_InteriorPt: " & 1 & "/" & r_bin
        
        ReDim A(1 To 1, 1 To n)
        ReDim B(1 To 1)
        B(1) = 1
        For i = 1 To n
            A(1, i) = 1
        Next i
        
        tmpBool = mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, C, c_max, c_min, iter_max, tol)

        If tmpBool = True Then
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
            Debug.Print "mPortOpt: EF_InteriorPt: MVP not found. Try giving better initial weights."
            Application.StatusBar = False
            Exit Sub
        End If

    End If
    
    'Solve for portfolio further up the efficient frontier
    ReDim A(1 To 2, 1 To n)
    ReDim B(1 To 2)
    B(2) = 1
    For i = 1 To n
        A(1, i) = x_r(i)
        A(2, i) = 1
    Next i
    For k = k_start To r_bin
        DoEvents
        Application.StatusBar = "EF_InteriorPt: " & k & "/" & r_bin

        r_tgt = r_min + (k - 1) * (r_max - r_min) / (r_bin - 1)
        B(1) = r_tgt

        'Use previous portoflio instead of re-initializing
        tmpBool = mQPSolve.IPM(w, QQ, q, A, B, w_max, w_min, C, c_max, c_min, iter_max, tol)
        If tmpBool = True Then
            For i = 1 To n
                ws(i, k) = w(i)
                mv(k, 1) = mv(k, 1) + w(i) * x_r(i)
                For j = 1 To n
                    mv(k, 2) = mv(k, 2) + w(i) * x_covar(i, j) * w(j)
                Next j
            Next i
        Else
            Debug.Print "mPortOpt: EF_InteriorPt: solution not found when r(" & k & ")=" & r_tgt
        End If

    Next k
    
    Erase q, QQ, A, B, w
    If IsNull(C) = False Then Erase C, c_max, c_min
    Application.StatusBar = False
End Sub


Sub Calc_Mean_Variance(ret As Double, var As Double, w() As Double, x_r() As Double, x_covar() As Double)
Dim i As Long, j As Long, n As Long
    n = UBound(w, 1)
    ret = 0
    var = 0
    For i = 1 To n
        ret = ret + w(i) * x_r(i)
        For j = 1 To n
            var = var + w(i) * x_covar(i, j) * w(j)
        Next j
    Next i
End Sub


Private Sub Construct_Constraints(C As Variant, c_max As Variant, c_min As Variant, _
    Optional x_sector As Variant = Null, Optional sector_list As Variant = Null, _
    Optional sector_max As Variant = Null, Optional sector_min As Variant = Null, _
    Optional x_ctry As Variant = Null, Optional ctry_list As Variant = Null, _
    Optional ctry_max As Variant = Null, Optional ctry_min As Variant = Null)
Dim i As Long, j As Long, n As Long, n_sector As Long, n_ctry As Long
Dim v_tmp As Variant

    If IsNull(x_sector) = False And IsNull(x_ctry) = True Then
        n = UBound(x_sector)
        n_sector = UBound(sector_list)
        ReDim C(1 To n_sector, 1 To n)
        For i = 1 To n
            v_tmp = x_sector(i)
            For j = 1 To n_sector
                If v_tmp = sector_list(j) Then
                    C(j, i) = 1
                    Exit For
                End If
            Next j
        Next i
        c_max = sector_max
        c_min = sector_min

    ElseIf IsNull(x_sector) = True And IsNull(x_ctry) = False Then
    
        n = UBound(x_ctry)
        n_ctry = UBound(ctry_list)
        ReDim C(1 To n_ctry, 1 To n)
        For i = 1 To n
            v_tmp = x_ctry(i)
            For j = 1 To n_ctry
                If v_tmp = ctry_list(j) Then
                    C(j, i) = 1
                    Exit For
                End If
            Next j
        Next i
        c_max = ctry_max
        c_min = ctry_min

    ElseIf IsNull(x_sector) = False And IsNull(x_ctry) = False Then
    
        n = UBound(x_sector)
        n_sector = UBound(sector_list)
        n_ctry = UBound(ctry_list)
        ReDim C(1 To n_sector + n_ctry, 1 To n)
        For i = 1 To n
            v_tmp = x_sector(i)
            For j = 1 To n_sector
                If v_tmp = sector_list(j) Then
                    C(j, i) = 1
                    Exit For
                End If
            Next j
            v_tmp = x_ctry(i)
            For j = 1 To n_ctry
                If v_tmp = ctry_list(j) Then
                    C(n_sector + j, i) = 1
                    Exit For
                End If
            Next j
        Next i
        c_max = sector_max
        c_min = sector_min
        ReDim Preserve c_max(1 To n_sector + n_ctry)
        ReDim Preserve c_min(1 To n_sector + n_ctry)
        For j = 1 To n_ctry
            c_max(n_sector + j) = ctry_max(j)
            c_min(n_sector + j) = ctry_min(j)
        Next j

    End If
    
End Sub

'Trim weights that are out of bound
Private Sub Trim_Wgt(w() As Double, Optional w_max As Variant = Null, Optional w_min As Variant = Null)
Dim i As Long, n As Long
    n = UBound(w, 1)
    If IsNull(w_min) = False Then
        For i = 1 To n
            If w(i) <= w_min Then w(i) = w_min + 0.0000000001
        Next i
    End If
    If IsNull(w_max) = False Then
        For i = 1 To n
            If w(i) >= w_max Then w(i) = w_max - 0.0000000001
        Next i
    End If
End Sub

'Initialize weights to "roughly" statisfy inequality constraints
Private Sub Init_wgt(n As Long, w() As Double, _
        Optional w_max As Variant = Null, Optional w_min As Variant = Null, _
        Optional C As Variant = Null, Optional cmax As Variant, Optional cmin As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n_c As Long
Dim cw() As Double, wmax() As Double, wmin() As Double, ic() As Long, dw() As Double, iArr() As Long
Dim tmp_x As Double
    
    ReDim w(1 To n)
    For i = 1 To n
        w(i) = 1# / n
    Next i

    If IsNull(C) = False Then
        n_c = UBound(C, 1)
        
        For k = 1 To 20
        
            'check if current wgts suffice
            m = 0
            cw = modMath.M_Dot(C, w)
            For j = 1 To n_c
                If cw(j) < cmin(j) Or cw(j) > cmax(j) Then
                    m = 1
                    Exit For
                End If
            Next j
            If m = 0 Then Exit For
            If m = 1 Then
                ReDim ic(1 To n_c)
                ReDim dw(1 To n)
                ReDim iArr(1 To n)
                For j = 1 To n_c
                    If cw(j) < cmin(j) Or cw(j) > cmax(j) Then
                        tmp_x = (cmax(j) + cmin(j)) / 2 - cw(j)
                        For i = 1 To n
                            If C(j, i) = 1 Then ic(j) = ic(j) + 1
                        Next i
                        tmp_x = tmp_x / ic(j)
                        For i = 1 To n
                            If C(j, i) = 1 Then 'w(i) = w(i) + tmp_x
                                dw(i) = dw(i) + tmp_x
                                iArr(i) = iArr(i) + 1
                            End If
                        Next i
                    End If
                Next j
                For i = 1 To n
                    If iArr(i) > 0 Then w(i) = w(i) + dw(i) / iArr(i)
                Next i
            End If
            Erase cw, ic, dw, iArr
        
        Next k
    End If

    Call Trim_Wgt(w, w_max, w_min)
    
End Sub
