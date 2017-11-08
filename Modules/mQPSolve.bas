Attribute VB_Name = "mQPSolve"
Option Explicit

'**********************************************
'***** Quadratic Optimizer
'**********************************************

'=== Interior Point Method
'Dr. Abebe Geletu
'https://www.tu-ilmenau.de/fileadmin/media/simulation/Lehre/Vorlesungsskripte/Lecture_materials_Abebe/QPs_with_IPM_and_ASM.pdf
'Solve for x() that minimize { (1/2) *( x^T QQ x) +q^T x }
's.t. Ax=B and x>0, with optional constraints x=[x_min, x_max], Cx=[c_min, c_max]
'x() is input as the initial guess, modified on output as the optimized solution
'Function returns TRUE if convergence is achieved before iter_max
Function IPM(x() As Double, QQ() As Double, q() As Double, A() As Double, B() As Double, _
        Optional x_max As Variant = Null, Optional x_min As Variant = Null, _
        Optional C As Variant = Null, Optional c_max As Variant = Null, Optional c_min As Variant = Null, _
        Optional iter_max As Long = 1000, Optional tol As Double = 0.0000000001) As Boolean
Dim i As Long, j As Long, k As Long, m As Long, iterate As Long
Dim n As Long, n_c As Long, n_ieq As Long, nn As Double, last_row As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_u As Double, tmp_v As Double
Dim B_norm As Double, q_norm As Double
Dim chk1 As Double, chk2 As Double, chk3 As Double
Dim f() As Double, Jacob() As Double, f1() As Double
Dim lambda() As Double, d() As Double
Dim mu As Double, mu_prev As Double, alpha As Double, alpha_z As Double, sigma As Double
Dim isConstraint As Boolean, isMax As Boolean, isMin As Boolean
Dim xmax() As Double, s() As Double, z_s() As Double, lambda_s() As Double
Dim f1s() As Double, f2s() As Double, f3s() As Double, ds() As Double, f11() As Double
Dim xmin() As Double, r() As Double, z_r() As Double, lambda_r() As Double
Dim f1r() As Double, f2r() As Double, f3r() As Double, dr() As Double
Dim d2s() As Double, d2r() As Double, d3s() As Double, d3r() As Double
Dim isMaxC As Boolean, isMinC As Boolean
Dim cx() As Double, tmp_vec() As Double, tmp_vec2() As Double
Dim lambda_u() As Double, z_u() As Double, s_u() As Double, cmax() As Double
Dim lambda_l() As Double, z_l() As Double, s_l() As Double, cmin() As Double
Dim f1u() As Double, f2u() As Double, f3u() As Double, d1u() As Double, d2u() As Double, d3u() As Double
Dim f1l() As Double, f2l() As Double, f3l() As Double, d1l() As Double, d2l() As Double, d3l() As Double

    IPM = False
    
    n = UBound(QQ, 1)   'dimension of x
    n_c = UBound(B, 1)  'number of linear equality constraints Ax=b
    nn = n              'dimesion of x & slack variables
    ReDim lambda(1 To n_c) 'Lagrange Multiplier for linear constraints Ax=B
    For i = 1 To n_c  'Better way to initialize lambda?
        lambda(i) = 0
    Next i
    sigma = 0.75        'step size to shrink mu
    mu = 0              'Lagrange multiplier for log barrier, better way to initialize?

    'Check if there is an upper constraint on x
    isMax = False
    If IsNull(x_max) = False Then
        isMax = True
        nn = nn + n
        ReDim xmax(1 To n)
        If IsArray(x_max) Then
            xmax = x_max
        Else
            For i = 1 To n
                xmax(i) = x_max
            Next i
        End If
        For i = 1 To n
            If x(i) > xmax(i) Then
                Debug.Print "mQPSolve: IPM: Init Fail: violates maximum constraints."
            End If
        Next i
        ReDim s(1 To n)
        ReDim z_s(1 To n)
        ReDim lambda_s(1 To n)
        For i = 1 To n
            s(i) = max2(xmax(i) - x(i), 1# / n)  'slack variable x+s=x_max, s>=0
            z_s(i) = s(i)         'slack variable for log barrier on s, z_s>=0 z_s(i) = mu / s(i)
            mu = mu + s(i) * z_s(i)
            lambda_s(i) = 0       'Lagrange Multiplier for x+s=x_max
        Next i
    End If

    'Check if there is a lower constraint on x
    isMin = False
    If IsNull(x_min) = False Then
        isMin = True
        nn = nn + n
        ReDim xmin(1 To n)
        If IsArray(x_min) Then
            xmin = x_min
        Else
            For i = 1 To n
                xmin(i) = x_min
            Next i
        End If
        For i = 1 To n
            If x(i) < xmin(i) Then
                Debug.Print "mQPSolve: IPM: Init Fail: violates minimum constraints."
                Exit Function
            End If
        Next i
        ReDim r(1 To n)
        ReDim z_r(1 To n)
        ReDim lambda_r(1 To n)
        For i = 1 To n
            r(i) = max2(x(i) - xmin(i), 1# / n)   'slack variable x-r=x_min, r>=0
            z_r(i) = r(i)         'slack variable for log barrier on r, z_r>0 z_r(i) = mu / r(i)
            mu = mu + r(i) * z_r(i)
            lambda_r(i) = 0   'Lagrange Multiplier for x-r=x_min
        Next i
    End If

    'Check if there are linear inequality constraints: L<Cx<U
    isMaxC = False
    isMinC = False
    If IsNull(C) = False Then
        n_ieq = UBound(C, 1) 'number of linear inequality constraints
        cx = modMath.M_Dot(C, x)
        If IsNull(c_max) = False Then
            isMaxC = True
            nn = nn + n_ieq
            ReDim cmax(1 To n_ieq)
            If IsArray(c_max) Then
                cmax = c_max
            Else
                For i = 1 To n_ieq
                    cmax(i) = c_max
                Next i
            End If
            For i = 1 To n_ieq
                If cx(i) > cmax(i) Then
                    Debug.Print "mQPSolve: IPM: Init Fail: violates maximum constraints Cx<U."
                End If
            Next i
            ReDim s_u(1 To n_ieq)
            ReDim z_u(1 To n_ieq)
            ReDim lambda_u(1 To n_ieq)
            For i = 1 To n_ieq
                s_u(i) = max2(cmax(i) - cx(i), 1# / n_ieq)
                z_u(i) = s_u(i)         'z_u(i) = mu / s_u(i)
                mu = mu + s_u(i) * z_u(i)
                lambda_u(i) = 0
            Next i
        End If

        If IsNull(c_min) = False Then
            isMinC = True
            nn = nn + n_ieq
            ReDim cmin(1 To n_ieq)
            If IsArray(c_min) Then
                cmin = c_min
            Else
                For i = 1 To n_ieq
                    cmin(i) = c_min
                Next i
            End If
            For i = 1 To n_ieq
                If cx(i) < cmin(i) Then
                    Debug.Print "mQPSolve: IPM: Init Fail: violates minimum constraints Cx>L."
                End If
            Next i
            ReDim s_l(1 To n_ieq)
            ReDim z_l(1 To n_ieq)
            ReDim lambda_l(1 To n_ieq)
            For i = 1 To n_ieq
                s_l(i) = max2(cx(i) - cmin(i), 1# / n_ieq)
                z_l(i) = s_l(i)     'z_l(i) = mu / s_l(i)
                mu = mu + s_l(i) * z_l(i)
                lambda_l(i) = 0
            Next i
        End If

        Erase cx
    End If
    
    If nn > n Then mu = mu / (nn - n)   'Initialize mu to (sz/n)
    
    If isMax = True Or isMin = True Or isMaxC = True Or isMinC = True Then
        isConstraint = True
    Else
        isConstraint = False
    End If
    
    'Constant part of the Jacobian
    ReDim Jacob(1 To n + n_c, 1 To n + n_c)
    For i = 1 To n
        For j = 1 To n_c
            Jacob(i, n + j) = A(j, i)
            Jacob(n + j, i) = A(j, i)
        Next j
    Next i

    '=======================================================
    'If there are no constraints then returns exact solution
    '=======================================================
    If isConstraint = False Then
        For i = 1 To n
            Jacob(i, i) = -QQ(i, i)
            For j = i + 1 To n
                Jacob(i, j) = -QQ(i, j)
                Jacob(j, i) = -QQ(j, i)
            Next j
        Next i
        f = q
        ReDim Preserve f(1 To n + n_c)
        For i = 1 To n_c
            f(n + i) = B(i)
        Next i
        x = modMath.Solve_Linear_LDL(Jacob, f)
        ReDim Preserve x(1 To n)
        Erase Jacob, f
        IPM = True
        Exit Function
    End If
    '=======================================================
    '=======================================================
    
    'Pre-allocate memory
    ReDim f1(1 To n)
    ReDim f(1 To n + n_c)
    If isMax = True Then Call Init_Vec(n, f1s, f2s, f3s, ds, d2s, d3s)
    If isMin = True Then Call Init_Vec(n, f1r, f2r, f3r, dr, d2r, d3r)
    If isMaxC = True Then Call Init_Vec(n_ieq, f1u, f2u, f3u, d1u, d2u, d3u)
    If isMinC = True Then Call Init_Vec(n_ieq, f1l, f2l, f3l, d1l, d2l, d3l)
        
    'Start Iteration
    iterate = 0
    Do
        For i = 1 To n
            Jacob(i, i) = -QQ(i, i)
            For j = i + 1 To n
                Jacob(i, j) = -QQ(i, j)
                'Jacob(j, i) = -QQ(j, i)
            Next j
        Next i
        chk1 = 0
        chk2 = 0
        For i = 1 To n
            tmp_x = 0
            For j = 1 To n
                tmp_x = tmp_x + QQ(i, j) * x(j)
            Next j
            tmp_y = 0
            For j = 1 To n_c
                tmp_y = tmp_y + A(j, i) * lambda(j)
            Next j
            f(i) = (tmp_x - tmp_y + q(i))
            f1(i) = f(i)
        Next i
        For j = 1 To n_c
            tmp_x = 0
            For i = 1 To n
                tmp_x = tmp_x + A(j, i) * x(i)
            Next i
            f(n + j) = -(tmp_x - B(j))
            chk2 = max2(chk2, Abs(f(n + j)))
        Next j
        If isMax = True Then 'Maximum constraints on individual x(i)
            Call Calc_Grad_Max(n, sigma * mu, x, s, z_s, lambda_s, xmax, f1s, f2s, f3s, Jacob, False)
            chk1 = max2(chk1, MaxNorm_vec(f1s))
            chk2 = max2(chk2, MaxNorm_vec(f2s))
            For i = 1 To n
                f(i) = f(i) - (lambda_s(i) + f1s(i) + (f2s(i) * z_s(i) + f3s(i)) / s(i))
                f1(i) = f1(i) - lambda_s(i)
            Next i
        End If
        If isMin = True Then 'Minimum constraints on individual x(i)
            Call Calc_Grad_Max(n, sigma * mu, x, r, z_r, lambda_r, xmin, f1r, f2r, f3r, Jacob, True)
            chk1 = max2(chk1, MaxNorm_vec(f1r))
            chk2 = max2(chk2, MaxNorm_vec(f2r))
            For i = 1 To n
                f(i) = f(i) - (lambda_r(i) - f1r(i) + (f2r(i) * z_r(i) - f3r(i)) / r(i))
                f1(i) = f1(i) - lambda_r(i)
            Next i
        End If
        If isMaxC = True Or isMinC = True Then 'Modify Jacobian based on Cx=[c_min, c_max]
            ReDim tmp_vec(1 To n_ieq)
            If isMaxC = True Then
                For k = 1 To n_ieq
                    tmp_vec(k) = tmp_vec(k) + z_u(k) / s_u(k)
                Next k
            End If
            If isMinC = True Then
                For k = 1 To n_ieq
                    tmp_vec(k) = tmp_vec(k) + z_l(k) / s_l(k)
                Next k
            End If
            For i = 1 To n
                tmp_x = 0
                For k = 1 To n_ieq
                    tmp_x = tmp_x + tmp_vec(k) * C(k, i) ^ 2
                Next k
                Jacob(i, i) = Jacob(i, i) - tmp_x
                For j = i + 1 To n
                    tmp_x = 0
                    For k = 1 To n_ieq
                        tmp_x = tmp_x + tmp_vec(k) * C(k, i) * C(k, j)
                    Next k
                    Jacob(i, j) = Jacob(i, j) - tmp_x
                    'Jacob(j, i) = Jacob(i, j)
                Next j
            Next i
        End If
        If isMaxC = True Or isMinC = True Then 'Modify gradient based on Cx=[c_min, c_max]
            cx = modMath.M_Dot(C, x)
            ReDim tmp_vec(1 To n_ieq)
            ReDim tmp_vec2(1 To n_ieq)
            If isMaxC = True Then
                Call Calc_Grad_MaxC(n_ieq, sigma * mu, cx, s_u, z_u, lambda_u, cmax, f1u, f2u, f3u, False)
                chk1 = max2(chk1, MaxNorm_vec(f1u))
                chk2 = max2(chk2, MaxNorm_vec(f2u))
                 For i = 1 To n_ieq
                    tmp_vec(i) = tmp_vec(i) + lambda_u(i) + f1u(i) + (f2u(i) * z_u(i) + f3u(i)) / s_u(i)
                    tmp_vec2(i) = tmp_vec2(i) + lambda_u(i)
                Next i
            End If
            If isMinC = True Then
                Call Calc_Grad_MaxC(n_ieq, sigma * mu, cx, s_l, z_l, lambda_l, cmin, f1l, f2l, f3l, True)
                chk1 = max2(chk1, MaxNorm_vec(f1l))
                chk2 = max2(chk2, MaxNorm_vec(f2l))
                For i = 1 To n_ieq
                    tmp_vec(i) = tmp_vec(i) + lambda_l(i) - f1l(i) + (f2l(i) * z_l(i) - f3l(i)) / s_l(i)
                    tmp_vec2(i) = tmp_vec2(i) + lambda_l(i)
                Next i
            End If
            
            tmp_vec = modMath.M_Dot(C, tmp_vec, 1)
            tmp_vec2 = modMath.M_Dot(C, tmp_vec2, 1)
            For i = 1 To n
                f(i) = f(i) - tmp_vec(i)
                f1(i) = f1(i) - tmp_vec2(i)
            Next i
            Erase cx, tmp_vec
        End If
        
        chk1 = max2(chk1, MaxNorm_vec(f1))
        'symmetrize the Jacobian
        For i = 1 To n - 1
            For j = i + 1 To n
                Jacob(j, i) = Jacob(i, j)
            Next j
        Next i

        'Solve the symmetrized Jacobian
        d = modMath.Solve_Linear_LDL(Jacob, f)

        If isMax = True Then
            For i = 1 To n
                ds(i) = f2s(i) - d(i)
                d2s(i) = f1s(i) + (f3s(i) + z_s(i) * ds(i)) / s(i)
                d3s(i) = -(f3s(i) + z_s(i) * ds(i)) / s(i)
            Next i
        End If
        If isMin = True Then
            For i = 1 To n
                dr(i) = -f2r(i) + d(i)
                d2r(i) = -f1r(i) - (f3r(i) + z_r(i) * dr(i)) / r(i)
                d3r(i) = -(f3r(i) + z_r(i) * dr(i)) / r(i)
            Next i
        End If
        If isMaxC = True Or isMinC = True Then
            ReDim tmp_vec(1 To n_ieq)
            For i = 1 To n_ieq
                For j = 1 To n
                    tmp_vec(i) = tmp_vec(i) + C(i, j) * d(j)
                Next j
            Next
            If isMaxC = True Then
                For i = 1 To n_ieq
                    d1u(i) = f2u(i) - tmp_vec(i)
                    d2u(i) = f1u(i) + (f3u(i) + z_u(i) * d1u(i)) / s_u(i)
                    d3u(i) = -(f3u(i) + z_u(i) * d1u(i)) / s_u(i)
                Next i
            End If
            If isMinC = True Then
                For i = 1 To n_ieq
                    d1l(i) = -f2l(i) + tmp_vec(i)
                    d2l(i) = -f1l(i) - (f3l(i) + z_l(i) * d1l(i)) / s_l(i)
                    d3l(i) = -(f3l(i) + z_l(i) * d1l(i)) / s_l(i)
                Next i
            End If
        End If
        
        'Line search to find valid step size
        alpha = 1
        alpha_z = 1
        If isMax = True Then Call Backtrack_stepsize(alpha, s, ds, n)
        If isMin = True Then Call Backtrack_stepsize(alpha, r, dr, n)
        If isMaxC = True Then Call Backtrack_stepsize(alpha, s_u, d1u, n_ieq)
        If isMinC = True Then Call Backtrack_stepsize(alpha, s_l, d1l, n_ieq)
        If isMax = True Then Call Backtrack_stepsize(alpha_z, z_s, d3s, n)
        If isMin = True Then Call Backtrack_stepsize(alpha_z, z_r, d3r, n)
        If isMaxC = True Then Call Backtrack_stepsize(alpha_z, z_u, d3u, n_ieq)
        If isMinC = True Then Call Backtrack_stepsize(alpha_z, z_l, d3l, n_ieq)
        alpha = min2(1, 0.99 * alpha)
        alpha_z = min2(1, 0.99 * alpha_z)
        
        'Update values
        For i = 1 To n
            x(i) = x(i) + alpha * d(i)
        Next i
        For i = 1 To n_c
            lambda(i) = lambda(i) + alpha * d(n + i)
        Next i
        If isMax = True Then Call Update_Solution(n, alpha, alpha_z, s, z_s, lambda_s, ds, d3s, d2s)
        If isMin = True Then Call Update_Solution(n, alpha, alpha_z, r, z_r, lambda_r, dr, d3r, d2r)
        If isMaxC = True Then Call Update_Solution(n_ieq, alpha, alpha_z, s_u, z_u, lambda_u, d1u, d3u, d2u)
        If isMinC = True Then Call Update_Solution(n_ieq, alpha, alpha_z, s_l, z_l, lambda_u, d1l, d3l, d2l)
        
'        mu = mu * sigma
        mu_prev = mu
        mu = 0
        If isMax = True Then
            For i = 1 To n
                mu = mu + s(i) * z_s(i)
            Next i
        End If
        If isMin = True Then
            For i = 1 To n
                mu = mu + r(i) * z_r(i)
            Next i
        End If
        If isMaxC = True Then
            For i = 1 To n_ieq
                mu = mu + s_u(i) * z_u(i)
            Next i
        End If
        If isMinC = True Then
            For i = 1 To n_ieq
                mu = mu + s_l(i) * z_l(i)
            Next i
        End If
        
        mu = mu / (nn - n)
        sigma = min2(1, (mu / mu_prev) ^ 3)
        
        iterate = iterate + 1
        If iterate Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "QPSolve:IPM: " & iterate & "/" & iter_max
        End If

    Loop While iterate < iter_max And (mu > tol Or chk1 > 0.0001 Or chk2 > tol)

    IPM = True
    If iterate >= iter_max Then
        IPM = False
        Debug.Print "QPSolve:IPM: failed to converge."
    End If

    Erase Jacob, f, d, lambda
    Erase f1s, f2s, f3s, s, ds, d2s, d3s, z_s, lambda_s, xmax
    Erase f1r, f2r, f3r, r, dr, d2r, d3r, z_r, lambda_r, xmin
    Erase lambda_u, z_u, s_u, cmax, f1u, f2u, f3u, d1u, d2u, d3u
    Erase lambda_l, z_l, s_l, cmin, f1l, f2l, f3l, d1l, d2l, d3l
    Application.StatusBar = False
End Function

Private Sub Init_Vec(n As Long, f1() As Double, f2() As Double, f3() As Double, d1() As Double, d2() As Double, d3() As Double)
    ReDim f1(1 To n)
    ReDim f2(1 To n)
    ReDim f3(1 To n)
    ReDim d1(1 To n)
    ReDim d2(1 To n)
    ReDim d3(1 To n)
End Sub

Private Sub Calc_Grad_Max(n As Long, sigma_mu As Double, _
    x() As Double, s() As Double, z() As Double, lambda() As Double, xmax() As Double, _
    f1() As Double, f2() As Double, f3() As Double, Jacob() As Double, Optional isMin As Boolean = False)
    Dim i As Long
    If isMin = False Then
        For i = 1 To n
            Jacob(i, i) = Jacob(i, i) - z(i) / s(i)
            f1(i) = -lambda(i) - z(i)
            f2(i) = -x(i) - s(i) + xmax(i)
            f3(i) = s(i) * z(i) - sigma_mu
        Next i
    Else
        For i = 1 To n
            Jacob(i, i) = Jacob(i, i) - z(i) / s(i)
            f1(i) = lambda(i) - z(i)
            f2(i) = -x(i) + s(i) + xmax(i)
            f3(i) = s(i) * z(i) - sigma_mu
        Next i
    End If
End Sub

Private Sub Calc_Grad_MaxC(n_ieq As Long, sigma_mu As Double, _
    cx() As Double, s() As Double, z() As Double, lambda() As Double, cmax() As Double, _
    f1() As Double, f2() As Double, f3() As Double, Optional isMin As Boolean = False)
    Dim i As Long
    If isMin = False Then
        For i = 1 To n_ieq
            f1(i) = -lambda(i) - z(i)
            f2(i) = -cx(i) - s(i) + cmax(i)
            f3(i) = s(i) * z(i) - sigma_mu
        Next i
    Else
        For i = 1 To n_ieq
            f1(i) = lambda(i) - z(i)
            f2(i) = -cx(i) + s(i) + cmax(i)
            f3(i) = s(i) * z(i) - sigma_mu
        Next i
    End If
End Sub

Private Sub Update_Solution(n As Long, alpha As Double, alpha_z As Double, _
    x() As Double, z() As Double, lambda() As Double, _
    dx() As Double, dz() As Double, dlambda() As Double)
    Dim i As Long
    For i = 1 To n
        x(i) = x(i) + dx(i) * alpha
        z(i) = z(i) + dz(i) * alpha_z
        lambda(i) = lambda(i) + dlambda(i) * alpha
    Next i
End Sub

Private Sub Backtrack_stepsize(alpha As Double, z() As Double, d() As Double, n As Long)
    Dim i As Long, j As Long, m As Long
    j = 1
    Do
        If alpha < 0.0000000001 Then
            Debug.Print "QPSolve:IPM: Step size is becoming too small."
            Exit Do
        End If
        m = 1
        For i = j To n
            If (z(i) + alpha * d(i)) < 0 Then
                m = 0
                j = i
                Exit For
            End If
        Next i
        If m = 0 Then
            alpha = alpha * 0.5
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function max2(x As Double, y As Double) As Double
    max2 = x
    If y > x Then max2 = y
End Function

Private Function min2(x As Double, y As Double) As Double
    min2 = x
    If y < x Then min2 = y
End Function

Private Function MaxNorm_vec(x() As Double) As Double
    Dim i As Long
    MaxNorm_vec = 0
    For i = 1 To UBound(x)
        If Abs(x(i)) > MaxNorm_vec Then MaxNorm_vec = Abs(x(i))
    Next i
End Function





Private Sub Test_xx()
Dim x() As Double, q() As Double, QQ() As Double, A() As Double, B() As Double, C() As Double
    'min (x1^2 + 2 x2^2)
    's.t. x1+x2=1, x1>=0, x2>=0
    ReDim QQ(1 To 2, 1 To 2)
    ReDim q(1 To 2)
    ReDim A(1 To 1, 1 To 2)
    ReDim B(1 To 1)
    QQ(1, 1) = 2
    QQ(2, 2) = 4
    A(1, 1) = 1
    A(1, 2) = 1
    B(1) = 1
    ReDim x(1 To 2)
    x(1) = 0.5
    x(2) = 0.5
    Call mQPSolve.IPM(x, QQ, q, A, B, , 0)
    Debug.Print x(1) & ", " & x(2)
    
    'min (x1^2 + x2^2)
    's.t.   4x1 - 2x2 + 4 =0
    '       2x1 + 2x2 <= -2
    ReDim QQ(1 To 2, 1 To 2)
    ReDim q(1 To 2)
    ReDim A(1 To 1, 1 To 2)
    ReDim B(1 To 1)
    ReDim C(1 To 1, 1 To 2)
    QQ(1, 1) = 2
    QQ(2, 2) = 2
    A(1, 1) = 4
    A(1, 2) = -2
    B(1) = -4
    C(1, 1) = 2
    C(1, 2) = 2
    ReDim x(1 To 2)
    x(1) = -3
    x(2) = -3
    Call mQPSolve.IPM(x, QQ, q, A, B, , , C, -2)
    Debug.Print x(1) & ", " & x(2)
End Sub
