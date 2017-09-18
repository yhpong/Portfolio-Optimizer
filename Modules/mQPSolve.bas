Attribute VB_Name = "mQPSolve"
Option Explicit

'**********************************************
'***** Quadratic Optimizer
'**********************************************

'=== Interior Point Method
'Dr. Abebe Geletu
'https://www.tu-ilmenau.de/fileadmin/media/simulation/Lehre/Vorlesungsskripte/Lecture_materials_Abebe/QPs_with_IPM_and_ASM.pdf
'Solve for x() that minimize { (1/2) *( x^T QQ x) +q^T x }
' s.t. Ax=B and x>=0 and x<=x_max
'If x_max is left blank then no maximum constraint is imposed.
'x() is input as the initial guess
Sub IPM(x() As Double, QQ() As Double, q() As Double, A() As Double, B() As Double, _
        Optional x_max As Variant, Optional x_min As Variant, _
        Optional iter_max As Long = 1000, Optional tol As Double = 0.0000001)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_c As Long, iterate As Long
Dim r_tgt As Double, tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_u As Double
Dim f() As Double, f3() As Double, Jacob() As Double
Dim z() As Double, lambda() As Double, d() As Double
Dim mu As Double, mu_prev As Double, alpha As Double, sigma As Double
Dim s() As Double, z_s() As Double, lambda_s() As Double, ds() As Double
Dim f1s() As Double, f2s() As Double, f3s() As Double
Dim xmax() As Double, isMax As Boolean
Dim r() As Double, z_r() As Double, lambda_r() As Double, dr() As Double
Dim f1r() As Double, f2r() As Double, f3r() As Double
Dim xmin() As Double, isMin As Boolean

    n = UBound(QQ, 1)   'dimension of x
    n_c = UBound(B, 1)  'number of linear constraints Ax=b
    sigma = 0.5
    
    'Check that initial guess of x is valid
    For i = 1 To n
        If x(i) < 0 Then
            Debug.Print "mQPSolve: IPM: Init Fail: x cannot be negative."
            Exit Sub
        End If
    Next i
    
    'Lagrange multiplier for log barrier
    mu = 1 / n 'Better way to initialize mu?
    If mu < 0.001 Then mu = 0.001
    
    'slack variable for log barrier
    ReDim z(1 To n)
    For i = 1 To n
        z(i) = mu / x(i)
    Next i
    
    'Lagrange Multiplier for linear constraints Ax=B
    ReDim lambda(1 To n_c)
    For i = 1 To n_c  'Better way to initialize lambda?
        lambda(i) = 1
    Next i
    
    'Check if there is an upper constraint on x
    isMax = False
    If IsMissing(x_max) = False Then
        isMax = True
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
                Exit Sub
            End If
        Next i
        ReDim s(1 To n)
        ReDim z_s(1 To n)
        ReDim lambda_s(1 To n)
        For i = 1 To n
            s(i) = xmax(i) - x(i) 'slack variable x+s=x_max, s>=0
            z_s(i) = mu / s(i)    'slack variable for log barrier on s
            lambda_s(i) = 1       'Lagrange Multiplier for x+s=x_max
        Next i
    End If
    
    'Check if there is an lower constraint on x
    isMin = False
    If IsMissing(x_min) = False Then
        isMin = True
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
                Exit Sub
            End If
        Next i
        ReDim r(1 To n)
        ReDim z_r(1 To n)
        ReDim lambda_r(1 To n)
        For i = 1 To n
            r(i) = x(i) - xmin(i) 'slack variable x-r=x_min, r<=0
            z_r(i) = mu / r(i)    'slack variable for log barrier on r
            lambda_r(i) = 1       'Lagrange Multiplier for x-r=x_min
        Next i
    End If
    
    'Constant part of the Jacobian
    ReDim Jacob(1 To n + n_c, 1 To n + n_c)
    For i = 1 To n
        For j = i + 1 To n
            Jacob(i, j) = -QQ(i, j)
            Jacob(j, i) = -QQ(j, i)
        Next j
        For j = 1 To n_c
            Jacob(i, n + j) = A(j, i)
            Jacob(n + j, i) = A(j, i)
        Next j
    Next i
    
    'Start Iteration
    iterate = 0
    Do
        ReDim f3(1 To n)
        ReDim f(1 To n + n_c)
        For i = 1 To n
            Jacob(i, i) = -QQ(i, i) - z(i) / x(i)
            f3(i) = -(x(i) * z(i) - sigma * mu)
            tmp_x = 0
            For j = 1 To n
                tmp_x = tmp_x + QQ(i, j) * x(j)
            Next j
            tmp_y = 0
            For j = 1 To n_c
                tmp_y = tmp_y + A(j, i) * lambda(j)
            Next j
            f(i) = (tmp_x - tmp_y - z(i) + q(i)) - f3(i) / x(i)
        Next i
        For j = 1 To n_c
            tmp_x = 0
            For i = 1 To n
                tmp_x = tmp_x + A(j, i) * x(i)
            Next i
            f(n + j) = -(tmp_x - B(j))
        Next j
        
        If isMax = True Then
            ReDim f1s(1 To n)
            ReDim f2s(1 To n)
            ReDim f3s(1 To n)
            For i = 1 To n
                Jacob(i, i) = Jacob(i, i) - z_s(i) / s(i)
                f1s(i) = -lambda_s(i) - z_s(i)
                f2s(i) = -(x(i) + s(i) - xmax(i))
                f3s(i) = -(s(i) * z_s(i) - sigma * mu)
                f(i) = f(i) - lambda_s(i) - (f1s(i) - f3s(i) / s(i)) - z_s(i) * f2s(i) / s(i)
            Next i
        End If
        
        If isMin = True Then
            ReDim f1r(1 To n)
            ReDim f2r(1 To n)
            ReDim f3r(1 To n)
            For i = 1 To n
                Jacob(i, i) = Jacob(i, i) - z_r(i) / r(i)
                f1r(i) = lambda_r(i) - z_r(i)
                f2r(i) = -(x(i) - r(i) - xmin(i))
                f3r(i) = -(r(i) * z_r(i) - sigma * mu)
                f(i) = f(i) - lambda_r(i) + (f1r(i) - f3r(i) / r(i)) - z_r(i) * f2r(i) / r(i)
            Next i
        End If
        
        'Solve the symmetrized Jacobian
        d = modMath.Solve_Linear_LDL(Jacob, f)
        ReDim Preserve d(1 To 2 * n + n_c)
        For i = 1 To n
            d(n + n_c + i) = (f3(i) - z(i) * d(i)) / x(i)
        Next i
        
        If isMax = True Then
            ReDim ds(1 To 3 * n)
            For i = 1 To n
                ds(i) = f2s(i) - d(i)
                ds(n + i) = (-z_s(i) * d(i) - f3s(i) + z_s(i) * f2s(i)) / s(i) + f1s(i)
                ds(2 * n + i) = (f3s(i) - z_s(i) * ds(i)) / s(i)
            Next i
        End If
        
        If isMin = True Then
            ReDim dr(1 To 3 * n)
            For i = 1 To n
                dr(i) = -f2r(i) + d(i)
                dr(n + i) = (-z_r(i) * d(i) + f3r(i) + z_r(i) * f2r(i)) / r(i) - f1r(i)
                dr(2 * n + i) = (f3r(i) - z_r(i) * dr(i)) / r(i)
            Next i
        End If
        
        'Line search to find valid step size
        alpha = 1
        Do
            If alpha < 0.0000000001 Then
                Debug.Print "QPSolve:IPM: Step size is becoming too small."
                Exit Do
            End If
            m = 0
            For i = 1 To n
                tmp_x = x(i) + alpha * d(i)
                tmp_y = z(i) + alpha * d(n + n_c + i)
                If tmp_x < 0 Or tmp_y < 0 Then
                    m = 1
                    Exit For
                End If
                If isMax = True And m = 0 Then
                    tmp_x = s(i) + alpha * ds(i)
                    tmp_y = z_s(i) + alpha * ds(2 * n + i)
                    If tmp_x < 0 Or tmp_y < 0 Then
                        m = 1
                        Exit For
                    End If
                End If
                If isMin = True And m = 0 Then
                    tmp_x = r(i) + alpha * dr(i)
                    tmp_y = z_r(i) + alpha * dr(2 * n + i)
                    If tmp_x < 0 Or tmp_y < 0 Then
                        m = 1
                        Exit For
                    End If
                End If
            Next i
            If m = 1 Then
                alpha = alpha * 0.5
            Else
                Exit Do
            End If
        Loop

        alpha = 0.8 * alpha
        If alpha > 1 Then alpha = 1
        For i = 1 To n
            x(i) = x(i) + alpha * d(i)
            z(i) = z(i) + alpha * d(n + n_c + i)
        Next i
        For i = 1 To n_c
            lambda(i) = lambda(i) + alpha * d(n + i)
        Next i

        If isMax = True Then
            For i = 1 To n
                s(i) = s(i) + alpha * ds(i)
                z_s(i) = z_s(i) + alpha * ds(2 * n + i)
                lambda_s(i) = lambda_s(i) + alpha * ds(n + i)
            Next i
        End If
       
        If isMin = True Then
            For i = 1 To n
                r(i) = r(i) + alpha * dr(i)
                z_r(i) = z_r(i) + alpha * dr(2 * n + i)
                lambda_r(i) = lambda_r(i) + alpha * dr(n + i)
            Next i
        End If
       
        mu_prev = mu
        mu = 0
        For i = 1 To n
            mu = mu + x(i) * z(i)
        Next i
        k = n
        If isMax = True Then
            For i = 1 To n
                mu = mu + s(i) * z_s(i)
            Next i
            k = k + n
        End If
        If isMin = True Then
            For i = 1 To n
                mu = mu + r(i) * z_r(i)
            Next i
            k = k + n
        End If
        mu = mu / k

        iterate = iterate + 1
        If iterate Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "QPSolve:IPM: " & iterate & "/" & iter_max
        End If

        sigma = (mu / mu_prev) ^ 3
        If sigma > 1 Then sigma = 1
    Loop While iterate < iter_max And mu > tol
    Erase Jacob, f, f3, d, z, lambda
    Erase f1s, f2s, f3s, s, ds, z_s, lambda_s, xmax
    Erase f1r, f2r, f3r, r, dr, z_r, lambda_r, xmin
    If iterate >= iter_max Then Debug.Print "QPSolve:IPM: failed to converge."
    Application.StatusBar = False
End Sub






Private Sub Test_XX()
Dim x() As Double, q() As Double, QQ() As Double, A() As Double, B() As Double
    'min (x1^2 + 2 x2^2)
    's.t. x1+x2=1, x1>=0, x2>=0
    ReDim QQ(1 To 2, 1 To 2)
    QQ(1, 1) = 2
    QQ(2, 2) = 4
    ReDim q(1 To 2)
    ReDim A(1 To 1, 1 To 2)
    ReDim B(1 To 1)
    A(1, 1) = 1
    A(1, 2) = 1
    B(1) = 1
    ReDim x(1 To 2)
    x(1) = 0.5
    x(2) = 0.5
    Call mQPSolve.IPM(x, QQ, q, A, B)
    Debug.Print x(1) & ", " & x(2)
    
    
    'min (x1^2 + x2^2)
    's.t.   4x1 - 2x2 + 4 =0
    '       2x1 + 2x2 - 20<=0
    ReDim QQ(1 To 3, 1 To 3)
    ReDim q(1 To 3)
    ReDim A(1 To 2, 1 To 3)
    ReDim B(1 To 2)
    QQ(1, 1) = 2
    QQ(2, 2) = 2
    A(1, 1) = 4
    A(1, 2) = -2
    A(2, 1) = 2
    A(2, 2) = 2
    A(2, 3) = 1
    B(1) = -4
    B(2) = 20
    ReDim x(1 To 3)
    x(1) = 1
    x(2) = 4
    x(3) = 20 - 2 * x(1) - 2 * x(2)
    Call mQPSolve.IPM(x, QQ, q, A, B)
    Debug.Print x(1) & ", " & x(2) & ", " & x(3)
End Sub
