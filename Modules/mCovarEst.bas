Attribute VB_Name = "mCovarEst"
Option Explicit
'*** Input in this module is x(1 to T, 1 to n_stk)
'*** where x(t, i) the the return of stock i at time t
'*** T is the time horizon, n_stk is the number of stocks\
'*** Output is a covariance matrix of size n_stk x n_stk
'*** Requires: cCorex_Linear, modMath

'Sample Convariance
Function Sample(x() As Double) As Double()
    If UBound(x, 1) <= UBound(x, 2) Then
        Debug.Print "mCovarEst:Sample:Caution, # of obs(" & UBound(x, 1) & ") <= dimension (" & UBound(x, 2) & ")."
    End If
    Sample = modMath.Covariance_Matrix(x)
End Function

'Convariance estimated with total correlation  model
Function Corex(x() As Double, n_latent As Long) As Double()
Dim n_dimension As Long
Dim corex1 As cCorex_Linear
    n_dimension = UBound(x, 2)
    Set corex1 = New cCorex_Linear
    With corex1
        Call .Init(n_dimension, n_latent)
        Call .Train(x)
        Corex = .Covariance_Est
        Call .Restore_x(x)
        Call .Reset
    End With
    Set corex1 = Nothing
End Function

'Convariance estimated with single index  model
'requires input of x_index(1 to T), return of the chosen index which has the same time horizon as x()
Function SingleIndex(x() As Double, x_index() As Double) As Double()
Dim i As Long, j As Long, k As Long, n_dimension As Long, n_obs As Long
Dim tmp_x As Double, tmp_y As Double, s_index As Double
Dim xn() As Double, beta() As Double, eps As Double, betas() As Double, s_eps() As Double
Dim covar() As Double
    n_obs = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    'variance of market index
    tmp_x = 0
    tmp_y = 0
    For i = 1 To n_obs
        tmp_x = tmp_x + x_index(i)
        tmp_y = tmp_y + x_index(i) ^ 2
    Next i
    's_index = (tmp_y - (tmp_x / n_obs) * tmp_x) * n_obs / (n_obs - 1)
    s_index = (tmp_y - (tmp_x / n_obs) * tmp_x) / (n_obs - 1)
    
    'Linear regression of each stock vs index to get
    'betas & variance of residuals
    ReDim xn(1 To n_obs)
    ReDim betas(1 To n_dimension) 'beta to market index
    ReDim s_eps(1 To n_dimension) 'variance of residual
    For j = 1 To n_dimension
        If j Mod 50 = 0 Then DoEvents
        For i = 1 To n_obs
            xn(i) = x(i, j)
        Next i
        Call modMath.linear_regression_single(xn, x_index, beta)
        tmp_x = 0
        tmp_y = 0
        For i = 1 To n_obs
            eps = xn(i) - (beta(1) * x_index(i) + beta(2))
            tmp_x = tmp_x + eps
            tmp_y = tmp_y + eps ^ 2
        Next i
        s_eps(j) = (tmp_y - (tmp_x / n_obs) * tmp_x) / (n_obs - 1)
        betas(j) = beta(1)
    Next j
    
    'Compute the convariance matrix
    ReDim covar(1 To n_dimension, 1 To n_dimension)
    For i = 1 To n_dimension
        covar(i, i) = s_eps(i) + s_index * (betas(i) ^ 2)
        For j = i + 1 To n_dimension
            covar(i, j) = betas(i) * betas(j) * s_index
            covar(j, i) = covar(i, j)
        Next j
    Next i
    
    SingleIndex = covar
    Erase covar, xn, betas, s_eps
End Function


'"Honey, I Shrunk the Sample Covariance Matrix"
'Olivier Ledoit, Michael Wolf (2003)
Function Ledoit(x() As Double) As Double()
Dim i As Long, j As Long, k As Long, n_dimension As Long, n_obs As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim x_avg() As Double, covar() As Double, f() As Double
Dim correl_avg As Double, gamma As Double, pi As Double, rho As Double, shrink_factor As Double

    n_obs = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    '=== mean of each dimension
    ReDim x_avg(1 To n_dimension)
    For i = 1 To n_dimension
        For k = 1 To n_obs
            x_avg(i) = x_avg(i) + x(k, i)
        Next k
        x_avg(i) = x_avg(i) / n_obs
    Next i
    
    '=== Sample Covariance Matrix
    ReDim covar(1 To n_dimension, 1 To n_dimension)
    For i = 1 To n_dimension
        tmp_x = 0
        For k = 1 To n_obs
            tmp_x = tmp_x + (x(k, i) - x_avg(i)) ^ 2
        Next k
        covar(i, i) = tmp_x / (n_obs - 1)
        For j = i + 1 To n_dimension
            tmp_x = 0
            For k = 1 To n_obs
                tmp_x = tmp_x + (x(k, i) - x_avg(i)) * (x(k, j) - x_avg(j))
            Next k
            covar(i, j) = tmp_x / (n_obs - 1)
            covar(j, i) = covar(i, j)
        Next j
    Next i
    
    '=== Average pairwise correlation
    correl_avg = 0
    For i = 1 To n_dimension - 1
        For j = i + 1 To n_dimension
            correl_avg = correl_avg + covar(i, j) / Sqr(covar(i, i) * covar(j, j))
        Next j
    Next i
    correl_avg = correl_avg * 2 / (n_dimension * (n_dimension - 1))

    '=== Shrinkage Target
    ReDim f(1 To n_dimension, 1 To n_dimension)
    For i = 1 To n_dimension
        f(i, i) = covar(i, i)
        For j = i + 1 To n_dimension
            f(i, j) = correl_avg * Sqr(covar(i, i) * covar(j, j))
            f(j, i) = f(i, j)
        Next j
    Next i
    
    '=== Shrinkage Intensity
    gamma = 0
    pi = 0
    rho = 0
    For i = 1 To n_dimension - 1
        For j = i + 1 To n_dimension
            gamma = gamma + (f(i, j) - covar(i, j)) ^ 2
            tmp_x = 0
            For k = 1 To n_obs
                tmp_x = tmp_x + ((x(k, i) - x_avg(i)) * (x(k, j) - x_avg(j)) - covar(i, j)) ^ 2
            Next k
            pi = pi + tmp_x / n_obs
            tmp_x = 0
            tmp_y = 0
            For k = 1 To n_obs
                tmp_z = (x(k, i) - x_avg(i)) * (x(k, j) - x_avg(j)) - covar(i, j)
                tmp_x = tmp_x + ((x(k, i) - x_avg(i)) ^ 2 - covar(i, i)) * tmp_z
                tmp_y = tmp_y + ((x(k, j) - x_avg(j)) ^ 2 - covar(j, j)) * tmp_z
            Next k
            rho = rho + (tmp_x * Sqr(covar(j, j) / covar(i, i)) + tmp_y * Sqr(covar(i, i) / covar(j, j))) / n_obs
        Next j
    Next i
    gamma = gamma * 2
    pi = pi * 2
    rho = rho * correl_avg
    For i = 1 To n_dimension
        gamma = gamma + (f(i, i) - covar(i, i)) ^ 2
        tmp_x = 0
        For k = 1 To n_obs
            tmp_x = tmp_x + ((x(k, i) - x_avg(i)) ^ 2 - covar(i, i)) ^ 2
        Next k
        pi = pi + tmp_x / n_obs
        rho = rho + tmp_x / n_obs
    Next i
    
    shrink_factor = ((pi - rho) / gamma) / n_obs
    If shrink_factor > 1 Then shrink_factor = 1
    If shrink_factor < 0 Then shrink_factor = 0
    
    '=== Apply Shrinkage
    For i = 1 To n_dimension
        covar(i, i) = (1 - shrink_factor) * covar(i, i) + shrink_factor * f(i, i)
        If i < n_dimension Then
            For j = i + 1 To n_dimension
                covar(i, j) = (1 - shrink_factor) * covar(i, j) + shrink_factor * f(i, j)
                covar(j, i) = covar(i, j)
            Next j
        End If
    Next i
    
    Ledoit = covar
    Erase covar, x_avg, f
End Function

