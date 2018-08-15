Attribute VB_Name = "mDependant"
Option Explicit

Sub Test()
Dim i As Long, j As Long, k As Long, m As Long, n  As Long
Dim A() As Double, B() As Double
Dim A_trim() As Double, B_trim() As Double
Dim A_eq() As Double, B_eq() As Double
Dim mywkbk As Workbook
Dim depend_idx() As Long, depend_ratio() As Double
Set mywkbk = ActiveWorkbook
MsgBox "A"
With mywkbk.Sheets("Sheet1")
    m = 13
    n = 5
    ReDim A(1 To m, 1 To n)
    ReDim B(1 To m)
    For i = 1 To m
        B(i) = .Cells(1 + i, 8).Value
        For j = 1 To n
            A(i, j) = .Cells(1 + i, 1 + j).Value
        Next j
    Next i
    
    Call Find_Dependent_Rows(A, depend_idx, depend_ratio)
    For i = 1 To UBound(depend_idx)
        .Range("J" & 1 + i).Value = depend_idx(i)
        .Range("K" & 1 + i).Value = depend_ratio(i)
    Next i
    
    .Range("B19:H30").Clear
    If Remove_Dependent_Eq(A, B, A_trim, B_trim, "MIN", , True, A_eq, B_eq) = True Then
        .Range("B19").Resize(UBound(A_trim, 2), UBound(A_trim, 1)).Value = Application.WorksheetFunction.Transpose(A_trim)
        .Range("H19").Resize(UBound(B_trim, 1), 1).Value = Application.WorksheetFunction.Transpose(B_trim)
    
        .Range("B29").Resize(UBound(A_eq, 2), UBound(A_trim, 1)).Value = Application.WorksheetFunction.Transpose(A_eq)
        .Range("H29").Resize(UBound(B_eq, 1), 1).Value = Application.WorksheetFunction.Transpose(B_eq)
    End If
    
End With
Set mywkbk = Nothing
End Sub


'For a MxN matrix A(), returns depdend_idx(1:N) and depend_ratio(1:N), where
'depend_idx(i)=0 means that A(i,:) is independent
'depend_idx(i)=i means that A(i,:) is parent to its dependents
'depend_idx(i)<0 means that A(i,:) is all zeroes
'depend_idx(i)=k with 0<k<i means that A(i,:)=A(k,:)*depend_ratio(i)
Sub Find_Dependent_Rows(A() As Double, depend_idx() As Long, depend_ratio() As Double, _
        Optional tol As Double = 0.000000000001)
Dim i As Long, j As Long, k As Long, m As Long, n  As Long
Dim i_pivot As Long, kk As Long, i_count As Long, i_zero As Long
Dim tmp_x As Double, tmp_y As Double, x_pivot As Double

    m = UBound(A, 1)
    n = UBound(A, 2)
    ReDim depend_idx(1 To m)
    ReDim depend_ratio(1 To m)
    For k = 1 To m - 1
        If depend_idx(k) = 0 Then
            'Find a non-zero element in row-k
            i_pivot = -1
            For i = 1 To n
                If A(k, i) <> 0 Then
                    i_pivot = i
                    Exit For
                End If
            Next i
            If i_pivot <> -1 Then
                x_pivot = A(k, i_pivot)
                For kk = k + 1 To m
                    If depend_idx(kk) = 0 Then
                        tmp_x = A(kk, i_pivot)
                        i_count = 0: i_zero = 0
                        For i = 1 To n
                            If A(kk, i) = 0 Then i_zero = i_zero + 1
                            If Abs(x_pivot * A(kk, i) - tmp_x * A(k, i)) < tol Then
                                i_count = i_count + 1
                            Else
                                Exit For
                            End If
                        Next i
                        If i_zero = n Then
                            depend_idx(kk) = -1
                        ElseIf i_count = n Then
                            depend_idx(kk) = k
                            depend_ratio(kk) = tmp_x / x_pivot
                        End If
                    End If
                Next kk
            Else
                depend_idx(k) = -1
            End If
        End If
    Next k
    For i = 1 To m
        k = depend_idx(i)
        If 0 < k And k < i Then
            If depend_idx(k) = 0 Then
                depend_idx(k) = k
                depend_ratio(k) = 1
            End If
        End If
    Next i
End Sub


'Remove redundant rows from system of linear equations. Return FALSE if inconsistency is detected.
'strType    System      Action
'"EQ"       Ax=B        remove null or dependent rows, keeping row with smallest row index
'"MAX"      Ax<=B       remove null or dependent rows, keep only row that gives lowest bound
'"MIN"      Ax>=B       remove null or dependent rows, keep only row that gives hightest bound
'if separate_equal is set to TRUE, inequality constraints that sandwich the same value
'will be output as a separate set of equality constraints A_eq x = B_eq
Function Remove_Dependent_Eq(A() As Double, B() As Double, A_trim() As Double, B_trim() As Double, _
            Optional strType As String = "MAX", Optional tol As Double = 0.000000000001, _
            Optional separate_equal As Boolean = False, _
            Optional A_eq As Variant, Optional B_eq As Variant) As Boolean
Dim i As Long, j As Long, k As Long, m As Long, n  As Long, i_parent As Long
Dim m_trim As Long, m_eq As Long, isKeep() As Long, isKeep_eq() As Long
Dim depend_idx() As Long, depend_ratio() As Double, isProcess() As Long
Dim tight_bound As Double, tight_idx As Long
Dim tight_bound_r As Double, tight_idx_r As Long
Dim tmp_x As Double
Dim n_child As Long, child_list() As Long

    Remove_Dependent_Eq = False
    m = UBound(A, 1)
    n = UBound(A, 2)
    
    Call Find_Dependent_Rows(A, depend_idx, depend_ratio, tol)
    
    m_trim = 0
    ReDim isKeep(1 To m)
    If strType = "EQ" Then
    
        For i = 1 To m
            k = depend_idx(i)
            If k = 0 Or k = i Then
                m_trim = m_trim + 1
                isKeep(m_trim) = i
            ElseIf k < 0 Then
                If B(i) <> 0 Then
                    Debug.Print "Remove_Dependent_Eq: Equation " & i & " is invalid. Termintate."
                    Exit Function
                Else
                    Debug.Print "Remove_Dependent_Eq: Equation " & i & " is null. Ignored."
                End If
            ElseIf k < i Then
                If Abs(B(i) - B(k) * depend_ratio(i)) > tol * (Abs(B(i)) + Abs(B(k) * depend_ratio(i))) Then
                    Debug.Print "Remove_Dependent_Eq: Equation " & i & " and " & k & " are inconsistent."
                    Exit Function
                End If
            End If
        Next i
        
    ElseIf strType = "MAX" Or strType = "MIN" Then
        
        m_eq = 0: ReDim isKeep_eq(1 To m)
        ReDim isProcess(1 To m)
        For i = 1 To m
            If isProcess(i) = 0 Then
                i_parent = depend_idx(i)
                If i_parent = 0 Then
                    'row is independent
                    m_trim = m_trim + 1
                    isKeep(m_trim) = i
                    isProcess(i) = 1
                ElseIf i_parent < 0 Then
                    'row is null
                    If B(i) < 0 Then
                        Debug.Print "Remove_Dependent_Eq: Equation " & i & " is invalid."
                        Exit Function
                    Else
                        Debug.Print "Remove_Dependent_Eq: Equation " & i & " is null."
                    End If
                    isProcess(i) = 1
                ElseIf i_parent > 0 Then
                    
                    'List all dependent rows
                    n_child = 0
                    ReDim child_list(1 To m)
                    For j = i_parent To m
                        If depend_idx(j) = i_parent Then
                            n_child = n_child + 1
                            child_list(n_child) = j
                            isProcess(j) = 1
                        End If
                    Next j
                    ReDim Preserve child_list(1 To n_child)
                    
                    'Find tightest bound of all dependent rows
                    tight_bound = B(i_parent)
                    tight_idx = i_parent
                    For k = 2 To n_child
                        j = child_list(k)
                        If depend_ratio(j) > 0 Then
                            If strType = "MAX" Then
                                If (B(j) / depend_ratio(j)) < tight_bound Then
                                    tight_bound = B(j) / depend_ratio(j)
                                    tight_idx = j
                                End If
                            ElseIf strType = "MIN" Then
                                If (B(j) / depend_ratio(j)) > tight_bound Then
                                    tight_bound = B(j) / depend_ratio(j)
                                    tight_idx = j
                                End If
                            End If
                        End If
                    Next k
        
                    'Only include tighest bound
                    m_trim = m_trim + 1
                    isKeep(m_trim) = tight_idx

                    'Also check for reverse bounds
                    If strType = "MAX" Then
                        tight_bound_r = -Exp(70): tight_idx_r = -1
                    ElseIf strType = "MIN" Then
                        tight_bound_r = Exp(70): tight_idx_r = -1
                    End If
                    For k = 1 To n_child
                        j = child_list(k)
                        If depend_ratio(j) < 0 Then
                            If strType = "MAX" Then
                                 If (B(j) / depend_ratio(j)) > tight_bound_r Then
                                     tight_bound_r = B(j) / depend_ratio(j)
                                     tight_idx_r = j
                                 End If
                             ElseIf strType = "MIN" Then
                                 If (B(j) / depend_ratio(j)) < tight_bound_r Then
                                     tight_bound_r = B(j) / depend_ratio(j)
                                     tight_idx_r = j
                                 End If
                            End If
                        End If
                    Next k
                    
                    If tight_idx_r > 0 Then
                        tmp_x = tight_bound_r - tight_bound
                        If (strType = "MAX" And tmp_x > tol) Or _
                                    (strType = "MIN" And tmp_x < -tol) Then
                                Debug.Print "Remove_Dependent_Eq: Equation " & tight_idx & " and " & tight_idx_r & " are inconsistent."
                                Exit Function
                        End If
                        
                        If separate_equal = True And Abs(tmp_x) < tol Then
                            m_eq = m_eq + 1: isKeep_eq(m_eq) = tight_idx
                            isKeep(m_trim) = 0: m_trim = m_trim - 1
                        Else
                            m_trim = m_trim + 1
                            isKeep(m_trim) = tight_idx_r
                        End If
                    End If
                    
                    Erase child_list
                End If
            End If
        Next i
    End If
    
    'Read out independent rows
    If m_trim = 0 Then Exit Function
    Erase depend_idx, depend_ratio
    
    ReDim Preserve isKeep(1 To m_trim)
    ReDim A_trim(1 To n, 1 To m_trim)
    ReDim B_trim(1 To m_trim)
    For k = 1 To m_trim
        i = isKeep(k)
        For j = 1 To n
            A_trim(j, k) = A(i, j)
        Next j
        B_trim(k) = B(i)
    Next k
    Erase isKeep
    
    If separate_equal = True Then
        If m_eq > 0 Then
            ReDim Preserve isKeep_eq(1 To m_eq)
            ReDim A_eq(1 To n, 1 To m_eq)
            ReDim B_eq(1 To m_eq)
            For k = 1 To m_eq
                i = isKeep_eq(k)
                For j = 1 To n
                    A_eq(j, k) = A(i, j)
                Next j
                B_eq(k) = B(i)
            Next k
        Else
            ReDim A_eq(1 To n, 0 To 0)
            ReDim B_eq(0 To 0)
        End If
    End If
    
    Remove_Dependent_Eq = True
End Function

