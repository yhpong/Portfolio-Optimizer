VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNSGA_II_Efficient_Frontier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pn_popul As Long            'population size
Private pn_objective As Long        'number of objectives to minimize
Private pn_front As Long            'number of pareto fronts after non-dominated sort
Private pchrom_len As Long          'length of a chromosome
Private pchromosomes() As Long      'set of chromomsomes (1:pchrom_len, 1:pn_popul)
Private pchromosomes_B() As Double  '2nd set of chromomsomes (1:pchrom_len, 1:pn_popul)
Private pobjectives() As Double     'output objective functions
Private pfront() As Long            'index of pareto fronts
Private pcrowd_dist() As Double     'crowding distance
Private pobjectives_initial() As Double

'Problem specific global variables
Private puniv_return() As Double
Private puniv_covar() As Double
Private pmin_wgt As Double, pmax_wgt As Double

'==============================
'Problem specific outputs
'==============================
'Return the mean-variance curve of efficient frontier
Public Property Get EF_Curve() As Double()
Dim i As Long, j As Long, n As Long
Dim subset() As Long, subObjective() As Double
Dim y() As Double, y_tmp() As Double, sort_idx() As Long
    Call KthFront_Subset(1, subset, subObjective)
    n = UBound(subset)
    ReDim y_tmp(1 To n)
    For i = 1 To n
        y_tmp(i) = -subObjective(1, i)
    Next i
    Call modMath.Sort_Quick_A(y_tmp, 1, n, sort_idx)
    ReDim y(1 To n, 1 To 2)
    j = 0
    For i = 1 To n
        j = sort_idx(i)
        y(i, 1) = -subObjective(1, j)
        y(i, 2) = subObjective(2, j)
    Next i
    EF_Curve = y
    Erase y, y_tmp, sort_idx, subset, subObjective
End Property

'Return the portfolio wgts that correspond to the efficient frontier
Public Property Get Portfolio_Wgts() As Double()
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim w() As Double, w_tmp() As Double, y_tmp() As Double, sort_idx() As Long
Dim subset() As Long, subObjective() As Double
    Call KthFront_Subset(1, subset, subObjective)
    n = UBound(subset)
    ReDim y_tmp(1 To n)
    For i = 1 To n
        y_tmp(i) = -subObjective(1, i)
    Next i
    Call modMath.Sort_Quick_A(y_tmp, 1, n, sort_idx)
    
    m = UBound(puniv_return, 1)
    ReDim w(1 To UBound(subset), 1 To m)
    For i = 1 To n
        k = subset(sort_idx(i))
        ReDim w_tmp(1 To pchrom_len)
        For j = 1 To pchrom_len
            w_tmp(j) = pchromosomes_B(j, k)
        Next j
        Call Normalize_Wgt(w_tmp, pmin_wgt, pmax_wgt)
        For j = 1 To pchrom_len
            w(i, pchromosomes(j, k)) = w_tmp(j)
        Next j
    Next i
    Portfolio_Wgts = w
    Erase w, w_tmp, subset, subObjective, y_tmp, sort_idx
End Property
'=======================================================


Public Property Get front() As Long()
    front = pfront
End Property

Public Property Get objectives() As Double()
    objectives = pobjectives
End Property

Public Property Get objectives_initial() As Double()
    objectives_initial = pobjectives_initial
End Property

Public Property Get chromosomes() As Long()
    chromosomes = pchromosomes
End Property

Public Property Get chromosomes_B() As Double()
    chromosomes_B = pchromosomes_B
End Property

Sub Init(n_objective As Long, n_popul As Long, chrom_len As Long)
    pn_objective = n_objective
    pn_popul = n_popul
    pchrom_len = chrom_len
    ReDim pchromosomes(1 To pchrom_len, 1 To pn_popul) 'Index of stocks held by each individual
    ReDim pchromosomes_B(1 To pchrom_len, 1 To pn_popul)  'Weight of stocks held by each individual
    ReDim pobjectives(1 To pn_objective, 1 To pn_popul) '1: -ve return, 2:variance
End Sub

Sub Assign(chromosomes() As Long, chromosomes_B() As Double, objectives() As Double)
    pn_objective = UBound(objectives, 1)
    pn_popul = UBound(chromosomes, 2)
    pchrom_len = UBound(chromosomes, 1)
    pchromosomes = chromosomes
    pchromosomes_B = chromosomes_B
    pobjectives = objectives
End Sub

Sub Reset()
    pn_objective = 0
    pn_popul = 0
    pchrom_len = 0
    Erase pchromosomes, pchromosomes_B, pobjectives, pfront, pcrowd_dist
End Sub

Sub Read_Global(univ_return() As Double, univ_covar() As Double, min_wgt As Double, max_wgt As Double)
    puniv_return = univ_return
    puniv_covar = univ_covar
    pmin_wgt = min_wgt
    pmax_wgt = max_wgt
End Sub


'*** Main Loop of NSGA_ii
'*** This example has two set of chromosomes
'*** Modify inputs and objective functions specific to the problem
Sub Evolve(Optional tournament_size As Long = 5, Optional crossover_rate As Double = 0.8, Optional mutation_rate As Double = 0.1, Optional iterate_max As Long = 1000)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, n_popul As Long
Dim child() As Long, chromosomes_temp() As Long
Dim child_B() As Double, chromosomes_B_temp() As Double
Dim subset() As Long, subObjective() As Double, sub_crowd_dist() As Double
Dim iArr() As Long, objectives_temp() As Double, child_objectives() As Double
Dim GA2 As cNSGA_II_Efficient_Frontier

'Randomly generate initial population
Call Randomize_Chromosomes
Call Calc_Objective(pchromosomes, pchromosomes_B, pn_objective, pobjectives)
pobjectives_initial = pobjectives

'Sort Initial Population
Call NonDominated_Sort
Call Calc_Crowding_Dist_Full

Set GA2 = New cNSGA_II_Efficient_Frontier
For iterate = 1 To iterate_max
    
    If iterate Mod 20 = 0 Then
        DoEvents
        Application.StatusBar = "cNSGA_II: Evolve..." & iterate & "/" & iterate_max
    End If

    'Create Children population
    Call Make_new_Pop(child, child_B, tournament_size, crossover_rate, mutation_rate)
    Call Calc_Objective(child, child_B, pn_objective, child_objectives)
    
    'Union current generation and its children
    Call Union_Set(pchromosomes, child, chromosomes_temp)
    Call Union_Set(pchromosomes_B, child_B, chromosomes_B_temp)
    Call Union_Set(pobjectives, child_objectives, objectives_temp)
    
    '=== Perform Non-Dominated Sort on combined population
    With GA2
        Call .Assign(chromosomes_temp, chromosomes_B_temp, objectives_temp)
        Call .NonDominated_Sort
    End With
    '=====================================

    '=== Add individual to new population starting from 1st front
    k = 1
    n_popul = 0
    Do While n_popul < pn_popul
        Call GA2.KthFront_Subset(k, subset, subObjective)
        Call Calc_Crowding_Dist(subObjective, sub_crowd_dist)
        m = n_popul
        n = UBound(subset)
        If (m + n) > pn_popul Then 'Add individual with larger crowding distance
        
            n_popul = pn_popul
            Call modMath.Sort_Quick_A(sub_crowd_dist, 1, n, iArr)
            For j = 1 To pn_popul - m
                pfront(m + j) = k
                pcrowd_dist(m + j) = sub_crowd_dist(n - j + 1)
                For i = 1 To pn_objective
                    pobjectives(i, m + j) = subObjective(i, iArr(n - j + 1))
                Next i
                For i = 1 To pchrom_len
                    pchromosomes(i, m + j) = chromosomes_temp(i, subset(iArr(n - j + 1)))
                    pchromosomes_B(i, m + j) = chromosomes_B_temp(i, subset(iArr(n - j + 1)))
                Next i
            Next j
            
        ElseIf (m + n) <= pn_popul Then
        
            n_popul = n_popul + n
            For j = 1 To n
                pfront(m + j) = k
                pcrowd_dist(m + j) = sub_crowd_dist(j)
                For i = 1 To pn_objective
                    pobjectives(i, m + j) = subObjective(i, j)
                Next i
                For i = 1 To pchrom_len
                    pchromosomes(i, m + j) = chromosomes_temp(i, subset(j))
                    pchromosomes_B(i, m + j) = chromosomes_B_temp(i, subset(j))
                Next i
            Next j
            
        End If
        k = k + 1
    Loop
    '=====================================
    
     Call GA2.Reset
Next iterate

'=== Calculate objectives and sort the final population
Call Calc_Objective(pchromosomes, pchromosomes_B, pn_objective, pobjectives)
Call NonDominated_Sort
'===============================================

Erase child, chromosomes_temp, iArr, child_B, chromosomes_B_temp
Application.StatusBar = False
End Sub



'*** New implementaion from Tsung-Che Chiang, much faster
'*** http://web.ntnu.edu.tw/~tcchiang/publications/nsga3cpp/nsga3cpp.htm
'*** Input: objectives(number of objectives, size of population) to be minimized
'*** Output: front(1 to size of population) pareto front that each individual belongs to
Sub NonDominated_Sort()
Dim i As Long, j As Long, m As Long, n As Long, k As Long, p As Long, q As Long
Dim be_dominated As Boolean
Dim F_set() As Long
Dim num_assigned As Long, rank As Long
num_assigned = 0
rank = 1
ReDim pfront(1 To pn_popul) 'layer number of each solution
Do While num_assigned < pn_popul

    ReDim F_set(0 To 0) 'Set of current front
    
    For p = 1 To pn_popul
        If pfront(p) = 0 Then
    
            be_dominated = False
        
            j = 1
            Do While j <= UBound(F_set)
                q = F_set(j)
                n = 0
                k = 0
                i = 0
                For m = 1 To pn_objective
                    If pobjectives(m, p) < pobjectives(m, q) Then 'p wins
                        n = n + 1
                    ElseIf pobjectives(m, p) > pobjectives(m, q) Then 'q wins
                        k = k + 1
                    Else 'tie
                        i = i + 1
                    End If
                Next m
    
                If (k + i) = pn_objective And k > 0 Then 'p is dominated by q
                    be_dominated = True
                    Exit Do
                ElseIf (n + i) = pn_objective And n > 0 Then 'p dominates q
                    Call Erase_Array(F_set, j)
                    j = j - 1
                End If
    
                j = j + 1
            Loop
    
            If be_dominated = False Then
                Call Stack_Array(F_set, p)
            End If
    
        End If
    Next p

    For i = 1 To UBound(F_set)
        pfront(F_set(i)) = rank
    Next i
    num_assigned = num_assigned + UBound(F_set)
    pn_front = rank
    rank = rank + 1
Loop

Erase F_set
End Sub


''*** Original Algortihm published, not very fast
''*** Input: objectives(number of objectives, size of population) is to be minimized
''*** Output: front(1 to popul) label front that each individual belongs to
'Sub NonDominated_Sort(objectives() As Double, front() As Long)
'Dim i As Long, j As Long, m As Long, n As Long, k As Long, p As Long, q As Long
'Dim np As Long, ns As Long
'Dim popul As Long, objective_num As Long
'Dim F_set() As Long, S() As Long, H_Set() As Long
'Dim n_dominated() As Long, n_dominates() As Long
'
'objective_num = UBound(objectives, 1)
'popul = UBound(objectives, 2)
'
'ReDim front(1 To popul) 'layer number of each solution
'ReDim F_set(0 To 0) 'Set of current front
'ReDim S(1 To popul, 1 To popul) 'set of solutions dominated by p
'ReDim n_dominated(1 To popul) 'number of solutions that dominate p
'ReDim n_dominates(1 To popul) 'number of solutions that p dominates
'
''Construct 1st Pareto Front
'For p = 1 To popul
'    np = 0
'    ns = 0
'    For q = 1 To popul
'    'If p <> q Then
'        n = 0
'        k = 0
'        i = 0
'        For m = 1 To objective_num
'            If objectives(m, p) < objectives(m, q) Then 'p wins
'            n = n + 1
'            ElseIf objectives(m, p) > objectives(m, q) Then 'q wins
'            k = k + 1
'            Else 'tie
'            i = i + 1
'            End If
'        Next m
'
'        If (n + i) = objective_num And n > 0 Then 'p dominates q
'            ns = ns + 1
'            S(p, ns) = q
'        ElseIf (k + i) = objective_num And k > 0 Then 'p dominated by q
'            np = np + 1
'        End If
'    'End If
'    Next q
'    n_dominates(p) = ns
'    n_dominated(p) = np
'    If np = 0 Then
'        Call Stack_Array(F_set, p)
'        front(p) = 1
'    End If
'Next p
'
'k = 1
'Do While UBound(F_set) >= 1
'    ReDim H_Set(0 To 0)
'    For i = 1 To UBound(F_set)
'        p = F_set(i)
'        For j = 1 To n_dominates(p)
'            q = S(p, j)
'            n_dominated(q) = n_dominated(q) - 1
'            If n_dominated(q) = 0 Then
'                Call Stack_Array(H_Set, q)
'                front(q) = k + 1
'            End If
'        Next j
'    Next i
'    k = k + 1
'    ReDim F_set(0 To UBound(H_Set))
'    F_set = H_Set
'Loop
''=====================================
'
'Erase F_set, S, H_Set, n_dominated, n_dominates
'
'End Sub


'*** Extract the k-th front from the whole population
Sub KthFront_Subset(k As Long, subset() As Long, subObjective() As Double)
Dim i As Long, j As Long, m As Long, n_sub As Long
    n_sub = 0
    
    ReDim subset(1 To pn_popul)
    For i = 1 To pn_popul
        If pfront(i) = k Then
            n_sub = n_sub + 1
            subset(n_sub) = i
        End If
    Next i
    ReDim Preserve subset(1 To n_sub)
    
    ReDim subObjective(1 To pn_objective, 1 To n_sub)
    For i = 1 To n_sub
        j = subset(i)
        For m = 1 To pn_objective
            subObjective(m, i) = pobjectives(m, j)
        Next m
    Next i
End Sub


'*** Calculate crowding distance of every front in the whole population
Private Sub Calc_Crowding_Dist_Full()
Dim i As Long, n As Long, k As Long
Dim subset() As Long
Dim subObjective() As Double, sub_crowd_dist() As Double
    ReDim pcrowd_dist(1 To pn_popul)
    For k = 1 To pn_front
        Call KthFront_Subset(k, subset, subObjective)
        Call Calc_Crowding_Dist(subObjective, sub_crowd_dist)
        n = UBound(subset)
        For i = 1 To n
            pcrowd_dist(subset(i)) = sub_crowd_dist(i)
        Next i
    Next k
    Erase subset, subObjective, sub_crowd_dist
End Sub


'*** Calculate crowding distance of a single front
Private Sub Calc_Crowding_Dist(objectives() As Double, crowd_dist() As Double)
Dim i As Long, j As Long, m As Long
Dim popul As Long, objective_num As Long
Dim sort_index() As Long, obj_temp() As Double
Dim tmp_x As Double, INFINITY As Double

INFINITY = Exp(70)

objective_num = UBound(objectives, 1)
popul = UBound(objectives, 2)

ReDim crowd_dist(1 To popul)
ReDim obj_temp(1 To popul)

For m = 1 To objective_num

    For i = 1 To popul
        obj_temp(i) = objectives(m, i)
    Next i
    
    Call modMath.Sort_Quick_A(obj_temp, 1, popul, sort_index)
      
    crowd_dist(sort_index(1)) = INFINITY
    crowd_dist(sort_index(popul)) = INFINITY
    
    If popul > 2 Then
        tmp_x = obj_temp(popul) - obj_temp(1)
        For i = 2 To (popul - 1)
            j = sort_index(i)
            If tmp_x = 0 Then
                crowd_dist(j) = INFINITY
            Else
                crowd_dist(j) = crowd_dist(j) + (obj_temp(i + 1) - obj_temp(i - 1)) / tmp_x
            End If
        Next i
    End If
    
Next m
Erase sort_index, obj_temp
End Sub


'*** Partial order operator
'*** i dominates j if i is in a lower layer,
'*** or if they are in the same layer but i has a larger crowding distance
Private Function Crowded_Compare(i_rank As Long, j_rank As Long, i_dist As Double, j_dist As Double) As Boolean
    Crowded_Compare = False
    If (i_rank < j_rank) Or ((i_rank = j_rank) And (i_dist > j_dist)) Then
        Crowded_Compare = True 'i dominates j
    End If
End Function


'*** Sort the entire population base on partial order
Private Sub Crowded_Sort(front() As Long, crowd_dist() As Double, inLow As Long, inHi As Long, sort_index() As Long, Optional first_run As Long = 1)
Dim pivot_pt As Long, i As Long
Dim tmpSwap As Double
Dim tmpLow  As Long
Dim tmpHi   As Long
Dim tmp_i As Long

If first_run = 1 Then
    ReDim sort_index(LBound(front) To UBound(front))
    For i = LBound(front) To UBound(front)
        sort_index(i) = i
    Next i
End If

tmpLow = inLow
tmpHi = inHi

pivot_pt = (inLow + inHi) / 2

While (tmpLow <= tmpHi)
    
    While (Crowded_Compare(front(tmpLow), front(pivot_pt), crowd_dist(tmpLow), crowd_dist(pivot_pt)) = True _
        And tmpLow < inHi)
        tmpLow = tmpLow + 1
    Wend

    While (Crowded_Compare(front(tmpHi), front(pivot_pt), crowd_dist(tmpHi), crowd_dist(pivot_pt)) = False _
        And tmpHi > inLow)
        tmpHi = tmpHi - 1
    Wend
    
    If (tmpLow <= tmpHi) Then
        tmpSwap = crowd_dist(tmpLow)
        crowd_dist(tmpLow) = crowd_dist(tmpHi)
        crowd_dist(tmpHi) = tmpSwap
        
        tmp_i = front(tmpLow)
        front(tmpLow) = front(tmpHi)
        front(tmpHi) = tmp_i
        
        tmp_i = sort_index(tmpLow)
        sort_index(tmpLow) = sort_index(tmpHi)
        sort_index(tmpHi) = tmp_i
        
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
    End If

Wend

If (inLow < tmpHi) Then Crowded_Sort front, crowd_dist, inLow, tmpHi, sort_index, 0
If (tmpLow < inHi) Then Crowded_Sort front, crowd_dist, tmpLow, inHi, sort_index, 0

End Sub


'Generate a set of children from current population
Private Sub Make_new_Pop(chromosome_new() As Long, chromosome_B_new() As Double, _
        tournament_size As Long, crossover_rate As Double, mutation_rate As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long, p As Long
Dim tmp_x As Double
Dim intfather As Long, intmother As Long
Dim father() As Long, mother() As Long, child() As Long
Dim father_B() As Double, mother_B() As Double, child_B() As Double

'=== Initialize Memory
ReDim father(1 To pchrom_len)
ReDim mother(1 To pchrom_len)
ReDim child(1 To pchrom_len)
ReDim chromosome_new(1 To pchrom_len, 1 To pn_popul)

'In this particular case there are 2 sets of chromosomes for each individual:
'the list of stocks held and the weight of each stock
ReDim father_B(1 To pchrom_len)
ReDim mother_B(1 To pchrom_len)
ReDim child_B(1 To pchrom_len)
ReDim chromosome_B_new(1 To pchrom_len, 1 To pn_popul)
'=====================================
    
'=== Generate New Population
k = 0
Do While k < pn_popul
    
    '=== Select father
    intfather = TournamentSelection(tournament_size)
    For i = 1 To pchrom_len
        father(i) = pchromosomes(i, intfather)
        father_B(i) = pchromosomes_B(i, intfather)
    Next i
    '=======================================
        
    Randomize
    tmp_x = Rnd()
    If tmp_x <= crossover_rate Then
        '=== Perform Crossover
        intmother = TournamentSelection(tournament_size)
        For i = 1 To pchrom_len
            mother(i) = pchromosomes(i, intmother)
            mother_B(i) = pchromosomes_B(i, intmother)
        Next i
            
        Call Crossover(father, father_B, mother, mother_B, child, child_B)
        '==========================
    Else
        '=== Directly Copy father
        child = father
        child_B = father_B
        '==========================
    End If
    
    'Mutate
    tmp_x = Rnd()
    If tmp_x < mutation_rate Then Call Mutation(child, child_B)
    
    '=== Append child to new population
    k = k + 1
    For i = 1 To pchrom_len
        chromosome_new(i, k) = child(i)
        chromosome_B_new(i, k) = child_B(i)
    Next i
    '====================================
        
Loop

Erase father, mother, child, father_B, mother_B, child_B

End Sub

'==============================================
'======= Selection methods
'==============================================

'Tournament Selection
Private Function TournamentSelection(tournament_size As Long) As Long
Dim i As Long, j As Long, k As Long, popul As Long
Dim tmp_x As Double
Dim intArray() As Long
    intArray = index_array(1, pn_popul)
    intArray = Random_Pick(intArray, tournament_size)
    k = intArray(1)
    For i = 2 To tournament_size
        j = intArray(i)
        If (Crowded_Compare(pfront(j), pfront(k), pcrowd_dist(j), pcrowd_dist(k)) = True) Then k = j
    Next i
    TournamentSelection = k
End Function

'Roulette Selection
Private Function RouletteSelection(prob_C() As Double) As Long
Dim i As Long, n As Long, k As Long, popul As Long
Dim tmp_x As Double
popul = UBound(prob_C)
Randomize
tmp_x = Rnd()
For n = 1 To popul
    If tmp_x >= prob_C(n - 1) And tmp_x <= prob_C(n) Then
        k = n
        Exit For
    End If
Next n
RouletteSelection = k
End Function


'==============================================
'======= Mutation methods
'==============================================
Private Sub Mutation(father() As Long, father_B() As Double)
Dim i As Long
Dim tmp_x As Double
    Randomize
    For i = 1 To UBound(father_B)
        tmp_x = Rnd() 'mutate by 0.9 or 1.1
        'i = Random_Integer(1, pchrom_len) 'pick random gene to mutate
        If tmp_x < 0.5 Then
            father_B(i) = father_B(i) * 0.8
        Else
            father_B(i) = father_B(i) * 1.2
        End If
    Next i
End Sub


'==============================================
'======= Crossover methods
'==============================================
Private Sub Crossover_Uniform(father() As Long, mother() As Long, child() As Long)
Dim i As Long
Dim tmp_x As Double
    ReDim child(1 To pchrom_len)
    child = father
    Randomize
    For i = 1 To pchrom_len
        tmp_x = Rnd()
        If tmp_x < 0.5 Then child(i) = mother(i)
    Next i
End Sub

Private Sub Crossover_1Point(father() As Long, mother() As Long, child() As Long)
Dim i As Long, j As Long
ReDim child(1 To pchrom_len)
    child = father
    Randomize
    j = Random_Integer(1, pchrom_len)
    For i = j To chrom_len
        child(i) = mother(i)
    Next i
End Sub

Private Sub Crossover_2Point(father() As Long, mother() As Long, child() As Long)
Dim i As Long, j As Long, k As Long
    ReDim child(1 To pchrom_len)
    child = father
    Randomize
    i = Random_Integer(1, pchrom_len)
    j = Random_Integer(1, pchrom_len)
    If j > i Then
        k = i
        i = j
        j = k
    End If
    For k = i To j
        child(k) = mother(k)
    Next k
End Sub

Private Sub Crossover(father() As Long, father_B() As Double, mother() As Long, mother_B() As Double, _
                child() As Long, child_B() As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim tmp_x As Double
Dim isCommon As Boolean
Dim intArray() As Long
Dim non_commmon() As Long, non_common_B() As Double

ReDim child(1 To pchrom_len)
ReDim child_B(1 To pchrom_len)

ReDim non_common(1 To 2 * pchrom_len)
ReDim non_common_B(1 To 2 * pchrom_len)

'=== Commonly held stocks are included in child with father's weight
k = 0
n = 0
For i = 1 To pchrom_len
    isCommon = False
    m = father(i)
    tmp_x = father_B(i)
    For j = 1 To pchrom_len
        If mother(j) = m Then
            isCommon = True
            k = k + 1
            child(k) = m
            child_B(k) = tmp_x
            Exit For
        End If
    Next j
    If isCommon = False Then
        n = n + 1
        non_common(n) = m
        non_common_B(n) = tmp_x
    End If
Next i

'=== Fill empty slots randomly from the non-common genes
If k < pchrom_len Then

    For i = 1 To pchrom_len
        isCommon = False
        m = mother(i)
        tmp_x = mother_B(i)
        For j = 1 To k
            If child(j) = m Then
                isCommon = True
                Exit For
            End If
        Next j
        If isCommon = False Then
            n = n + 1
            non_common(n) = m
            non_common_B(n) = tmp_x
        End If
    Next i
    
    ReDim Preserve non_common(1 To n)
    ReDim Preserve non_common_B(1 To n)
    
    intArray = index_array(1, n)
    intArray = Random_Pick(intArray, pchrom_len - k)
    
    For i = k + 1 To pchrom_len
        child(i) = non_common(intArray(i - k))
        child_B(i) = non_common_B(intArray(i - k))
    Next i
    
End If

End Sub


'*** Union two set of populations
Private Sub Union_Popul(Set1() As Long, Set1_B() As Double, Set2() As Long, Set2_B() As Double, new_set() As Long, new_set_B() As Double)
    Call Union_Set(Set1, Set2, new_set)
    Call Union_Set(Set1_B, Set2_B, new_set_B)
End Sub





'***************************************
'*** Context dependent module **********
'***************************************

'Randomly generate chromosomes
Private Sub Randomize_Chromosomes()
Dim i As Long, n As Long, intArray() As Long, intArray2() As Long
    intArray = index_array(1, UBound(puniv_return, 1))
    For n = 1 To pn_popul
        Randomize
        intArray2 = Random_Pick(intArray, pchrom_len)
        For i = 1 To pchrom_len
            pchromosomes(i, n) = intArray2(i)
            pchromosomes_B(i, n) = Rnd()
        Next i
    Next n
End Sub


'Calculate objective functions of the whole population
Private Sub Calc_Objective(chromosomes() As Long, chromosomes_B() As Double, n_objective As Long, objectives() As Double)
Dim i As Long, j As Long, n As Long, chrom_len As Long, n_popul As Long
Dim tmp_x As Double
Dim port() As Long
Dim wgt() As Double
    chrom_len = UBound(chromosomes, 1)
    n_popul = UBound(chromosomes, 2)
    ReDim objectives(1 To n_objective, 1 To n_popul)
    ReDim port(1 To pchrom_len)
    ReDim wgt(1 To pchrom_len)
    '=== Calculate Variance and Return of each portfolio
    For n = 1 To n_popul

        For i = 1 To chrom_len
            port(i) = chromosomes(i, n)
            wgt(i) = chromosomes_B(i, n)
        Next i

        Call Normalize_Wgt(wgt, pmin_wgt, pmax_wgt)

        '=== Return of Portfolio
        tmp_x = 0
        For i = 1 To chrom_len
            tmp_x = tmp_x + wgt(i) * puniv_return(port(i))
        Next i
        objectives(1, n) = -tmp_x 'negative since we are doing minimization
        '==============================================

        '=== Variance of Portfolio
        tmp_x = 0
        For i = 1 To chrom_len
            For j = 1 To chrom_len
                tmp_x = tmp_x + wgt(i) * wgt(j) * puniv_covar(port(i), port(j))
            Next j
        Next i
        objectives(2, n) = tmp_x
        '==============================================
    Next n
End Sub


'Normalize weightings to meet weight contraints
Private Sub Normalize_Wgt(wgt() As Double, min_wgt As Double, max_wgt As Double)
Dim i As Long, j As Long, n As Long
Dim tmp_x As Double, fpp As Double
Dim exceed_max As Boolean
Dim portfolio_size As Long
Dim R_Set() As Long

portfolio_size = UBound(wgt, 1)

'=== Minimum Constraints
fpp = 1 - portfolio_size * min_wgt

tmp_x = 0
For i = 1 To portfolio_size
    tmp_x = tmp_x + wgt(i)
Next i
    
For i = 1 To portfolio_size
    wgt(i) = min_wgt + wgt(i) * fpp / tmp_x
Next i
'====================================
    
'=== Handle Maximum Constraints
ReDim R_Set(1 To portfolio_size)
exceed_max = False
For i = 1 To portfolio_size
    If R_Set(i) = 0 And wgt(i) > max_wgt Then
        exceed_max = True
        Exit For
    End If
Next i
    
Do While exceed_max = True
        
    exceed_max = False
        
    For i = 1 To portfolio_size
        If wgt(i) > max_wgt Then R_Set(i) = 1
    Next i
        
    tmp_x = 0
    fpp = 0
    For i = 1 To portfolio_size
        If R_Set(i) = 0 Then
            tmp_x = tmp_x + wgt(i)
            fpp = fpp + min_wgt
        ElseIf R_Set(i) = 1 Then
            fpp = fpp + max_wgt
        End If
    Next i
    fpp = 1 - fpp
        
    For i = 1 To portfolio_size
        If R_Set(i) = 0 Then
            wgt(i) = min_wgt + wgt(i) * fpp / tmp_x
            If wgt(i) > max_wgt Then exceed_max = True
        ElseIf R_Set(i) = 1 Then
            wgt(i) = max_wgt
        End If
    Next i
        
Loop
'==============================================

End Sub




'==============================================
'===== General Operations
'==============================================

'*** Union two matrix
Private Sub Union_Set(Set1 As Variant, Set2 As Variant, new_set As Variant)
Dim i, p As Long
Dim dimension As Long, N1 As Long, N2 As Long, n As Long
    dimension = UBound(Set1, 1)
    N1 = UBound(Set1, 2)
    N2 = UBound(Set2, 2)
    n = N1 + N2
    
    new_set = Set1
    ReDim Preserve new_set(1 To dimension, 1 To n)
    For p = 1 To N2
        For i = 1 To dimension
            new_set(i, N1 + p) = Set2(i, p)
        Next i
    Next p
End Sub

'*** Generate an interger array from m to n
Private Function index_array(m As Long, n As Long) As Long()
Dim i As Long
Dim intArray() As Long
    ReDim intArray(m To n)
    For i = m To n
        intArray(i) = i
    Next i
    index_array = intArray
End Function

'*** sum of a 1D array
Private Function sum_array(x() As Double) As Double
Dim i As Long, m As Long, n As Long
    m = LBound(x)
    n = UBound(x)
    sum_array = 0
    For i = m To n
        sum_array = sum_array + x(i)
    Next i
End Function

'*** Radomly shuffle a base-1 integer array
Private Function Shuffle(x() As Long) As Long()
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim y() As Long
    n = UBound(x)
    ReDim y(1 To n)
    y = x
    Randomize
    For i = n To 2 Step -1
        j = Random_Integer(1, i)
        k = y(j)
        y(j) = y(i)
        y(i) = k
    Next i
    Shuffle = y
End Function

'*** Radomly pick k-items from x(1 to n)
Private Function Random_Pick(x() As Long, k As Long) As Long()
Dim i As Long, j As Long, n As Long
Dim y() As Long
    n = UBound(x)
    ReDim y(1 To k)
    For i = 1 To k
        y(i) = x(i)
    Next i
    
    Randomize
    For i = k + 1 To n
        j = Random_Integer(1, i)
        If j <= k Then y(j) = x(i)
    Next i
    Random_Pick = y
End Function

'*** A Random integer between lower to upper, including end points
Private Function Random_Integer(lower As Long, upper As Long) As Long
    Random_Integer = Int(Rnd() * (upper - lower + 1)) + lower
End Function

'*** Add i to the last element of x(0 to n)
Private Sub Stack_Array(x() As Long, i As Long)
Dim n As Long
    n = UBound(x) + 1
    ReDim Preserve x(0 To n)
    x(n) = i
End Sub

'*** Remove the i-th element of x(0 to n)
Private Sub Erase_Array(x() As Long, i As Long)
Dim j As Long, n As Long
    n = UBound(x)
    If i = n Then
        ReDim Preserve x(0 To n - 1)
    ElseIf i < n Then
        For j = i To n - 1
            x(j) = x(j + 1)
        Next j
        ReDim Preserve x(0 To n - 1)
    End If
End Sub
