Attribute VB_Name = "模块8"
Option Explicit

Public Enum MK_RESULT_TYPE
    MK_S = 1
    MK_Z = 2
    MK_P = 3
    MK_SEN_SLOPE = 4
    MK_N = 5
End Enum

Function MK(rng As Range, retType As MK_RESULT_TYPE) As Variant
    Dim arr() As Double
    Dim n As Long, i As Long, j As Long, k As Long
    Dim S As Double, varS As Double, Z As Double, pValue As Double
    Dim slopes() As Double, idx As Long, nSlopes As Long
    Dim sortedArr() As Double, tieTerm As Double, tieCount As Long
    Dim v As Variant
    Dim cnt As Long
    
    v = rng.value
    If IsArray(v) = False Then
        ReDim arr(1 To 1)
        arr(1) = v
    Else
        cnt = 0
        ReDim arr(1 To UBound(v))
        For i = 1 To UBound(v)
            If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
                cnt = cnt + 1
                arr(cnt) = v(i, 1)
            End If
        Next i
        ReDim Preserve arr(1 To cnt)
    End If
    n = UBound(arr)
    
    If retType = MK_N Then
        MK = n
        Exit Function
    End If
    
    If n < 4 Then
        MK = CVErr(xlErrValue)
        Exit Function
    End If
    
    'MK S统计量
    S = 0
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j) > arr(i) Then
                S = S + 1
            ElseIf arr(j) < arr(i) Then
                S = S - 1
            End If
        Next j
    Next i
    
    '复制数组并排序
    ReDim sortedArr(1 To n)
    For i = 1 To n
        sortedArr(i) = arr(i)
    Next i
    Call QuickSort(sortedArr, 1, n)
    
    '====修复后的结tie修正====
    tieTerm = 0
    k = 1
    Do While k < n
        If sortedArr(k) = sortedArr(k + 1) Then
            tieCount = 1
            Do
                If k + tieCount > n Then Exit Do
                If sortedArr(k + tieCount) <> sortedArr(k) Then Exit Do
                tieCount = tieCount + 1
            Loop
            tieTerm = tieTerm + tieCount * (tieCount - 1) * (2 * tieCount + 5)
            k = k + tieCount
        Else
            k = k + 1
        End If
    Loop
    
    varS = (n * (n - 1) * (2 * n + 5) - tieTerm) / 18
    
    'Z统计量
    If S > 0 Then
        Z = (S - 1) / Sqr(varS)
    ElseIf S < 0 Then
        Z = (S + 1) / Sqr(varS)
    Else
        Z = 0
    End If
    pValue = 2 * (1 - WorksheetFunction.Norm_S_Dist(Abs(Z), True))
    
    'Sen斜率
    nSlopes = n * (n - 1) / 2
    ReDim slopes(1 To nSlopes)
    idx = 0
    For i = 1 To n - 1
        For j = i + 1 To n
            idx = idx + 1
            slopes(idx) = (arr(j) - arr(i)) / (j - i)
        Next j
    Next i
    Dim senSlope As Double
    senSlope = WorksheetFunction.Median(slopes)
    
    Select Case retType
        Case MK_S
            MK = S
        Case MK_Z
            MK = Z
        Case MK_P
            MK = pValue
        Case MK_SEN_SLOPE
            MK = senSlope
    End Select
End Function

Private Sub QuickSort(arr() As Double, L As Long, R As Long)
    Dim i As Long, j As Long
    Dim pivot As Double, tmp As Double
    i = L: j = R
    pivot = arr((L + R) \ 2)
    Do While i <= j
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If L < j Then Call QuickSort(arr, L, j)
    If i < R Then Call QuickSort(arr, i, R)
End Sub

