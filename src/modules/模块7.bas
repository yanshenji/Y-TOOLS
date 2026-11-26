Attribute VB_Name = "模块7"
Option Explicit

'内部工具：截取数组1~n有效片段
Private Function GetValidArr(vals() As Double, n As Long) As Variant
    Dim temp() As Double
    ReDim temp(1 To n)
    Dim i As Long
    For i = 1 To n
        temp(i) = vals(i)
    Next i
    GetValidArr = temp
End Function

' 单值移动极差MR法计算Cpk 单元格UDF
' =Cpk_MR(数据区域, LSL, USL)
Public Function Cpk_MR(dataRng As Range, LSL As Variant, USL As Variant) As Variant
    Dim arr As Variant
    Dim i As Long, n As Long
    Dim xBar As Double, sigmaWithin As Double
    Dim vals() As Double
    Dim MRlist() As Double, MRbar As Double
    Dim upperCpk As Double, lowerCpk As Double
    
    arr = dataRng.value
    n = 0
    ReDim vals(1 To dataRng.Rows.Count)
    
    '提取有效数值
    For i = 1 To UBound(arr)
        If IsNumeric(arr(i, 1)) And Not IsEmpty(arr(i, 1)) Then
            n = n + 1
            vals(n) = arr(i, 1)
        End If
    Next i
    
    If n < 2 Then
        Cpk_MR = "样本不足"
        Exit Function
    End If
    
    '修复：传入有效片段
    xBar = WorksheetFunction.Average(GetValidArr(vals, n))
    
    ReDim MRlist(1 To n - 1)
    For i = 1 To n - 1
        MRlist(i) = Abs(vals(i + 1) - vals(i))
    Next i
    MRbar = WorksheetFunction.Average(MRlist)
    sigmaWithin = MRbar / 1.128
    
    If sigmaWithin <= 0 Then
        Cpk_MR = "σ=0无波动"
        Exit Function
    End If
    
    upperCpk = 999
    lowerCpk = 999
    If IsNumeric(USL) And Not IsEmpty(USL) Then
        upperCpk = (USL - xBar) / (3 * sigmaWithin)
    End If
    If IsNumeric(LSL) And Not IsEmpty(LSL) Then
        lowerCpk = (xBar - LSL) / (3 * sigmaWithin)
    End If
    
    Cpk_MR = WorksheetFunction.Min(upperCpk, lowerCpk)
End Function

'配套 Ppk 单元格UDF
Public Function Ppk_Calc(dataRng As Range, LSL As Variant, USL As Variant) As Variant
    Dim arr As Variant
    Dim i As Long, n As Long
    Dim xBar As Double, sigmaTotal As Double
    Dim vals() As Double
    Dim upper As Double, lower As Double
    
    arr = dataRng.value
    n = 0
    ReDim vals(1 To dataRng.Rows.Count)
    For i = 1 To UBound(arr)
        If IsNumeric(arr(i, 1)) And Not IsEmpty(arr(i, 1)) Then
            n = n + 1
            vals(n) = arr(i, 1)
        End If
    Next i
    If n < 2 Then
        Ppk_Calc = "样本不足"
        Exit Function
    End If
    
    xBar = WorksheetFunction.Average(GetValidArr(vals, n))
    sigmaTotal = WorksheetFunction.StDev_S(GetValidArr(vals, n))
    
    If sigmaTotal <= 0 Then
        Ppk_Calc = "σ=0无波动"
        Exit Function
    End If
    
    upper = 999
    lower = 999
    If IsNumeric(USL) And Not IsEmpty(USL) Then upper = (USL - xBar) / (3 * sigmaTotal)
    If IsNumeric(LSL) And Not IsEmpty(LSL) Then lower = (xBar - LSL) / (3 * sigmaTotal)
    Ppk_Calc = WorksheetFunction.Min(upper, lower)
End Function

'数组入参版本（供宏内部调用，传入已经过滤好的干净数组）
Public Function Cpk_MR_FromArray(vals() As Double, LSL As Variant, USL As Variant) As Variant
    Dim n As Long
    n = UBound(vals) - LBound(vals) + 1
    If n < 2 Then
        Cpk_MR_FromArray = "样本不足"
        Exit Function
    End If
    
    Dim xBar As Double, MRbar As Double, sigmaWithin As Double
    xBar = WorksheetFunction.Average(vals)
    
    Dim MRlist() As Double
    ReDim MRlist(1 To n - 1)
    Dim i As Long
    For i = 1 To n - 1
        MRlist(i) = Abs(vals(i + 1) - vals(i))
    Next i
    MRbar = WorksheetFunction.Average(MRlist)
    sigmaWithin = MRbar / 1.128
    
    If sigmaWithin <= 0 Then
        Cpk_MR_FromArray = "σ=0无波动"
        Exit Function
    End If
    
    Dim upperCpk As Double, lowerCpk As Double
    upperCpk = 999
    lowerCpk = 999
    If IsNumeric(USL) Then upperCpk = (USL - xBar) / (3 * sigmaWithin)
    If IsNumeric(LSL) Then lowerCpk = (xBar - LSL) / (3 * sigmaWithin)
    Cpk_MR_FromArray = WorksheetFunction.Min(upperCpk, lowerCpk)
End Function

'Cpk等级
Public Function Cpk_Rank(cpkVal As Variant) As String
    If Not IsNumeric(cpkVal) Then
        Cpk_Rank = cpkVal
        Exit Function
    End If
    Dim v As Double
    v = cpkVal
    If v < 0.67 Then
        Cpk_Rank = "D级?严重不足"
    ElseIf v < 1 Then
        Cpk_Rank = "C级?能力不足"
    ElseIf v < 1.33 Then
        Cpk_Rank = "B级?临界尚可"
    ElseIf v < 1.67 Then
        Cpk_Rank = "A级?能力充足"
    Else
        Cpk_Rank = "AA级?能力富余"
    End If
End Function

