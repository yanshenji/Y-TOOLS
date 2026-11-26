Attribute VB_Name = "函数"
Option Explicit ' 强制变量声明，编译时捕获错误

'======================================================================
' 宏功能：逐单元格对比【标准区域】和【数值区域】，校验数值是否符合标准
'           并自动为数值区域单元格设置字体/填充颜色
' 传入参数：
'   StdRange   As Range   - 标准值区域（存放校验规则的单元格）
'   valueRange As Range   - 数值区域（需要校验的数值单元格）
' 依赖函数：cmpStd 函数（核心标准对比逻辑）
'======================================================================
Sub cmpStd222(StdRange As Range, valueRange As Range)

    '==================== 变量声明 ====================
    Dim maxCount As Long       ' 两个区域的最小单元格数量（防止区域长度不一致）
    Dim std As String          ' 存储当前遍历的标准值
    Dim value_ As Variant      ' 存储原始数值（变体类型，兼容空值/文本）
    Dim value As Double        ' 转换后的数值（用于标准对比）
    Dim i As Long              ' 循环遍历计数器
    Dim iscankao2 As Boolean   ' 标记：当前标准是否包含「参考2」
    Dim isany As Boolean       ' 标记：当前标准是否包含「抽检/自检」
    Dim b As Boolean           ' 存储cmpStd函数的对比结果
    Dim isEmptyCell As Boolean ' 标记：单元格是否为空（区分空值与0）
    
    '==================== 初始化与性能优化 ====================
    ' 关闭屏幕更新/事件，提升大区域校验速度
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 全局错误捕获
    On Error GoTo ex
    
    ' 校验输入区域有效性
    If StdRange Is Nothing Or valueRange Is Nothing Then
        MsgBox "输入区域不能为空！", vbExclamation
        GoTo Finalize
    End If
    
    ' 取两个区域的最小单元格数，避免越界
    maxCount = Application.WorksheetFunction.Min(StdRange.Cells.Count, valueRange.Cells.Count)
    If maxCount = 0 Then
        MsgBox "输入区域无有效单元格！", vbExclamation
        GoTo Finalize
    End If

    '==================== 核心循环：逐单元格校验 ====================
    For i = 1 To maxCount
        ' 1. 读取标准值与待校验值，去除首尾空格
        std = Trim(StdRange.Cells(i).value)
        value_ = valueRange.Cells(i).value
        isEmptyCell = IsEmpty(value_) ' 标记空单元格，区分0与空值
        
        ' 2. 数值类型转换，非数字触发错误处理
        On Error GoTo ee
        If isEmptyCell Then
            value = 0 ' 空值统一转为0，便于规则判断
        Else
            value = CDbl(value_) ' 强制转换为双精度，非数字触发错误
        End If
        
        '==================== 重置单元格格式（清除原有样式） ====================
        With valueRange.Cells(i)
            .Interior.pattern = xlPatternNone
            .Font.ColorIndex = xlColorIndexAutomatic
            .Font.TintAndShade = 0
        End With
        
        '==================== 规则1：IsNull规则（要求数值必须为0/空） ====================
        If InStr(std, "IsNull") > 0 Then
            ' 非0且非空 → 填充橙色标记异常
            If value <> 0 And Not isEmptyCell Then
                With valueRange.Cells(i).Interior
                    .pattern = xlPatternSolid
                    .Color = RGB(243, 84, 26)
                    .TintAndShade = 0
                End With
            End If
            GoTo continue ' 跳过后续规则
        End If
       
        '==================== 规则标记：提前判断特殊关键字 ====================
        iscankao2 = InStr(std, "参考2") > 0
        isany = InStr(std, "抽检") > 0 Or InStr(std, "自检") > 0
        
        '==================== 规则2：有效标准校验（排除无意义文本） ====================
        If (std <> "" And std <> "/" And std <> "-" _
            And InStr(std, "实测") = 0 And InStr(std, "监测") = 0 And InStr(std, "参考") = 0 _
            Or iscankao2) Then
            
            '==================== 子规则：数值=0/空单元格的情况 ====================
            If value = 0 Or isEmptyCell Then
                If Not isany Then ' 非抽检/自检项才填充颜色
                    If iscankao2 Then
                        ' 参考2项 → 填充蓝色
                        With valueRange.Cells(i).Interior
                            .pattern = xlPatternSolid
                            .Color = RGB(37, 109, 183)
                            .TintAndShade = 0
                        End With
                    ElseIf InStr(std, "NotNull") > 0 Then
                        ' NotNull规则 → 空/0填充橙色（必填项未填）
                        With valueRange.Cells(i).Interior
                            .pattern = xlPatternSolid
                            .Color = RGB(255, 136, 16)
                            .TintAndShade = 0
                        End With
                    Else
                        ' 普通必填项 → 空/0填充绿色（未检测）
                        With valueRange.Cells(i).Interior
                            .pattern = xlPatternSolid
                            .Color = RGB(169, 208, 142)
                            .TintAndShade = 0
                        End With
                    End If
                End If
            
            '==================== 子规则：数值≠0 → 核心标准对比 ====================
            Else
                ' 清理标准值中的无效字符（关键字、单位、符号）
                std = Replace(std, "NotNull", "")
                std = Replace(std, "抽检", ""): std = Replace(std, "自检", "")
                std = Replace(std, "参考2", "")
                std = Replace(std, "g/L", ""): std = Replace(std, "N", "")
                std = Replace(std, "%", ""): std = Replace(std, "us/cm2", "")
                std = Replace(std, "(", ""): std = Replace(std, ")", "")
                std = Replace(std, "（", ""): std = Replace(std, "）", "")
                std = Trim(std) ' 再次去除空格，避免空字符串
                
                ' 清理后标准为空 → 默认符合
                If std = "" Or Len(std) <= 1 Then
                    b = True
                Else
                    b = cmpStd(std, value) ' 调用核心对比函数
                End If
                
                '==================== 对比结果：设置字体颜色 ====================
                If Not b Then
                    With valueRange.Cells(i).Font
                        .Color = vbRed ' 不符合标红
                        .TintAndShade = 0
                    End With
                End If
            End If
        End If
        
continue:
        ' 重置错误处理，避免影响下一次循环
        On Error GoTo 0
        On Error GoTo ee
    Next i
    
    ' 校验完成提示
    MsgBox "校验完成！", vbInformation

'==================== 异常处理1：数值转换失败（非数字） ====================
ee:
    ' 非数字 → 填充红色标记
    With valueRange.Cells(i).Interior
        .pattern = xlPatternSolid
        .Color = vbRed
        .TintAndShade = 0
    End With
    ' 批量处理可注释以下MsgBox，改为日志记录
    'MsgBox "数值格式错误！单元格：" & valueRange.Cells(i).Address & " 值：" & valueRange.Cells(i).Value
    Resume continue

'==================== 异常处理2：全局未知错误 ====================
ex:
    MsgBox "发生未知错误：" & Err.Description & vbCrLf & "错误代码：" & Err.Number, vbCritical

'==================== 收尾：恢复Excel设置 ====================
Finalize:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'======================================================================
' 函数功能：核心标准对比逻辑，支持全角/半角符号、范围、负数
' 传入参数：
'   criteriaExpression As String - 标准表达式（如">=10"、"5~10"、"≤20"、"-5"）
'   value As Double              - 待校验数值
' 返回值：Boolean - True符合标准，False不符合
'======================================================================
Function cmpStd(criteriaExpression As String, value As Double) As Boolean
    Dim criteriaValue As Double
    Dim criteriaParts() As String
    Dim lowerLimit As Double
    Dim upperLimit As Double
    Dim expr As String
    
    ' 预处理：统一全角符号为半角，去除首尾空格
    expr = Trim(criteriaExpression)
    expr = Replace(expr, "≥", ">=")
    expr = Replace(expr, "≤", "<=")
    expr = Replace(expr, "＞", ">")
    expr = Replace(expr, "＜", "<")
    
    ' 处理大于等于
    If InStr(expr, ">=") = 1 Then
        criteriaValue = CDbl(Replace(expr, ">=", ""))
        cmpStd = (value >= criteriaValue)
    
    ' 处理小于等于
    ElseIf InStr(expr, "<=") = 1 Then
        criteriaValue = CDbl(Replace(expr, "<=", ""))
        cmpStd = (value <= criteriaValue)
    
    ' 处理大于
    ElseIf InStr(expr, ">") = 1 Then
        criteriaValue = CDbl(Replace(expr, ">", ""))
        cmpStd = (value > criteriaValue)
    
    ' 处理小于
    ElseIf InStr(expr, "<") = 1 Then
        criteriaValue = CDbl(Replace(expr, "<", ""))
        cmpStd = (value < criteriaValue)
    
    ' 处理范围（- 连接，排除单负号负数）
    ElseIf InStr(expr, "-") > 0 And Left(expr, 1) <> "-" Then
        criteriaParts = Split(expr, "-")
        lowerLimit = CDbl(Trim(criteriaParts(0)))
        upperLimit = CDbl(Trim(criteriaParts(1)))
        cmpStd = (value >= lowerLimit And value <= upperLimit)
    
    ' 处理范围（~ 连接）
    ElseIf InStr(expr, "~") > 0 Then
        criteriaParts = Split(expr, "~")
        lowerLimit = CDbl(Trim(criteriaParts(0)))
        upperLimit = CDbl(Trim(criteriaParts(1)))
        cmpStd = (value >= lowerLimit And value <= upperLimit)
    
    ' 处理单值等于/负数
    Else
        On Error Resume Next
        criteriaValue = CDbl(expr)
        If Err.Number = 0 Then
            cmpStd = (value = criteriaValue)
        Else
            cmpStd = False ' 无法解析的标准，返回不符合
        End If
        On Error GoTo 0
    End If
End Function
' 【校验自定义函数】cmpStd2
' 功能：逐单元格校验标准区域与数值区域，返回校验结果
' 参数：
'   StdRange     - 必选，标准区域（存放校验规则的单元格）
'   valueRange   - 必选，数值区域（需要校验的数值单元格）
'   returnType   - 可选，返回值类型：1=True/False 2=合格/不合格 3=符合/不符合 4=1/2，默认=1
' 返回：根据参数返回对应格式的校验结果
'======================================================================
Function cmpStd2(StdRange As Range, valueRange As Range, Optional returnType As Integer = 1) As Variant
    Dim maxCount As Long, std As String, value_ As Variant, value As Double
    Dim i As Long, iscankao2 As Boolean, isany As Boolean, isPass As Boolean
    Dim isEmptyCell As Boolean
    
    ' 初始化默认：合格
    isPass = True
    
    ' 错误处理
    On Error GoTo ex
    maxCount = Application.WorksheetFunction.Min(StdRange.Cells.Count, valueRange.Cells.Count)
    If maxCount = 0 Then GoTo ReturnResult

    ' 核心校验循环
    For i = 1 To maxCount
        std = Trim(StdRange.Cells(i).value)
        value_ = valueRange.Cells(i).value
        isEmptyCell = IsEmpty(value_)
        
        ' 数值转换
        If isEmptyCell Then
            value = 0
        Else
            On Error Resume Next
            value = CDbl(value_)
            If Err.Number <> 0 Then
                isPass = False
                GoTo continue
            End If
            On Error GoTo ex
        End If
        
        ' 规则1：IsNull 必须为0
        If InStr(std, "IsNull") > 0 Then
            If value <> 0 And Not isEmptyCell Then
                isPass = False
            End If
            GoTo continue
        End If
        
        ' 标记关键字
        iscankao2 = InStr(std, "参考2") > 0
        isany = InStr(std, "抽检") > 0 Or InStr(std, "自检") > 0
        
        ' 有效标准判断
        If (std <> "" And std <> "/" And std <> "-" _
            And InStr(std, "实测") = 0 And InStr(std, "监测") = 0 And InStr(std, "参考") = 0 _
            Or iscankao2) Then
            
            If value = 0 Or isEmptyCell Then GoTo continue
            
            ' 清理标准字符
            std = Replace(std, "NotNull", "")
            std = Replace(std, "抽检", "")
            std = Replace(std, "自检", "")
            std = Replace(std, "参考2", "")
            std = Replace(std, "g/L", "")
            std = Replace(std, "N", "")
            std = Replace(std, "%", "")
            std = Replace(std, "us/cm2", "")
            std = Replace(std, "(", "")
            std = Replace(std, ")", "")
            std = Replace(std, "（", "")
            std = Replace(std, "）", "")
            std = Trim(std)
            
            ' 对比标准
            If Not (std = "" Or Len(std) <= 1) Then
                If Not cmpStd(std, value) And Not iscankao2 Then
                    isPass = False
                End If
            End If
        End If
continue:
    Next i

' 根据参数返回对应格式结果
ReturnResult:
    Select Case returnType
        Case 1: cmpStd2 = isPass
        Case 2: cmpStd2 = IIf(isPass, "合格", "不合格")
        Case 3: cmpStd2 = IIf(isPass, "符合", "不符合")
        Case 4: cmpStd2 = IIf(isPass, 1, 2)
        Case Else: cmpStd2 = isPass
    End Select
    Exit Function

ex:
    cmpStd2 = "参数错误"
End Function



