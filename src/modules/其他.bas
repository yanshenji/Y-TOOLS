Attribute VB_Name = "其他"
Option Explicit

' 声明Windows API函数
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

' 常量定义
Private Const CF_TEXT = 1
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40

'将某个文件夹中的所有excel文件中的条件格式删除并且保存
Sub DeleteConditionalFormatting()
    ' 定义文件夹路径，这里以 "C:\YourFolder\" 为例，需替换为实际文件夹路径
    Dim folderPath As String
    folderPath = "C:\质量证书\碳酸钴\"

    ' 获取文件夹中所有 Excel 文件
    Dim fileName As String
    fileName = Dir(folderPath & "*.xls*")

    ' 遍历每个文件
    Do While fileName <> ""
        ' 拼接完整文件路径
        Dim filePath As String
        filePath = folderPath & fileName

        ' 打开工作簿
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath)

        ' 遍历工作簿中的每个工作表
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            ' 检查工作表中是否有条件格式
            If ws.Cells.FormatConditions.Count > 0 Then
                ' 删除工作表中的所有条件格式
                ws.Cells.FormatConditions.Delete
            End If
        Next ws

        ' 保存并关闭工作簿
        wb.Save
        wb.Close

        ' 获取下一个文件
        fileName = Dir
    Loop
    MsgBox "所有 Excel 文件中的条件格式已删除并保存！"
End Sub
Sub CurrentExcelToPDF()
    Dim currentWorkbook As Workbook
    Set currentWorkbook = ActiveWorkbook ' 获取当前工作簿
    
    Dim pdfPath As String
    ' 构造 PDF 文件路径，将文件扩展名改为 .pdf
    pdfPath = Left(currentWorkbook.FullName, InStrRev(currentWorkbook.FullName, ".") - 1) & ".pdf"
    
    ' 检查文件是否已经存在
    If Dir(pdfPath) <> "" Then
        MsgBox "PDF 文件已存在，是否要覆盖？", vbYesNo + vbQuestion, "文件存在"
        If MsgBox("PDF 文件已存在，是否要覆盖？", vbYesNo + vbQuestion, "文件存在") = vbYes Then
            Kill pdfPath ' 删除已存在的 PDF 文件
        Else
            Exit Sub ' 用户选择不覆盖，退出子程序
        End If
    End If
    
    ' 将工作簿导出为 PDF
    currentWorkbook.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    MsgBox "当前 Excel 文件已成功转换为 PDF 并保存在相同路径中！", vbInformation
End Sub




Sub ConvertExcelToPDF()
    ' 定义文件夹路径，需要转换的 Excel 文件所在文件夹路径
    Dim folderPath As String
    folderPath = "C:\质量证书\碳酸钴\" ' 替换为你的文件夹路径
    
    ' 定义输出文件夹路径，转换后的 PDF 文件将保存在这里
    Dim outputPath As String
    outputPath = "C:\质量证书\碳酸钴\" ' 替换为你希望保存 PDF 的文件夹路径
    
    ' 确保输出文件夹存在
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If
    
    ' 获取文件夹中所有 Excel 文件
    Dim fileName As String
    fileName = Dir(folderPath & "*.xls*")
    
    ' 检查文件夹中是否有文件
    If fileName = "" Then
        MsgBox "在指定文件夹中没有找到任何 Excel 文件。", vbExclamation
        Exit Sub
    End If
    
    ' 遍历每个文件
    Do While fileName <> ""
        ' 拼接完整文件路径
        Dim filePath As String
        filePath = folderPath & fileName
        
        ' 打开工作簿
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath)
        
        ' 定义输出 PDF 文件的完整路径
        Dim pdfPath As String
        pdfPath = outputPath & Left(fileName, InStrRev(fileName, ".") - 1) & ".pdf"
        
        ' 将工作簿导出为 PDF
        wb.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        ' 关闭工作簿
        wb.Close SaveChanges:=False
        
        ' 获取下一个文件
        fileName = Dir
    Loop
    
    MsgBox "所有 Excel 文件已成功转换为 PDF 并保存在指定文件夹中！", vbInformation
End Sub

Sub ExportCurrentSheetToPDF()
    Dim filePath As String
    Dim fileName As String
    
    ' 检查工作簿是否已保存（确定工作表的所在路径）
    If ActiveWorkbook.Path = "" Then
        MsgBox "请先保存当前工作簿（工作表的路径依赖于工作簿）", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 获取当前工作表所属工作簿的保存路径（即工作表的所在路径）
    filePath = ActiveWorkbook.Path & "\"
    
    ' 以当前工作表名称作为PDF文件名
    
    fileName = ActiveSheet.Name & ".pdf"
    
    If fileName Like "Sheet*" Then
        fileName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & ".pdf"
    End If
    
    
    ' 执行导出（导出到工作表所在的工作簿路径）
    On Error Resume Next
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath & fileName, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' 反馈结果
    If Err.Number = 0 Then
        MsgBox "PDF已导出至工作表所在路径：" & vbCrLf & filePath & fileName, vbInformation, "成功"
    Else
        MsgBox "导出失败：" & Err.Description, vbCritical, "错误"
    End If
    On Error GoTo 0
End Sub

Sub 连接可见单元格并复制()
    Dim selectedRange As Range
    Dim cell As Range
    Dim result As String
    
    ' 检查是否有选择区域
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "请先选择包含可见单元格的区域", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 遍历可见单元格并连接内容
    result = ""
    For Each cell In selectedRange
        If cell.value <> "" Then
            If result = "" Then
                result = cell.value
            Else
                result = result & "、" & cell.value
            End If
        End If
    Next cell
    
    If result <> "" Then
        ' 显示结果在输入框中，方便手动复制
        InputBox "请选中以下内容并按Ctrl+C复制：", "连接结果", result
    Else
        MsgBox "所选可见单元格中没有有效内容", vbExclamation, "提示"
    End If
    
    Set selectedRange = Nothing
End Sub

Sub ImprovedChemicalFormatter()
'自动识别化学公式中哪些部分应该是上下标
    Dim cell As Range
    Dim i As Integer
    Dim formulaText As String
    Dim char As String
    Dim nextChar As String
    Dim prevChar As String
    Dim j As Integer
    
    ' 检查是否有选中的单元格
    If Selection.Cells.Count = 0 Then
        MsgBox "请先选择包含化学式的单元格区域", vbExclamation
        Exit Sub
    End If
    
    ' 遍历选中区域的每个单元格
    For Each cell In Selection
        ' 跳过空单元格
        If cell.value = "" Then GoTo NextCell
        
        formulaText = cell.value
        
        ' 先重置所有字符的上下标格式
        cell.Font.Superscript = False
        cell.Font.Subscript = False
        
        ' 处理下标：元素符号(字母)后面的数字通常是下标
        For i = 1 To Len(formulaText)
            char = Mid(formulaText, i, 1)
            
            ' 如果当前字符是数字
            If IsNumeric(char) Then
                ' 检查前一个字符是否是字母(元素符号)
                If i > 1 Then
                    prevChar = Mid(formulaText, i - 1, 1)
                    If (prevChar >= "A" And prevChar <= "Z") Or (prevChar >= "a" And prevChar <= "z") Then
                        ' 检查后面是否有+或-，如果有则当前数字可能是上标
                        If i < Len(formulaText) Then
                            nextChar = Mid(formulaText, i + 1, 1)
                            If nextChar <> "+" And nextChar <> "-" Then
                                ' 设置为下标
                                cell.Characters(Start:=i, Length:=1).Font.Subscript = True
                            End If
                        Else
                            ' 设置为下标
                            cell.Characters(Start:=i, Length:=1).Font.Subscript = True
                        End If
                    End If
                End If
            End If
        Next i
        
        ' 处理上标：电荷符号及前面的数字
        For i = 1 To Len(formulaText)
            char = Mid(formulaText, i, 1)
            
            ' 如果当前字符是+或-（电荷符号）
            If char = "+" Or char = "-" Then
                ' 设置符号为上标
                cell.Characters(Start:=i, Length:=1).Font.Superscript = True
                
                ' 检查前面的数字是否应该是上标（电荷数）
                j = i - 1
                Do While j >= 1 And IsNumeric(Mid(formulaText, j, 1))
                    ' 确保这个数字不是元素符号后的下标
                    If j = 1 Or Not ((Mid(formulaText, j - 1, 1) >= "A" And Mid(formulaText, j - 1, 1) <= "Z") Or _
                                   (Mid(formulaText, j - 1, 1) >= "a" And Mid(formulaText, j - 1, 1) <= "z")) Then
                        cell.Characters(Start:=j, Length:=1).Font.Superscript = True
                    Else
                        Exit Do ' 如果是元素符号后的数字，则是下标，退出循环
                    End If
                    j = j - 1
                Loop
            End If
        Next i
        
NextCell:
    Next cell
    
    MsgBox "化学式上下标格式处理完成！", vbInformation
End Sub

Sub 隐藏所有工作表CE列()
    Dim ws As Worksheet  ' 声明工作表对象变量
    
    ' 遍历活动工作簿中的所有工作表
    For Each ws In ActiveWorkbook.Worksheets
        ' 检查是否为工作表类型（排除图表等其他类型）
        If ws.Type = xlWorksheet Then
            ' 隐藏C到E列（Columns("C:E")表示从C列到E列的连续区域）
            ws.Columns("C:E").Hidden = True
        End If
    Next ws
    
    ' 操作完成提示
    MsgBox "已完成所有工作表C到E列的隐藏操作！", vbInformation, "提示"
End Sub
Sub 统计可见单元格数值并复制()
    Dim selectedRange As Range
    Dim col As Range
    Dim cell As Range
    Dim valArr() As Double
    Dim arrIndex As Integer
    Dim avgVal As Double, maxVal As Double, minVal As Double
    Dim result As String
    Dim colHeader As String ' 存储列的表头文本
    
    ' 检查是否选择了区域
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选择需要统计的列（包含表头）", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 遍历选中的每一列
    For Each col In Selection.Columns
        ' 获取当前列的可见单元格（排除隐藏/筛选行）
        On Error Resume Next
        Set selectedRange = col.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' 获取当前列的表头（第一行单元格内容）
        colHeader = col.Cells(1).value
        ' 处理表头为空的情况
        If colHeader = "" Then colHeader = "未命名列"
        
        If Not selectedRange Is Nothing Then
            ' 初始化数组索引
            arrIndex = 0
            ReDim valArr(0)
            
            ' 遍历可见单元格，筛选大于0的数值（跳过表头行）
            For Each cell In selectedRange
                ' 跳过表头行（第一行），仅处理数据行
                If cell.Row > 1 And IsNumeric(cell.value) And cell.value > 0 Then
                    ReDim Preserve valArr(arrIndex)
                    valArr(arrIndex) = cell.value
                    arrIndex = arrIndex + 1
                End If
            Next cell
            
            ' 生成统计结果
            If arrIndex > 0 Then
                avgVal = Application.WorksheetFunction.Average(valArr)
                maxVal = Application.WorksheetFunction.Max(valArr)
                minVal = Application.WorksheetFunction.Min(valArr)
                
                ' 拼接结果（用表头文本作为前缀）
                If result = "" Then
                    result = colHeader & " 均值" & Format(avgVal, "0.0000") & "g/L(" & _
                             Format(minVal, "0.0000") & "-" & Format(maxVal, "0.0000") & "g/L)"
                Else
                    result = result & vbCrLf & colHeader & " 均值" & Format(avgVal, "0.0000") & "g/L(" & _
                             Format(minVal, "0.0000") & "-" & Format(maxVal, "0.0000") & "g/L)"
                End If
            Else
                ' 无有效数据提示
                result = IIf(result = "", colHeader & " 无大于0的可见数值", result & vbCrLf & colHeader & " 无大于0的可见数值")
            End If
        Else
            ' 无可视单元格提示
            result = IIf(result = "", colHeader & " 无可视单元格", result & vbCrLf & colHeader & " 无可视单元格")
        End If
    Next col
    
    ' 复制结果到剪贴板
    If result <> "" Then
        Dim objData As Object
        Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        objData.SetText result
        objData.PutInClipboard
        Set objData = Nothing
        
        ' 弹窗展示结果（确认后可直接粘贴）
        InputBox "统计结果已自动复制到剪贴板：", "统计完成", result
    Else
        MsgBox "未找到有效数据", vbExclamation, "提示"
    End If
    
    ' 释放对象
    Set selectedRange = Nothing
End Sub


Function 保留两位有效数字_四舍六入五成双(num As Double) As Double
    Dim exponent As Integer
    Dim mantissa As Double
    Dim thirdDigit As Integer
    Dim secondDigit As Integer
    
    If num = 0 Then
        保留两位有效数字_四舍六入五成双 = 0
        Exit Function
    End If
    
    ' 计算指数
    exponent = Int(Log(Abs(num)) / Log(10))
    
    ' 计算尾数
    mantissa = num / (10 ^ exponent)
    
    ' 获取第三位数字用于判断
    thirdDigit = CInt(Mid(CStr(mantissa * 1000), 4, 1))
    secondDigit = CInt(Mid(CStr(mantissa * 100), 3, 1))
    
    ' 四舍六入五成双规则
    If thirdDigit < 5 Then
        ' 小于5，舍去
        mantissa = Int(mantissa * 10) / 10
    ElseIf thirdDigit > 5 Then
        ' 大于5，进1
        mantissa = (Int(mantissa * 10) + 1) / 10
    Else
        ' 等于5的情况
        If secondDigit Mod 2 = 0 Then
            ' 第二位是偶数，舍去
            mantissa = Int(mantissa * 10) / 10
        Else
            ' 第二位是奇数，进1
            mantissa = (Int(mantissa * 10) + 1) / 10
        End If
    End If
    
    ' 返回数值结果
    保留两位有效数字_四舍六入五成双 = mantissa * (10 ^ exponent)
End Function

' 四舍五入版本的保留两位有效数字函数
Function 保留两位有效数字_四舍五入(num As Double) As Double
    Dim exponent As Integer
    Dim mantissa As Double
    
    If num = 0 Then
        保留两位有效数字_四舍五入 = 0
        Exit Function
    End If
    
    ' 计算指数
    exponent = Int(Log(Abs(num)) / Log(10))
    
    ' 计算尾数
    mantissa = num / (10 ^ exponent)
    
    ' 四舍五入到两位有效数字
    mantissa = Round(mantissa, 1)
    
    ' 返回数值结果
    保留两位有效数字_四舍五入 = mantissa * (10 ^ exponent)
End Function


Sub 检测数据转窄表()

    Dim srcSht As Worksheet, destSht As Worksheet
    Dim headerRow As Long, lastCol As Long, lastRow As Long
    Dim keepFields As Variant
    Dim keepColIndex As Variant
    Dim testColIndex As Variant
    Dim i As Long, j As Long, outRow As Long, k As Long
    Dim arrKeep() As Variant
    Dim destSheetName As String
    Dim badChars As Variant
    Dim ch As Variant
    
    Set srcSht = ActiveSheet
    headerRow = 1
    
    '=====固定保留字段，不要修改顺序=====
    keepFields = Array("送样日期", "送样时间", "报出时间", "分析时长", _
                       "复核人", "样品名称", "录入人", "判定", "样品类别")
    
    '获取源表边界
    lastCol = srcSht.Cells(headerRow, srcSht.Columns.Count).End(xlToLeft).Column
    lastRow = srcSht.Cells(srcSht.Rows.Count, 1).End(xlUp).Row
    
    '去除sheet名称非法字符 \ / ? * [ ] :
    badChars = Array("\", "/", ":", "*", "?", "[", "]")
    destSheetName = srcSht.Name
    For Each ch In badChars
        destSheetName = Replace(destSheetName, ch, "_")
    Next ch
    destSheetName = destSheetName & "窄表"
    'sheet名最长31字符
    If Len(destSheetName) > 31 Then destSheetName = Left(destSheetName, 31)
    
    '尝试获取或新建sheet
    On Error Resume Next
    Set destSht = ThisWorkbook.Worksheets(destSheetName)
    On Error GoTo 0
    
    If destSht Is Nothing Then
        On Error Resume Next
        Set destSht = ThisWorkbook.Worksheets.Add(After:=srcSht)
        destSht.Name = destSheetName
        On Error GoTo 0
    End If
    
    '新建sheet失败的兜底方案：输出到全新独立工作簿
    If destSht Is Nothing Then
        Dim newWb As Workbook
        Set newWb = Workbooks.Add(xlWBATWorksheet)
        Set destSht = newWb.Sheets(1)
        MsgBox "无法在当前文件新建工作表，结果输出至【新工作簿】", vbExclamation
    End If
    
    destSht.Cells.Clear
    
    '====区分：保留列下标、检测项目列下标====
    ReDim keepColIndex(0 To UBound(keepFields))
    ReDim testColIndex(0 To lastCol - 1)
    Dim testCnt As Long: testCnt = 0
    Dim keepCnt As Long: keepCnt = UBound(keepFields) + 1
    
    For j = 1 To lastCol
        Dim hText As String
        hText = Trim(srcSht.Cells(headerRow, j).value)
        Dim isKeep As Boolean: isKeep = False
        For k = 0 To UBound(keepFields)
            If hText = keepFields(k) Then
                keepColIndex(k) = j
                isKeep = True
                Exit For
            End If
        Next k
        If Not isKeep Then
            testColIndex(testCnt) = j
            testCnt = testCnt + 1
        End If
    Next j
    
    '写入窄表表头
    For k = 0 To UBound(keepFields)
        destSht.Cells(1, k + 1).value = keepFields(k)
    Next k
    destSht.Cells(1, keepCnt + 1).value = "检测项目"
    destSht.Cells(1, keepCnt + 2).value = "检测结果"
    outRow = 2
    
    '循环每一条样品数据
    For i = 2 To lastRow
        ReDim arrKeep(0 To keepCnt - 1)
        For k = 0 To keepCnt - 1
            arrKeep(k) = srcSht.Cells(i, keepColIndex(k)).value
        Next k
        
        For j = 0 To testCnt - 1
            Dim colNo As Long
            colNo = testColIndex(j)
            Dim resVal As Variant
            resVal = srcSht.Cells(i, colNo).value
            
            '如果需要保留空行，注释If和End If
            If Not IsEmpty(resVal) Then
                For k = 0 To keepCnt - 1
                    destSht.Cells(outRow, k + 1).value = arrKeep(k)
                Next k
                destSht.Cells(outRow, keepCnt + 1).value = Trim(srcSht.Cells(headerRow, colNo).value)
                destSht.Cells(outRow, keepCnt + 2).value = resVal
                outRow = outRow + 1
            End If
        Next j
    Next i
    
    destSht.Columns.AutoFit
    MsgBox "转换完成！共生成 " & outRow - 2 & " 条检测行", vbInformation
End Sub




' ============================================================
' Grubbs 异常值检验 (Grubbs' Test for Outliers)
' 适用：正态分布小样本中的异常值检测
' 公式：G = max(|xi - mean|) / stdev
' 临界值：G(α,n) = (n-1)/√n * √(t2/(n-2+t2))
'   其中 t = T_INV.2T(α/n, n-2)，双侧检验
' ============================================================


' ------------------------------------------------------------
' 1. 主函数：返回数据集中所有异常值
'    用法：在工作表中 =GrubbsOutliers(A1:A20, 0.05)
'    或在 VBA 中调用
' ------------------------------------------------------------
Public Function GrubbsOutliers(dataRange As Range, Optional alpha As Double = 0.05) As Variant
    Dim rawArr As Variant
    rawArr = dataRange.value
    
    ' 提取有效数值到一维数组
    Dim data() As Double
    Dim n As Long
    n = 0
    
    Dim i As Long
    For i = 1 To UBound(rawArr, 1)
        If IsNumeric(rawArr(i, 1)) And Not IsEmpty(rawArr(i, 1)) Then
            n = n + 1
        End If
    Next i
    
    If n < 3 Then
        GrubbsOutliers = "错误：至少需要 3 个数据点"
        Exit Function
    End If
    
    ReDim data(1 To n)
    Dim j As Long: j = 0
    For i = 1 To UBound(rawArr, 1)
        If IsNumeric(rawArr(i, 1)) And Not IsEmpty(rawArr(i, 1)) Then
            j = j + 1
            data(j) = CDbl(rawArr(i, 1))
        End If
    Next i
    
    ' 迭代检测：每次移除一个最极端值，直到不再显著
    Dim outliers As Collection
    Set outliers = New Collection
    
    Dim meanVal As Double, stdVal As Double
    Dim gCalc As Double, Gcrit As Double
    Dim maxIdx As Long
    Dim gTemp As Double
    Dim removed As Boolean
    
    Do
        removed = False
        n = UBound(data)
        If n <= 3 Then Exit Do
        
        meanVal = WorksheetFunction.Average(data)
        stdVal = WorksheetFunction.StDev_S(data)
        
        If stdVal = 0 Then Exit Do   ' 所有值相同，无法检验
        
        gCalc = 0
        maxIdx = 0
        For i = 1 To n
            gTemp = Abs(data(i) - meanVal) / stdVal
            If gTemp > gCalc Then
                gCalc = gTemp
                maxIdx = i
            End If
        Next i
        
        Gcrit = GrubbsCrit(n, alpha)
        
        If gCalc > Gcrit Then
            outliers.Add data(maxIdx)
            data = RemoveAt(data, maxIdx)
            removed = True
        End If
    Loop While removed
    
    ' 输出结果
    If outliers.Count = 0 Then
        GrubbsOutliers = "未检测到异常值"
    Else
        Dim res() As Double
        ReDim res(1 To outliers.Count, 1 To 1)
        For i = 1 To outliers.Count
            res(i, 1) = outliers(i)
        Next i
        GrubbsOutliers = res
    End If
End Function


' ------------------------------------------------------------
' 2. 计算单个数据点的 G 值
'    用法：=GrubbsG(A1, $A$1:$A$20)
' ------------------------------------------------------------
Public Function GrubbsG(singleValue As Double, dataRange As Range) As Variant
    Dim meanVal As Double, stdVal As Double
    meanVal = WorksheetFunction.Average(dataRange)
    stdVal = WorksheetFunction.StDev_S(dataRange)
    
    If stdVal = 0 Then
        GrubbsG = "标准差为 0"
    Else
        GrubbsG = Abs(singleValue - meanVal) / stdVal
    End If
End Function


' ------------------------------------------------------------
' 3. 返回 Grubbs 临界值
'    用法：=GrubbsCrit(20, 0.05)
' ------------------------------------------------------------
Public Function GrubbsCrit(n As Long, Optional alpha As Double = 0.05) As Variant
    If n <= 2 Then
        GrubbsCrit = "n 必须大于 2"
        Exit Function
    End If
    
    Dim tCrit As Double
    ' 双侧 Grubbs 检验：t 分布自由度 n-2，尾部概率 α/n
    On Error Resume Next
    tCrit = WorksheetFunction.T_Inv_2T(alpha / n, n - 2)
    On Error GoTo 0
    
    If tCrit = 0 Then
        ' 兼容旧版 Excel：TINV 即为双侧 t 分位数
        tCrit = WorksheetFunction.TInv(alpha / n, n - 2)
    End If
    
    GrubbsCrit = ((n - 1) / Sqr(n)) * Sqr(tCrit ^ 2 / (n - 2 + tCrit ^ 2))
End Function


' ------------------------------------------------------------
' 4. 辅助函数：从一维数组中移除指定位置的元素
' ------------------------------------------------------------
Private Function RemoveAt(arr() As Double, idx As Long) As Double()
    Dim res() As Double
    Dim i As Long, k As Long
    ReDim res(1 To UBound(arr) - 1)
    k = 0
    For i = 1 To UBound(arr)
        If i <> idx Then
            k = k + 1
            res(k) = arr(i)
        End If
    Next i
    RemoveAt = res
End Function


' ------------------------------------------------------------
' 5. 演示宏：选中数据区域后运行，结果输出到立即窗口
'    （按 Ctrl+G 在 VBA 编辑器中查看立即窗口）
' ------------------------------------------------------------
Public Sub DemoGrubbs()
    Dim sel As Range
    Set sel = Selection
    
    If sel Is Nothing Then
        MsgBox "请先选中一列数据", vbExclamation
        Exit Sub
    End If
    
    Dim alpha As Double
    alpha = 0.05
    
    Dim result As Variant
    result = GrubbsOutliers(sel, alpha)
    
    Debug.Print "===== Grubbs 检验结果 (α=" & alpha & ") ====="
    Debug.Print "样本量: " & sel.Cells.Count
    
    If IsArray(result) Then
        Dim i As Long
        Debug.Print "检测到 " & UBound(result) & " 个异常值："
        For i = 1 To UBound(result)
            Debug.Print "  #" & i & "  " & result(i, 1)
        Next i
    Else
        Debug.Print result
    End If
    
    Debug.Print "============================================="
End Sub


' ------------------------------------------------------------
' 6. 批量标注宏：在数据右侧一列标注每个点是否为异常值
'    假设数据从 A1 开始纵向排列
' ------------------------------------------------------------
Public Sub FlagOutliersInSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBox "A 列数据不足 3 行", vbExclamation
        Exit Sub
    End If
    
    Dim dataRng As Range
    Set dataRng = ws.Range("A1:A" & lastRow)
    
    Dim outliers As Collection
    Set outliers = New Collection
    
    ' 先跑一遍完整检验拿到异常值列表
    Dim res As Variant
    res = GrubbsOutliers(dataRng, 0.05)
    
    If IsArray(res) Then
        Dim i As Long
        For i = 1 To UBound(res)
            outliers.Add res(i, 1)
        Next i
    End If
    
    ' 写表头
    ws.Range("B1").value = "Grubbs 判定"
    
    ' 逐行标注
    Dim cellVal As Double
    Dim j As Long
    Dim isOutlier As Boolean
    
    For i = 1 To lastRow
        cellVal = ws.Cells(i, "A").value
        isOutlier = False
        For j = 1 To outliers.Count
            If cellVal = outliers(j) Then
                isOutlier = True
                Exit For
            End If
        Next j
        ws.Cells(i, "B").value = IIf(isOutlier, "异常值", "正常")
    Next i
    
    MsgBox "标注完成，共 " & outliers.Count & " 个异常值", vbInformation
End Sub

Sub GrubbsMarkYellow()
    '格拉布斯Grubbs检验，选中区域，离群异常值单元格标黄
    Dim rng As Range, cell As Range
    Dim arr() As Double
    Dim n As Long, i As Long
    Dim xBar As Double, S As Double
    Dim G As Double, Gcrit As Double
    Dim alpha As Double
    Dim tVal As Double
    
    alpha = 0.05   '显著性水平，可改为0.01
    Set rng = Selection
    If rng.Areas.Count > 1 Then
        MsgBox "请选择连续一块数据区域！"
        Exit Sub
    End If
    
    '读取有效数值
    ReDim arr(1 To rng.Cells.Count)
    n = 0
    For Each cell In rng
        If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
            n = n + 1
            arr(n) = cell.value
        End If
    Next cell
    If n < 3 Then
        MsgBox "有效样本数必须≥3！"
        Exit Sub
    End If
    ReDim Preserve arr(1 To n)
    
    '均值、样本标准差
    xBar = WorksheetFunction.Average(arr)
    S = WorksheetFunction.StDev_S(arr)
    If S < 0.000000000001 Then Exit Sub '全部数据相同无离群
    
    'Grubbs临界值公式 Gcrit=(n?1)/Sqr(n)*Sqr(t^2/(n?2+t^2))
    tVal = WorksheetFunction.T_Inv_2T(alpha / n, n - 2)
    Gcrit = (n - 1) / Sqr(n) * Sqr(tVal ^ 2 / ((n - 2) + tVal ^ 2))
    
    '遍历原选区，计算每个点G统计量，大于临界值标黄
    For Each cell In rng
        If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
            G = Abs(cell.value - xBar) / S
            If G > Gcrit Then
                cell.Interior.Color = RGB(255, 255, 0) '黄色
            Else
                cell.Interior.ColorIndex = xlColorIndexNone '清除底色
            End If
        Else
            cell.Interior.ColorIndex = xlColorIndexNone
        End If
    Next cell
    
End Sub

