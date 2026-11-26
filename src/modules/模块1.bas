Attribute VB_Name = "模块1"
Sub 筛选当前列为当前单元格值()
    Dim rngCurrentCell As Range
    Dim strCurrentCellValue As String

    ' 获取当前单元格
    Set rngCurrentCell = Selection
    ' 获取当前单元格的值
    strCurrentCellValue = rngCurrentCell.value

    ' 筛选当前列为当前单元格值
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=rngCurrentCell.Column, Criteria1:="=" & strCurrentCellValue
End Sub

Sub 筛选当前列非空值()
    Dim lastRow As Long

    ' 筛选非空值
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=Selection.Column, Criteria1:="<>"
End Sub
Sub 筛选当前列数值()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
     Dim maxNum As Double
    ' 获取当前活动工作表
    Set ws = ActiveSheet

    ' 获取当前选中单元格所在的列
    Dim currentColumn As Long
    currentColumn = Selection.Column

    ' 确定筛选范围（从第一行到最后一行，当前列）
    lastRow = ws.Cells(ws.Rows.Count, currentColumn).End(xlUp).Row
    Set rng = ws.Range(ws.Cells(1, currentColumn), ws.Cells(lastRow, currentColumn))

     ' 找到当前列的最大值
    maxNum = Application.WorksheetFunction.Max(rng)
    ' 筛选非空值
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=Selection.Column, Criteria1:=">0", Operator:=xlAnd, Criteria2:="<=" & maxNum
End Sub

Sub 筛选最大日期近30天的数据()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxDate As Date
    Dim dateRange As Range

    Set ws = ActiveSheet ' 或指定工作表名称
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' 获取第一列最后一行的行号
    Set dateRange = ws.Range("A1:A" & lastRow) ' 包含标题行的日期范围

    
    ' 找出最大日期
    maxDate = Application.WorksheetFunction.Max(ws.Range("A2:A" & lastRow)) ' 假设数据从第2行开始

    ' 筛选近30天的数据
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=1, _
        Criteria1:=">=" & DateAdd("d", -30, maxDate), _
        Criteria2:="<=" & maxDate, _
        Operator:=xlAnd

End Sub

Sub 清除之前的筛选()
' 清除之前的筛选
    On Error Resume Next
    ActiveSheet.AutoFilterMode = False
    On Error GoTo 0
End Sub

Sub 清除当前列筛选()
    Dim rngCurrentCell As Range

    ' 获取当前单元格
    Set rngCurrentCell = Selection
     ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=rngCurrentCell.Column
End Sub

Sub 保留最后50条可见的数据()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim visibleRows As Long
    Dim i As Long
    Dim n As Integer
    
    ' 设置当前工作表
    Set ws = ActiveSheet
    
    ' 找到当前列最后一个有数据的行
    lastRow = ws.Cells(ws.Rows.Count, ActiveCell.Column).End(xlUp).Row
    
    ' 确定当前筛选后的可见行数
    visibleRows = ws.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1 ' 减去标题行
    
    ' 如果可见行数大于50，则隐藏多余的行
    If visibleRows > 50 Then
        ' 定义要操作的范围
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(2, ActiveCell.Column), ws.Cells(lastRow, ActiveCell.Column)) ' 从第二行开始，假设第一行是标题行
        
        ' 隐藏除最后50条可见数据之外的所有行
        For i = 2 To lastRow
            If ws.Rows(i).Hidden = False Then
                ws.Rows(i).Hidden = True
                n = n + 1
                ' 如果隐藏的行数达到 visibleRows - 50，则退出循环
                If n >= visibleRows - 50 Then Exit For
            End If
        Next i
    End If
    
    MsgBox "操作完成！只保留了最后50条可见数据。", vbInformation
End Sub
Sub 筛选本月数据()
 Dim ws As Worksheet
    Dim lastRow As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim dateColumn As Range
    Dim cell As Range
    Dim i As Long
    Dim visibleCount As Long

    ' 设置当前工作表
    Set ws = ActiveSheet

    ' 检查单元格(1,1)是否为“送样日期”
    'If ws.Cells(1, 1).Value <> "送样日期" And ws.Cells(1, 1).Value <> "取样日期" Then
        'MsgBox "单元格(1,1)不是'送样日期'，无法进行筛选！", vbExclamation
        'Exit Sub
    'End If

    ' 计算筛选日期范围
    startDate = DateSerial(Year(Date), Month(Date) - 1, 28) ' 上月29日
    endDate = DateSerial(Year(Date), Month(Date), 28) ' 本月28日

    ' 找到第一列最后一个有数据的行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 定义要筛选的范围
    Set dateColumn = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)) ' 从第二行开始，假设第一行是标题行
     ' 应用筛选
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=1, Criteria1:=">" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
End Sub
Sub 筛选上月数据()
 Dim ws As Worksheet
    Dim lastRow As Long
    Dim startDate As Date
    Dim endDate As Date
    Dim dateColumn As Range
    Dim cell As Range
    Dim i As Long
    Dim visibleCount As Long

    ' 设置当前工作表
    Set ws = ActiveSheet

    ' 检查单元格(1,1)是否为“送样日期”
    'If ws.Cells(1, 1).Value <> "送样日期" And ws.Cells(1, 1).Value <> "取样日期" Then
        'MsgBox "单元格(1,1)不是'送样日期'，无法进行筛选！", vbExclamation
        'Exit Sub
    'End If

    ' 计算筛选日期范围
    startDate = DateSerial(Year(Date), Month(Date) - 2, 28) ' 上月29日
    endDate = DateSerial(Year(Date), Month(Date) - 1, 28) ' 本月28日

    ' 找到第一列最后一个有数据的行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 定义要筛选的范围
    Set dateColumn = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)) ' 从第二行开始，假设第一行是标题行
     ' 应用筛选
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=1, Criteria1:=">" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
End Sub

Sub 隐藏符合条件的列()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 或者设置为特定工作表，如 ThisWorkbook.Worksheets("Sheet1")
    
    ' 获取实际使用的列数
    Dim lastColumn As Long
    lastColumn = ws.usedRange.Columns.Count
    
    Application.ScreenUpdating = False ' 禁用屏幕刷新
    Dim colIndex As Long
    For colIndex = 1 To lastColumn
        ' 假设工作表中数据从第 1 行到第 100 行，根据实际情况修改行范围
        Dim isAllEmptyExceptFirstRow As Boolean
        isAllEmptyExceptFirstRow = True
        
        Dim rowIndex As Long
        For rowIndex = 2 To ws.usedRange.Rows.Count
            ' 检查除第一行外的单元格是否为空
            If ws.Cells(rowIndex, colIndex).value <> "" Then
                isAllEmptyExceptFirstRow = False
                Exit For
            End If
        Next rowIndex
        
        ' 如果除第一行外其他行全为空，则隐藏该列
        If isAllEmptyExceptFirstRow Then
            ws.Columns(colIndex).EntireColumn.Hidden = True
        End If
    Next colIndex
    Application.ScreenUpdating = True ' 恢复屏幕刷新
End Sub


Sub 根据Ctrl向下查找隐藏列()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 或者设置为特定工作表，如 ThisWorkbook.Worksheets("Sheet1")
    
    Application.ScreenUpdating = False ' 禁用屏幕刷新
    
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    Dim lastColumn As Long
    lastColumn = usedRange.Columns.Count
    
    Dim colIndex As Long
    For colIndex = 1 To lastColumn
        ' 获取当前列中最后一个非空单元格的行号（模拟 Ctrl + 向下查找）
        Dim lastCellRow As Long
        lastCellRow = ws.Cells(1, colIndex).End(xlDown).Row
        
        ' 检查该行号是否大于 UsedRange 的行数
        If lastCellRow > usedRange.Rows.Count Then
            ws.Columns(colIndex).Hidden = True
        End If
    Next colIndex
    
    Application.ScreenUpdating = True ' 恢复屏幕刷新
End Sub

Sub 隐藏符合条件的列2()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 或者设置为特定工作表，如 ThisWorkbook.Worksheets("Sheet1")
    
    Application.ScreenUpdating = False ' 禁用屏幕刷新
    
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    Dim lastColumn As Long
    lastColumn = usedRange.Columns.Count
    
    Dim colIndex As Long
    For colIndex = 1 To lastColumn
        Dim isAllEmptyExceptFirstRow As Boolean
        isAllEmptyExceptFirstRow = True
        
        Dim rowIndex As Long
        For rowIndex = 2 To usedRange.Rows.Count
            If ws.Cells(rowIndex, colIndex).value <> "" Then
                isAllEmptyExceptFirstRow = False
                Exit For
            End If
        Next rowIndex
        
        If isAllEmptyExceptFirstRow Then
            ws.Columns(colIndex).Hidden = True
        End If
    Next colIndex
    
    Application.ScreenUpdating = True ' 恢复屏幕刷新
End Sub

Sub 隐藏符合条件的列_选中()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 或者设置为特定工作表，如 ThisWorkbook.Worksheets("Sheet1")
     ' 【新增】保存Excel原始计算模式
    Dim originalCalculation As XlCalculation
    originalCalculation = Application.Calculation
    ' 【关键】临时切换为手动计算，大幅提速
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False ' 禁用屏幕刷新
     
    
    ' 获取选中的范围
    Dim selectedRange As Range
    Set selectedRange = Selection
    
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    Dim lastColumn As Long
    'lastColumn = usedRange.Columns.Count
    
    lastColumn = ws.Cells(1, 16384).End(xlToLeft).Column
    ' 获取选中范围的起始列索引
    Dim startColumn As Long
    startColumn = selectedRange.Column
    
    Dim colIndex As Long
    For colIndex = 1 To lastColumn
        ' 当前选中列的绝对列索引
        Dim currentColumn As Long
        currentColumn = startColumn + colIndex - 1
        
        ' 假设选中范围中数据从第 1 行到选中范围的最后一行
        Dim isAllEmptyExceptFirstRow As Boolean
        isAllEmptyExceptFirstRow = True
        
        Dim rowIndex As Long
        For rowIndex = 1 To selectedRange.Rows.Count
            ' 检查选中范围内的单元格是否为空
            If ws.Cells(selectedRange.Row + rowIndex - 1, currentColumn).value <> "" Then
                isAllEmptyExceptFirstRow = False
                Exit For
            End If
        Next rowIndex
        
        ' 如果除第一行外其他行全为空，则隐藏该列
        If isAllEmptyExceptFirstRow Then
            ws.Columns(currentColumn).Hidden = True
        End If
    Next colIndex
    Application.ScreenUpdating = True ' 恢复屏幕刷新
    ' 【关键】恢复成你原来的计算模式（自动/手动都能复原）
    Application.Calculation = originalCalculation
    
    
End Sub

Sub 隐藏符合条件的行_选中()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 或者设置为特定工作表，如 ThisWorkbook.Worksheets("Sheet1")
    
    'Application.ScreenUpdating = False ' 禁用屏幕刷新
    
    ' 获取选中的范围
    Dim selectedRange As Range
    Set selectedRange = Selection
    
    Dim usedRange As Range
    Set usedRange = ws.usedRange
    
    Dim lastRow As Long
    'lastColumn = usedRange.Columns.Count
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' 获取选中范围的起始列索引
    Dim startColumn As Long
    startColumn = selectedRange.Column
    
    Dim rowIndex As Long
    For rowIndex = 1 To lastRow
        ' 当前选中列的绝对列索引
       
        ' 假设选中范围中数据从第 1 行到选中范围的最后一行
        Dim isAllEmptyExceptFirstRow As Boolean
        isAllEmptyExceptFirstRow = True
        
        Dim colIndex As Long
        For colIndex = 1 To selectedRange.Columns.Count
            ' 检查选中范围内的单元格是否为空
            If ws.Cells(rowIndex, selectedRange.Column + colIndex - 1).value <> "" Then
                isAllEmptyExceptFirstRow = False
                Exit For
            End If
        Next colIndex
        
        ' 如果除第一行外其他行全为空，则隐藏该列
        If isAllEmptyExceptFirstRow Then
            ws.Rows(rowIndex).Hidden = True
        End If
    Next rowIndex
    
    Application.ScreenUpdating = True ' 恢复屏幕刷新
End Sub
Sub 取消隐藏列()
     Dim ws As Worksheet
    ' 设置工作表
    Set ws = ActiveSheet
    ' 取消工作表中所有列的隐藏状态
    ws.Cells.EntireColumn.Hidden = False
    'MsgBox "所有列已取消隐藏！"
End Sub


Function one()
    one = 1
End Function
