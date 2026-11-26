Attribute VB_Name = "模块2"
Sub 标记异常_Boxplot()
    Dim ws As Worksheet
    Dim cell As Range
    Dim Q1 As Double, Q3 As Double, IQR As Double
    Dim lowerBound As Double, upperBound As Double
    Dim colIndex As Integer
    Dim lastRow As Long
    Dim currentColumn As Range
    Dim visibleCells As Range
    Dim maxValue As Double
    Dim selectedColumns As Range
    Dim area As Range
    Dim visibleColumns As Range
    Dim col As Range
    
    ' 设置工作表
    Set ws = ActiveSheet
    
    ' 初始化可见列范围
    Set visibleColumns = Nothing
    
    ' 遍历选中列的每个列，只保留可见的列
    For Each col In Selection.Columns
        If Not col.Hidden Then ' 判断列是否可见
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    ' 如果没有可见列，提示用户并退出
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    ' 获取可见列与数据区域的交集
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False ' 暂时关闭屏幕更新，提高效率
    
    ' 遍历每一列
    For Each area In selectedColumns.Areas
        For colIndex = 1 To area.Columns.Count
            ' 获取当前列有数据的最后一行行号
            lastRow = ws.Cells(ws.Rows.Count, area.Columns(colIndex).Column).End(xlUp).Row
            
            ' 获取当前列从第一行到最后一行有数据的单元格范围
            Set currentColumn = ws.Range(ws.Cells(1, area.Columns(colIndex).Column), ws.Cells(lastRow, area.Columns(colIndex).Column))
            
            ' 获取当前列的可见单元格
            On Error Resume Next ' 忽略错误，防止没有筛选时出错
            Set visibleCells = currentColumn.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If visibleCells Is Nothing Then
                ' 如果没有可见单元格，跳过该列
                GoTo ContinueNextColumn
            End If
            
            ' 计算四分位数和最大值
            On Error Resume Next
            Q1 = Application.WorksheetFunction.Quartile(visibleCells, 1)
            Q3 = Application.WorksheetFunction.Quartile(visibleCells, 3)
            On Error GoTo 0
            
            IQR = Q3 - Q1
            maxValue = Application.WorksheetFunction.Max(visibleCells)
            
            ' 计算异常值的边界
            lowerBound = Application.WorksheetFunction.Max(Q1 - 1.5 * IQR, 0.1 * Q1)
            upperBound = Q3 + 1.5 * IQR
            
            ' 遍历可见单元格，标记异常值
            For Each cell In visibleCells
                If (cell.value < lowerBound And cell.value >= 0) Or (cell.value > upperBound And cell.value <= maxValue) Then
                    cell.Interior.Color = RGB(255, 255, 0) ' 填充为黄色
                Else
                    cell.Interior.ColorIndex = xlNone ' 清除填充颜色
                End If
            Next cell
            
ContinueNextColumn:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True ' 恢复屏幕更新
    
    MsgBox "异常值已标记完成！lowerBound：" & lowerBound & "upperBound：" & upperBound
End Sub

Sub 标记异常_3西格玛()
    Dim ws As Worksheet
    Dim cell As Range
    Dim mean As Double, stdDev As Double ' 均值、标准差（替换原四分位数变量）
    Dim lowerBound As Double, upperBound As Double ' 3σ上下限
    Dim colIndex As Integer
    Dim lastRow As Long
    Dim currentColumn As Range
    Dim visibleCells As Range
    Dim selectedColumns As Range
    Dim area As Range
    Dim visibleColumns As Range
    Dim col As Range
    
    ' 设置工作表
    Set ws = ActiveSheet
    
    ' 初始化可见列范围
    Set visibleColumns = Nothing
    
    ' 遍历选中列的每个列，只保留可见的列
    For Each col In Selection.Columns
        If Not col.Hidden Then ' 判断列是否可见
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    ' 如果没有可见列，提示用户并退出
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    ' 获取可见列与数据区域的交集
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False ' 暂时关闭屏幕更新，提高效率
    
    ' 遍历每一列
    For Each area In selectedColumns.Areas
        For colIndex = 1 To area.Columns.Count
            ' 获取当前列有数据的最后一行行号
            lastRow = ws.Cells(ws.Rows.Count, area.Columns(colIndex).Column).End(xlUp).Row
            
            ' 获取当前列从第一行到最后一行有数据的单元格范围
            Set currentColumn = ws.Range(ws.Cells(1, area.Columns(colIndex).Column), ws.Cells(lastRow, area.Columns(colIndex).Column))
            
            ' 获取当前列的可见单元格（适配筛选状态）
            On Error Resume Next
            Set visibleCells = currentColumn.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If visibleCells Is Nothing Then
                GoTo ContinueNextColumn
            End If
            
            ' ===================== 核心修改：3西格玛计算 =====================
            On Error Resume Next
            ' 1. 计算数据均值
            mean = Application.WorksheetFunction.Average(visibleCells)
            ' 2. 计算样本标准差（工业生产数据推荐用 StDev_S）
            stdDev = Application.WorksheetFunction.StDev_S(visibleCells)
            On Error GoTo 0
            
            ' 防错：如果计算失败（无有效数值），跳过该列
            If stdDev = 0 Or IsEmpty(mean) Then
                GoTo ContinueNextColumn
            End If
            
            ' 3. 计算3σ上下报警限
            lowerBound = mean - 3 * stdDev
            upperBound = mean + 3 * stdDev
            ' =================================================================
            
            ' 遍历可见单元格，标记3σ异常值
            For Each cell In visibleCells
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    ' 判断：数值超出 均值±3倍标准差 → 标记为异常
                    If cell.value < lowerBound Or cell.value > upperBound Then
                        cell.Interior.Color = RGB(255, 255, 0) ' 黄色标记异常
                    Else
                        cell.Interior.ColorIndex = xlNone ' 清除原有颜色
                    End If
                End If
            Next cell
            
ContinueNextColumn:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True ' 恢复屏幕更新
    
    MsgBox "3西格玛异常值标记完成！"
End Sub

Sub 标记异常_Grubbs_单侧最大最小()
    Dim ws As Worksheet
    Dim cell As Range
    Dim xBar As Double, S As Double
    Dim G_max As Double, G_min As Double
    Dim Gcrit_one As Double, alpha As Double, tVal_one As Double
    Dim colIndex As Integer
    Dim lastRow As Long
    Dim currentColumn As Range
    Dim visibleCells As Range
    Dim selectedColumns As Range
    Dim area As Range
    Dim visibleColumns As Range
    Dim col As Range
    Dim arr() As Double
    Dim n As Long, i As Long
    Dim valMax As Double, valMin As Double
    Dim infoStr As String   '保存各列临界值汇总
    
    alpha = 0.05
    infoStr = ""
    Set ws = ActiveSheet
    
    Set visibleColumns = Nothing
    For Each col In Selection.Columns
        If Not col.Hidden Then
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each area In selectedColumns.Areas
        For colIndex = 1 To area.Columns.Count
            Dim thisColNum As Long
            thisColNum = area.Columns(colIndex).Column
            
            lastRow = ws.Cells(ws.Rows.Count, thisColNum).End(xlUp).Row
            Set currentColumn = ws.Range(ws.Cells(1, thisColNum), ws.Cells(lastRow, thisColNum))
            
            On Error Resume Next
            Set visibleCells = currentColumn.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If visibleCells Is Nothing Then
                infoStr = infoStr & "第" & thisColNum & "列：无可见数据，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            n = 0
            ReDim arr(1 To visibleCells.Cells.Count)
            For Each cell In visibleCells
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    n = n + 1
                    arr(n) = cell.value
                End If
            Next cell
            If n < 3 Then
                infoStr = infoStr & "第" & thisColNum & "列：样本n<3，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            ReDim Preserve arr(1 To n)
            
            xBar = WorksheetFunction.Average(arr)
            S = WorksheetFunction.StDev_S(arr)
            If S < 0.000000000001 Then
                infoStr = infoStr & "第" & thisColNum & "列：全部数值相同，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            valMax = WorksheetFunction.Max(arr)
            valMin = WorksheetFunction.Min(arr)
            
            tVal_one = WorksheetFunction.T_Inv_2T(alpha, n - 2)
            Gcrit_one = (n - 1) / Sqr(n) * Sqr(tVal_one ^ 2 / ((n - 2) + tVal_one ^ 2))
            
            infoStr = infoStr & "第" & thisColNum & "列 n=" & n & vbCrLf _
            & "  Max=" & Round(valMax, 6) & "  Min=" & Round(valMin, 6) & vbCrLf _
            & "  单侧Gcrit=" & Round(Gcrit_one, 4) & vbCrLf & vbCrLf
            
            '遍历单元格标记
            For Each cell In visibleCells
                cell.Interior.ColorIndex = xlNone
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    G_max = (cell.value - xBar) / S
                    G_min = (xBar - cell.value) / S
                    If G_max > Gcrit_one Or G_min > Gcrit_one Then
                        cell.Interior.Color = RGB(255, 255, 0)
                    End If
                End If
            Next cell
            
ContinueNextColumn:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True
    infoStr = "【提示】格拉布斯检验前提：数据近似正态；非正态建议改用IQR箱线法！" & vbCrLf & vbCrLf & infoStr
    MsgBox "Grubbs单侧(最大/最小)标记完成！" & vbCrLf & vbCrLf & infoStr
End Sub

Sub 标记异常_CUSUM_Vmask()
    Dim ws As Worksheet
    Dim cell As Range
    Dim colIndex As Integer
    Dim lastRow As Long
    Dim currentColumn As Range
    Dim visibleCells As Range
    Dim selectedColumns As Range
    Dim area As Range
    Dim visibleColumns As Range
    Dim col As Range
    Dim arr() As Double
    Dim n As Long, i As Long
    Dim xBar As Double, S As Double
    Dim k As Double, h As Double
    Dim Cplus As Double, Cminus As Double
    Dim alarmCnt As Long
    Dim infoStr As String
    
    'CUSUM V?mask参数 k=0.5 , h=5；可修改
    k = 0.5
    h = 5
    
    infoStr = ""
    Set ws = ActiveSheet
    
    Set visibleColumns = Nothing
    For Each col In Selection.Columns
        If Not col.Hidden Then
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each area In selectedColumns.Areas
        For colIndex = 1 To area.Columns.Count
            Dim thisColNum As Long
            thisColNum = area.Columns(colIndex).Column
            
            lastRow = ws.Cells(ws.Rows.Count, thisColNum).End(xlUp).Row
            Set currentColumn = ws.Range(ws.Cells(1, thisColNum), ws.Cells(lastRow, thisColNum))
            
            On Error Resume Next
            Set visibleCells = currentColumn.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If visibleCells Is Nothing Then
                infoStr = infoStr & "第" & thisColNum & "列：无可见数据，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            n = 0
            ReDim arr(1 To visibleCells.Cells.Count)
            For Each cell In visibleCells
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    n = n + 1
                    arr(n) = cell.value
                End If
            Next cell
            If n < 5 Then
                infoStr = infoStr & "第" & thisColNum & "列：样本n<5，跳过(CUSUM至少5点)" & vbCrLf
                GoTo ContinueNextColumn
            End If
            ReDim Preserve arr(1 To n)
            
            xBar = WorksheetFunction.Average(arr)
            S = WorksheetFunction.StDev_S(arr)
            If S < 0.000000000001 Then
                infoStr = infoStr & "第" & thisColNum & "列：全部数值相同，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            Cplus = 0
            Cminus = 0
            alarmCnt = 0
            i = 0
            '遍历可见单元格循环CUSUM迭代
            For Each cell In visibleCells
                cell.Interior.ColorIndex = xlNone
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    i = i + 1
                    Dim xi As Double
                    xi = cell.value
                    Cplus = WorksheetFunction.Max(0, xi - xBar - k * S + Cplus)
                    Cminus = WorksheetFunction.Max(0, xBar - xi - k * S + Cminus)
                    
                    If Cplus > h * S Or Cminus > h * S Then
                        cell.Interior.Color = RGB(255, 255, 0)
                        alarmCnt = alarmCnt + 1
                    End If
                End If
            Next cell
            
            infoStr = infoStr & "第" & thisColNum & "列 n=" & n & vbCrLf _
            & "  xBar=" & Round(xBar, 6) & " σ=" & Round(S, 6) & vbCrLf _
            & "  k=" & k & " h=" & h & " 报警点数=" & alarmCnt & vbCrLf & vbCrLf
            
ContinueNextColumn:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True
    infoStr = "【提示】CUSUM累积和：本程序使用样本均值作为目标μ0；在线监控建议输入历史基准μ0、σ！" & vbCrLf & vbCrLf & infoStr
    MsgBox "CUSUM V?mask标记完成！" & vbCrLf & vbCrLf & infoStr
End Sub

Sub 标记异常_EWMA()
    Dim ws As Worksheet
    Dim cell As Range
    Dim colIndex As Integer
    Dim lastRow As Long
    Dim currentColumn As Range
    Dim visibleCells As Range
    Dim selectedColumns As Range
    Dim area As Range
    Dim visibleColumns As Range
    Dim col As Range
    Dim arr() As Double
    Dim n As Long, i As Long
    Dim xBar As Double, S As Double
    Dim lambda As Double, L As Double
    Dim Zi As Double
    Dim UCL As Double, LCL As Double
    Dim alarmCnt As Long
    Dim infoStr As String
    
    'EWMA 参数 λ 平滑系数；L控制限倍数；常用 λ=0.2 L=3
    lambda = 0.2
    L = 3
    
    infoStr = ""
    Set ws = ActiveSheet
    
    Set visibleColumns = Nothing
    For Each col In Selection.Columns
        If Not col.Hidden Then
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each area In selectedColumns.Areas
        For colIndex = 1 To area.Columns.Count
            Dim thisColNum As Long
            thisColNum = area.Columns(colIndex).Column
            
            lastRow = ws.Cells(ws.Rows.Count, thisColNum).End(xlUp).Row
            Set currentColumn = ws.Range(ws.Cells(1, thisColNum), ws.Cells(lastRow, thisColNum))
            
            On Error Resume Next
            Set visibleCells = currentColumn.SpecialCells(xlCellTypeVisible)
            On Error GoTo 0
            
            If visibleCells Is Nothing Then
                infoStr = infoStr & "第" & thisColNum & "列：无可见数据，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            n = 0
            ReDim arr(1 To visibleCells.Cells.Count)
            For Each cell In visibleCells
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    n = n + 1
                    arr(n) = cell.value
                End If
            Next cell
            If n < 5 Then
                infoStr = infoStr & "第" & thisColNum & "列：样本n<5，跳过(EWMA至少5点)" & vbCrLf
                GoTo ContinueNextColumn
            End If
            ReDim Preserve arr(1 To n)
            
            xBar = WorksheetFunction.Average(arr)
            S = WorksheetFunction.StDev_S(arr)
            If S < 0.000000000001 Then
                infoStr = infoStr & "第" & thisColNum & "列：全部数值相同，跳过" & vbCrLf
                GoTo ContinueNextColumn
            End If
            
            Zi = xBar       'Z0初始值等于目标μ0
            alarmCnt = 0
            i = 0
            For Each cell In visibleCells
                cell.Interior.ColorIndex = xlNone
                If IsNumeric(cell.value) And Not IsEmpty(cell.value) Then
                    i = i + 1
                    Dim xi As Double
                    xi = cell.value
                    'EWMA迭代
                    Zi = lambda * xi + (1 - lambda) * Zi
                    '动态随i变化的上下控制限
                    Dim factor As Double
                    factor = Sqr(lambda / (2 - lambda) * (1 - (1 - lambda) ^ (2 * i)))
                    UCL = xBar + L * S * factor
                    LCL = xBar - L * S * factor
                    
                    If Zi > UCL Or Zi < LCL Then
                        cell.Interior.Color = RGB(255, 255, 0)
                        alarmCnt = alarmCnt + 1
                    End If
                End If
            Next cell
            
            infoStr = infoStr & "第" & thisColNum & "列 n=" & n & vbCrLf _
            & "  xBar=" & Round(xBar, 6) & " σ=" & Round(S, 6) & vbCrLf _
            & "  λ=" & lambda & " L=" & L & " 报警点数=" & alarmCnt & vbCrLf & vbCrLf
            
ContinueNextColumn:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True
    infoStr = "【提示】EWMA指数加权移动平均；本程序使用样本均值作为μ0；正式SPC请使用历史稳态μ0、σ！" & vbCrLf & vbCrLf & infoStr
    MsgBox "EWMA?SPC标记完成！" & vbCrLf & vbCrLf & infoStr
End Sub

Sub 一键异常检测_五法对比()
    Dim wsSrc As Worksheet, wsNew As Worksheet
    Dim col As Range, area As Range
    Dim visibleColumns As Range
    Dim selectedColumns As Range
    Dim srcRng As Range
    Dim srcCols As Long, lastRowNew As Long
    Dim offsetCol As Long
    Dim methodNames As Variant
    Dim i As Long
    Dim destRng As Range
    
    ' ===== 1. 检查选中 =====
    If Selection Is Nothing Then
        MsgBox "请先在源表选中数据区域（可多选列）！", vbExclamation
        Exit Sub
    End If
    Set wsSrc = ActiveSheet
    
    ' ===== 2. 遍历选中列，跳过隐藏列（参考之前Boxplot/Grubbs写法）=====
    Set visibleColumns = Nothing
    srcCols = 0
    For Each col In Selection.Columns
        If Not col.Hidden Then
            srcCols = srcCols + 1
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    ' ===== 3. 与UsedRange取交集 =====
    On Error Resume Next
    Set selectedColumns = Intersect(wsSrc.usedRange, visibleColumns.EntireColumn)
    On Error GoTo 0
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    
    ' ===== 4. 只取筛选后的可见单元格 =====
    On Error Resume Next
    Set srcRng = selectedColumns.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If srcRng Is Nothing Then
        MsgBox "没有可见数据！", vbExclamation
        Exit Sub
    End If
    
    ' ===== 5. 新建结果sheet =====
    Set wsNew = Sheets.Add(After:=wsSrc)
    wsNew.Name = "异常检测_" & Format(Now(), "hhmmss")
    
    methodNames = Array("3西格玛", "Grubbs单侧", "CUSUM_Vmask", "EWMA", "Boxplot_IQR")
    offsetCol = 1
    
    Application.ScreenUpdating = False
    
    ' ===== 6. 逐方法复制数据块并调用对应Sub =====
    For i = 0 To UBound(methodNames)
        ' 写方法名标题
        wsNew.Cells(1, offsetCol).value = "【" & methodNames(i) & "】"
        wsNew.Cells(1, offsetCol).Font.Bold = True
        
        ' 复制可见数据（自动跳过隐藏行）
        srcRng.Copy Destination:=wsNew.Cells(2, offsetCol)
        
        ' 在新sheet上定位刚复制出的数据块
        lastRowNew = wsNew.Cells(wsNew.Rows.Count, offsetCol).End(xlUp).Row
        Set destRng = wsNew.Range(wsNew.Cells(2, offsetCol), _
                                   wsNew.Cells(lastRowNew, offsetCol + srcCols - 1))
        destRng.Select
        
        ' 调用对应检测Sub
        On Error Resume Next
        Select Case methodNames(i)
            Case "3西格玛":      Call 标记异常_3西格玛
            Case "Grubbs单侧":   Call 标记异常_Grubbs_单侧最大最小
            Case "CUSUM_Vmask":  Call 标记异常_CUSUM_Vmask
            Case "EWMA":         Call 标记异常_EWMA
            Case "Boxplot_IQR":  Call 标记异常_Boxplot
        End Select
        On Error GoTo 0
        
        ' 下一块往右挪：原列数 + 2列间隔
        offsetCol = offsetCol + srcCols + 2
    Next i
    
    wsNew.Cells(1, 1).Select
    Application.ScreenUpdating = True
    
    MsgBox "五法对比完成！新sheet【" & wsNew.Name & "】" & vbCrLf & vbCrLf & _
           "已自动：跳过隐藏列 / 只复制筛选后的可见行" & vbCrLf & _
           "从左到右：3σ → Grubbs → CUSUM → EWMA → Boxplot(IQR)", vbInformation
End Sub


Sub MK趋势检验_Sen斜率()
    ' Mann-Kendall趋势检验 + Sen斜率估计
    ' 选中多列 → 逐列独立计算 → 弹窗汇总Z/p值/Sen斜率/趋势判定
    Dim ws As Worksheet
    Dim col As Range, area As Range
    Dim visibleColumns As Range, selectedColumns As Range
    Dim thisColNum As Long, lastRow As Long
    Dim arr() As Double, sortedArr() As Double
    Dim n As Long, i As Long, j As Long, k As Long
    Dim S As Double, varS As Double, Z As Double, pValue As Double
    Dim senSlope As Double
    Dim slopes() As Double, nSlopes As Long, idx As Long
    Dim alpha As Double, tieCount As Long, tieTerm As Double
    Dim trend As String, infoStr As String
    
    alpha = 0.05
    infoStr = ""
    Set ws = ActiveSheet
    
    ' 收集可见列（跳过隐藏列）
    Set visibleColumns = Nothing
    For Each col In Selection.Columns
        If Not col.Hidden Then
            If visibleColumns Is Nothing Then
                Set visibleColumns = col
            Else
                Set visibleColumns = Union(visibleColumns, col)
            End If
        End If
    Next col
    If visibleColumns Is Nothing Then
        MsgBox "没有找到可见的选中列！", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set selectedColumns = Intersect(ws.usedRange, visibleColumns.EntireColumn)
    On Error GoTo 0
    If selectedColumns Is Nothing Then
        MsgBox "未找到有效的数据列！", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    For Each area In selectedColumns.Areas
        Dim colIndex As Long
        For colIndex = 1 To area.Columns.Count
            thisColNum = area.Columns(colIndex).Column
            lastRow = ws.Cells(ws.Rows.Count, thisColNum).End(xlUp).Row
            
            ' 逐行读取可见数据
            n = 0
            ReDim arr(1 To lastRow)
            For i = 1 To lastRow
                If Not ws.Rows(i).Hidden Then
                    If IsNumeric(ws.Cells(i, thisColNum).value) And Not IsEmpty(ws.Cells(i, thisColNum).value) Then
                        n = n + 1
                        arr(n) = ws.Cells(i, thisColNum).value
                    End If
                End If
            Next i
            If n < 4 Then
                infoStr = infoStr & "第" & thisColNum & "列：n<4，跳过" & vbCrLf
                GoTo ContinueNextCol
            End If
            ReDim Preserve arr(1 To n)
            
            '==== MK 统计量 S ====
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
            
            '==== 结修正：复制排序后数组识别结 ====
            ReDim sortedArr(1 To n)
            For i = 1 To n
                sortedArr(i) = arr(i)
            Next i
            For i = 1 To n - 1
                For j = i + 1 To n
                    If sortedArr(i) > sortedArr(j) Then
                        Dim tmp As Double
                        tmp = sortedArr(i)
                        sortedArr(i) = sortedArr(j)
                        sortedArr(j) = tmp
                    End If
                Next j
            Next i
            
            tieTerm = 0
            k = 1
            Do While k < n
                If sortedArr(k) = sortedArr(k + 1) Then
                    tieCount = 1
                    Do While k + tieCount <= n And sortedArr(k + tieCount) = sortedArr(k)
                        tieCount = tieCount + 1
                    Loop
                    tieTerm = tieTerm + tieCount * (tieCount - 1) * (2 * tieCount + 5)
                    k = k + tieCount
                Else
                    k = k + 1
                End If
            Loop
            
            varS = (n * (n - 1) * (2 * n + 5) - tieTerm) / 18
            
            '==== Z 统计量 ====
            If S > 0 Then
                Z = (S - 1) / Sqr(varS)
            ElseIf S < 0 Then
                Z = (S + 1) / Sqr(varS)
            Else
                Z = 0
            End If
            
            '==== 双侧 p 值 ====
            pValue = 2 * (1 - WorksheetFunction.Norm_S_Dist(Abs(Z), True))
            
            '==== Sen 斜率：所有点对斜率的中位数 ====
            nSlopes = n * (n - 1) / 2
            ReDim slopes(1 To nSlopes)
            idx = 0
            For i = 1 To n - 1
                For j = i + 1 To n
                    idx = idx + 1
                    slopes(idx) = (arr(j) - arr(i)) / (j - i)
                Next j
            Next i
            senSlope = WorksheetFunction.Median(slopes)
            
            '==== 趋势判定 ====
            If pValue < alpha Then
                If Z > 0 Then
                    trend = "▲ 显著上升"
                Else
                    trend = "▼ 显著下降"
                End If
            Else
                trend = "— 无显著趋势"
            End If
            
            infoStr = infoStr & "第" & thisColNum & "列 n=" & n & vbCrLf _
                & "  S=" & Round(S, 1) & "  Z=" & Round(Z, 4) & "  p=" & Round(pValue, 4) & vbCrLf _
                & "  Sen斜率=" & Round(senSlope, 6) & "  " & trend & vbCrLf & vbCrLf
            
ContinueNextCol:
        Next colIndex
    Next area
    
    Application.ScreenUpdating = True
    infoStr = "【Mann-Kendall趋势检验 + Sen斜率】α=" & alpha & vbCrLf & vbCrLf & infoStr
    MsgBox infoStr, vbInformation
End Sub


