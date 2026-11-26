Attribute VB_Name = "模块6"
Option Explicit

' 安全检测 Chart 是否有指定轴（防止直接调用 .Axes 报错）
Private Function HasChartAxis(ByVal ch As Chart, ByVal axisType As XlAxisType, Optional ByVal axisGroup As XlAxisGroup = xlPrimary) As Boolean
On Error Resume Next
Dim ax As Axis
Set ax = ch.Axes(axisType, axisGroup)
If Err.Number <> 0 Then
Err.Clear
HasChartAxis = False
Else
HasChartAxis = Not ax Is Nothing
End If
On Error GoTo 0
End Function

Public Sub 添加辅助线()
    辅助线窗口.Show
End Sub

Public Function 添加辅助线函数_final(ByVal seriesName As String, ByVal yValue As Double, Optional ByVal showSuccessMessage As Boolean = True) As Boolean
    添加辅助线函数_final = AddHelperSeriesToSelectedChart(seriesName, yValue, showSuccessMessage)
End Function

Public Function AddHelperLineFinal(ByVal seriesName As String, ByVal yValue As Double, Optional ByVal showSuccessMessage As Boolean = True) As Boolean
    AddHelperLineFinal = 添加辅助线函数_final(seriesName, yValue, showSuccessMessage)
End Function

Private Function AddHelperSeriesToSelectedChart(ByVal seriesName As String, ByVal yValue As Double, ByVal showSuccessMessage As Boolean) As Boolean
    Dim cht As Chart
    Dim xValues As Variant
    Dim helperXValues As Variant
    Dim helperYValues As Variant
    Dim pointCount As Long
    Dim i As Long
    Dim minX As Double
    Dim maxX As Double

    If TypeName(Selection) <> "ChartArea" Then
        MsgBox "请先选中一个图表区域！", vbExclamation
        Exit Function
    End If

    Set cht = Selection.Parent

    If seriesName = "" Then seriesName = "标准"

    If cht.SeriesCollection.Count = 0 Then
        MsgBox "图表中没有数据系列，无法添加辅助线！", vbExclamation
        Exit Function
    End If

    xValues = cht.SeriesCollection(1).XValues
    minX = WorksheetFunction.Min(xValues)
    maxX = WorksheetFunction.Max(xValues)
    pointCount = UBound(xValues) - LBound(xValues) + 1

    ReDim helperXValues(1 To pointCount)
    ReDim helperYValues(1 To pointCount)

    If pointCount = 1 Then
        helperXValues(1) = minX
        helperYValues(1) = yValue
    Else
        For i = 1 To pointCount
            helperXValues(i) = minX + (maxX - minX) * (i - 1) / (pointCount - 1)
            helperYValues(i) = yValue
        Next i
    End If

    With cht.SeriesCollection.NewSeries()
        .Name = seriesName
        .XValues = helperXValues
        .Values = helperYValues
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
    End With

    AddHelperSeriesToSelectedChart = True

    If showSuccessMessage Then
        MsgBox "辅助线已添加到图表！", vbInformation
    End If
End Function

Sub 画标记折线图_final()
    '===================== 错误处理：程序报错时弹出提示 =====================
    On Error GoTo ErrHandler

    '===================== 变量定义（详细注释每个变量用途） =====================
    Dim wb As Workbook       ' 当前工作簿对象
    Dim ws As Worksheet      ' 当前工作表对象
    Dim sel As Range         ' 用户选中的单元格区域（数据列）
    Dim sampleCategoryColumn As Integer  ' 样品类别列号
    Dim sampleNameColumn As Integer      ' 样品名称列号
    Dim titleRow As Long                 ' 表头（标题）所在行号
    Dim dateCol As Long                  ' 日期列号（X轴数据）

    ' 绑定当前激活的工作簿和工作表
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    '===================== 基础校验：检查用户是否选中了单元格区域 =====================
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中一个或多个列（可不连续）。", vbExclamation
        Exit Sub  ' 未选中区域，直接退出程序
    End If
    Set sel = Selection  ' 将用户选中区域赋值给变量

    '===================== 自动识别表头行（兼容表头不在第1行） =====================
    ' 有自动筛选时用筛选范围首行；否则用选中区域第一行，而非固定第1行
    If ws.AutoFilterMode Then
        titleRow = ws.AutoFilter.Range.Row
    Else
        titleRow = sel.Rows(1).Row
    End If

    '===================== 查找日期列（X轴数据源） =====================
    dateCol = 0  ' 初始化日期列号为0（未找到）
    Dim iCol As Long
    ' 遍历所有已使用的列，查找表头包含「日期」的列
    For iCol = 1 To ws.usedRange.Columns.Count
        If ws.Cells(titleRow, iCol) Like "*日期*" Then
            dateCol = iCol  ' 找到日期列，记录列号
            Exit For        ' 找到后退出循环
        End If
    Next iCol
    ' 未找到日期列，弹出错误提示并退出
    If dateCol = 0 Then
        MsgBox "未找到[取样日期]列，请确认表头。", vbCritical
        Exit Sub
    End If

    '===================== 构建有效数据列（排除日期列，去重存储选中列） =====================
    Dim dictCols As Object
    Set dictCols = CreateObject("Scripting.Dictionary")  ' 创建字典：用于去重存储有效数据列号
    Dim area As Range, colRng As Range

    ' 遍历用户选中的所有区域（支持不连续选区）
    For Each area In sel.Areas
        ' 遍历区域中的每一列
        For Each colRng In area.Columns
            ' 排除日期列 + 只保留【可见列】（隐藏列跳过，保留 final 原有行为）
            If colRng.Column <> dateCol And colRng.EntireColumn.Hidden = False Then
                ' 字典去重：列号不存在则添加
                If Not dictCols.Exists(colRng.Column) Then dictCols.Add colRng.Column, True
            End If
        Next colRng
    Next area
    ' 无有效数据列，弹出提示并退出
    If dictCols.Count = 0 Then
        MsgBox "没有有效的数据列！", vbExclamation
        Exit Sub
    End If

    '===================== 查找样品信息列（用于图表标题） =====================
    sampleCategoryColumn = 0  ' 初始化样品类别列号
    sampleNameColumn = 0      ' 初始化样品名称列号
    Dim col As Integer

    ' 第一步：查找表头为「样品类别」的列
    For col = 1 To ws.usedRange.Columns.Count
        If StrComp(ws.Cells(titleRow, col).value, "样品类别", vbTextCompare) = 0 Then
            sampleCategoryColumn = col
            Exit For
        End If
    Next col

    ' 数据起始行：表头行的下一行
    Dim startDataRow As Long: startDataRow = titleRow + 1

    '===================== 新建工作表（存放生成的所有图表） =====================
    Dim newWs As Worksheet
    Set newWs = wb.Worksheets.Add(Before:=ws)  ' 在当前工作表前新建表
    ' 工作表名称：折线图+时间戳（防止重名）
    Dim baseName As String: baseName = "折线图_" & Format(Now, "yyyymmdd_hhnnss")
    On Error Resume Next  ' 忽略重命名报错
    newWs.Name = Left(baseName, 31)  ' Excel表名最大31字符，截取
    On Error GoTo 0  ' 恢复默认错误处理

    '===================== 图表布局参数（一行2个排列） =====================
    Dim chartsPerRow As Integer: chartsPerRow = 2   ' 每行显示2个图表
    Dim chartW As Double: chartW = 16.33 * 28.35    ' 图表宽度（转换为磅）
    Dim chartH As Double: chartH = 6.66 * 28.35 * 0.9 ' 图表高度
    Dim marginLeft As Double: marginLeft = 20        ' 左边距
    Dim marginTop As Double: marginTop = 20          ' 上边距
    Dim hGap As Double: hGap = 15                    ' 图表水平间距
    Dim vGap As Double: vGap = 15 * 0.8              ' 图表垂直间距

    '===================== 图表配色（固定样式） =====================
    Dim lineColor As Long, markerColor As Long, fillColor As Long
    lineColor = RGB(91, 155, 213)       ' 折线颜色
    markerColor = lineColor             ' 标记边框颜色
    fillColor = RGB(237, 125, 49)       ' 标记填充颜色

    '===================== 循环生成图表（遍历所有有效数据列） =====================
    Dim idx As Long: idx = 0  ' 图表计数
    Dim yCol As Variant       ' 遍历字典中的数据列号

    For Each yCol In dictCols.Keys
        ' 1. 获取数据最后一行（取日期列和数据列的最大行号）
        Dim lastRowX As Long, lastRowY As Long, lastRow As Long
        lastRowX = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row  ' 日期列最后一行
        lastRowY = ws.Cells(ws.Rows.Count, CLng(yCol)).End(xlUp).Row ' 数据列最后一行
        lastRow = IIf(lastRowX > lastRowY, lastRowX, lastRowY)     ' 有效数据最后一行
        If lastRow < startDataRow Then GoTo NextYCol  ' 无数据，跳过当前列

        ' 2. 逐行收集可见且为数值的X/Y点（兼容Y列含文本、空值、隐藏行）
        '    原 SpecialCells 方式在Y列出现文本（如“未测”“N/A”）时 Min/Max 报错
        '    导致整段宏中断；逐行过滤后非数值点自动跳过，其余点正常绘制
        Dim xList As Collection: Set xList = New Collection
        Dim yList As Collection: Set yList = New Collection
        Dim R As Long
        For R = startDataRow To lastRow
            If Not ws.Rows(R).Hidden Then   ' 只取非隐藏行（含自动筛选隐藏的行）
                Dim xv As Variant, yv As Variant
                xv = ws.Cells(R, dateCol).value
                yv = ws.Cells(R, yCol).value
                ' 仅收集：Y值非空且为数值（文本自动跳过，避免图表报错）
                If Len(Trim(CStr(yv))) > 0 Then
                    If IsNumeric(yv) Then
                        xList.Add xv
                        yList.Add CDbl(yv)
                    End If
                End If
            End If
        Next R
        If xList.Count = 0 Then GoTo NextYCol  ' 无可绘制数据点，跳过

        ' 3. 集合转数组（供图表数据源与坐标范围计算）
        Dim nPts As Long: nPts = xList.Count
        Dim xArr() As Variant, yArr() As Variant
        ReDim xArr(1 To nPts)
        ReDim yArr(1 To nPts)
        Dim k As Long
        For k = 1 To nPts
            xArr(k) = xList(k)
            yArr(k) = yList(k)
        Next k

        ' 4. 计算数据极值（基于实际绘制的数值点，不再对文本报错）
        Dim dataMin As Double, dataMax As Double, dataSpan As Double
        dataMin = Application.Min(yArr)  ' 数据最小值（容错调用）
        dataMax = Application.Max(yArr)  ' 数据最大值
        dataSpan = dataMax - dataMin                       ' 数据跨度

        ' 防止数据全相同导致跨度为0
        If dataSpan = 0 Then
            dataSpan = IIf(dataMin = 0, 1, dataMin * 0.1)
        End If

        ' 5. 自动计算图表位置（两列排列，自动换行）
        idx = idx + 1
        Dim chartLeft As Double, chartTop As Double
        chartLeft = marginLeft + ((idx - 1) Mod chartsPerRow) * (chartW + hGap)  ' 左边距
        chartTop = marginTop + Int((idx - 1) / chartsPerRow) * (chartH + vGap)   ' 上边距

        ' 6. 创建图表对象
        Dim chtObj As ChartObject
        Set chtObj = newWs.ChartObjects.Add(chartLeft, chartTop, chartW, chartH)
        Dim ch As Chart: Set ch = chtObj.Chart
        ch.ChartType = xlLineMarkers  ' 设置图表类型：带标记的折线图
        ch.PlotVisibleOnly = True     ' 仅绘制可见数据

        ' 7. 清空图表默认生成的系列
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop

        ' 8. 添加数据系列（设置X/Y轴、样式、颜色）
        Dim seriesName As String
        seriesName = CStr(ws.Cells(titleRow, yCol).value)  ' 系列名称=列标题
        If Len(seriesName) = 0 Then seriesName = "Col" & yCol  ' 表头为空时回退为列号
        With ch.SeriesCollection.NewSeries
            .Name = seriesName           ' 系列名称
            ' 数据源用工作表区域而非数组（规避运行时错误5）：
            ' 数组赋值在【X格空白】或【单点序列】时，WPS/Excel 会报“无效的过程调用或参数”；
            ' 区域赋值由 PlotVisibleOnly 自动跳过隐藏行、文本单元格自动留空，不会触发该错误
            .xValues = ws.Range(ws.Cells(startDataRow, dateCol), ws.Cells(lastRow, dateCol))  ' X轴数据（日期）
            .Values = ws.Range(ws.Cells(startDataRow, yCol), ws.Cells(lastRow, yCol))  ' Y轴数据（数值）
            .Format.Line.ForeColor.RGB = lineColor  ' 折线颜色
            .MarkerStyle = xlMarkerStyleCircle      ' 标记样式：圆形
            .MarkerForegroundColor = markerColor    ' 标记边框色
            .MarkerBackgroundColor = fillColor      ' 标记填充色
        End With

        ' 9. 智能设置Y轴范围（自适应数据，不修改刻度单位）
        Dim yAxis As Axis
        Dim userMin As Double, userMax As Double, autoUnit As Double
        Set yAxis = ch.Axes(xlValue)  ' 绑定Y轴

        ' 初步设置轴范围（预留5%边距）
        userMin = dataMin - dataSpan * 0.05
        userMax = dataMax + dataSpan * 0.05

        With yAxis
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = userMin
            .MaximumScale = userMax
        End With

        ch.Refresh
        DoEvents
        autoUnit = yAxis.MajorUnit  ' 获取自动生成的刻度单位

        ' 规整轴范围（对齐刻度，防止除0错误）
        If autoUnit <> 0 Then
            userMin = Application.WorksheetFunction.RoundDown(userMin / autoUnit, 0) * autoUnit
            userMax = Application.WorksheetFunction.RoundUp(userMax / autoUnit, 0) * autoUnit
        End If

        ' 应用最终轴范围（不修改刻度单位）
        With yAxis
            .MinimumScale = userMin
            .MaximumScale = userMax
        End With

        '===================== 图表标题、样式设置（原版样式） =====================
        ch.HasTitle = True  ' 显示图表标题
        Dim chartTitle As String
        Dim catVal As Variant, nameVal As Variant
        chartTitle = ""

        ' 修复运行时错误5：VBA 的 And 不短路，列号=0 时若直接求值 Cells(row, 0)
        ' 会报"无效的过程调用或参数"，故单元格读取必须放进"列存在"的嵌套判断。
        ' 值取该类列最后一条数据的值（End(xlUp)）；
        ' 该格为表头字/空/错误值/未知样品时自动回退，不会中断
        If sampleCategoryColumn <> 0 Then
            catVal = ws.Cells(ws.Rows.Count, sampleCategoryColumn).End(xlUp).value
        End If
        If sampleCategoryColumn <> 0 And Not IsError(catVal) And Len(Trim(CStr(catVal))) > 0 _
           And CStr(catVal) <> "未知样品" And CStr(catVal) <> "样品类别" Then
            chartTitle = "[" & ws.Name & " - " & CStr(catVal) & " - " & seriesName & "]趋势图"
        End If

        ' 回退1：用"样品名称"列值拼标题
        If Len(chartTitle) = 0 Then
            For col = 1 To ws.usedRange.Columns.Count
                If StrComp(ws.Cells(titleRow, col).value, "样品名称", vbTextCompare) = 0 Then
                    sampleNameColumn = col
                    Exit For
                End If
            Next col
            If sampleNameColumn <> 0 Then
                nameVal = ws.Cells(ws.Rows.Count, sampleNameColumn).End(xlUp).value
                If Not IsError(nameVal) And Len(Trim(CStr(nameVal))) > 0 And CStr(nameVal) <> "样品名称" Then
                    chartTitle = "[" & ws.Name & " - " & CStr(nameVal) & " - " & seriesName & "]趋势图"
                End If
            End If
        End If

        ' 回退2：纯表名+系列名（兼容无样品类别/名称列的报表）
        If Len(chartTitle) = 0 Then
            chartTitle = "[" & ws.Name & " - " & seriesName & "]趋势图"
        End If

        ch.chartTitle.Text = chartTitle  ' 设置标题文字
        ' 标题字体样式
        With ch.chartTitle.Font
            .Name = "等线": .Size = 14
        End With

        ' 坐标轴标题、字体设置
        ch.Axes(xlCategory).HasTitle = False  ' 隐藏X轴标题
        ch.Axes(xlValue).HasTitle = False     ' 隐藏Y轴标题
        ch.Axes(xlCategory).TickLabels.Font.Name = "等线"  ' X轴刻度字体
        ch.Axes(xlValue).TickLabels.Font.Name = "等线"     ' Y轴刻度字体

        ' Y轴网格线样式（灰色虚线；加轴存在性判断）
        If HasChartAxis(ch, xlValue) Then
            With ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line
                .ForeColor.RGB = RGB(220, 220, 220)
                .DashStyle = msoLineDash
            End With
        End If

        '===================== 图例设置（核心修改） =====================
        ' 方案1：【默认】完全隐藏图例（推荐）
        ch.HasLegend = False

        ' 方案2：显示图例并放置在左上角（取消下面注释，注释方案1即可）
        'ch.HasLegend = True
        'ch.Legend.Position = xlLegendPositionTopLeft

        ' 图表背景、边框样式（纯白背景，蓝色边框）
        ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ch.ChartArea.Border.Color = 16711680
        ' 给 ChartArea 加蓝色边框（必须按这个顺序写）

        ' X轴日期格式（X非日期时不强制套日期格式）
        On Error Resume Next
        If IsDate(xArr(1)) Then
            If HasChartAxis(ch, xlCategory) Then
                With ch.Axes(xlCategory, xlPrimary)
                    .TickLabels.NumberFormat = "m月d日"
                End With
            End If
        End If
        On Error GoTo 0

NextYCol:                                          ' 跳过当前列，继续下一列
    Next yCol

    '===================== 程序完成提示 =====================
    MsgBox "完成！共生成 " & idx & " 张图表（两列排列）", vbInformation
    Exit Sub

    '===================== 报错处理 =====================
ErrHandler:
    MsgBox "执行出错：" & Err.Number & " - " & Err.Description, vbCritical
End Sub


Sub 画控制图_final()
    '===================== 错误处理：程序报错时弹出提示 =====================
    On Error GoTo ErrHandler

    '===================== 变量定义 =====================
    Dim wb As Workbook       ' 当前工作簿对象
    Dim ws As Worksheet      ' 当前工作表对象
    Dim sel As Range         ' 用户选中的单元格区域（数据列）
    Dim sampleCategoryColumn As Integer  ' 样品类别列号
    Dim sampleNameColumn As Integer      ' 样品名称列号
    Dim titleRow As Long                 ' 表头（标题）所在行号
    Dim dateCol As Long                  ' 日期列号（X轴数据）

    ' 绑定当前激活的工作簿和工作表
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    '===================== 基础校验 =====================
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中一个或多个列（可不连续）。", vbExclamation
        Exit Sub
    End If
    Set sel = Selection

    '===================== 自动识别表头行（兼容表头不在第1行） =====================
    If ws.AutoFilterMode Then
        titleRow = ws.AutoFilter.Range.Row
    Else
        titleRow = sel.Rows(1).Row
    End If

    '===================== 查找日期列（X轴数据源） =====================
    dateCol = 0
    Dim iCol As Long
    For iCol = 1 To ws.usedRange.Columns.Count
        If ws.Cells(titleRow, iCol) Like "*日期*" Then
            dateCol = iCol
            Exit For
        End If
    Next iCol
    If dateCol = 0 Then
        MsgBox "未找到[取样日期]列，请确认表头。", vbCritical
        Exit Sub
    End If

    '===================== 构建有效数据列（排除日期列 + 隐藏列） =====================
    Dim dictCols As Object
    Set dictCols = CreateObject("Scripting.Dictionary")
    Dim area As Range, colRng As Range
    For Each area In sel.Areas
        For Each colRng In area.Columns
            If colRng.Column <> dateCol And colRng.EntireColumn.Hidden = False Then
                If Not dictCols.Exists(colRng.Column) Then dictCols.Add colRng.Column, True
            End If
        Next colRng
    Next area
    If dictCols.Count = 0 Then
        MsgBox "没有有效的数据列！", vbExclamation
        Exit Sub
    End If

    '===================== 查找样品信息列（用于图表标题） =====================
    sampleCategoryColumn = 0
    sampleNameColumn = 0
    Dim col As Integer
    For col = 1 To ws.usedRange.Columns.Count
        If StrComp(ws.Cells(titleRow, col).value, "样品类别", vbTextCompare) = 0 Then
            sampleCategoryColumn = col
            Exit For
        End If
    Next col

    ' 数据起始行：表头行的下一行
    Dim startDataRow As Long: startDataRow = titleRow + 1

    '===================== 新建工作表（存放所有控制图） =====================
    Dim newWs As Worksheet
    Set newWs = wb.Worksheets.Add(Before:=ws)
    Dim baseName As String: baseName = "控制图_" & Format(Now, "yyyymmdd_hhnnss")
    On Error Resume Next
    newWs.Name = Left(baseName, 31)
    On Error GoTo 0

    '===================== 图表布局参数（一行2个排列，与 _final 一致） =====================
    Dim chartsPerRow As Integer: chartsPerRow = 2
    Dim chartW As Double: chartW = 16.33 * 28.35
    Dim chartH As Double: chartH = 6.66 * 28.35 * 0.9
    Dim marginLeft As Double: marginLeft = 20
    Dim marginTop As Double: marginTop = 20
    Dim hGap As Double: hGap = 15
    Dim vGap As Double: vGap = 15 * 0.8

    '===================== 图表配色 =====================
    Dim lineColor As Long, markerColor As Long, fillColor As Long
    lineColor = RGB(91, 155, 213)       ' 数据线：蓝
    markerColor = lineColor             ' 数据标记边框
    fillColor = RGB(237, 125, 49)       ' 数据标记填充
    Dim clColor As Long, uclColor As Long, oocColor As Long
    clColor = RGB(0, 176, 80)           ' CL 中心线：绿
    uclColor = RGB(192, 0, 0)           ' UCL/LCL 控制限：红
    oocColor = RGB(255, 0, 0)           ' 超限点：亮红

    '===================== 控制图参数（单值 X 图，移动极差法估 σ） =====================
    Const D2 As Double = 1.128          ' n=2 移动极差的 d2 常数（国标 SPC 用）
    Const SIGMA_MULTIPLIER As Double = 3 ' 3σ 控制限
    ' 辅助列区：新表第100行起，每列占用6列（CL/UCL/LCL/超限/日期/数据），
    ' 全部系列引用新表区域，避免跨表引用与数组赋值在 WPS 的兼容问题
    Dim helperRow As Long: helperRow = 100
    Dim helperCols As Long: helperCols = 6

    '===================== 循环生成控制图 =====================
    Dim idx As Long: idx = 0
    Dim yCol As Variant
    Dim colIdx As Long: colIdx = 0

    For Each yCol In dictCols.Keys
        colIdx = colIdx + 1

        ' 1. 数据最后一行
        Dim lastRowX As Long, lastRowY As Long, lastRow As Long
        lastRowX = ws.Cells(ws.Rows.Count, dateCol).End(xlUp).Row
        lastRowY = ws.Cells(ws.Rows.Count, CLng(yCol)).End(xlUp).Row
        lastRow = IIf(lastRowX > lastRowY, lastRowX, lastRowY)
        If lastRow < startDataRow Then GoTo NextYCol

        ' 2. 逐行收集可见且为数值的点（Y列含文本/空/隐藏行时自动跳过）
        Dim xList As Collection: Set xList = New Collection
        Dim yList As Collection: Set yList = New Collection
        Dim R As Long
        For R = startDataRow To lastRow
            If Not ws.Rows(R).Hidden Then
                Dim xv As Variant, yv As Variant
                xv = ws.Cells(R, dateCol).value
                yv = ws.Cells(R, yCol).value
                If Len(Trim(CStr(yv))) > 0 Then
                    If IsNumeric(yv) Then
                        xList.Add xv
                        yList.Add CDbl(yv)
                    End If
                End If
            End If
        Next R
        If xList.Count < 2 Then GoTo NextYCol  ' 控制图至少需2个点（否则算不出移动极差）

        ' 3. 集合转数组
        Dim nPts As Long: nPts = xList.Count
        Dim xArr() As Variant, yArr() As Variant
        ReDim xArr(1 To nPts)
        ReDim yArr(1 To nPts)
        Dim k As Long
        For k = 1 To nPts
            xArr(k) = xList(k)
            yArr(k) = yList(k)
        Next k

        ' 4. 计算控制限（单值 X 图，移动极差法）：
        '    CL = 均值；MR = 相邻两点极差均值；σ = MR / d2；UCL/LCL = CL ± 3σ
        Dim cl As Double, uc As Double, lc As Double
        Dim sumY As Double, sumMR As Double
        sumY = 0: sumMR = 0
        For k = 1 To nPts
            sumY = sumY + yArr(k)
        Next k
        cl = sumY / nPts
        For k = 2 To nPts
            sumMR = sumMR + Abs(yArr(k) - yArr(k - 1))
        Next k
        uc = cl + SIGMA_MULTIPLIER * (sumMR / (nPts - 1)) / D2
        lc = cl - SIGMA_MULTIPLIER * (sumMR / (nPts - 1)) / D2
        ' 注：若全部数据相同（MR=0），σ=0，UCL=LCL=CL，属正常退化情形

        ' 5. 超限点判定（超出 UCL/LCL 即标红）
        Dim oocCount As Long: oocCount = 0
        Dim oocFlag() As Boolean
        ReDim oocFlag(1 To nPts)
        For k = 1 To nPts
            oocFlag(k) = (yArr(k) > uc) Or (yArr(k) < lc)
            If oocFlag(k) Then oocCount = oocCount + 1
        Next k

        ' 6. 自动计算图表位置
        idx = idx + 1
        Dim chartLeft As Double, chartTop As Double
        chartLeft = marginLeft + ((idx - 1) Mod chartsPerRow) * (chartW + hGap)
        chartTop = marginTop + Int((idx - 1) / chartsPerRow) * (chartH + vGap)

        ' 7. 创建图表
        Dim chtObj As ChartObject
        Set chtObj = newWs.ChartObjects.Add(chartLeft, chartTop, chartW, chartH)
        Dim ch As Chart: Set ch = chtObj.Chart
        ch.ChartType = xlLineMarkers
        ch.PlotVisibleOnly = True

        ' 8. 清空默认系列
        Do While ch.SeriesCollection.Count > 0
            ch.SeriesCollection(1).Delete
        Loop

        ' 9. 写入辅助列（每列 6 列：CL/UCL/LCL/超限值/日期/数据），全部系列引用新表区域
        Dim hBaseCol As Long: hBaseCol = 1 + (colIdx - 1) * helperCols
        Dim hCL As Range, hUC As Range, hLC As Range, hOOC As Range, hX As Range, hY As Range
        Set hCL = newWs.Range(newWs.Cells(helperRow, hBaseCol), newWs.Cells(helperRow + nPts - 1, hBaseCol))
        Set hUC = newWs.Range(newWs.Cells(helperRow, hBaseCol + 1), newWs.Cells(helperRow + nPts - 1, hBaseCol + 1))
        Set hLC = newWs.Range(newWs.Cells(helperRow, hBaseCol + 2), newWs.Cells(helperRow + nPts - 1, hBaseCol + 2))
        Set hOOC = newWs.Range(newWs.Cells(helperRow, hBaseCol + 3), newWs.Cells(helperRow + nPts - 1, hBaseCol + 3))
        Set hX = newWs.Range(newWs.Cells(helperRow, hBaseCol + 4), newWs.Cells(helperRow + nPts - 1, hBaseCol + 4))
        Set hY = newWs.Range(newWs.Cells(helperRow, hBaseCol + 5), newWs.Cells(helperRow + nPts - 1, hBaseCol + 5))
        hCL.value = cl
        hUC.value = uc
        hLC.value = lc
        hX.value = xArr
        hY.value = yArr
        Dim oocArr() As Variant
        ReDim oocArr(1 To nPts)
        For k = 1 To nPts
            If oocFlag(k) Then
                oocArr(k) = yArr(k)
            Else
                oocArr(k) = Empty
            End If
        Next k
        hOOC.value = oocArr   ' 超限位置写值、其余留空（图表自动跳过空值）

        ' 10. 数据系列
        Dim seriesName As String
        seriesName = CStr(ws.Cells(titleRow, yCol).value)
        If Len(seriesName) = 0 Then seriesName = "Col" & yCol
        With ch.SeriesCollection.NewSeries
            .Name = seriesName
            .xValues = hX
            .Values = hY
            .Format.Line.ForeColor.RGB = lineColor
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
            .MarkerForegroundColor = markerColor
            .MarkerBackgroundColor = fillColor
        End With

        ' 11. 控制限系列（CL 实线绿，UCL/LCL 虚线红）
        With ch.SeriesCollection.NewSeries
            .Name = "CL"
            .xValues = hX
            .Values = hCL
            .Format.Line.ForeColor.RGB = clColor
            .MarkerStyle = xlMarkerStyleNone
        End With
        With ch.SeriesCollection.NewSeries
            .Name = "UCL"
            .xValues = hX
            .Values = hUC
            .Format.Line.ForeColor.RGB = uclColor
            .Format.Line.DashStyle = msoLineDash
            .MarkerStyle = xlMarkerStyleNone
        End With
        With ch.SeriesCollection.NewSeries
            .Name = "LCL"
            .xValues = hX
            .Values = hLC
            .Format.Line.ForeColor.RGB = uclColor
            .Format.Line.DashStyle = msoLineDash
            .MarkerStyle = xlMarkerStyleNone
        End With

        ' 12. 超限点系列（红色大标记，只画超限点）
        With ch.SeriesCollection.NewSeries
            .Name = "超限点"
            .xValues = hX
            .Values = hOOC
            .Format.Line.Visible = msoFalse
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 9
            .MarkerForegroundColor = oocColor
            .MarkerBackgroundColor = oocColor
        End With

        ' 13. Y轴范围：覆盖数据与控制限，预留5%边距并对齐刻度
        Dim spanMin As Double, spanMax As Double, dataSpan As Double
        spanMin = Application.Min(Application.Min(yArr), lc, uc)
        spanMax = Application.Max(Application.Max(yArr), lc, uc)
        dataSpan = spanMax - spanMin
        If dataSpan = 0 Then
            dataSpan = IIf(spanMin = 0, 1, Abs(spanMin) * 0.1)
        End If
        Dim yAxis As Axis
        Dim userMin As Double, userMax As Double, autoUnit As Double
        Set yAxis = ch.Axes(xlValue)
        userMin = spanMin - dataSpan * 0.05
        userMax = spanMax + dataSpan * 0.05
        With yAxis
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = userMin
            .MaximumScale = userMax
        End With
        ch.Refresh
        DoEvents
        autoUnit = yAxis.MajorUnit
        If autoUnit <> 0 Then
            userMin = Application.WorksheetFunction.RoundDown(userMin / autoUnit, 0) * autoUnit
            userMax = Application.WorksheetFunction.RoundUp(userMax / autoUnit, 0) * autoUnit
        End If
        With yAxis
            .MinimumScale = userMin
            .MaximumScale = userMax
        End With

        ' 14. 标题（样品类别/名称回退 + 超限个数注记）
        ch.HasTitle = True
        Dim chartTitle As String
        Dim catVal As Variant, nameVal As Variant
        chartTitle = ""
        If sampleCategoryColumn <> 0 Then
            catVal = ws.Cells(ws.Rows.Count, sampleCategoryColumn).End(xlUp).value
        End If
        If sampleCategoryColumn <> 0 And Not IsError(catVal) And Len(Trim(CStr(catVal))) > 0 _
           And CStr(catVal) <> "未知样品" And CStr(catVal) <> "样品类别" Then
            chartTitle = "[" & ws.Name & " - " & CStr(catVal) & " - " & seriesName & "]控制图"
        End If
        If Len(chartTitle) = 0 Then
            For col = 1 To ws.usedRange.Columns.Count
                If StrComp(ws.Cells(titleRow, col).value, "样品名称", vbTextCompare) = 0 Then
                    sampleNameColumn = col
                    Exit For
                End If
            Next col
            If sampleNameColumn <> 0 Then
                nameVal = ws.Cells(ws.Rows.Count, sampleNameColumn).End(xlUp).value
                If Not IsError(nameVal) And Len(Trim(CStr(nameVal))) > 0 And CStr(nameVal) <> "样品名称" Then
                    chartTitle = "[" & ws.Name & " - " & CStr(nameVal) & " - " & seriesName & "]控制图"
                End If
            End If
        End If
        If Len(chartTitle) = 0 Then
            chartTitle = "[" & ws.Name & " - " & seriesName & "]控制图"
        End If
        If oocCount > 0 Then
            chartTitle = chartTitle & "（超限 " & oocCount & " 点）"
        End If
        ch.chartTitle.Text = chartTitle
        With ch.chartTitle.Font
            .Name = "等线": .Size = 14
        End With

        ' 15. 坐标轴标题、字体
        ch.Axes(xlCategory).HasTitle = False
        ch.Axes(xlValue).HasTitle = False
        ch.Axes(xlCategory).TickLabels.Font.Name = "等线"
        ch.Axes(xlValue).TickLabels.Font.Name = "等线"

        ' 16. 网格线（浅灰虚线）
        If HasChartAxis(ch, xlValue) Then
            With ch.Axes(xlValue, xlPrimary).MajorGridlines.Format.Line
                .ForeColor.RGB = RGB(220, 220, 220)
                .DashStyle = msoLineDash
            End With
        End If

        ' 17. 图例：左上角（控制图需区分 数据/CL/UCL/LCL/超限点）
        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionTopLeft
        ch.Legend.Font.Name = "等线"
        ch.Legend.Font.Size = 9

        ' 18. 背景白底 + 蓝色边框
        ch.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ch.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        ch.ChartArea.Border.Color = 16711680

        ' 19. X轴日期格式（仅当X是日期时应用）
        On Error Resume Next
        If IsDate(xArr(1)) Then
            If HasChartAxis(ch, xlCategory) Then
                With ch.Axes(xlCategory, xlPrimary)
                    .TickLabels.NumberFormat = "m月d日"
                End With
            End If
        End If
        On Error GoTo 0

NextYCol:
    Next yCol

    '===================== 完成提示 =====================
    MsgBox "完成！共生成 " & idx & " 张控制图（两列排列）", vbInformation
    Exit Sub

    '===================== 报错处理 =====================
ErrHandler:
    MsgBox "执行出错：" & Err.Number & " - " & Err.Description, vbCritical
End Sub
