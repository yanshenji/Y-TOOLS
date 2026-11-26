Attribute VB_Name = "模块3"
Sub 画标记折线图()
    ' 声明变量
    Dim ws As Worksheet
    Dim selectedRng As Range
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim xRng As Range, yRng As Range
    Dim titleRow As Long, lastRow As Long, startRow As Long
    Dim selectedCol As Integer, curRow As Long
    Dim sampleCol As Integer
    Dim chartTitle As String
    Dim lineColor As Long, markerColor As Long, fillColor As Long
    
    ' 1. 基础校验
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中数据单元格！", vbExclamation
        Exit Sub
    End If
    Set selectedRng = Selection
    Set ws = ActiveSheet
    selectedCol = selectedRng.Column
    curRow = selectedRng.Row
    
    ' 2. 识别标题行
    If ws.AutoFilterMode Then
        titleRow = ws.AutoFilter.Range.Row
    Else
        titleRow = 1
    End If
    startRow = titleRow + 1
    
    ' 3. 获取数据最后一行
    lastRow = ws.Cells(ws.Rows.Count, selectedCol).End(xlUp).Row
    If lastRow < startRow Then
        MsgBox "选中列无有效数据！", vbExclamation
        Exit Sub
    End If
    
    ' 4. 获取可见数据区域（筛选模式）
    On Error Resume Next
    Set yRng = ws.Range(ws.Cells(startRow, selectedCol), ws.Cells(lastRow, selectedCol)).SpecialCells(xlCellTypeVisible)
    Set xRng = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, 1)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If yRng Is Nothing Or xRng Is Nothing Then
        MsgBox "无可见数据！", vbExclamation
        Exit Sub
    End If
    
    ' 5. 查找样品列
    sampleCol = 0
    Dim col As Integer
    For col = 1 To ws.usedRange.Columns.Count
        If StrComp(ws.Cells(titleRow, col).value, "样品类别", vbTextCompare) = 0 Or _
           StrComp(ws.Cells(titleRow, col).value, "样品名称", vbTextCompare) = 0 Then
            sampleCol = col
            Exit For
        End If
    Next col
    
    ' 6. 配色方案
    lineColor = RGB(91, 155, 213)
    markerColor = lineColor
    fillColor = RGB(237, 125, 49)
    
    ' 7. 创建图表
    Dim chartLeft As Double, chartTop As Double, chartW As Double, chartH As Double
    chartW = 16.33 * 28.35
    chartH = 6.66 * 28.35
    chartLeft = selectedRng.Left + selectedRng.Width
    chartTop = selectedRng.Top + selectedRng.Height
    
    Set chartObj = ws.ChartObjects.Add(chartLeft, chartTop, chartW, chartH)
    Set cht = chartObj.Chart
    cht.ChartType = xlLineMarkers
    
    ' 清空默认系列
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    
    ' 添加数据系列
    With cht.SeriesCollection.NewSeries
        .Name = ws.Cells(titleRow, selectedCol).value
        .xValues = xRng
        .Values = yRng
        .Format.Line.ForeColor.RGB = lineColor
        .MarkerStyle = xlMarkerStyleCircle
        .MarkerForegroundColor = markerColor
        .MarkerBackgroundColor = fillColor
    End With
    
    ' 8. 图表标题
    If sampleCol <> 0 Then
        chartTitle = ws.Name & " - " & ws.Cells(curRow, sampleCol).value & " - " & ws.Cells(titleRow, selectedCol).value & " 趋势图"
    Else
        chartTitle = ws.Name & " - " & ws.Cells(titleRow, selectedCol).value & " 趋势图"
    End If
    cht.HasTitle = True
    cht.chartTitle.Text = chartTitle
    With cht.chartTitle.Font
        .Name = "等线"
        .Size = 14
    End With
    
    ' 9. 坐标轴基础样式
    cht.Axes(xlCategory).HasTitle = False
    cht.Axes(xlValue).HasTitle = False
    cht.Axes(xlCategory).TickLabels.Font.Name = "等线"
    cht.Axes(xlValue).TickLabels.Font.Name = "等线"
    
    ' 网格线
    With cht.Axes(xlValue).MajorGridlines.Format.Line
        .ForeColor.RGB = RGB(220, 220, 220)
        .DashStyle = msoLineDash
    End With

' ===================== 【核心修正】先设最大/最小值 → 再读取WPS自动单位 =====================
    Dim yAxis As Axis
    Dim dataMin As Double, dataMax As Double
    Dim userMin As Double, userMax As Double
    Dim autoUnit As Double
    
    Set yAxis = cht.Axes(xlValue)
    
    ' 第一步：计算数据真实最大/最小值
    dataMin = Application.WorksheetFunction.Min(yRng)
    dataMax = Application.WorksheetFunction.Max(yRng)
    
    ' 第二步：手动设置坐标轴范围（预留5%余量）
    userMin = dataMin - (dataMax - dataMin) * 0.05
    userMax = dataMax + (dataMax - dataMin) * 0.05
    
    ' 第三步：锁定最大最小值，强制WPS计算自动单位
    With yAxis
        .MinimumScaleIsAuto = False
        .MaximumScaleIsAuto = False
        .MinimumScale = userMin
        .MaximumScale = userMax
    End With
    
    ' 刷新图表，让WPS自动计算出最优 MajorUnit
    cht.Refresh
    DoEvents
    
    ' 第四步：读取WPS自动生成的最优刻度单位（零报错）
    autoUnit = yAxis.MajorUnit
    
    ' 第五步：规整边界（对齐刻度，无零头）
    userMin = Application.WorksheetFunction.RoundDown(userMin / autoUnit, 0) * autoUnit
    userMax = Application.WorksheetFunction.RoundUp(userMax / autoUnit, 0) * autoUnit
    
    ' 第六步：应用最终设置
    With yAxis
        .MinimumScale = userMin
        .MaximumScale = userMax
        '.MajorUnit = autoUnit
        '.TickLabels.NumberFormat = IIf(autoUnit < 1, "0.0", "0")
    End With
' ==========================================================================================
    
    ' 10. 图表布局
    cht.Legend.Position = xlLegendPositionBottom
    cht.ChartArea.Format.Fill.Visible = msoTrue
    cht.PlotArea.Format.Fill.Visible = msoTrue
    cht.ChartArea.Format.Line.Visible = msoFalse
    
    ' 清理对象
    Set xRng = Nothing: Set yRng = Nothing
    Set cht = Nothing: Set chartObj = Nothing
    Set ws = Nothing: Set selectedRng = Nothing
    
    'MsgBox "绘制完成！" & vbCrLf & "自动刻度单位：" & autoUnit & vbCrLf & "范围：" & userMin & " ~ " & userMax
End Sub

Sub 画标记折线图_0()
    ' 声明变量
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim xValues As Range
    Dim yValues As Range
    Dim seriesIndex As Integer
    Dim seriesName As String
    Dim markerColor As Long
    Dim lineColor As Long
    Dim markerFillColor As Long
    Dim selectedCell As Range
    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim sampleCategoryColumn As Integer ' 样品类别列
    Dim sampleNameColumn As Integer ' 样品名称列
    Dim currentRow As Long ' 当前行号
    Dim titleRow As Long ' 标题行号

    If Not TypeName(Selection) = "Range" Then
        MsgBox "请先选中一个单元格区域。"
        Exit Sub
    End If
    
    ' 设置工作表
    Set ws = ActiveSheet ' 使用当前活动工作表
    
    ' 确定当前选中单元格
    Set selectedCell = Selection

    ' 设置图表的大小（转换为磅，1厘米=28.35磅）
    chartWidth = 16.33 * 28.35
    chartHeight = 6.66 * 28.35

    ' 计算图表左上角位置，使其左上角位于选中单元格的右下角
    chartLeft = selectedCell.Left + selectedCell.Width
    chartTop = selectedCell.Top + selectedCell.Height

    ' 确定当前选中列（假设用户已选中一列数据）
    Dim selectedColumn As Integer
    selectedColumn = selectedCell.Column

    ' 设置标题行（检查是否进行筛选）
    If ws.AutoFilterMode Then
        'If ws.FilterMode Then
            ' 如果进行了筛选，标题行是筛选的第一行
            Dim filterRange As Range
            Set filterRange = ws.AutoFilter.Range
            titleRow = filterRange.Row
        'Else
            ' 如果有筛选但处于非筛选状态，标题行仍是第1行
            'titleRow = 1
        'End If
    Else
        ' 如果没有筛选，标题行是第1行
        titleRow = 1
    End If

    ' 设置数据区域（从标题行下一行开始，到最后一行非空单元格）
    Dim startDataRow As Long
    startDataRow = titleRow + 1
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, selectedColumn).End(xlUp).Row

    Set xValues = ws.Range(ws.Cells(startDataRow, 1), ws.Cells(lastRow, 1)) ' 第一列为 X 值
    Set yValues = ws.Range(ws.Cells(startDataRow, selectedColumn), ws.Cells(lastRow, selectedColumn)) ' 当前选中列为 Y 值

    ' 设置系列名称（从当前选中列的标题行获取）
    seriesName = ws.Cells(titleRow, selectedColumn).value

    ' 设置线条颜色（这里设置为蓝色，可以修改为你想要的颜色代码）
    lineColor = RGB(91, 155, 213)
    ' 设置标记颜色（与线条颜色一致）
    markerColor = lineColor
    ' 设置标记填充颜色（主题颜色的下一个颜色）
    markerFillColor = RGB(237, 125, 49) ' 示例：橙色作为主题颜色的下一个颜色，你可以根据需要修改

    ' 添加图表对象
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartWidth, Top:=chartTop, Height:=chartHeight)
    Set cht = chartObj.Chart

    ' 设置图表类型为带数据标记的折线图
    cht.ChartType = xlLineMarkers

    ' 清除图表中的默认系列
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' 添加数据系列
    With cht.SeriesCollection.NewSeries()
        .Name = seriesName ' 设置系列名称
        .xValues = xValues ' 设置 x 轴数据
        .Values = yValues ' 设置 y 轴数据
        .Format.Line.ForeColor.RGB = lineColor ' 设置线条颜色
        .MarkerStyle = xlMarkerStyleCircle ' 设置标记为圆形
        .MarkerForegroundColor = markerColor ' 设置标记线条颜色
        .MarkerBackgroundColor = markerFillColor ' 设置标记填充颜色
    End With

    ' 设置图表标题
    cht.HasTitle = True
    
    ' 查找“样品类别”列或“样品名称”列
    sampleCategoryColumn = 0
    sampleNameColumn = 0
    
    ' 遍历工作表的标题行（当前标题行是titleRow），查找“样品类别”或“样品名称”列
    Dim col As Integer
    For col = 1 To ws.usedRange.Columns.Count
        If StrComp(ws.Cells(titleRow, col).value, "样品类别", vbTextCompare) = 0 Then
            sampleCategoryColumn = col
            Exit For
        End If
    Next col
    
    If sampleCategoryColumn = 0 Then
        For col = 1 To ws.usedRange.Columns.Count
            If StrComp(ws.Cells(titleRow, col).value, "样品名称", vbTextCompare) = 0 Then
                sampleNameColumn = col
                Exit For
            End If
        Next col
    End If
    
    ' 当前行号（选中单元格的行号）
    currentRow = selectedCell.Row
    
    ' 组合标题
    Dim chartTitle As String
    If sampleCategoryColumn <> 0 Then
        chartTitle = "[" & ws.Name & " - " & ws.Cells(currentRow, sampleCategoryColumn).value & " - " & ws.Cells(titleRow, selectedColumn).value & "]趋势图"
    ElseIf sampleNameColumn <> 0 Then
        chartTitle = "[" & ws.Name & " - " & ws.Cells(currentRow, sampleNameColumn).value & " - " & ws.Cells(titleRow, selectedColumn).value & "]趋势图"
    Else
        chartTitle = "[" & ws.Name & " - " & ws.Cells(titleRow, selectedColumn).value & "]趋势图"
    End If
    
    cht.chartTitle.Text = chartTitle ' 设置图表标题内容
    cht.chartTitle.Font.Size = 14 ' 设置图表标题字体大小为 14
    cht.chartTitle.Font.Name = "等线" ' 设置图表标题字体为“等线”

    ' 取消坐标轴标题
    cht.Axes(xlCategory).HasTitle = False ' 取消 X 轴标题
    cht.Axes(xlValue).HasTitle = False ' 取消 Y 轴标题

    ' 设置坐标轴标签字体
    cht.Axes(xlCategory).TickLabels.Font.Name = "等线" ' 设置 X 轴标签字体为“等线”
    cht.Axes(xlValue).TickLabels.Font.Name = "等线" ' 设置 Y 轴标签字体为“等线”

    ' 设置主要网格线
    cht.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200) ' 设置 Y 轴主要网格线颜色
    cht.Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineDash ' 设置 Y 轴主要网格线为虚线样式

    ' 可选：调整图表布局和样式
    cht.Legend.Position = xlLegendPositionBottom ' 设置图例位置到下方
    
    ' 设置图表背景为白色
    cht.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' 白色背景
    cht.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)  ' 绘图区也设置为白色
    
   ' 方法1：使用RGB值的数值形式（WPS有时对这种形式支持更好）
    'With chartObj.ShapeRange.Line
        '.Visible = msoTrue
        '.ForeColor.RGB = 16711680 ' 这是RGB(0,0,255)的数值等效值
        '.Weight = 0.5
        '.Transparency = 0
    'End With
    
    ' 方法2：通过图表区格式设置
    With cht.ChartArea.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)  ' 蓝色的数值表示
        .Weight = 0.5
    End With
    
    ' 方法3：使用WPS兼容的主题色索引（12通常代表蓝色）
    'chartObj.ShapeRange.Line.ForeColor.Index = 12
    
    On Error GoTo 0

    ' 清理对象
    Set xValues = Nothing
    Set yValues = Nothing
    Set chartObj = Nothing
    Set cht = Nothing
    Set ws = Nothing
End Sub

Sub 添加辅助线0()
    ' 声明变量
    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim seriesName As String
    Dim yValue As Double
    Dim xValues As Variant
    Dim minX As Double, maxX As Double
    Dim helperXValues As Variant
    Dim helperYValues As Variant
    Dim pointCount As Integer

    ' 设置工作表
    Set ws = ActiveSheet ' 使用当前活动工作表

    ' 确保用户选中了一个图表区域
    If Not TypeOf Selection Is ChartArea Then
        MsgBox "请先选中一个图表区域！", vbExclamation
        Exit Sub
    End If

    ' 获取选中的图表对象
    Set cht = Selection.Parent ' 从选中的图表区域获取图表对象

    ' 弹出输入框获取系列名称和 y 值
    seriesName = InputBox("请输入辅助线的系列名称（留空则默认为[标准]）：", "输入系列名称")
    If seriesName = "" Then seriesName = "标准"

    On Error Resume Next ' 捕获输入非数字的错误
    yValue = CDbl(InputBox("请输入辅助线的 Y 值：", "输入 Y 值"))
    If Err.Number <> 0 Then
        MsgBox "输入的 Y 值无效，请输入一个数字！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' 确定 x 轴数据范围（假设与第一个系列的 x 轴数据相同）
    If cht.SeriesCollection.Count > 0 Then
        xValues = cht.SeriesCollection(1).xValues
        minX = WorksheetFunction.Min(xValues)
        maxX = WorksheetFunction.Max(xValues)
        pointCount = UBound(xValues) - LBound(xValues) + 1 ' 获取第一个系列的点数
    Else
        MsgBox "图表中没有数据系列，无法添加辅助线！", vbExclamation
        Exit Sub
    End If

    ' 生成辅助线的 x 值和 y 值（根据第一个系列的点数）
    ReDim helperXValues(1 To pointCount)
    ReDim helperYValues(1 To pointCount)
    Dim i As Integer
    For i = 1 To pointCount
        helperXValues(i) = minX + (maxX - minX) * (i - 1) / (pointCount - 1)
        helperYValues(i) = yValue
    Next i

    ' 添加辅助线系列
    With cht.SeriesCollection.NewSeries()
        .Name = seriesName
        .xValues = helperXValues
        .Values = helperYValues
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0) ' 设置辅助线为红色
        .Format.Line.DashStyle = msoLineDash ' 设置辅助线为虚线样式
        .Format.Line.Weight = 1.5 ' 设置辅助线宽度
        .MarkerStyle = xlMarkerStyleNone ' 不显示标记
    End With

    ' 提示完成
    MsgBox "辅助线已添加到图表！", vbInformation

    ' 清理对象
    Set ws = Nothing
    Set cht = Nothing
End Sub

Function 添加辅助线函数(seriesName As String, yValue As Double)
    添加辅助线函数 = 模块6.添加辅助线函数_final(seriesName, yValue)
End Function

Sub 添加辅助线()
    模块6.添加辅助线
End Sub

Sub 画钴杂比标记折线图()
    ' 声明变量
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim xValues() As Variant
    Dim yValues() As Variant
    Dim seriesIndex As Integer
    Dim seriesName As String
    Dim markerColor As Long
    Dim lineColor As Long
    Dim markerFillColor As Long
    Dim selectedCell As Range
    Dim chartWidth As Double
    Dim chartHeight As Double
    Dim chartLeft As Double
    Dim chartTop As Double
    Dim sampleCategoryColumn As Integer ' 样品类别列
    Dim sampleNameColumn As Integer ' 样品名称列
    Dim currentRow As Integer ' 当前行号
    Dim titleRow As Long ' 标题行号
    Dim coColumn As Integer ' 钴值列号
    Dim 杂质Column As Integer ' 杂质当前列号

    If Not TypeName(Selection) = "Range" Then
        MsgBox "请先选中一个单元格区域。"
        Exit Sub
    End If
    
    ' 设置工作表
    Set ws = ActiveSheet ' 使用当前活动工作表
    
    ' 确定当前选中单元格
    Set selectedCell = Selection

    ' 设置图表的大小（转换为磅，1厘米=28.35磅）
    chartWidth = 16.33 * 28.35
    chartHeight = 6.66 * 28.35

    ' 计算图表左上角位置，使其左上角位于选中单元格的右下角
    chartLeft = selectedCell.Left + selectedCell.Width
    chartTop = selectedCell.Top + selectedCell.Height

    ' 确定当前选中列（杂质列）
    杂质Column = selectedCell.Column

    ' 设置标题行（检查是否进行筛选）
    If ws.AutoFilterMode Then
        Dim filterRange As Range
        Set filterRange = ws.AutoFilter.Range
        titleRow = filterRange.Row
    Else
        titleRow = 1
    End If

    ' 查找钴值列（标题中包含"Co"的列）
    coColumn = 0
    Dim col As Integer
    For col = 1 To ws.usedRange.Columns.Count
        If InStr(1, ws.Cells(titleRow, col).value, "Co", vbTextCompare) > 0 Then
            coColumn = col
            Exit For
        End If
    Next col

    If coColumn = 0 Then
        MsgBox "未找到标题中包含'Co'的列，请检查数据标题。"
        Exit Sub
    End If

    ' 检查Co值是否全为0
    Dim coRange As Range
    Set coRange = ws.Range(ws.Cells(titleRow + 1, coColumn), ws.Cells(ws.Rows.Count, coColumn).End(xlUp))
    Dim coValue As Variant
    Dim isCoZero As Boolean
    isCoZero = True
    For Each coValue In coRange
        If coValue.value <> 0 Then
            isCoZero = False
            Exit For
        End If
    Next coValue

    If isCoZero Then
        MsgBox "Co值全为0，无法绘制钴杂比图表。"
        Exit Sub
    End If

    ' 设置数据区域（从标题行下一行开始，到最后一行非空单元格）
    Dim startDataRow As Long
    startDataRow = titleRow + 1
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 杂质Column).End(xlUp).Row

    ' 获取可见区域的X值
    Dim visibleXRange As Range
    On Error Resume Next ' 忽略错误，如果没有任何可见单元格
    Set visibleXRange = ws.Range(ws.Cells(startDataRow, 1), ws.Cells(lastRow, 1)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' 恢复错误处理

    If visibleXRange Is Nothing Then
        MsgBox "没有找到可见的X值数据。"
        Exit Sub
    End If

    ' 将可见区域的X值转换为数组
    Dim xAreas As Areas
    Dim xAreaValues As Variant
    Dim xList As New Collection
    Dim xArea As Range
    Dim xRow As Long

    Set xAreas = visibleXRange.Areas
    For Each xArea In xAreas
        xAreaValues = xArea.value
        If IsArray(xAreaValues) Then
            For xRow = 1 To UBound(xAreaValues, 1)
                xList.Add xAreaValues(xRow, 1)
            Next xRow
        Else
            xList.Add xAreaValues
        End If
    Next xArea

    ReDim xValues(1 To xList.Count)
    For xRow = 1 To xList.Count
        xValues(xRow) = xList(xRow)
    Next xRow

    ' 获取可见区域的钴值和杂质值
    Dim visibleCoRange As Range
    Dim visible杂质Range As Range
    On Error Resume Next ' 忽略错误，如果没有任何可见单元格
    Set visibleCoRange = ws.Range(ws.Cells(startDataRow, coColumn), ws.Cells(lastRow, coColumn)).SpecialCells(xlCellTypeVisible)
    Set visible杂质Range = ws.Range(ws.Cells(startDataRow, 杂质Column), ws.Cells(lastRow, 杂质Column)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0 ' 恢复错误处理

    If visibleCoRange Is Nothing Or visible杂质Range Is Nothing Then
        MsgBox "没有找到可见的钴值或杂质值数据。"
        Exit Sub
    End If

    ' 计算钴杂比值并创建数组
    Dim coAreas As Areas
    Dim 杂质Areas As Areas
    Dim coArea As Range
    Dim 杂质Area As Range
    Dim coAreaValues As Variant
    Dim 杂质AreaValues As Variant
    Dim cobaltRatioList As New Collection
    Dim coRow As Long
    Dim 杂质Row As Long

    Set coAreas = visibleCoRange.Areas
    Set 杂质Areas = visible杂质Range.Areas

    For Each coArea In coAreas
        Set杂质Area = 杂质Areas.Item(coAreas.Item(coAreas.IndexOf(coArea)).Row - startDataRow + 1)
        coAreaValues = coArea.value
       杂质AreaValues = 杂质Area.value

        If IsArray(coAreaValues) Then
            For coRow = 1 To UBound(coAreaValues, 1)
                If 杂质AreaValues(coRow, 1) <> 0 Then
                    cobaltRatioList.Add coAreaValues(coRow, 1) / 杂质AreaValues(coRow, 1)
                Else
                    cobaltRatioList.Add 0
                End If
            Next coRow
        Else
            If 杂质AreaValues <> 0 Then
                cobaltRatioList.Add coAreaValues / 杂质AreaValues
            Else
                cobaltRatioList.Add 0
            End If
        End If
    Next coArea

    ReDim 钴杂比Values(1 To cobaltRatioList.Count)
    For i = 1 To cobaltRatioList.Count
        钴杂比Values(i) = cobaltRatioList(i)
    Next i

    ' 确保钴杂比值数组和X值数组长度一致
    If UBound(xValues) <> UBound(钴杂比Values) Then
        MsgBox "X值和钴杂比值数量不一致，无法绘制图表。"
        Exit Sub
    End If

    ' 设置系列名称（钴杂比）
    seriesName = "钴杂比 (Co/" & ws.Cells(titleRow, 杂质Column).value & ")"

    ' 设置线条颜色（这里设置为蓝色，可以修改为你想要的颜色代码）
    lineColor = RGB(91, 155, 213)
    ' 设置标记颜色（与线条颜色一致）
    markerColor = lineColor
    ' 设置标记填充颜色（主题颜色的下一个颜色）
    markerFillColor = RGB(237, 125, 49) ' 示例：橙色作为主题颜色的下一个颜色，你可以根据需要修改

    ' 添加图表对象
    Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Width:=chartWidth, Top:=chartTop, Height:=chartHeight)
    Set cht = chartObj.Chart

    ' 设置图表类型为带数据标记的折线图
    cht.ChartType = xlLineMarkers

    ' 清除图表中的默认系列
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    ' 添加数据系列
    With cht.SeriesCollection.NewSeries()
        .Name = seriesName ' 设置系列名称
        .xValues = xValues ' 设置 x 轴数据（仅可见区域）
        .Values = 钴杂比Values ' 设置钴杂比值作为 Y 值
        .Format.Line.ForeColor.RGB = lineColor ' 设置线条颜色
        .MarkerStyle = xlMarkerStyleCircle ' 设置标记为圆形
        .MarkerForegroundColor = markerColor ' 设置标记线条颜色
        .MarkerBackgroundColor = markerFillColor ' 设置标记填充颜色
    End With

    ' 设置图表标题
    cht.HasTitle = True
    
    ' 查找“样品类别”列或“样品名称”列
    sampleCategoryColumn = 0
    sampleNameColumn = 0
    
    ' 遍历工作表的标题行（当前标题行是titleRow），查找“样品类别”或“样品名称”列
    For col = 1 To ws.usedRange.Columns.Count
        If StrComp(ws.Cells(titleRow, col).value, "样品类别", vbTextCompare) = 0 Then
            sampleCategoryColumn = col
            Exit For
        End If
    Next col
    
    If sampleCategoryColumn = 0 Then
        For col = 1 To ws.usedRange.Columns.Count
            If StrComp(ws.Cells(titleRow, col).value, "样品名称", vbTextCompare) = 0 Then
                sampleNameColumn = col
                Exit For
            End If
        Next col
    End If
    
    ' 当前行号（选中单元格的行号）
    currentRow = selectedCell.Row
    
    ' 组合标题
    Dim chartTitle As String
    If sampleCategoryColumn <> 0 Then
        chartTitle = "[" & ws.Name & " - " & ws.Cells(currentRow, sampleCategoryColumn).value & " - 钴杂比趋势图]"
    ElseIf sampleNameColumn <> 0 Then
        chartTitle = "[" & ws.Name & " - " & ws.Cells(currentRow, sampleNameColumn).value & " - 钴杂比趋势图]"
    Else
        chartTitle = "[" & ws.Name & " - 钴杂比趋势图]"
    End If
    
    cht.chartTitle.Text = chartTitle ' 设置图表标题内容
    cht.chartTitle.Font.Size = 14 ' 设置图表标题字体大小为 14
    cht.chartTitle.Font.Name = "等线" ' 设置图表标题字体为“等线”

    ' 取消坐标轴标题
    cht.Axes(xlCategory).HasTitle = False ' 取消 X 轴标题
    cht.Axes(xlValue).HasTitle = False ' 取消 Y 轴标题

    ' 设置坐标轴标签字体
    cht.Axes(xlCategory).TickLabels.Font.Name = "等线" ' 设置 X 轴标签字体为“等线”
    cht.Axes(xlValue).TickLabels.Font.Name = "等线" ' 设置 Y 轴标签字体为“等线”

    ' 设置主要网格线
    cht.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(200, 200, 200) ' 设置 Y 轴主要网格线颜色
    cht.Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineDash ' 设置 Y 轴主要网格线为虚线样式

    ' 可选：调整图表布局和样式
    cht.Legend.Position = xlLegendPositionBottom ' 设置图例位置到下方

    ' 设置图表轮廓颜色
    With chartObj.ShapeRange.Line
        .ForeColor.RGB = RGB(0, 0, 255) ' 设置图表轮廓为蓝色
        .Weight = 2.5 ' 设置边框线条粗细，可根据需要调整
    End With

    ' 清理对象
    Set chartObj = Nothing
    Set cht = Nothing
    Set ws = Nothing
End Sub
