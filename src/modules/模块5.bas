Attribute VB_Name = "模块5"
' 智能格式化：≥1保留两位小数，<1保留两位有效数字
Private Function FormatSmartNumber(ByVal v As Double) As String
    Dim av As Double, p As Double, factor As Double
    If v = 0 Then
        FormatSmartNumber = "0"
        Exit Function
    End If
    av = Abs(v)
    If av >= 1 Then
        FormatSmartNumber = Format(v, "0.00")
    Else
        p = Int(Log(av) / Log(10#))
        factor = 10 ^ (2 - p - 1)
        FormatSmartNumber = CStr(Round(v * factor) / factor)
    End If
End Function

' 在首行查找匹配的列名
Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerNames As Variant) As Long
    Dim lastCol As Long, i As Long, j As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For j = LBound(headerNames) To UBound(headerNames)
        For i = 1 To lastCol
            If Trim(CStr(ws.Cells(1, i).value)) = headerNames(j) Then
                FindHeaderCol = i
                Exit Function
            End If
        Next i
    Next j
    FindHeaderCol = 0
End Function

' 向上取整为 Long
Private Function RoundUpLong(ByVal x As Double) As Long
    Dim i As Long
    i = Int(x)
    If x > i Then i = i + 1
    If i < 1 Then i = 1
    RoundUpLong = i
End Function

' 生成安全且唯一的工作表名；同名自动追加 2、3…
Private Function MakeSafeSheetName(ByVal baseName As String, ByVal wb As Workbook) As String
    Dim S As String, bad As Variant, i As Long, n As Long, candidate As String, ws As Worksheet
    S = Trim(baseName)
    bad = Array("[", "]", ":", "*", "?", "/", "\", Chr(10), Chr(13))
    For i = LBound(bad) To UBound(bad)
        S = Replace(S, bad(i), "_")
    Next i
    If Len(S) = 0 Then S = "图表"
    If Len(S) > 31 Then S = Left(S, 31)
    candidate = S: n = 1
    Do
        On Error Resume Next: Set ws = wb.Sheets(candidate): On Error GoTo 0
        If ws Is Nothing Then
            MakeSafeSheetName = candidate
            Exit Function
        Else
            n = n + 1
            candidate = Left(S, 31 - Len(CStr(n))) & CStr(n)
            Set ws = Nothing
        End If
    Loop
End Function

' 是否包含逻辑连接符（逗号=OR 或 &=AND）
Private Function HasLogicalOps(ByVal expr As String) As Boolean
    HasLogicalOps = (InStr(1, expr, ",") > 0) Or (InStr(1, expr, "&") > 0)
End Function

' 名称逻辑匹配（支持 OR 用逗号, AND 用 &；支持 * ?；大小写不敏感）
Private Function NameMatches(ByVal textVal As String, ByVal expr As String) As Boolean
    Dim orParts As Variant, andParts As Variant
    Dim i As Long, j As Long, t As String, pat As String, okAnd As Boolean

    If Len(Trim(expr)) = 0 Or expr = "*" Then
        NameMatches = True
        Exit Function
    End If

    textVal = LCase$(CStr(textVal))
    orParts = Split(expr, ",")

    For i = LBound(orParts) To UBound(orParts)
        andParts = Split(orParts(i), "&")
        okAnd = True
        For j = LBound(andParts) To UBound(andParts)
            t = LCase$(Trim(andParts(j)))
            If t <> "" Then
                If InStr(t, "*") = 0 And InStr(t, "?") = 0 Then
                    pat = "*" & t & "*"
                Else
                    pat = t
                End If
                If Not (textVal Like pat) Then
                    okAnd = False
                    Exit For
                End If
            End If
        Next j
        If okAnd Then
            NameMatches = True
            Exit Function
        End If
    Next i

    NameMatches = False
End Function

' 在当前(dataRng, 已做日期筛选后)基础上，对“样品名称列”追加逻辑过滤（行隐藏）
Private Sub ApplyNameLogicalFilter(ByVal ws As Worksheet, ByVal dataRng As Range, ByVal nameCol As Long, ByVal expr As String)
    Dim lastRow As Long, R As Long, c As Long
    If Len(Trim(expr)) = 0 Or expr = "*" Then Exit Sub
    lastRow = dataRng.Row + dataRng.Rows.Count - 1
    c = nameCol
    For R = dataRng.Row + 1 To lastRow
        If ws.Rows(R).Hidden = False Then
            If Not NameMatches(ws.Cells(R, c).value, expr) Then
                ws.Rows(R).Hidden = True
            End If
        End If
    Next R
End Sub

' ================= 主过程 =================
Sub GenerateCharts_AndExportData_ByRule()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsConfig As Worksheet
    On Error Resume Next: Set wsConfig = wb.Sheets("配置表"): On Error GoTo 0
    If wsConfig Is Nothing Then
        MsgBox "未找到配置表！", vbCritical
        Exit Sub
    End If

    ' 参数（从 B2 开始）
    Dim chartsPerRow As Long, chartWidth As Double, chartHeight As Double, titleFontSize As Long
    chartsPerRow = CLng(wsConfig.Range("B2").value)
    chartWidth = CDbl(wsConfig.Range("B3").value)
    chartHeight = CDbl(wsConfig.Range("B4").value)
    titleFontSize = CLng(wsConfig.Range("B5").value)

    Dim lastRuleRow As Long: lastRuleRow = wsConfig.Cells(wsConfig.Rows.Count, 1).End(xlUp).Row

    Dim R As Long, sheetName As String, wsData As Worksheet
    Dim wsOutput As Worksheet, wsDataOut As Worksheet
    Dim dataRng As Range, filteredRng As Range, headerRow As Range
    Dim nameCol As Long, dateCol As Long, dateColAlt As Long
    Dim sampleType As String, pattern As String, startDate As Variant, endDate As Variant, itemList As String
    Dim autoList As String, c As Range, items As Variant, it As Variant
    Dim chartObj As ChartObject
    Dim outRow As Long, outCol As Long, totalMade As Long

    For R = 7 To lastRuleRow
        sheetName = Trim(CStr(wsConfig.Cells(R, 1).value))
        If sheetName = "" Then GoTo NextRule

        Set wsData = wb.Sheets(sheetName)
        ' 复位任何手动隐藏（避免上一轮残留）
        wsData.Rows.Hidden = False

        nameCol = FindHeaderCol(wsData, Array("样品名称", "樣品名稱"))
        dateCol = FindHeaderCol(wsData, Array("送样日期", "送樣日期"))
        dateColAlt = FindHeaderCol(wsData, Array("取样日期", "取樣日期"))
        If nameCol = 0 Then GoTo NextRule
        If dateCol = 0 And dateColAlt = 0 Then GoTo NextRule
        If dateCol = 0 Then dateCol = dateColAlt

        ' 读取规则
        sampleType = Trim(CStr(wsConfig.Cells(R, 2).value))
        pattern = CStr(wsConfig.Cells(R, 3).value)      ' 逻辑表达式或普通通配
        startDate = wsConfig.Cells(R, 4).value
        endDate = wsConfig.Cells(R, 5).value
        itemList = CStr(wsConfig.Cells(R, 6).value)

        ' —— 新建输出工作表：图表、数据 ——（唯一命名）
        Dim chartSheetName As String, dataSheetName As String
        Dim baseChart As String, baseData As String
        If sampleType = "" Then sampleType = pattern
        baseChart = sheetName & "_" & sampleType & "_图表"
        baseData = sheetName & "_" & sampleType & "_数据"
        chartSheetName = MakeSafeSheetName(baseChart, wb)
        dataSheetName = MakeSafeSheetName(baseData, wb)

        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = chartSheetName

        Set wsDataOut = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsDataOut.Name = dataSheetName

        outRow = 1: outCol = 1

        ' 数据区域
        Set dataRng = wsData.Range("A1").CurrentRegion
        Set headerRow = dataRng.Rows(1)
        If wsData.AutoFilterMode Then wsData.AutoFilterMode = False

        ' 先做日期 AutoFilter（快且稳）
        If Not IsEmpty(startDate) And Not IsEmpty(endDate) Then
            dataRng.AutoFilter Field:=dateCol, Criteria1:=">=" & CDbl(startDate), _
                Operator:=xlAnd, Criteria2:="<=" & CDbl(endDate)
        ElseIf Not IsEmpty(startDate) Then
            dataRng.AutoFilter Field:=dateCol, Criteria1:=">=" & CDbl(startDate)
        ElseIf Not IsEmpty(endDate) Then
            dataRng.AutoFilter Field:=dateCol, Criteria1:="<=" & CDbl(endDate)
        End If

        ' 名称：有逻辑符→自定义过滤；无→原 AutoFilter
        If Len(Trim(pattern)) > 0 And pattern <> "*" Then
            If HasLogicalOps(pattern) Then
                ApplyNameLogicalFilter wsData, dataRng, nameCol, pattern
            Else
                dataRng.AutoFilter Field:=nameCol, Criteria1:=pattern
            End If
        End If

        ' 排序（按日期升序）
        dataRng.Sort Key1:=dataRng.Columns(dateCol), Order1:=xlAscending, Header:=xlYes

        ' 可见区
        Set filteredRng = Nothing
        On Error Resume Next
        Set filteredRng = dataRng.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If filteredRng Is Nothing Then GoTo NextRule

        ' —— 导出筛选后的数据到“数据”工作表 ——
        filteredRng.Copy
        wsDataOut.Range("A1").PasteSpecial xlPasteAll
        Application.CutCopyMode = False

        ' 检测项目：为空则自动（从第7列起）
        If Trim(itemList) = "" Then
            autoList = ""
            For Each c In headerRow.Cells
                If c.Column > 6 And Len(Trim(CStr(c.value))) > 0 Then
                    Dim yAll As Range
                    Set yAll = wsData.Range(wsData.Cells(2, c.Column), wsData.Cells(dataRng.Rows.Count, c.Column))
                    If Application.WorksheetFunction.Count(Intersect(filteredRng, yAll)) > 0 Then
                        autoList = autoList & Trim(CStr(c.value)) & ","
                    End If
                End If
            Next c
            If Right(autoList, 1) = "," Then autoList = Left(autoList, Len(autoList) - 1)
            If Len(autoList) = 0 Then GoTo NextRule
            items = Split(autoList, ",")
        Else
            items = Split(itemList, ",")
        End If

        ' X 轴（日期列）：表头以下 ∩ 可见行
        Dim dateColRange As Range, xVals As Range
        Set dateColRange = wsData.Range(wsData.Cells(2, dateCol), wsData.Cells(dataRng.Rows.Count, dateCol))
        Set xVals = Intersect(filteredRng, dateColRange)
        If Not xVals Is Nothing Then
            On Error Resume Next: Set xVals = xVals.SpecialCells(xlCellTypeVisible): On Error GoTo 0
        End If
        If xVals Is Nothing Then GoTo NextRule

        ' 逐项目作图
        Dim itName As String, col As Variant, yColRange As Range, yRng As Range
        For Each it In items
            itName = Trim(CStr(it))
            col = Application.Match(itName, headerRow, 0)
            If Not IsError(col) Then
                Set yColRange = wsData.Range(wsData.Cells(2, col), wsData.Cells(dataRng.Rows.Count, col))
                Set yRng = Intersect(filteredRng, yColRange)
                If Not yRng Is Nothing Then
                    On Error Resume Next: Set yRng = yRng.SpecialCells(xlCellTypeVisible): On Error GoTo 0
                    If Not yRng Is Nothing Then
                        If Application.WorksheetFunction.Count(yRng) > 0 Then
                            Dim maxVal As Double, minVal As Double, avgVal As Double
                            maxVal = Application.WorksheetFunction.Max(yRng)
                            minVal = Application.WorksheetFunction.Min(yRng)
                            avgVal = Application.WorksheetFunction.Average(yRng)

                            Set chartObj = wsOutput.ChartObjects.Add( _
                                Left:=(outCol - 1) * (chartWidth + 10), _
                                Top:=(outRow - 1) * (chartHeight + 10), _
                                Width:=chartWidth, Height:=chartHeight)

                            With chartObj.Chart
                                .ChartType = xlXYScatterLines
                                .PlotVisibleOnly = True
                                .HasTitle = True
                                .chartTitle.Text = sheetName & " | " & sampleType & " - " & itName & _
                                    vbCr & "Max=" & FormatSmartNumber(maxVal) & _
                                    ", Min=" & FormatSmartNumber(minVal) & _
                                    ", Avg=" & FormatSmartNumber(avgVal)
                                On Error Resume Next
                                .chartTitle.Format.TextFrame2.TextRange.Font.Size = titleFontSize
                                On Error GoTo 0

                                .SeriesCollection.NewSeries
                                With .SeriesCollection(.SeriesCollection.Count)
                                    .Name = itName
                                    .xValues = xVals
                                    .Values = yRng
                                    .MarkerStyle = xlMarkerStyleCircle
                                    .MarkerSize = 5
                                    .Format.Fill.ForeColor.RGB = RGB(255, 140, 0) ' 仅标记橙色
                                    .Format.Line.Weight = 1.25                      ' 线条保持默认色
                                End With

                                ' —— 背景白色 ——
                                With .ChartArea.Format.Fill
                                    .Visible = msoTrue: .Solid
                                    .ForeColor.RGB = vbWhite
                                End With
                                With .PlotArea.Format.Fill
                                    .Visible = msoTrue: .Solid
                                    .ForeColor.RGB = vbWhite
                                End With

                                ' —— 主网格线=自动颜色 ——
                                If .HasAxis(xlValue, xlPrimary) Then
                                    .Axes(xlValue, xlPrimary).HasMajorGridlines = True
                                    On Error Resume Next
                                    .Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = xlColorIndexAutomatic
                                    .Axes(xlValue, xlPrimary).MajorGridlines.Format.Line.Visible = msoTrue
                                    On Error GoTo 0
                                End If

                                ' —— X 轴：跨度≈10段抽稀；方向 Upward/Automatic ——
                                If .HasAxis(xlCategory, xlPrimary) Then
                                    Dim minDate As Double, maxDate As Double, span As Double
                                    minDate = Application.Min(xVals): maxDate = Application.Max(xVals)
                                    span = maxDate - minDate
                                    If span > 0 Then .Axes(xlCategory, xlPrimary).MajorUnit = RoundUpLong(span / 10)
                                    With .Axes(xlCategory, xlPrimary).TickLabels
                                        .NumberFormat = "yyyy-mm-dd"
                                        On Error Resume Next
                                        .Orientation = -4171   ' Upward
                                        If Err.Number <> 0 Then Err.Clear: .Orientation = -4105 ' Automatic
                                        On Error GoTo 0
                                    End With
                                End If

                                ' 去图例
                                On Error Resume Next: .Legend.Delete: On Error GoTo 0
                            End With

                            totalMade = totalMade + 1
                            outCol = outCol + 1
                            If outCol > chartsPerRow Then outCol = 1: outRow = outRow + 1
                        End If
                    End If
                End If
            End If
        Next it

        If wsData.AutoFilterMode Then wsData.AutoFilterMode = False
NextRule:
    Next R

    MsgBox "筛选数据已另存并绘图完成。共生成图表 " & totalMade & " 张。", vbInformation
End Sub


