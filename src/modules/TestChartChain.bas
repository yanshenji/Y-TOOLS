Attribute VB_Name = "TestChartChain"
Option Explicit

Private Const TEST_SHEET_NAME As String = "xlflow_TestChartChain"

Public Sub Test_Module6_ProvidesHelperLineEntryPoint()
    Dim testWs As Worksheet
    Dim chartObj As ChartObject
    Dim result As Variant
    Dim helperValues As Variant
    Dim i As Long

    On Error GoTo CleanFail

    Set testWs = CreateChartTestSheet()
    Set chartObj = CreateLineChart(testWs)
    chartObj.Activate
    chartObj.Chart.ChartArea.Select

    result = Application.Run("ChartInterop.AddHelperLineFinalProxy", "标准", 12#, False)

    XlflowAssert.AssertTrue VarType(result) = vbBoolean, "新入口应返回 Boolean 结果"
    XlflowAssert.AssertTrue CBool(result), "新入口应返回 True"
    XlflowAssert.AssertTrue chartObj.Chart.SeriesCollection.Count = 2, "应新增一条辅助线系列"
    XlflowAssert.AssertTrue InStr(1, CStr(chartObj.Chart.SeriesCollection(2).Name), "标准", vbTextCompare) > 0, "辅助线系列名称应包含标准"

    helperValues = chartObj.Chart.SeriesCollection(2).Values
    For i = LBound(helperValues) To UBound(helperValues)
        XlflowAssert.AssertTrue CDbl(helperValues(i)) = 12#, "辅助线各点的 Y 值应保持一致"
    Next i

CleanExit:
    DeleteSheetIfExists TEST_SHEET_NAME
    Exit Sub

CleanFail:
    DeleteSheetIfExists TEST_SHEET_NAME
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function CreateChartTestSheet() As Worksheet
    Dim ws As Worksheet

    DeleteSheetIfExists TEST_SHEET_NAME
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = TEST_SHEET_NAME

    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "指标"
    ws.Range("A2").Value = DateSerial(2026, 1, 1)
    ws.Range("A3").Value = DateSerial(2026, 1, 2)
    ws.Range("A4").Value = DateSerial(2026, 1, 3)
    ws.Range("B2").Value = 10
    ws.Range("B3").Value = 11
    ws.Range("B4").Value = 13
    ws.Range("A2:A4").NumberFormat = "m/d/yyyy"

    Set CreateChartTestSheet = ws
End Function

Private Function CreateLineChart(ByVal ws As Worksheet) As ChartObject
    Dim chartObj As ChartObject

    Set chartObj = ws.ChartObjects.Add(Left:=120, Width:=320, Top:=40, Height:=180)
    chartObj.Chart.ChartType = xlLineMarkers

    With chartObj.Chart.SeriesCollection.NewSeries
        .Name = "=""指标"""
        .XValues = ws.Range("A2:A4")
        .Values = ws.Range("B2:B4")
    End With

    Set CreateLineChart = chartObj
End Function

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
