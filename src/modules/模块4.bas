Attribute VB_Name = "模块4"
' 定义全局变量
Dim currentCategoryIndex As Integer
Dim uniqueCategories As Collection
Dim categoryColumn As Long
Dim isInitialized As Boolean
Dim currentSheet As Worksheet

Sub 初始化筛选器()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim categories As Range
    Dim cell As Range

    Set currentSheet = ActiveSheet ' 记录当前工作表
    currentCategoryIndex = 1 ' 初始化索引
    Set uniqueCategories = New Collection
    categoryColumn = 0
    isInitialized = False ' 初始化状态为未初始化

    Set ws = currentSheet

    ' 查找“样品类别”列
    For i = 1 To ws.Rows(1).Cells.Count
        If ws.Cells(1, i).value = "样品类别" Then
            categoryColumn = i
            Exit For
        End If
    Next i

    If categoryColumn = 0 Then
        MsgBox "未找到[样品类别]列！", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, categoryColumn).End(xlUp).Row ' 获取“样品类别”列最后一行的行号
    Set categories = ws.Range(ws.Cells(2, categoryColumn), ws.Cells(lastRow, categoryColumn)) ' 获取“样品类别”列的数据范围

    On Error Resume Next ' 忽略重复值错误
    For Each cell In categories
        If Len(cell.value) > 0 Then
            uniqueCategories.Add cell.value, CStr(cell.value)
        End If
    Next cell
    On Error GoTo 0 ' 恢复错误处理

    If uniqueCategories.Count = 0 Then
        MsgBox "[样品类别]列中没有数据！", vbExclamation
        Exit Sub
    End If

    isInitialized = True ' 已完成初始化
    Set currentSheet = ws
    MsgBox "已找到[样品类别]列，并初始化筛选器！", vbInformation
End Sub

Sub 下一个样品类别()
    Dim ws As Worksheet

    Set ws = ActiveSheet ' 获取当前活动工作表

      ' 判断是否需要初始化
    If Not isInitialized Then
        初始化筛选器 ' 进行初始化
    Else
        ' 检查 currentSheet 是否为 Nothing，如果是，则重新初始化
        If currentSheet Is Nothing Then
            初始化筛选器
        Else
            ' 检查当前活动工作表是否与记录的工作表一致
            If ws.Name <> currentSheet.Name Then
                初始化筛选器
            End If
        End If
    End If

    If categoryColumn = 0 Then
        MsgBox "未找到[样品类别]列！", vbExclamation
        Exit Sub
    End If

    If uniqueCategories Is Nothing Then
        MsgBox "未获取到[样品类别]唯一值！", vbExclamation
        Exit Sub
    End If

    ' 更新索引并处理循环
    currentCategoryIndex = currentCategoryIndex + 1
    If currentCategoryIndex > uniqueCategories.Count Then
        currentCategoryIndex = 1 ' 循环回到第一个类别
    End If

    ' 应用筛选器
    ActiveSheet.Range("$A:$ZZ").AutoFilter Field:=categoryColumn, Criteria1:=uniqueCategories(currentCategoryIndex)

    MsgBox "已筛选到类别：" & uniqueCategories(currentCategoryIndex), vbInformation
End Sub
