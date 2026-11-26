# 图表链收敛 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将辅助线能力从 `模块3` 收敛到 `模块6`，并把 `辅助线窗口` 的调用链切到新主线，同时保留旧入口作为过渡壳。

**Architecture:** 保持 `模块6` 作为图表主线，不做大拆分，只新增一个公共辅助线入口和一个私有辅助线实现。`模块3` 不立即删除旧函数，而是把旧辅助线入口改成转调 `模块6`，这样可以在最小改动下收直调用关系并降低回归风险。

**Tech Stack:** VBA, UserForm, xlflow, XlflowAssert

---

### Task 1: 为辅助线迁移建立回归测试

**Files:**
- Create: `src/modules/Test图表链.bas`
- Modify: `xlflow.toml`（仅当需要纳入测试模块路径时）
- Test: `src/modules/Test图表链.bas`

- [ ] **Step 1: 写一个失败测试，验证新主线入口存在且窗体调用目标可替换**

```vb
Attribute VB_Name = "Test图表链"
Option Explicit

Public Sub Test_模块6_提供辅助线公共入口()
    Dim result As Boolean

    On Error Resume Next
    Application.Run "模块6.添加辅助线函数_final", "标准", 1#
    result = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    XlflowAssert.AssertTrue result, "模块6 应提供可调用的辅助线公共入口"
End Sub
```

- [ ] **Step 2: 运行测试，确认它先失败**

Run: `xlflow test --filter Test图表链.Test_模块6_提供辅助线公共入口 --json`

Expected: FAIL，原因是 `模块6.添加辅助线函数_final` 尚不存在，或 `Application.Run` 无法解析该入口。

- [ ] **Step 3: 如果 `Application.Run` 方案不稳定，则退回到更小的存在性测试**

```vb
Public Sub Test_模块6_入口命名已迁移()
    XlflowAssert.AssertTrue True, "先占位，等公共入口落地后替换为真实行为测试"
    XlflowAssert.AssertFail "迁移前不应通过该测试"
End Sub
```

- [ ] **Step 4: 再次运行测试，确认仍然是红灯**

Run: `xlflow test --filter Test图表链 --json`

Expected: FAIL，至少有一个测试明确失败，且失败原因是“新入口尚未实现”。

### Task 2: 在模块6中实现新的辅助线主线入口

**Files:**
- Modify: `src/modules/模块6.bas`
- Test: `src/modules/Test图表链.bas`

- [ ] **Step 1: 在 `模块6.bas` 中加入新的公共入口与私有实现**

```vb
Public Sub 添加辅助线()
    辅助线窗口.Show
End Sub

Public Function 添加辅助线函数_final(ByVal seriesName As String, ByVal yValue As Double) As Boolean
    添加辅助线函数_final = AddHelperSeriesToSelectedChart(seriesName, yValue)
End Function

Private Function AddHelperSeriesToSelectedChart(ByVal seriesName As String, ByVal yValue As Double) As Boolean
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

    If Len(seriesName) = 0 Then seriesName = "标准"

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

    For i = 1 To pointCount
        helperXValues(i) = minX + (maxX - minX) * (i - 1) / (pointCount - 1)
        helperYValues(i) = yValue
    Next i

    With cht.SeriesCollection.NewSeries
        .Name = seriesName
        .XValues = helperXValues
        .Values = helperYValues
        .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .Format.Line.DashStyle = msoLineDash
        .Format.Line.Weight = 1.5
        .MarkerStyle = xlMarkerStyleNone
    End With

    AddHelperSeriesToSelectedChart = True
End Function
```

- [ ] **Step 2: 保持实现尽量与旧逻辑一致，不顺手改行为**

Do not change:
- 辅助线默认名称仍为 `标准`
- 折线样式仍为红色虚线
- 仍以第一个系列的 X 轴范围生成辅助线

- [ ] **Step 3: 运行聚焦测试，确认新入口变绿**

Run: `xlflow test --filter Test图表链.Test_模块6_提供辅助线公共入口 --json`

Expected: PASS 或至少从“入口不存在”转为更接近真实行为的失败；如果仍失败，优先修入口命名或签名。

### Task 3: 把窗体调用切换到模块6

**Files:**
- Modify: `src/forms/code/辅助线窗口.bas`
- Test: `src/modules/Test图表链.bas`

- [ ] **Step 1: 修改窗体按钮事件，只调用模块6的新入口**

```vb
Private Sub btnOK_Click()
    Dim seriesName As String
    Dim yValue As Double

    seriesName = txtSeriesName.Value

    If Not IsNumeric(txtYValue.Value) Then
        MsgBox "输入的 Y 值无效，请输入一个数字！", vbExclamation
        txtYValue.Value = ""
        txtYValue.SetFocus
        Exit Sub
    End If

    yValue = CDbl(txtYValue.Value)

    Call 模块6.添加辅助线函数_final(seriesName, yValue)

    Unload Me
End Sub
```

- [ ] **Step 2: 不修改窗体交互文案和校验行为**

Keep unchanged:
- 数字校验
- 错误提示文本
- `Unload Me` 时机

- [ ] **Step 3: 运行 lint，确保窗体代码引用合法**

Run: `xlflow lint --json`

Expected: PASS，没有新的未声明变量或过程未找到错误。

### Task 4: 把模块3旧入口改成过渡壳

**Files:**
- Modify: `src/modules/模块3.bas`
- Test: `src/modules/Test图表链.bas`

- [ ] **Step 1: 将旧辅助线入口改为转调模块6**

```vb
Function 添加辅助线函数(seriesName As String, yValue As Double)
    添加辅助线函数 = 模块6.添加辅助线函数_final(seriesName, yValue)
End Function

Sub 添加辅助线()
    模块6.添加辅助线
End Sub
```

- [ ] **Step 2: 保留 `添加辅助线0`，但不再把它作为主入口继续演进**

Do not delete yet:
- `添加辅助线0`

Reason:
- 先保证兼容
- 等回归验证完成后，再考虑彻底移除旧壳

- [ ] **Step 3: 运行影响较小的一轮测试，确认过渡层不破坏现有窗体链路**

Run: `xlflow test --filter Test图表链 --json`

Expected: PASS，至少证明：
- 新主线入口存在
- 窗体目标已迁移
- 旧入口仍可转调

### Task 5: 做源码校验与最小验证闭环

**Files:**
- Modify: `src/modules/模块6.bas`
- Modify: `src/forms/code/辅助线窗口.bas`
- Modify: `src/modules/模块3.bas`
- Modify: `src/modules/Test图表链.bas`

- [ ] **Step 1: 跑静态校验**

Run: `xlflow lint --json`

Expected: PASS

- [ ] **Step 2: 跑源码分析**

Run: `xlflow analyze --json`

Expected: PASS，或仅出现已有且可解释的非阻断诊断。

- [ ] **Step 3: 回写到受管工作簿**

Run: `xlflow push --backend file --json`

Expected: PASS，并明确接受 `vbe_validation="not_performed"`（如果是 headless file backend）。

- [ ] **Step 4: 说明剩余验证空缺**

Record explicitly:
- 已完成：源码迁移、静态校验、入口收敛
- 未完成：Excel 交互态下“选中图表后点击窗体添加辅助线”的真实运行验证

- [ ] **Step 5: 如需进一步实机验证，再启用 session 流程**

Run:

```bash
xlflow session start --json
xlflow push --fast --session --no-save --json
xlflow run 模块6.添加辅助线 --session --json
```

Expected:
- 能打开窗体入口
- 若需要人工选中图表，这一步只作为半自动验证，不在无人值守阶段强行宣称通过
