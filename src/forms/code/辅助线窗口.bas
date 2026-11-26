' 用户窗体代码
Private Sub btnCancel_Click()
    ' 取消按钮点击事件
    Unload Me
End Sub

Private Sub btnOK_Click()
    ' 确定按钮点击事件
    Dim seriesName As String
    Dim yValue As Double

    seriesName = txtSeriesName.value

    ' 验证 Y 值输入
    If Not IsNumeric(txtYValue.value) Then
        MsgBox "输入的 Y 值无效，请输入一个数字！", vbExclamation
        txtYValue.value = ""
        txtYValue.SetFocus
        Exit Sub
    End If

    yValue = CDbl(txtYValue.value)

    ' 将值传递给图表主线模块
    Call 模块6.添加辅助线函数_final(seriesName, yValue)

    ' 关闭窗体
    Unload Me
End Sub
