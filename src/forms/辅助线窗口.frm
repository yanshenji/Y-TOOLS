VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 辅助线窗口 
   Caption         =   "辅助线窗口"
   ClientHeight    =   2295
   ClientLeft      =   42
   ClientTop       =   392
   ClientWidth     =   3976
   OleObjectBlob   =   "辅助线窗口.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "辅助线窗口"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

    ' 将值传递给添加辅助线的子程序
    Call 添加辅助线函数(seriesName, yValue)

    ' 关闭窗体
    Unload Me
End Sub

