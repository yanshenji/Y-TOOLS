VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 产品质量证明书 
   Caption         =   "UserForm1"
   ClientHeight    =   3540
   ClientLeft      =   42
   ClientTop       =   392
   ClientWidth     =   4725
   OleObjectBlob   =   "产品质量证明书.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "产品质量证明书"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 在 VBA 编辑器中新建一个用户表单（UserForm1），并添加以下控件：
' - 3 个标签（Label）：批号、金属量、每包重量
' - 3 个文本框（TextBox）：批号（BatchNumberTextBox）、金属量（MetalQuantityTextBox）、每包重量（PackageWeightTextBox）
' - 2 个按钮（CommandButton）：取消（CancelButton）、确定（OKButton）

' UserForm1 的代码
Private Sub UserForm_Initialize()
    ' 初始化下拉框选项
    With Me.PackageWeightComboBox
        .AddItem "500"
        .AddItem "750"
        .AddItem "1000"
        .value = "1000" ' 默认选择500
    End With
End Sub

Private Sub OKButton_Click()
    Dim batchNumber As String
    Dim metalQuantity As Double
    Dim PackageWeight As Double
    
    ' 获取用户输入
    batchNumber = Me.BatchNumberTextBox.value
    metalQuantity = CDbl(Me.MetalQuantityTextBox.value)
    PackageWeight = CDbl(Me.PackageWeightComboBox.value)
    
    ' 调用函数生成质量证书
    Dim result As String
    result = GenerateQualityCertificate(batchNumber, metalQuantity, PackageWeight)
    
    ' 显示结果
    MsgBox result
    
    ' 关闭表单
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    ' 关闭表单
    Me.Hide
End Sub



