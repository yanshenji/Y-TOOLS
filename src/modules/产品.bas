Attribute VB_Name = "产品"
Function GenerateQualityCertificateFunction(batchNumber As String, 包数 As Double, 净重 As Double, _
    detectionDataWorkbookName As String, detectionDataSheetName As String, outputPath As String, _
    qualityCertTemplatePath As String) As String
    Dim detectionDataWorkbook As Workbook
    Dim qualityCertTemplateWorkbook As Workbook
    Dim detectionDataSheet As Worksheet
    Dim qualityCertSheet As Worksheet
    Dim cell As Range
    Dim foundCell As Range
    Dim qualityCertWorkbook As Workbook
    Dim qualityCertSheetName As String
    Dim templatePlaceholder As String
    Dim dataField As String
    Dim headerRow As Long
    Dim batchColumn As Long
    Dim columnNumber As Long
    Dim savePath As String
    Dim printCertificate As Boolean
    Dim exportAsPDF As Boolean
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    ' 打开检测数据台账
    'On Error Resume Next
    'Set detectionDataWorkbook = Workbooks(detectionDataWorkbookName)
    'If detectionDataWorkbook Is Nothing Then
    '    Set detectionDataWorkbook = Workbooks.Open(Environ("USERPROFILE") & "\Documents\" & detectionDataWorkbookName)
    'End If
    
    'If detectionDataWorkbook Is Nothing Then
    '    result = "无法打开检测数据台账!"
    '    GenerateQualityCertificateFunction = result
    '    Exit Function
    'End If
    Set detectionDataWorkbook = Workbooks(detectionDataWorkbookName)
    Set detectionDataSheet = detectionDataWorkbook.Sheets(detectionDataSheetName)
    
    ' 打开质量证书模板
    On Error Resume Next
    Set qualityCertTemplateWorkbook = Workbooks.Open(qualityCertTemplatePath)
    
    If qualityCertTemplateWorkbook Is Nothing Then
        result = "无法打开质量证书模板!"
        detectionDataWorkbook.Close SaveChanges:=False
        GenerateQualityCertificateFunction = result
        Exit Function
    End If
    
    ' 查找表头行（第一列包含“日期”的行）
    With detectionDataSheet.Columns(1)
        Set cell = .Find(What:="日期", LookIn:=xlValues, LookAt:=xlPart)
        If Not cell Is Nothing Then
            headerRow = cell.Row
        Else
            result = "未找到检测数据台账中的表头行!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificateFunction = result
            Exit Function
        End If
    End With
    
    ' 查找批号列（表头行中包含“批号”的列）
    With detectionDataSheet.Rows(headerRow)
        Set cell = .Find(What:="批号", LookIn:=xlValues, LookAt:=xlPart)
        If Not cell Is Nothing Then
            batchColumn = cell.Column
        Else
            result = "未找到检测数据台账中的批号列!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificateFunction = result
            Exit Function
        End If
    End With
    
    ' 根据批号查找检测数据
    With detectionDataSheet
        Set foundCell = .Columns(batchColumn).Find(What:=batchNumber, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, MatchCase:=False)
        If foundCell Is Nothing Then
            result = "未找到批号为 " & batchNumber & " 的检测数据!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificateFunction = result
            Exit Function
        End If
    End With
    
    ' 复制模板并创建新工作簿
    qualityCertTemplateWorkbook.Sheets(1).Copy
    Set qualityCertWorkbook = ActiveWorkbook
    Set qualityCertSheet = qualityCertWorkbook.Sheets(1)
    
    ' 填充质量证书数据（根据模板中的占位符来填充）
    With qualityCertSheet
        ' 替换模板中的批号占位符（假设模板中批号占位符为 {{BATCH_NUMBER}}）
        .Cells.Replace What:="{{批号}}", Replacement:=batchNumber, LookAt:=xlPart
        
        ' 替换模板中的重量占位符（假设模板中重量占位符为 {{WEIGHT}}）
        .Cells.Replace What:="{{包数}}", Replacement:=包数, LookAt:=xlPart
        ' 替换模板中的重量占位符（假设模板中重量占位符为 {{WEIGHT}}）
        .Cells.Replace What:="{{净重}}", Replacement:=净重, LookAt:=xlPart
        
        ' 遍历表头中的每个字段，查找模板中的占位符并替换
        For Each cell In detectionDataSheet.Rows(headerRow).Cells
            If cell.value <> "" Then
                dataField = cell.value
                templatePlaceholder = "{{" & UCase(dataField) & "}}"
                
                If Len(templatePlaceholder) > 2 Then
                    columnNumber = cell.Column
                    .Cells.Replace What:=templatePlaceholder, Replacement:=detectionDataSheet.Cells(foundCell.Row, columnNumber).value, LookAt:=xlPart
                End If
            End If
        Next cell
    End With
    
    ' 确定保存路径
    If Not Dir(outputPath, vbDirectory) <> "" Then
        MkDir outputPath
    End If
    
    ' 设置证书保存名称和路径
    qualityCertSheetName = "质量证明书 - " & batchNumber & 包数 & " 包 " & 净重 & "kg"
    qualityCertWorkbook.Sheets(1).Name = qualityCertSheetName
    savePath = outputPath & qualityCertSheetName & ".xlsx"
    
    ' 保存质量证书工作簿
    qualityCertWorkbook.SaveAs savePath
    
    ' 询问用户是否打印或导出为PDF
    'Dim userChoice As VbMsgBoxResult
    'userChoice = MsgBox("证书已生成，您希望如何操作？", vbQuestion + vbYesNoCancel + vbDefaultButton1, "操作选择")
    
    'printCertificate = (userChoice = vbYes)
    'exportAsPDF = (userChoice = vbNo)
    
    'If printCertificate Then
        ' 打印质量证书
    '    qualityCertWorkbook.PrintOutCopies = 1
    '    qualityCertWorkbook.PrintOut Collate:=True
   ' ElseIf exportAsPDF Then
        ' 导出为PDF
        Dim pdfPath As String
        pdfPath = outputPath & qualityCertSheetName & ".pdf"
        qualityCertSheet.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    'End If
    
    ' 设置函数返回值为证书保存路径
    result = "质量证书已生成并保存在：" & savePath
    GenerateQualityCertificateFunction = result
    
    ' 关闭工作簿
    qualityCertWorkbook.Close SaveChanges:=False
    'detectionDataWorkbook.Close SaveChanges:=False
    qualityCertTemplateWorkbook.Close SaveChanges:=False
    
    Exit Function
ErrorHandler:
    result = "发生错误：" & Err.Description
    GenerateQualityCertificateFunction = result
End Function


Sub TestGenerateQualityCertificateFunction()
    Dim result As String
    result = GenerateQualityCertificateFunction( _
        "TSG-2025062106", _
        12, _
        6000, _
        "01 2025年主产品及铁渣检测台账.xlsx", _
        "碳酸钴", _
        "C:\质量证书\碳酸钴\", _
        "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\碳酸钴产品质量证明书模板.xlsx" _
    )
    MsgBox result
End Sub

Function GenerateQualityCertificate(batchNumber As String, metalQuantity As Double, PackageWeight As Double) As String
    Dim detectionDataWorkbook As Workbook
    Dim qualityCertTemplateWorkbook As Workbook
    Dim detectionDataSheet As Worksheet
    Dim qualityCertSheet As Worksheet
    Dim cell As Range
    Dim foundCell As Range
    Dim qualityCertWorkbook As Workbook
    Dim qualityCertSheetName As String
    Dim templatePlaceholder As String
    Dim dataField As String
    Dim headerRow As Long
    Dim batchColumn As Long
    Dim columnNumber As Long
    Dim savePath As String
    Dim pdfPath As String
    Dim result As String
    Dim packageCount As Double
    Dim netWeight As Double
    Dim productType As String
    Dim qualityCertTemplatePath As String
    
    On Error GoTo ErrorHandler
    
    ' 根据批号前缀识别产品类型
    productType = GetProductType(batchNumber)
    
    ' 如果用户未输入每包重量，则使用默认值
    If PackageWeight = 0 Then
        PackageWeight = GetDefaultPackageWeight(productType)
    End If
    
    ' 计算净重和包数
    netWeight = metalQuantity / GetDefaultMetalContent(productType)
    packageCount = netWeight / PackageWeight
    
    workbookName = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\01 2025年主产品及铁渣检测台账.xlsx"
    ' 打开检测数据工作簿
    On Error Resume Next
    Set detectionDataWorkbook = Workbooks(workbookName)
    
    If detectionDataWorkbook Is Nothing Then
        ' 如果工作簿未打开，则尝试打开
        MsgBox "未打开工作簿"
    End If
    
    On Error GoTo ErrorHandler
    Set detectionDataSheet = detectionDataWorkbook.Sheets(GetDefaultSheetName(productType)) ' 假设工作表名为“检测数据”
    
    ' 打开质量证书模板
    On Error Resume Next
    qualityCertTemplatePath = GetTemplatePath(productType)
    Set qualityCertTemplateWorkbook = Workbooks.Open(qualityCertTemplatePath)
    
    If qualityCertTemplateWorkbook Is Nothing Then
        result = "无法打开质量证书模板!"
        detectionDataWorkbook.Close SaveChanges:=False
        GenerateQualityCertificate = result
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' 查找表头行（第一列包含“日期”的行）
    With detectionDataSheet.Columns(1)
        Set cell = .Find(What:="日期", LookIn:=xlValues, LookAt:=xlPart)
        If cell Is Nothing Then
            result = "未找到检测数据工作簿中的表头行!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificate = result
            Exit Function
        End If
        headerRow = cell.Row
    End With
    
    ' 查找批号列（表头行中包含“批号”的列）
    With detectionDataSheet.Rows(headerRow)
        Set cell = .Find(What:="批号", LookIn:=xlValues, LookAt:=xlPart)
        If cell Is Nothing Then
            result = "未找到检测数据工作簿中的批号列!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificate = result
            Exit Function
        End If
        batchColumn = cell.Column
    End With
    
    ' 根据批号查找检测数据
    With detectionDataSheet
        Set foundCell = .Columns(batchColumn).Find(What:=batchNumber, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, MatchCase:=False)
        If foundCell Is Nothing Then
            result = "未找到批号为 " & batchNumber & " 的检测数据!"
            detectionDataWorkbook.Close SaveChanges:=False
            qualityCertTemplateWorkbook.Close SaveChanges:=False
            GenerateQualityCertificate = result
            Exit Function
        End If
    End With
    
    ' 复制模板并创建新工作簿
    qualityCertTemplateWorkbook.Sheets(1).Copy
    Set qualityCertWorkbook = ActiveWorkbook
    Set qualityCertSheet = qualityCertWorkbook.Sheets(1)
    
    ' 填充质量证书数据（根据模板中的占位符来填充）
    With qualityCertSheet
        ' 替换模板中的批号占位符（假设模板中批号占位符为 {{批号}}）
        .Cells.Replace What:="{{批号}}", Replacement:=batchNumber, LookAt:=xlPart
        
        ' 替换模板中的包数占位符（假设模板中包数占位符为 {{包数}}）
        .Cells.Replace What:="{{包数}}", Replacement:=packageCount, LookAt:=xlPart
        
        ' 替换模板中的净重占位符（假设模板中净重占位符为 {{净重}}）
        .Cells.Replace What:="{{净重}}", Replacement:=netWeight, LookAt:=xlPart
        
        ' 遍历表头中的每个字段，查找模板中的占位符并替换
        For Each cell In detectionDataSheet.Rows(headerRow).Cells
            If cell.value <> "" Then
                dataField = cell.value
                templatePlaceholder = "{{" & UCase(dataField) & "}}"
                
                If Len(templatePlaceholder) > 2 Then
                    columnNumber = cell.Column
                    .Cells.Replace What:=templatePlaceholder, Replacement:=detectionDataSheet.Cells(foundCell.Row, columnNumber).value, LookAt:=xlPart
                End If
            End If
        Next cell
    End With
    
    ' 确定保存路径
    Dim outputPath As String
    outputPath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\"
    If Not Dir(outputPath, vbDirectory) <> "" Then
        MkDir outputPath
    End If
    
    ' 设置证书保存名称和路径
    qualityCertSheetName = "COA - " & batchNumber & " - " & packageCount & " 包 - " & netWeight & "kg"
    'qualityCertWorkbook.Sheets(1).Name = qualityCertSheetName
    savePath = outputPath & qualityCertSheetName & ".xlsx"
    pdfPath = outputPath & qualityCertSheetName & ".pdf"
    
    ' 保存质量证书工作簿
    qualityCertWorkbook.SaveAs savePath
    
    ' 导出为PDF
    qualityCertSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    
    ' 设置函数返回值为证书保存路径
    result = "质量证书已生成并保存在：" & vbCrLf & _
             "Excel 文件： " & savePath & vbCrLf & _
             "PDF 文件： " & pdfPath
    GenerateQualityCertificate = result
    
    ' 关闭工作簿
    qualityCertWorkbook.Close SaveChanges:=False
    'detectionDataWorkbook.Close SaveChanges:=False
    qualityCertTemplateWorkbook.Close SaveChanges:=False
    
    Exit Function
ErrorHandler:
    result = "发生错误：" & Err.Description
    GenerateQualityCertificate = result
    
    ' 确保关闭已打开的工作簿
    If Not qualityCertWorkbook Is Nothing Then
        qualityCertWorkbook.Close SaveChanges:=False
    End If
    'If Not detectionDataWorkbook Is Nothing Then
    '    detectionDataWorkbook.Close SaveChanges:=False
   ' End If
    If Not qualityCertTemplateWorkbook Is Nothing Then
        qualityCertTemplateWorkbook.Close SaveChanges:=False
    End If
End Function

Function GetProductType(batchNumber As String) As String
    ' 根据批号前缀识别产品类型
    Dim productPrefixes As Object
    Set productPrefixes = CreateObject("Scripting.Dictionary")
    
    productPrefixes.Add "LHG", "氯化钴"
    productPrefixes.Add "TSG", "碳酸钴"
    productPrefixes.Add "YHG", "氧化钴"
    productPrefixes.Add "ALLOY", "2108电钴"
    productPrefixes.Add "STANDARD", "109电钴"
    productPrefixes.Add "HR06A0", "四钴"
    
    Dim productPrefix As Variant
    For Each productPrefix In productPrefixes.Keys
        If Left(batchNumber, Len(productPrefix)) = productPrefix Then
            GetProductType = productPrefixes(productPrefix)
            Exit Function
        End If
    Next productPrefix
    
    GetProductType = "未知产品"
End Function

Function GetDefaultMetalContent(productType As String) As Double
    ' 设置每种产品的默认金属含量
    Select Case productType
        Case "氯化钴"
            GetDefaultMetalContent = 0.242
        Case "碳酸钴"
            GetDefaultMetalContent = 0.465
        Case "氧化钴"
            GetDefaultMetalContent = 0.715
        Case "2108电钴"
            GetDefaultMetalContent = 1
        Case "109电钴"
            GetDefaultMetalContent = 1
        Case "四钴"
            GetDefaultMetalContent = 0.725
        Case Else
            GetDefaultMetalContent = 0
    End Select
End Function

Function GetDefaultPackageWeight(productType As String) As Double
    ' 设置每种产品的默认每包重量
    Select Case productType
        Case "氯化钴"
            GetDefaultPackageWeight = 1000
        Case "碳酸钴"
            GetDefaultPackageWeight = 500
        Case "氧化钴"
            GetDefaultPackageWeight = 500
        Case "2108电钴"
            GetDefaultPackageWeight = 1000
        Case "109电钴"
            GetDefaultPackageWeight = 1000
        Case "四钴"
            GetDefaultPackageWeight = 1000
        Case Else
            GetDefaultPackageWeight = 0
    End Select
End Function

Function GetTemplatePath(productType As String) As String
    ' 设置每种产品的默认模板路径
    Select Case productType
        Case "氯化钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\氯化钴产品质量证明书模板.xlsx"
        Case "碳酸钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\碳酸钴产品质量证明书模板.xlsx"
        Case "氧化钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\氧化钴产品质量证明书模板.xlsx"
        Case "2108电钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\2108电钴钴产品质量证明书模板.xlsx"
        Case "109电钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\109电钴产品质量证明书模板.xlsx"
        Case "四钴"
            GetTemplatePath = "\\10.34.100.50\质量部\赣州寒锐检测数据（每日更新）\产品\产品质量证明书\产品质量证明书模板\四钴产品质量证明书模板.xlsx"
        Case Else
            GetTemplatePath = ""
    End Select
End Function

Function GetDefaultSheetName(productType As String) As String
    ' 设置每种产品的默认模板路径
    Select Case productType
        Case "氯化钴"
            GetDefaultSheetName = "氯化钴"
        Case "碳酸钴"
            GetDefaultSheetName = "碳酸钴"
        Case "氧化钴"
            GetDefaultSheetName = "氧化钴"
        Case "2108电钴"
            GetDefaultSheetName = "（2108）电钴检测记录表 "
        Case "109电钴"
            GetDefaultSheetName = "（109）电钴检测记录表"
        Case "四钴"
            GetDefaultSheetName = "四氧化三钴"
        Case Else
            GetDefaultSheetName = ""
    End Select
End Function


Sub test1323()
    'MsgBox GetDefaultSheetName("氯化钴")
 MsgBox GenerateQualityCertificate("LHG-2025062201", 2420, 1000)
End Sub

