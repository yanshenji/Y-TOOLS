# 工具宏：VBA → Excel-DNA 迁移清单

> 生成日期：2026-11-26 ｜ 基线提交：`98968c9` ｜ 范围：`src/modules`、`src/forms`、`src/workbook`

## 1. 总览

| 项目 | 数量 |
|---|---:|
| 标准模块 | 15 |
| 窗体 | 2（`产品质量证明书`、`辅助线窗口`） |
| 文档模块 | 2（`ThisWorkbook`、`Sheet2`，均只有 `Option Explicit`，无事件代码） |
| 代码行数 | 约 5,700 |
| 过程总数 | 119 |

按迁移目标统计：

| 迁移目标 | 数量 | 说明 |
|---|---:|---|
| UDF（`[ExcelFunction]`） | 11 | 工作表函数，纯计算，最先迁 |
| 命令（`[ExcelCommand]` + Ribbon 按钮） | 65 | 原来的菜单/按钮宏 |
| 内部方法 | 21 | Private 辅助函数、只在代码内部调用的 Public 函数 |
| WinForms | 5 | 两个窗体的事件代码 |
| 删除 | 9 | 空过程、演示代码、过渡壳 |
| 待确认 | 4 | 和别的过程功能重复，需要你决定保留哪个 |
| 改写为测试 | 4 | `TestChartChain` 模块 |

## 2. 迁移批次

| 批次 | 内容 | 过程数 | 能否在 Linux 容器验证 |
|---|---|---:|---|
| **B0** | 工程骨架：`.csproj`、`.dna`、Ribbon XML、配置文件、日志、COM 访问封装 | — | 能编译；加载要在 Windows 上做 |
| **B1** | 统计/校验 UDF：`函数`、`模块7`、`模块8`、`其他` 中的 Grubbs 与有效数字函数 | 11 UDF + 4 内部方法 | **能**：核心算法抽成纯 C#，用 xUnit 测 |
| **B2** | 简单命令：`Openfiles`（打开共享盘文件）、`模块1`（筛选/隐藏）、`模块4`（样品类别筛选器）、`其他`（PDF 导出、复制可见单元格、转窄表、条件格式清理） | 53 | 只能编译；行为要在 Excel 里试 |
| **B3** | 业务流程：`模块2`（五种异常检测与 MK 趋势）、`产品`（质量证明书生成）+ `产品质量证明书` 窗体 | 7 命令 + 6 内部方法 + 3 窗体事件 + 1 待确认 | 算法部分能测；Excel 交互部分不能 |
| **B4** | 图表链：`模块3`、`模块5`、`模块6` + `辅助线窗口` | 5 命令 + 11 内部方法 + 2 窗体事件 + 3 待确认 | 只能编译 |

建议每批迁完先在 Windows 上和 `.xlam` 对拍一遍，确认结果一致再进入下一批。

## 3. 映射规则

| VBA | Excel-DNA / C# | 说明 |
|---|---|---|
| `Public Function`（工作表调用） | `[ExcelFunction(Name = "原函数名")]` 静态方法 | 保留原函数名，已有工作簿里的公式不用改。中文函数名（如 `保留两位有效数字_四舍五入`）需要在 B1 先实测能否注册 |
| `Range` 参数（UDF） | `object[,]`（取值）或 `[ExcelArgument(AllowReference = true)]` | 只读值的函数一律用 `object[,]`，不在 UDF 里碰 COM |
| `WorksheetFunction.Average/StDev_S/Min/Median` | LINQ / 自写 | 计算逻辑完全脱离 Excel，可在 Linux 上测 |
| `WorksheetFunction.T_Inv_2T / Norm_S_Dist` | MathNet.Numerics（`StudentT.InvCDF`、`Normal.CDF`） | 需要和 Excel 结果对拍到 1e-9 |
| `Enum` 参数（`模块8.MK_RESULT_TYPE`） | `int` 参数 + C# 内部枚举 | 工作表只能传数字 |
| `Public Sub`（无参数） | `[ExcelCommand]` + Ribbon 按钮 | 用 `ExcelDnaUtil.Application`（`dynamic` 或 Interop）操作 `Selection`、`ActiveSheet` |
| `MsgBox / InputBox` | `MessageBox` / 自定义对话框；`Application.InputBox(Type:=8)` 仍通过 COM 调用 | 选区输入保留 Excel 原生体验 |
| `UserForm` | WinForms 窗体 | 两个窗体都只有 3–5 个控件，直接重画 |
| `Scripting.Dictionary`、`Collection` | `Dictionary<,>`、`List<>` | |
| `Declare PtrSafe ... user32/kernel32`（`其他.bas` 剪贴板 API） | `System.Windows.Forms.Clipboard` | 声明了但**没有被调用**，实际用的是 `DataObject.PutInClipboard` |
| 硬编码路径 `\\10.34.100.50\质量部\...`、`C:\质量证书\...` | `工具宏.config.json`（随 `.xll` 部署） | 共 30+ 处，集中到配置后，换年份或换路径不用改代码 |
| `On Error Resume Next` / `GoTo` | `try/catch` + 统一错误提示 | 命令入口统一包一层：异常时弹框并写日志，不让 Excel 崩 |

## 4. 需要你确认的问题

1. **重复功能保留哪个**
   - `产品.GenerateQualityCertificateFunction`（170 行，无调用方） vs `产品.GenerateQualityCertificate`（窗体在用）→ 建议只迁后者。
   - `模块3.画标记折线图` / `画标记折线图_0` vs `模块6.画标记折线图_final` → 建议只迁 `_final`。
   - `模块3.添加辅助线0` vs `模块6` 辅助线主线 → 建议只迁 `模块6`。
   - 也就是说，原计划 `2026-10-23-chart-chain-consolidation.md` 在 VBA 里已经基本做完（模块3 现在只是转调模块6），迁移时直接以模块6 为准。
2. **`Cpk_Rank` 的返回文字**：源码里是 `"D级?严重不足"`，`?` 应该是编码转换时丢的字（原来可能是 `—` 或 `：`）。现有工作簿里有没有公式依赖这个字符串？
3. **有效数字函数**：`保留两位有效数字_四舍五入` 用的是 VBA `Round`，它其实是银行家舍入（2.25→2.2），和函数名说的“四舍五入”不一致。C# 版是保持原行为，还是修成真正的四舍五入？
4. **`Openfiles` 里的年份文件**（2023–2026 共 24 个“打开 XX 数据”）：要原样做成 24 个按钮，还是改成“按年份 + 数据类型”的下拉菜单，每年只改配置？
5. **Excel 版本与位数**：Excel 是 32 位还是 64 位？决定打 `-AddIn.xll` 还是 `-AddIn64.xll`，也可以两个都打。
6. **过渡期并存**：`.xlam` 和 `.xll` 同时加载时，同名 UDF 会冲突。建议每批迁完，就把 `.xlam` 里对应的函数删掉或改名。

## 5. 源码里顺带发现的问题（迁移时修）

- `Openfiles.打开中控检测结果_分类`：`subFile1` 被 `Open` 了两次，`subFile3` 和 `subFile1` 是同一个文件（101-105）。
- `其他.DeleteConditionalFormatting` 的目录是占位符 `C:\YourFolder\`。
- `其他.bas` 顶部 8 个 Win32 `Declare` 没有被使用。
- `日报` 模块的两个过程是空的。
- 大量过程依赖 `Selection` / `ActiveSheet`，C# 里每个命令入口都要先检查选区类型，否则选中图表时会抛 COM 异常。

## 6. 逐过程清单

依赖列含义：交互 = MsgBox/InputBox/窗体；文件 = 打开、保存、导出；图表 = ChartObject/Series；选区 = Selection/ActiveSheet；格式 = 填色/字体/条件格式；跨模块 = Application.Run/Call。行数含声明行。

### 产品质量证明书

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `UserForm_Initialize` | Sub | 9 | — | WinForms | B3 |  |
| `OKButton_Click` | Sub | 20 | 交互 | WinForms | B3 |  |
| `CancelButton_Click` | Sub | 4 | — | WinForms | B3 |  |

### 辅助线窗口

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `btnCancel_Click` | Sub | 4 | — | WinForms | B4 |  |
| `btnOK_Click` | Sub | 23 | 交互、跨模块 | WinForms | B4 |  |

### ChartInterop

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `AddHelperLineFinalProxy` | Function | 3 | — | 删除 | — | VBA Application.Run 互操作壳 |

### Openfiles

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `打开2026中控质量报表` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025中控质量报表` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开自动判定` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025外拉氯化锰溶液数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2026中控数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025中控数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2024中控数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2026成品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025成品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2024成品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2023成品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2026辅料数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025辅料数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2026副产品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025副产品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2024副产品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2023副产品数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2026铁渣数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025铁渣数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开2025外购外拉数据` | Sub | 7 | 文件 | 命令 | B2 |  |
| `打开中控检测结果_分类` | Sub | 27 | 文件 | 命令 | B2 |  |
| `保存并关闭中控检测结果_分类` | Sub | 48 | 交互、文件 | 命令 | B2 |  |
| `打开1期中控质量报表` | Sub | 8 | 文件 | 命令 | B2 |  |
| `打开赣州寒锐质量统计报表_取数` | Sub | 7 | 文件 | 命令 | B2 |  |

### TestChartChain

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `Test_Module6_ProvidesHelperLineEntryPoint` | Sub | 34 | 图表、跨模块 | 改写测试 | — | 改为 C# 集成测试（需 Windows+Excel） |
| `CreateChartTestSheet` | Function | 19 | 选区、格式 | 改写测试 | — |  |
| `CreateLineChart` | Function | 14 | 图表 | 改写测试 | — |  |
| `DeleteSheetIfExists` | Sub | 10 | 选区 | 改写测试 | — |  |

### 产品

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `GenerateQualityCertificateFunction` | Function | 170 | 交互、文件、选区 | 待确认 | B3 | 与 GenerateQualityCertificate 功能重叠，无调用方；确认是否废弃 |
| `TestGenerateQualityCertificateFunction` | Sub | 13 | 交互 | 删除 | — | 手工测试 |
| `GenerateQualityCertificate` | Function | 184 | 交互、文件、选区 | 内部方法 | B3 | 由窗体调用；路径改配置 |
| `GetProductType` | Function | 22 | 字典 | 内部方法 | B3 | 批号前缀 → 产品类型映射，可单测 |
| `GetDefaultMetalContent` | Function | 19 | — | 内部方法 | B3 |  |
| `GetDefaultPackageWeight` | Function | 19 | — | 内部方法 | B3 |  |
| `GetTemplatePath` | Function | 19 | — | 内部方法 | B3 | UNC 模板路径 → 配置 |
| `GetDefaultSheetName` | Function | 19 | — | 内部方法 | B3 |  |
| `test1323` | Sub | 4 | 交互 | 删除 | — | 手工测试 |

### 其他

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `DeleteConditionalFormatting` | Sub | 38 | 交互、文件、格式 | 命令 | B2 |  |
| `CurrentExcelToPDF` | Sub | 24 | 交互、文件、选区 | 命令 | B2 |  |
| `ConvertExcelToPDF` | Sub | 50 | 交互、文件 | 命令 | B2 |  |
| `ExportCurrentSheetToPDF` | Sub | 40 | 交互、文件、选区 | 命令 | B2 |  |
| `连接可见单元格并复制` | Sub | 36 | 交互、选区 | 命令 | B2 |  |
| `ImprovedChemicalFormatter` | Sub | 82 | 交互、选区、格式 | 命令 | B2 |  |
| `隐藏所有工作表CE列` | Sub | 15 | 交互、选区 | 命令 | B2 |  |
| `统计可见单元格数值并复制` | Sub | 84 | 交互、选区、字典 | 命令 | B2 |  |
| `保留两位有效数字_四舍六入五成双` | Function | 42 | — | UDF | B1 | 原实现用 CStr 取位，存在浮点/科学计数法风险，迁移时用 decimal 重写并对拍 |
| `保留两位有效数字_四舍五入` | Function | 21 | — | UDF | B1 | VBA Round 是银行家舍入，需确认期望行为 |
| `检测数据转窄表` | Sub | 115 | 交互、选区 | 命令 | B2 |  |
| `GrubbsOutliers` | Function | 81 | — | UDF | B1 | 返回数组 → 动态数组溢出 |
| `GrubbsG` | Function | 11 | — | UDF | B1 |  |
| `GrubbsCrit` | Function | 19 | — | UDF | B1 | T_Inv_2T → MathNet.Numerics StudentT |
| `RemoveAt` | Function | 13 | — | 内部方法 | B1 |  |
| `DemoGrubbs` | Sub | 30 | 交互、选区 | 删除 | — | 演示代码 |
| `FlagOutliersInSheet` | Sub | 51 | 交互、选区 | 命令 | B2 |  |
| `GrubbsMarkYellow` | Sub | 56 | 交互、选区、格式 | 命令 | B2 |  |

### 函数

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `cmpStd222` | Sub | 164 | 交互、格式 | 命令 | B2 | 带参数 Sub，无法从菜单直接调用；改为“选区→标准区”交互命令 |
| `cmpStd` | Function | 60 | — | UDF | B1 | 纯字符串解析，可完整单测 |
| `cmpStd2` | Function | 90 | — | UDF | B1 | Range 参数 → object[,]；依赖 cmpStd |

### 日报

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `运行一期中控质量日报` | Sub | 3 | — | 删除 | — | 空过程 |
| `运行15期中控质量日报` | Sub | 3 | — | 删除 | — | 空过程 |

### 模块1

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `筛选当前列为当前单元格值` | Sub | 12 | 选区 | 命令 | B2 |  |
| `筛选当前列非空值` | Sub | 6 | 选区 | 命令 | B2 |  |
| `筛选当前列数值` | Sub | 21 | 选区 | 命令 | B2 |  |
| `筛选最大日期近30天的数据` | Sub | 21 | 选区 | 命令 | B2 |  |
| `清除之前的筛选` | Sub | 6 | 选区 | 命令 | B2 |  |
| `清除当前列筛选` | Sub | 7 | 选区 | 命令 | B2 |  |
| `保留最后50条可见的数据` | Sub | 35 | 交互、选区 | 命令 | B2 |  |
| `筛选本月数据` | Sub | 31 | 交互、选区 | 命令 | B2 |  |
| `筛选上月数据` | Sub | 31 | 交互、选区 | 命令 | B2 |  |
| `隐藏符合条件的列` | Sub | 31 | 选区 | 命令 | B2 |  |
| `根据Ctrl向下查找隐藏列` | Sub | 26 | 选区 | 命令 | B2 |  |
| `隐藏符合条件的列2` | Sub | 32 | 选区 | 命令 | B2 |  |
| `隐藏符合条件的列_选中` | Sub | 56 | 选区 | 命令 | B2 |  |
| `隐藏符合条件的行_选中` | Sub | 46 | 选区 | 命令 | B2 |  |
| `取消隐藏列` | Sub | 8 | 交互、选区 | 命令 | B2 |  |
| `one` | Function | 3 | — | 删除 | — | 测试残留 |

### 模块2

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `标记异常_Boxplot` | Sub | 98 | 交互、选区、格式 | 命令 | B3 |  |
| `标记异常_3西格玛` | Sub | 104 | 交互、选区、格式 | 命令 | B3 |  |
| `标记异常_Grubbs_单侧最大最小` | Sub | 117 | 交互、选区、格式 | 命令 | B3 |  |
| `标记异常_CUSUM_Vmask` | Sub | 123 | 交互、选区、格式 | 命令 | B3 |  |
| `标记异常_EWMA` | Sub | 127 | 交互、选区、格式 | 命令 | B3 |  |
| `一键异常检测_五法对比` | Sub | 102 | 交互、选区、格式、跨模块 | 命令 | B3 |  |
| `MK趋势检验_Sen斜率` | Sub | 160 | 交互、选区 | 命令 | B3 |  |

### 模块3

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `画标记折线图` | Sub | 173 | 交互、图表、选区、格式 | 待确认 | B4 | 与 模块6.画标记折线图_final 重复 |
| `画标记折线图_0` | Sub | 191 | 交互、图表、选区、格式 | 待确认 | B4 | 旧版 |
| `添加辅助线0` | Sub | 75 | 交互、图表、选区 | 待确认 | B4 | 旧版辅助线，已被模块6取代？ |
| `添加辅助线函数` | Function | 3 | — | 删除 | — | 过渡壳，转调模块6 |
| `添加辅助线` | Sub | 3 | — | 删除 | — | 与模块6同名过渡壳 |
| `画钴杂比标记折线图` | Sub | 289 | 交互、图表、选区、格式 | 命令 | B4 |  |

### 模块4

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `初始化筛选器` | Sub | 47 | 交互、选区 | 命令 | B2 |  |
| `下一个样品类别` | Sub | 41 | 交互、选区 | 命令 | B2 |  |

### 模块5

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `FormatSmartNumber` | Function | 15 | — | 内部方法 | B4 |  |
| `FindHeaderCol` | Function | 13 | — | 内部方法 | B4 |  |
| `RoundUpLong` | Function | 7 | — | 内部方法 | B4 |  |
| `MakeSafeSheetName` | Function | 22 | — | 内部方法 | B4 |  |
| `HasLogicalOps` | Function | 3 | — | 内部方法 | B4 |  |
| `NameMatches` | Function | 37 | — | 内部方法 | B4 |  |
| `ApplyNameLogicalFilter` | Sub | 13 | — | 内部方法 | B4 |  |
| `GenerateCharts_AndExportData_ByRule` | Sub | 231 | 交互、图表、选区、格式 | 命令 | B4 |  |

### 模块6

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `HasChartAxis` | Function | 12 | 图表 | 内部方法 | B4 |  |
| `添加辅助线` | Sub | 3 | 交互 | 命令 | B4 | 打开辅助线窗口 |
| `添加辅助线函数_final` | Function | 3 | — | 内部方法 | B4 |  |
| `AddHelperLineFinal` | Function | 3 | — | 内部方法 | B4 |  |
| `AddHelperSeriesToSelectedChart` | Function | 58 | 交互、图表、选区 | 内部方法 | B4 |  |
| `画标记折线图_final` | Sub | 328 | 交互、图表、选区、格式、字典 | 命令 | B4 |  |
| `画控制图_final` | Sub | 390 | 交互、图表、选区、格式、字典 | 命令 | B4 |  |

### 模块7

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `GetValidArr` | Function | 9 | — | 内部方法 | B1 |  |
| `Cpk_MR` | Function | 51 | — | UDF | B1 |  |
| `Ppk_Calc` | Function | 35 | — | UDF | B1 |  |
| `Cpk_MR_FromArray` | Function | 32 | — | 内部方法 | B1 | Double() 参数，工作表不可调用 |
| `Cpk_Rank` | Function | 19 | — | UDF | B1 | 源码中“?”疑为编码丢字（原应为“-”或“：”），需确认 |

### 模块8

| 过程 | 类型 | 行数 | 依赖 | 迁移目标 | 批次 | 备注 |
|---|---|---:|---|---|---|---|
| `MK` | Function | 109 | 跨模块 | UDF | B1 | Enum 参数 → int；Norm_S_Dist → MathNet |
| `QuickSort` | Sub | 16 | 跨模块 | 内部方法 | B1 | 用 Array.Sort 替代 |
