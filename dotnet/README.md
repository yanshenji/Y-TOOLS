# 工具宏 · Excel-DNA 版

VBA 加载项「工具宏」的 C# 重写版，迁移计划见 [`docs/migration/excel-dna-inventory.md`](../docs/migration/excel-dna-inventory.md)。

## 目录

| 项目 | 说明 |
|---|---|
| `src/YTools.Core` | 纯算法（netstandard2.0），不依赖 Excel，可在任何平台测试 |
| `src/YTools.AddIn` | Excel-DNA 加载项（net48）：函数注册、功能区、命令 |
| `tests/YTools.Core.Tests` | xUnit 单元测试 |

## 构建

```bash
cd dotnet
dotnet test                                  # 跑单元测试
dotnet build src/YTools.AddIn -c Release     # 生成加载项
```

产物在 `src/YTools.AddIn/bin/Release/net48/publish/`：

- `YTools-packed.xll`：**32 位** Excel / WPS 用这个
- `YTools64-packed.xll`：64 位备用

依赖库已经打进 `.xll` 里，部署时只需要拷这一个文件。

## 安装

1. 把 `YTools-packed.xll` 拷到本机固定目录，例如 `D:\工具宏\`。
2. 如果文件是从网上或聊天工具下载的，先右键 → 属性 → 勾选「解除锁定」，否则 Excel 会静默拦截。
3. Excel：文件 → 选项 → 加载项 → 管理「Excel 加载项」→ 转到 → 浏览 → 选择 `.xll`。
   WPS：开发工具 → 加载项（或 COM 加载项）→ 浏览 → 选择 `.xll`。

## 与 `.xlam` 并存

迁移期间 `.xlam` 和 `.xll` 会同时加载。两边有**同名工作表函数**（如 `Cpk_MR`）时，公式到底调用哪一个，需要实测确认（见下方冒烟测试第 4 项）。
确认之后，再把 `.xlam` 中对应的函数删除，或在模块顶部加 `Option Private Module`，让它们不再出现在工作表里。VBA 宏内部对这些函数的调用不受影响。

## 冒烟测试（Excel 与 WPS 各做一遍）

| # | 操作 | 期望 |
|---|---|---|
| 1 | 加载 `.xll` | 无报错 |
| 2 | 任意单元格输入 `=工具宏版本()` | 显示 `工具宏 0.1.0.0 (32 位)` |
| 3 | 看功能区 | 出现「工具宏」选项卡，点「环境信息」弹出对话框 |
| 3b | 如果没有选项卡 | 到「加载项」选项卡 → 工具宏 → 环境信息（备用入口） |
| 4 | 只开 `.xll`、不开 `.xlam`，打开一份用了 `Cpk_MR` 等函数的旧报表 | 公式正常计算 |
| 5 | `=保留两位有效数字_四舍六入五成双(2.36)` | `2.4`（旧 VBA 版是 2.3） |
| 6 | `=GrubbsCrit(10)` | `2.2900...` |
| 7 | `=MK(A1:A5,1)`，A1:A5 填 1~5 | `10` |
| 8 | `=Cpk_Rank(0.5)` | `D级-严重不足` |
| 9 | `=GrubbsOutliers(...)` 返回多个值时 | 能向下溢出，或按数组公式输入能正常显示 |

请把「环境信息」对话框截图，以及每项是否通过，反馈回来。WPS 个人版和专业版的结果可能不同，用的是哪个版本请注明。
