using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 其他.bas 的数据处理宏与 函数.cmpStd222。</summary>
    public static class DataCommands
    {
        private const string Menu = "工具宏·数据";

        [ExcelCommand(Name = "工具宏_化学式上下标", MenuName = Menu, MenuText = "化学式上下标")]
        public static void ChemicalFormat() => Cmd.Run("化学式上下标", () =>
        {
            dynamic sel = Xl.TrimToUsed(Xl.RequireRangeSelection("请先选择包含化学式的单元格区域"));
            if (sel == null) return;
            int n = 0;
            foreach (dynamic cell in sel.Cells)
            {
                if ((bool)cell.HasFormula) continue; // 公式单元格无法逐字设置格式
                if (!(cell.Value2 is string text) || text.Length == 0) continue;
                cell.Font.Superscript = false;
                cell.Font.Subscript = false;
                foreach (var kv in ChemicalFormula.Analyze(text))
                {
                    dynamic f = cell.Characters[kv.Key, 1].Font;
                    if (kv.Value == ScriptKind.Subscript) f.Subscript = true; else f.Superscript = true;
                }
                n++;
            }
            Ui.Info($"化学式上下标格式处理完成！共 {n} 个单元格。");
        });

        [ExcelCommand(Name = "工具宏_检测数据转窄表", MenuName = Menu, MenuText = "检测数据转窄表")]
        public static void ToNarrowTable() => Cmd.Run("转窄表", () =>
        {
            var src = Xl.RequireActiveSheet();
            int lastCol = Xl.LastCol(src, 1), lastRow = Xl.LastRow(src, 1);
            if (lastRow < 2) throw new UserMessage("没有数据（以 A 列判断最后一行）。");
            // 用 .Value 而不是 .Value2：保留日期类型，写到窄表后仍显示为日期
            object[,] wide = Xl.Read(src.Range[src.Cells[1, 1], src.Cells[lastRow, lastCol]], useValue2: false);
            var narrow = NarrowTable.Transform(wide);

            // 原 VBA 写的是 ThisWorkbook（加载项自身），实际总是落到新工作簿；这里输出到当前工作簿
            dynamic wb = src.Parent;
            string name = SheetOps.SafeSheetName((string)src.Name, "窄表");
            dynamic dest = null;
            foreach (dynamic s in wb.Worksheets) if ((string)s.Name == name) { dest = s; break; }
            if (dest == null) { dest = wb.Worksheets.Add(After: src); dest.Name = name; }
            dest.Cells.Clear();
            Xl.Write(dest.Range["A1"], narrow);
            dest.Columns.AutoFit();
            dest.Activate();
            Ui.Info($"转换完成！共生成 {narrow.GetLength(0) - 1} 条检测行");
        });

        [ExcelCommand(Name = "工具宏_连接可见单元格", MenuName = Menu, MenuText = "连接可见单元格并复制")]
        public static void JoinVisible() => Cmd.Run("连接并复制", () =>
        {
            dynamic sel = Xl.TrimToUsed(Xl.RequireRangeSelection());
            List<(int Row, int Col, object Value)> cells = sel == null ? new List<(int, int, object)>() : Xl.VisibleCells(sel);
            var text = SheetOps.JoinNonBlank(cells.Select(c => c.Value).Where(v => !(v is CellError)));
            if (text.Length == 0) throw new UserMessage("所选可见单元格中没有有效内容");
            Ui.CopyText(text);
            Ui.Info("已复制到剪贴板：\n\n" + text);
        }, freezeScreen: false);

        [ExcelCommand(Name = "工具宏_统计可见数值", MenuName = Menu, MenuText = "统计可见数值并复制（均值/范围）")]
        public static void StatsVisible() => Cmd.Run("统计并复制", () =>
        {
            dynamic sel = Xl.RequireRangeSelection("请先选择需要统计的列（包含表头）");
            var lines = new List<string>();
            foreach (dynamic col in sel.Columns)
            {
                object[,] head = Xl.Read(col.Cells[1, 1]);
                string header = head[0, 0] is CellError ? "" : Vba.CStr(head[0, 0]);
                dynamic used = Xl.TrimToUsed(col);
                List<(int Row, int Col, object Value)> cells = used == null ? new List<(int Row, int Col, object Value)>() : Xl.VisibleCells(used);
                if (cells.Count == 0) { lines.Add((header == "" ? "未命名列" : header) + " 无可视单元格"); continue; }
                lines.Add(SheetOps.StatsLine(header, cells.Where(c => c.Row > 1).Select(c => c.Value)));
            }
            var text = string.Join("\r\n", lines);
            Ui.CopyText(text);
            Ui.Info("统计结果已复制到剪贴板：\n\n" + text);
        });

        [ExcelCommand(Name = "工具宏_标准校验着色", MenuName = Menu, MenuText = "按标准校验并着色…")]
        public static void StandardColoring() => Cmd.Run("标准校验", () =>
        {
            dynamic std = Ui.PickRange("第 1 步：选择【标准】区域");
            if (std == null) return;
            dynamic val = Ui.PickRange("第 2 步：选择【数值】区域（与标准逐格对应）");
            if (val == null) return;
            int n = ApplyStandardColors(std, val);
            Ui.Info($"校验完成！共 {n} 格。");
        });

        /// <summary>按 cmpStd222 的规则给数值区域着色，返回处理的格数。</summary>
        internal static int ApplyStandardColors(dynamic stdRange, dynamic valueRange)
        {
            dynamic app = Xl.App;
            app.EnableEvents = false;
            int count = System.Math.Min((int)stdRange.Cells.Count, (int)valueRange.Cells.Count);
            object[,] stdArr = Xl.Read(stdRange), valArr = Xl.Read(valueRange);
            var stds = Flatten(stdArr);
            var vals = Flatten(valArr);
            for (int i = 0; i < count; i++)
            {
                dynamic cell = valueRange.Cells[i + 1];
                cell.Interior.Pattern = Xl.xlPatternNone;
                cell.Font.ColorIndex = Xl.xlColorIndexAutomatic;
                switch (StandardMarks.Classify(stds[i], vals[i]))
                {
                    case StdMark.NotNumeric: Fill(cell, 255, 0, 0); break;
                    case StdMark.IsNullViolated: Fill(cell, 243, 84, 26); break;
                    case StdMark.Cankao2Missing: Fill(cell, 37, 109, 183); break;
                    case StdMark.NotNullMissing: Fill(cell, 255, 136, 16); break;
                    case StdMark.Untested: Fill(cell, 169, 208, 142); break;
                    case StdMark.OutOfSpec: cell.Font.Color = Xl.Rgb(255, 0, 0); break;
                }
            }
            return count;
        }

        private static void Fill(dynamic cell, int r, int g, int b)
        {
            cell.Interior.Pattern = Xl.xlPatternSolid;
            cell.Interior.Color = Xl.Rgb(r, g, b);
        }

        private static object[] Flatten(object[,] a) => a.Cast<object>().ToArray(); // C# 多维数组按行优先枚举，与 Range.Cells(i) 一致
    }
}
