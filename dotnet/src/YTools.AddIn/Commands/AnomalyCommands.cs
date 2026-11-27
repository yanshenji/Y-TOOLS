using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 模块2：异常检测五法、五法对比、MK 趋势检验。</summary>
    public static class AnomalyCommands
    {
        private const string Menu = "工具宏·异常检测";

        private enum Method { ThreeSigma, GrubbsOneSided, Cusum, Ewma, Boxplot }

        private static readonly (Method M, string Name)[] Methods =
        {
            (Method.ThreeSigma, "3西格玛"), (Method.GrubbsOneSided, "Grubbs单侧"), (Method.Cusum, "CUSUM_Vmask"),
            (Method.Ewma, "EWMA"), (Method.Boxplot, "Boxplot_IQR"),
        };

        private static readonly Dictionary<Method, (string Done, string Tip)> Texts = new Dictionary<Method, (string, string)>
        {
            [Method.ThreeSigma] = ("3西格玛异常值标记完成！", ""),
            [Method.GrubbsOneSided] = ("Grubbs单侧(最大/最小)标记完成！", "【提示】格拉布斯检验前提：数据近似正态；非正态建议改用IQR箱线法！"),
            [Method.Cusum] = ("CUSUM V-mask标记完成！", "【提示】CUSUM累积和：本程序使用样本均值作为目标μ0；在线监控建议输入历史基准μ0、σ！"),
            [Method.Ewma] = ("EWMA-SPC标记完成！", "【提示】EWMA指数加权移动平均；本程序使用样本均值作为μ0；正式SPC请使用历史稳态μ0、σ！"),
            [Method.Boxplot] = ("异常值已标记完成！", ""),
        };

        [ExcelCommand(Name = "工具宏_异常_3西格玛", MenuName = Menu, MenuText = "3西格玛")]
        public static void ThreeSigma() => RunOnSelection(Method.ThreeSigma);

        [ExcelCommand(Name = "工具宏_异常_Grubbs单侧", MenuName = Menu, MenuText = "Grubbs 单侧(最大/最小)")]
        public static void GrubbsOneSided() => RunOnSelection(Method.GrubbsOneSided);

        [ExcelCommand(Name = "工具宏_异常_CUSUM", MenuName = Menu, MenuText = "CUSUM")]
        public static void Cusum() => RunOnSelection(Method.Cusum);

        [ExcelCommand(Name = "工具宏_异常_EWMA", MenuName = Menu, MenuText = "EWMA")]
        public static void Ewma() => RunOnSelection(Method.Ewma);

        [ExcelCommand(Name = "工具宏_异常_箱线图", MenuName = Menu, MenuText = "箱线图(IQR)")]
        public static void Boxplot() => RunOnSelection(Method.Boxplot);

        private static void RunOnSelection(Method m) => Cmd.Run(Methods.First(x => x.M == m).Name, () =>
        {
            dynamic sel = Xl.RequireRangeSelection();
            var ws = sel.Worksheet;
            List<int> cols = VisibleSelectedColumns(sel);
            var lines = cols.Select(c => $"第{c}列 " + RunColumn(ws, c, m).Info).ToList();
            var (done, tip) = Texts[m];
            Ui.Info(done + "\n\n" + (tip.Length > 0 ? tip + "\n\n" : "") + string.Join("\n\n", lines));
        });

        /// <summary>一列：读可见单元格 → 计算 → 着色，返回结果。</summary>
        private static ColumnResult RunColumn(dynamic ws, int col, Method m)
        {
            int last = Xl.LastRow(ws, col);
            List<(int Row, int Col, object Value)> cells = Xl.VisibleCells(ws.Range[ws.Cells[1, col], ws.Cells[last, col]]);
            var values = cells.Select(c => c.Value).ToList();
            ColumnResult r;
            switch (m)
            {
                case Method.ThreeSigma: r = Anomaly.ThreeSigma(values); break;
                case Method.GrubbsOneSided: r = Anomaly.GrubbsOneSided(values); break;
                case Method.Cusum: r = Anomaly.Cusum(values); break;
                case Method.Ewma: r = Anomaly.Ewma(values); break;
                default: r = Anomaly.Boxplot(values, out _, out _); break;
            }
            if (r.Marks != null) ApplyMarks(ws, col, cells.Select(c => c.Row).ToList(), r.Marks);
            return r;
        }

        /// <summary>选区中未隐藏、且在已用区域内的列号（原 VBA：Intersect(UsedRange, 可见列)）。</summary>
        private static List<int> VisibleSelectedColumns(dynamic sel)
        {
            dynamic ws = sel.Worksheet;
            dynamic used = ws.UsedRange;
            int u1 = used.Column, u2 = u1 + (int)used.Columns.Count - 1;
            var cols = new SortedSet<int>();
            foreach (dynamic area in sel.Areas)
                foreach (dynamic col in area.Columns)
                {
                    int c = col.Column;
                    if (c >= u1 && c <= u2 && !(bool)col.EntireColumn.Hidden) cols.Add(c);
                }
            if (cols.Count == 0) throw new UserMessage("没有找到可见的选中列！");
            return cols.ToList();
        }

        /// <summary>连续同类标记合并成一个区域设置底色。</summary>
        private static void ApplyMarks(dynamic ws, int col, IList<int> rows, CellMark[] marks)
        {
            int i = 0;
            while (i < rows.Count)
            {
                int j = i;
                while (j + 1 < rows.Count && marks[j + 1] == marks[i] && rows[j + 1] == rows[j] + 1) j++;
                if (marks[i] != CellMark.Keep)
                {
                    dynamic r = ws.Range[ws.Cells[rows[i], col], ws.Cells[rows[j], col]];
                    if (marks[i] == CellMark.Flag) r.Interior.Color = Xl.Rgb(255, 255, 0);
                    else r.Interior.ColorIndex = Xl.xlColorIndexNone;
                }
                i = j + 1;
            }
        }

        /// <summary>
        /// 一键五法对比：把可见数据复制 5 份到新工作表，每份用一种方法标记，最后统一汇总一次
        /// （原 VBA 每种方法各弹一次提示框，共 5 次）。
        /// </summary>
        [ExcelCommand(Name = "工具宏_异常_五法对比", MenuName = Menu, MenuText = "一键五法对比")]
        public static void CompareAll() => Cmd.Run("五法对比", () =>
        {
            dynamic sel = Xl.RequireRangeSelection("请先在源表选中数据区域（可多选列）！");
            dynamic src = sel.Worksheet;
            List<int> cols = VisibleSelectedColumns(sel);
            dynamic union = null;
            foreach (int c in cols) union = union == null ? src.Columns[c] : Xl.App.Union(union, src.Columns[c]);
            dynamic block = Xl.App.Intersect(src.UsedRange, union);
            dynamic visible;
            try { visible = block.SpecialCells(Xl.xlCellTypeVisible); }
            catch (System.Runtime.InteropServices.COMException) { throw new UserMessage("没有可见数据！"); }

            dynamic dest = src.Parent.Worksheets.Add(After: src);
            dest.Name = "异常检测_" + DateTime.Now.ToString("HHmmss");
            int width = cols.Count, offset = 1;
            var summary = new List<string>();
            foreach (var (m, name) in Methods)
            {
                dest.Cells[1, offset].Value = "【" + name + "】";
                dest.Cells[1, offset].Font.Bold = true;
                visible.Copy(Destination: dest.Cells[2, offset]); // 只复制可见行
                int flagged = 0;
                for (int c = offset; c < offset + width; c++) flagged += RunColumn(dest, c, m).Flagged;
                summary.Add($"{name}：标记 {flagged} 个");
                offset += width + 2;
            }
            dest.Activate();
            dest.Cells[1, 1].Select();
            Ui.Info($"五法对比完成！新工作表【{dest.Name}】\n\n已自动：跳过隐藏列 / 只复制筛选后的可见行\n从左到右：3σ → Grubbs → CUSUM → EWMA → Boxplot(IQR)\n\n"
                    + string.Join("\n", summary));
        });

        [ExcelCommand(Name = "工具宏_MK趋势检验", MenuName = Menu, MenuText = "MK 趋势检验 + Sen 斜率")]
        public static void MannKendallTrend() => Cmd.Run("MK 趋势检验", () =>
        {
            dynamic sel = Xl.RequireRangeSelection();
            dynamic ws = sel.Worksheet;
            List<int> cols = VisibleSelectedColumns(sel);
            var lines = new List<string>();
            foreach (int c in cols)
            {
                int last = Xl.LastRow(ws, c);
                List<(int Row, int Col, object Value)> cells = Xl.VisibleCells(ws.Range[ws.Cells[1, c], ws.Cells[last, c]]);
                lines.Add($"第{c}列 " + Anomaly.MannKendallSummary(cells.Select(x => x.Value).ToList()));
            }
            Ui.Info("【Mann-Kendall趋势检验 + Sen斜率】α=0.05\n\n" + string.Join("\n\n", lines));
        });
    }
}
