using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace YTools.Core
{
    /// <summary>
    /// 不依赖 Excel 的表格运算。区域值统一是 0 基 object[,]（加载项层负责把 Excel 的 1 基数组转过来）。
    /// </summary>
    public static class SheetOps
    {
        /// <summary>VBA 的 cell.Value &lt;&gt; "" 判断：空单元格与空字符串都算空。</summary>
        public static bool IsBlank(object v) => v == null || (v is string s && s.Length == 0);

        /// <summary>区域中整列为空的列（返回 0 基列号）。fromRow 为开始检查的行（0 基）。</summary>
        public static List<int> BlankColumns(object[,] block, int fromRow = 0)
        {
            var res = new List<int>();
            int rows = block.GetLength(0), cols = block.GetLength(1);
            for (int c = 0; c < cols; c++)
            {
                bool blank = true;
                for (int r = fromRow; r < rows && blank; r++) blank = IsBlank(block[r, c]);
                if (blank) res.Add(c);
            }
            return res;
        }

        /// <summary>区域中整行为空的行（0 基行号）。</summary>
        public static List<int> BlankRows(object[,] block)
        {
            var res = new List<int>();
            int rows = block.GetLength(0), cols = block.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                bool blank = true;
                for (int c = 0; c < cols && blank; c++) blank = IsBlank(block[r, c]);
                if (blank) res.Add(r);
            }
            return res;
        }

        /// <summary>只保留最后 keep 条可见行：返回需要隐藏的行（按传入顺序，从前往后）。</summary>
        public static List<int> RowsToHideKeepingLast(IList<int> visibleDataRows, int keep = 50)
        {
            int extra = visibleDataRows.Count - keep;
            return extra > 0 ? visibleDataRows.Take(extra).ToList() : new List<int>();
        }

        /// <summary>用“、”连接非空值（原 连接可见单元格并复制）。</summary>
        public static string JoinNonBlank(IEnumerable<object> values, string separator = "、")
            => string.Join(separator, values.Where(v => !IsBlank(v)).Select(Vba.CStr));

        /// <summary>
        /// 一列可见数值的统计行：“表头 均值x.xxxxg/L(min-maxg/L)”，只统计 &gt;0 的数值。
        /// </summary>
        public static string StatsLine(string header, IEnumerable<object> dataCells)
        {
            if (string.IsNullOrEmpty(header)) header = "未命名列";
            var vals = new List<double>();
            foreach (var v in dataCells)
            {
                // 原 VBA 遇到文本单元格会报“类型不匹配”中断，这里直接跳过
                if (v == null || v is bool) continue;
                if (Vba.TryCDbl(v, out var d) && d > 0) vals.Add(d);
            }
            if (vals.Count == 0) return header + " 无大于0的可见数值";
            string F(double x) => x.ToString("0.0000", CultureInfo.InvariantCulture);
            return $"{header} 均值{F(vals.Average())}g/L({F(vals.Min())}-{F(vals.Max())}g/L)";
        }

        /// <summary>工作表名去掉非法字符并截到 31 个字符。</summary>
        public static string SafeSheetName(string name, string suffix = "")
        {
            var sb = new StringBuilder(name ?? "");
            foreach (var ch in new[] { '\\', '/', ':', '*', '?', '[', ']' }) sb.Replace(ch, '_');
            var s = sb + suffix;
            return s.Length > 31 ? s.Substring(0, 31) : s;
        }
    }

    /// <summary>检测数据宽表 → 窄表（原 其他.检测数据转窄表）。</summary>
    public static class NarrowTable
    {
        public static readonly string[] KeepFields =
            { "送样日期", "送样时间", "报出时间", "分析时长", "复核人", "样品名称", "录入人", "判定", "样品类别" };

        /// <summary>
        /// 第 0 行为表头。保留字段原样带出；其余每一列都视为检测项目，非空值各生成一行。
        /// 表里缺少某个保留字段时该列留空（原 VBA 会报错中断）。
        /// </summary>
        public static object[,] Transform(object[,] wide)
        {
            int rows = wide.GetLength(0), cols = wide.GetLength(1);
            var keepIdx = Enumerable.Repeat(-1, KeepFields.Length).ToArray();
            var testIdx = new List<int>();
            for (int c = 0; c < cols; c++)
            {
                string h = Vba.Trim(Vba.CStr(wide[0, c] is CellError ? "" : wide[0, c]));
                int k = Array.IndexOf(KeepFields, h);
                if (k >= 0 && keepIdx[k] < 0) keepIdx[k] = c;
                else if (k < 0) testIdx.Add(c);
            }

            var outRows = new List<object[]>();
            int width = KeepFields.Length + 2;
            var header = new object[width];
            for (int k = 0; k < KeepFields.Length; k++) header[k] = KeepFields[k];
            header[KeepFields.Length] = "检测项目";
            header[KeepFields.Length + 1] = "检测结果";
            outRows.Add(header);

            for (int r = 1; r < rows; r++)
                foreach (int c in testIdx)
                {
                    var v = wide[r, c];
                    if (v == null) continue;
                    var line = new object[width];
                    for (int k = 0; k < KeepFields.Length; k++) line[k] = keepIdx[k] >= 0 ? wide[r, keepIdx[k]] : null;
                    line[KeepFields.Length] = Vba.Trim(Vba.CStr(wide[0, c] is CellError ? "" : wide[0, c]));
                    line[KeepFields.Length + 1] = v;
                    outRows.Add(line);
                }

            var res = new object[outRows.Count, width];
            for (int r = 0; r < outRows.Count; r++)
                for (int c = 0; c < width; c++) res[r, c] = outRows[r][c];
            return res;
        }
    }
}
