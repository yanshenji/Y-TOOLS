using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace YTools.Core
{
    /// <summary>画图命令（原 模块3 / 模块5 / 模块6）中与 Excel 无关的计算。</summary>
    public static class ChartMath
    {
        /// <summary>
        /// 收集可画的点：跳过隐藏行；Y 去空格后非空且为数值才保留（文本如“未测”“N/A”自动跳过）。
        /// 返回保留下来的行下标（0 基）。
        /// </summary>
        public static List<int> PlottableRows(IList<object> y, IList<bool> hidden)
        {
            var rows = new List<int>();
            for (int i = 0; i < y.Count; i++)
            {
                if (hidden != null && hidden[i]) continue;
                var v = y[i];
                if (v == null || v is CellError) continue;
                if (Vba.Trim(Vba.CStr(v)).Length == 0) continue;
                if (Vba.IsNumeric(v)) rows.Add(i);
            }
            return rows;
        }

        /// <summary>单值-移动极差（I-MR）控制限：CL=均值，UCL/LCL=CL±3·MR̄/1.128。</summary>
        public static (double CL, double UCL, double LCL, bool[] OutOfControl) IndividualsLimits(double[] y)
        {
            if (y.Length < 2) throw new ArgumentException("控制图至少需要 2 个点。");
            double cl = y.Average();
            double mrBar = 0;
            for (int i = 1; i < y.Length; i++) mrBar += Math.Abs(y[i] - y[i - 1]);
            mrBar /= y.Length - 1;
            double ucl = cl + 3 * mrBar / Capability.D2, lcl = cl - 3 * mrBar / Capability.D2;
            return (cl, ucl, lcl, y.Select(v => v > ucl || v < lcl).ToArray());
        }

        /// <summary>Y 轴初始范围：数据范围上下各留 5%；数据全相同时按 10%（或 1）展开。</summary>
        public static (double Min, double Max) PaddedRange(double min, double max)
        {
            double span = max - min;
            if (span == 0) span = min == 0 ? 1 : Math.Abs(min) * 0.1;
            return (min - span * 0.05, max + span * 0.05);
        }

        /// <summary>
        /// 把范围对齐到刻度单位：下限向下取整、上限向上取整。
        /// 原 VBA 用 ROUNDDOWN/ROUNDUP（向零截断 / 远离零），数据为负时会把轴边界收到数据点以内，截掉负值点；
        /// 数据为正时两者结果相同。
        /// </summary>
        public static (double Min, double Max) AlignToUnit(double min, double max, double unit)
        {
            if (unit <= 0) return (min, max);
            return (Math.Floor(min / unit) * unit, Math.Ceiling(max / unit) * unit);
        }

        /// <summary>
        /// 图表标题：优先“[表名 - 样品类别 - 项目]后缀”，其次样品名称，最后只有表名和项目。
        /// 样品类别为空、错误值、“未知样品”或表头字样时跳过。
        /// </summary>
        public static string Title(string sheetName, object category, object sampleName, string series, string suffix)
        {
            string Ok(object v, params string[] reject)
            {
                if (v == null || v is CellError) return null;
                var s = Vba.CStr(v);
                return Vba.Trim(s).Length == 0 || reject.Contains(s) ? null : s;
            }
            var who = Ok(category, "未知样品", "样品类别") ?? Ok(sampleName, "样品名称");
            return who != null ? $"[{sheetName} - {who} - {series}]{suffix}" : $"[{sheetName} - {series}]{suffix}";
        }

        /// <summary>钴杂比 = Co / 杂质；杂质为 0 或非数值时记 0（与原 VBA 一致）。</summary>
        public static double[] CobaltRatio(IList<object> co, IList<object> impurity)
        {
            if (co.Count != impurity.Count) throw new ArgumentException("X值和钴杂比值数量不一致，无法绘制图表。");
            var r = new double[co.Count];
            for (int i = 0; i < r.Length; i++)
            {
                bool okImp = Vba.TryCDbl(impurity[i], out var imp) && !(impurity[i] is string s && s.Trim().Length == 0);
                bool okCo = Vba.TryCDbl(co[i], out var c);
                r[i] = okImp && okCo && imp != 0 ? c / imp : 0;
            }
            return r;
        }

        /// <summary>辅助线的点：X 在第一条系列的最小~最大值之间等分，Y 全为给定值。</summary>
        public static (double[] X, double[] Y) HelperLine(IList<double> xs, double y)
        {
            int n = xs.Count;
            if (n == 0) throw new ArgumentException("图表中没有数据系列，无法添加辅助线！");
            double min = xs.Min(), max = xs.Max();
            var hx = new double[n];
            for (int i = 0; i < n; i++) hx[i] = n == 1 ? min : min + (max - min) * i / (n - 1);
            return (hx, Enumerable.Repeat(y, n).ToArray());
        }

        /// <summary>图表在网格中的位置（每行 perRow 张，从左上角开始）。</summary>
        public static (double Left, double Top) GridPosition(int index, int perRow, double w, double h,
                                                              double left0 = 20, double top0 = 20, double hGap = 15, double vGap = 12)
            => (left0 + (index % perRow) * (w + hGap), top0 + (index / perRow) * (h + vGap));
    }

    /// <summary>配置表一行规则（原 模块5.GenerateCharts_AndExportData_ByRule，配置表从第 7 行起）。</summary>
    public sealed class ChartRule
    {
        public string SheetName, SampleType, Pattern, Items;
        public double? StartDate, EndDate;

        public string Label => string.IsNullOrEmpty(SampleType) ? Pattern : SampleType;
        public string[] ItemList => (Items ?? "").Split(',').Select(s => s.Trim()).Where(s => s.Length > 0).ToArray();

        public static ChartRule FromRow(object[] cells)
        {
            object At(int i) => i < cells.Length ? cells[i] : null;
            string S(int i) => At(i) == null || At(i) is CellError ? "" : Vba.CStr(At(i));
            double? D(int i) => At(i) == null || At(i) is CellError ? (double?)null : Vba.TryCDbl(At(i), out var d) ? d : (double?)null;
            var sheet = Vba.Trim(S(0));
            if (sheet.Length == 0) return null;
            return new ChartRule { SheetName = sheet, SampleType = Vba.Trim(S(1)), Pattern = S(2), StartDate = D(3), EndDate = D(4), Items = S(5) };
        }
    }

    public static class RuleCharts
    {
        /// <summary>样品名称条件中含“,”（或）或“&amp;”（且）时用自定义匹配，否则交给自动筛选。</summary>
        public static bool HasLogicalOps(string expr) => expr.Contains(",") || expr.Contains("&");

        /// <summary>
        /// 样品名称匹配：“,”分隔为“或”，“&amp;”分隔为“且”；不含通配符的词按“包含”处理；不区分大小写。
        /// 通配符同 VBA Like：* ? # [字符组]。
        /// </summary>
        public static bool NameMatches(string text, string expr)
        {
            if (string.IsNullOrWhiteSpace(expr) || expr == "*") return true;
            text = (text ?? "").ToLowerInvariant();
            foreach (var orPart in expr.Split(','))
            {
                bool ok = true;
                foreach (var raw in orPart.Split('&'))
                {
                    var t = Vba.Trim(raw).ToLowerInvariant();
                    if (t.Length == 0) continue;
                    var pat = t.IndexOf('*') < 0 && t.IndexOf('?') < 0 ? "*" + t + "*" : t;
                    if (!Like(text, pat)) { ok = false; break; }
                }
                if (ok) return true;
            }
            return false;
        }

        /// <summary>VBA Like 运算符（Option Compare Binary）。</summary>
        public static bool Like(string text, string pattern)
        {
            var sb = new StringBuilder("^");
            for (int i = 0; i < pattern.Length; i++)
            {
                char c = pattern[i];
                switch (c)
                {
                    case '*': sb.Append(".*"); break;
                    case '?': sb.Append('.'); break;
                    case '#': sb.Append("[0-9]"); break;
                    case '[':
                        int end = pattern.IndexOf(']', i + 1);
                        if (end < 0) { sb.Append(@"\["); break; }
                        var body = pattern.Substring(i + 1, end - i - 1);
                        if (body.StartsWith("!")) body = "^" + body.Substring(1);
                        sb.Append('[').Append(body.Replace(@"\", @"\\")).Append(']');
                        i = end;
                        break;
                    default: sb.Append(Regex.Escape(c.ToString())); break;
                }
            }
            return Regex.IsMatch(text, sb.Append('$').ToString(), RegexOptions.Singleline);
        }

        /// <summary>标题里的数值：≥1 保留两位小数，&lt;1 保留两位有效数字。</summary>
        public static string SmartNumber(double v)
        {
            if (v == 0) return "0";
            double av = Math.Abs(v);
            if (av >= 1) return Math.Round((decimal)v, 2, MidpointRounding.AwayFromZero).ToString("0.00", CultureInfo.InvariantCulture);
            double p = Math.Floor(Math.Log10(av));
            double factor = Math.Pow(10, 1 - p);
            return Vba.CStr(Math.Round(v * factor, MidpointRounding.ToEven) / factor);
        }

        /// <summary>去掉非法字符、截到 31 个字符；与已有工作表重名时在末尾加序号 2、3…</summary>
        public static string UniqueSheetName(string baseName, ICollection<string> existing)
        {
            var s = Vba.Trim(baseName ?? "");
            foreach (var bad in new[] { "[", "]", ":", "*", "?", "/", "\\", "\n", "\r" }) s = s.Replace(bad, "_");
            if (s.Length == 0) s = "图表";
            if (s.Length > 31) s = s.Substring(0, 31);
            var set = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase); // Excel 表名不区分大小写
            string candidate = s;
            for (int n = 2; set.Contains(candidate); n++)
                candidate = s.Substring(0, Math.Min(s.Length, 31 - n.ToString().Length)) + n;
            return candidate;
        }

        /// <summary>日期轴主刻度：跨度/10 向上取整，至少 1 天。</summary>
        public static long DateMajorUnit(double span) => Math.Max(1, (long)Math.Ceiling(span / 10));
    }
}
