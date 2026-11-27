using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace YTools.Core
{
    /// <summary>单元格标记：保持不动 / 清除底色 / 标黄。</summary>
    public enum CellMark { Keep, Clear, Flag }

    /// <summary>一列的检测结果：每个可见单元格的标记，以及用于提示框的说明文字。</summary>
    public sealed class ColumnResult
    {
        public CellMark[] Marks { get; set; }
        public int Flagged => Marks?.Count(m => m == CellMark.Flag) ?? 0;
        public string Info { get; set; }
        public bool Skipped { get; set; }
    }

    /// <summary>
    /// 异常检测五法（原 模块2）。输入为一列中的可见单元格值（从第 1 行到最后一行，已排除隐藏行）。
    /// </summary>
    public static class Anomaly
    {
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        /// <summary>VBA Round（银行家舍入）后按 CStr 输出，用于提示文字与原版一致。</summary>
        public static string R(double x, int digits) => Math.Round(x, digits, MidpointRounding.ToEven).ToString("G15", Inv);

        // WorksheetFunction 作用在区域上时只认数值单元格（文本型数字、逻辑值都不算）
        private static double[] RangeNumbers(IList<object> cells) => cells.OfType<double>().ToArray();

        // 各方法逐格循环里用的 VBA IsNumeric 判定（文本型数字也算）
        private static bool IsNum(object v) => !Vba.IsEmpty(v) && !(v is bool) && Vba.IsNumeric(v);

        /// <summary>
        /// 箱线图 IQR：下限 = max(Q1-1.5·IQR, 0.1·Q1)，上限 = Q3+1.5·IQR（QUARTILE 包含端点）。
        /// 非数值单元格只清除底色（原 VBA 会把空单元格当 0 比较，下限 &gt; 0 时空格也被标黄）。
        /// </summary>
        public static ColumnResult Boxplot(IList<object> cells, out double lower, out double upper)
        {
            lower = upper = 0;
            var nums = RangeNumbers(cells);
            if (nums.Length == 0) return new ColumnResult { Skipped = true, Info = "无数值，跳过" };
            double q1 = QuartileInc(nums, 1), q3 = QuartileInc(nums, 3), iqr = q3 - q1;
            double lo = Math.Max(q1 - 1.5 * iqr, 0.1 * q1), hi = q3 + 1.5 * iqr, max = nums.Max();
            lower = lo; upper = hi;
            var marks = cells.Select(c => c is double v
                ? ((v < lo && v >= 0) || (v > hi && v <= max) ? CellMark.Flag : CellMark.Clear)
                : CellMark.Clear).ToArray();
            return new ColumnResult { Marks = marks, Info = $"lowerBound={R(lo, 6)}  upperBound={R(hi, 6)}" };
        }

        /// <summary>QUARTILE / QUARTILE.INC：包含端点的线性插值。</summary>
        public static double QuartileInc(double[] x, int quart)
        {
            var s = x.OrderBy(v => v).ToArray();
            double pos = (s.Length - 1) * quart / 4.0;
            int lo = (int)Math.Floor(pos);
            double frac = pos - lo;
            return lo + 1 < s.Length ? s[lo] + frac * (s[lo + 1] - s[lo]) : s[lo];
        }

        /// <summary>3σ：均值 ± 3·STDEV.S 之外标黄；非数值单元格不动。</summary>
        public static ColumnResult ThreeSigma(IList<object> cells)
        {
            var nums = RangeNumbers(cells);
            if (nums.Length < 2) return new ColumnResult { Skipped = true, Info = "样本不足，跳过" };
            double mean = Stats.Average(nums), sd = Stats.StDevS(nums);
            if (sd == 0) return new ColumnResult { Skipped = true, Info = "全部数值相同，跳过" };
            double lo = mean - 3 * sd, hi = mean + 3 * sd;
            var marks = cells.Select(c => c is double v ? (v < lo || v > hi ? CellMark.Flag : CellMark.Clear) : CellMark.Keep).ToArray();
            return new ColumnResult { Marks = marks, Info = $"均值={R(mean, 6)} σ={R(sd, 6)}  范围 {R(lo, 6)} ~ {R(hi, 6)}" };
        }

        /// <summary>
        /// Grubbs 单侧（最大/最小）：|x-均值|/S 超过临界值标黄，临界值 t 取 T.INV.2T(α, n-2)（与原 VBA 一致）。
        /// </summary>
        public static ColumnResult GrubbsOneSided(IList<object> cells, double alpha = 0.05)
        {
            var nums = cells.Where(IsNum).Select(Vba.CDbl).ToArray();
            int n = nums.Length;
            if (n < 3) return new ColumnResult { Skipped = true, Info = "样本n<3，跳过" };
            double mean = Stats.Average(nums), sd = Stats.StDevS(nums);
            if (sd < 1e-12) return new ColumnResult { Skipped = true, Info = "全部数值相同，跳过" };
            double t = Stats.TInv2T(alpha, n - 2);
            double crit = (n - 1) / Math.Sqrt(n) * Math.Sqrt(t * t / ((n - 2) + t * t));
            var marks = cells.Select(c =>
            {
                if (!IsNum(c)) return CellMark.Clear;
                double v = Vba.CDbl(c);
                return (v - mean) / sd > crit || (mean - v) / sd > crit ? CellMark.Flag : CellMark.Clear;
            }).ToArray();
            return new ColumnResult
            {
                Marks = marks,
                Info = $"n={n}\n  Max={R(nums.Max(), 6)}  Min={R(nums.Min(), 6)}\n  单侧Gcrit={R(crit, 4)}"
            };
        }

        /// <summary>CUSUM（表格法，目标 μ0 取样本均值）：C+ 或 C- 超过 h·σ 的点标黄。</summary>
        public static ColumnResult Cusum(IList<object> cells, double k = 0.5, double h = 5)
        {
            var nums = cells.Where(IsNum).Select(Vba.CDbl).ToArray();
            int n = nums.Length;
            if (n < 5) return new ColumnResult { Skipped = true, Info = "样本n<5，跳过(CUSUM至少5点)" };
            double mean = Stats.Average(nums), sd = Stats.StDevS(nums);
            if (sd < 1e-12) return new ColumnResult { Skipped = true, Info = "全部数值相同，跳过" };
            double cp = 0, cm = 0;
            var marks = new CellMark[cells.Count];
            for (int i = 0; i < cells.Count; i++)
            {
                marks[i] = CellMark.Clear;
                if (!IsNum(cells[i])) continue;
                double x = Vba.CDbl(cells[i]);
                cp = Math.Max(0, x - mean - k * sd + cp);
                cm = Math.Max(0, mean - x - k * sd + cm);
                if (cp > h * sd || cm > h * sd) marks[i] = CellMark.Flag;
            }
            var r = new ColumnResult { Marks = marks };
            r.Info = $"n={n}\n  xBar={R(mean, 6)} σ={R(sd, 6)}\n  k={k.ToString(Inv)} h={h.ToString(Inv)} 报警点数={r.Flagged}";
            return r;
        }

        /// <summary>EWMA（μ0 取样本均值，控制限随 i 变化）：Z 超出控制限的点标黄。</summary>
        public static ColumnResult Ewma(IList<object> cells, double lambda = 0.2, double L = 3)
        {
            var nums = cells.Where(IsNum).Select(Vba.CDbl).ToArray();
            int n = nums.Length;
            if (n < 5) return new ColumnResult { Skipped = true, Info = "样本n<5，跳过(EWMA至少5点)" };
            double mean = Stats.Average(nums), sd = Stats.StDevS(nums);
            if (sd < 1e-12) return new ColumnResult { Skipped = true, Info = "全部数值相同，跳过" };
            double z = mean;
            int i = 0;
            var marks = new CellMark[cells.Count];
            for (int c = 0; c < cells.Count; c++)
            {
                marks[c] = CellMark.Clear;
                if (!IsNum(cells[c])) continue;
                i++;
                z = lambda * Vba.CDbl(cells[c]) + (1 - lambda) * z;
                double f = Math.Sqrt(lambda / (2 - lambda) * (1 - Math.Pow(1 - lambda, 2 * i)));
                if (z > mean + L * sd * f || z < mean - L * sd * f) marks[c] = CellMark.Flag;
            }
            var r = new ColumnResult { Marks = marks };
            r.Info = $"n={n}\n  xBar={R(mean, 6)} σ={R(sd, 6)}\n  λ={lambda.ToString(Inv)} L={L.ToString(Inv)} 报警点数={r.Flagged}";
            return r;
        }

        /// <summary>MK 趋势检验 + Sen 斜率的一列汇总文字（原 MK趋势检验_Sen斜率）。</summary>
        public static string MannKendallSummary(IList<object> cells, double alpha = 0.05)
        {
            var x = cells.Where(IsNum).Select(Vba.CDbl).ToArray();
            int n = x.Length;
            if (n < 4) return "n<4，跳过";
            double s = MannKendall.S(x), z = MannKendall.Z(x), p = MannKendall.P(x), sen = MannKendall.SenSlope(x);
            string trend = p < alpha ? (z > 0 ? "▲ 显著上升" : "▼ 显著下降") : "— 无显著趋势";
            return $"n={n}\n  S={R(s, 1)}  Z={R(z, 4)}  p={R(p, 4)}\n  Sen斜率={R(sen, 6)}  {trend}";
        }
    }
}
