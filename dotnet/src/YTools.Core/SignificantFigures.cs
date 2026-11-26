using System;

namespace YTools.Core
{
    /// <summary>
    /// 保留两位有效数字。按十进制（decimal）计算，避免 2.35 这类数在二进制浮点里存不精确导致的误判。
    /// 与原 VBA 版的差异见 docs/migration/excel-dna-inventory.md §8。
    /// </summary>
    public static class SignificantFigures
    {
        /// <summary>真正的四舍五入：舍弃部分 ≥5 即进一（对应原 保留两位有效数字_四舍五入）。</summary>
        public static double RoundHalfUp(double num, int digits = 2) => Round(num, digits, MidpointRounding.AwayFromZero);

        /// <summary>
        /// GB/T 8170 四舍六入五成双：舍弃部分 &lt;5 舍、&gt;5 进；恰为 5（其后全为 0）时，前一位奇进偶舍。
        /// </summary>
        public static double RoundHalfEven(double num, int digits = 2) => Round(num, digits, MidpointRounding.ToEven);

        private static double Round(double num, int digits, MidpointRounding mode)
        {
            if (num == 0 || double.IsNaN(num) || double.IsInfinity(num)) return num;
            if (digits < 1) throw new ArgumentOutOfRangeException(nameof(digits));

            decimal d;
            try { d = (decimal)num; } // double→decimal 保留 15 位有效数字，即用户看到的十进制值
            catch (OverflowException) { return RoundDouble(num, digits, mode); }
            if (d == 0) return RoundDouble(num, digits, mode); // 小于 1e-28，超出 decimal 精度

            int e = Exponent(d);          // 10^e ≤ |d| < 10^(e+1)
            int decimals = digits - 1 - e; // 需要保留的小数位数
            decimal r;
            if (decimals >= 0)
            {
                if (decimals > 28) return RoundDouble(num, digits, mode);
                r = Math.Round(d, decimals, mode);
            }
            else
            {
                decimal p = Pow10(-decimals);
                r = Math.Round(d / p, 0, mode) * p;
            }
            return (double)r;
        }

        private static int Exponent(decimal d)
        {
            d = Math.Abs(d);
            int e = 0;
            while (d >= 10m) { d /= 10m; e++; }
            while (d < 1m) { d *= 10m; e--; }
            return e;
        }

        private static decimal Pow10(int n)
        {
            decimal p = 1m;
            for (int i = 0; i < n; i++) p *= 10m;
            return p;
        }

        // 超出 decimal 范围时的兜底（实际业务数据不会走到这里）
        private static double RoundDouble(double num, int digits, MidpointRounding mode)
        {
            double e = Math.Floor(Math.Log10(Math.Abs(num)));
            double scale = Math.Pow(10, digits - 1 - e);
            return Math.Round(num * scale, mode) / scale;
        }
    }
}
