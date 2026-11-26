using System;
using System.Linq;

namespace YTools.Core
{
    /// <summary>与 Excel 工作表函数一致的基础统计。</summary>
    public static class Stats
    {
        public static double Average(double[] x)
        {
            if (x.Length == 0) throw new VbaRuntimeException("AVERAGE：没有数值（#DIV/0!）");
            double sum = 0;
            foreach (var v in x) sum += v;
            return sum / x.Length;
        }

        /// <summary>STDEV.S：样本标准差。</summary>
        public static double StDevS(double[] x)
        {
            if (x.Length < 2) throw new VbaRuntimeException("STDEV.S：至少需要 2 个数值（#DIV/0!）");
            double mean = Average(x), ss = 0;
            foreach (var v in x) ss += (v - mean) * (v - mean);
            return Math.Sqrt(ss / (x.Length - 1));
        }

        public static double Median(double[] x)
        {
            if (x.Length == 0) throw new VbaRuntimeException("MEDIAN：没有数值");
            var s = x.OrderBy(v => v).ToArray();
            int n = s.Length;
            return n % 2 == 1 ? s[n / 2] : (s[n / 2 - 1] + s[n / 2]) / 2.0;
        }

        /// <summary>T.INV.2T(p, df)：双尾 t 分布反函数。</summary>
        public static double TInv2T(double p, double df)
        {
            if (!(p > 0 && p <= 1) || df < 1) throw new VbaRuntimeException("T.INV.2T 参数无效");
            return MathNet.Numerics.Distributions.StudentT.InvCDF(0, 1, Math.Floor(df), 1 - p / 2);
        }

        /// <summary>NORM.S.DIST(z, TRUE)。</summary>
        public static double NormSDist(double z) => MathNet.Numerics.Distributions.Normal.CDF(0, 1, z);
    }
}
