using System;

namespace YTools.Core
{
    /// <summary>过程能力：Cpk（移动极差法）、Ppk、Cpk 等级。移植自 模块7。</summary>
    public static class Capability
    {
        public const string InsufficientSample = "样本不足";
        public const string NoVariation = "σ=0无波动";

        /// <summary>d2 常数（n=2 的移动极差）。</summary>
        public const double D2 = 1.128;

        /// <summary>原 VBA 中 Cpk 等级文字里的分隔符（源码中该字符已乱码为“?”）。</summary>
        public const string RankSeparator = "-";

        /// <summary>=Cpk_MR(数据, LSL, USL)：单值移动极差法 Cpk，只取区域第一列。</summary>
        public static object CpkMR(object[,] data, object lsl, object usl)
            => CpkMRFromValues(Vba.NumericFirstColumn(data), lsl, usl);

        public static object CpkMRFromValues(double[] vals, object lsl, object usl)
        {
            int n = vals.Length;
            if (n < 2) return InsufficientSample;

            double xBar = Stats.Average(vals);
            var mr = new double[n - 1];
            for (int i = 0; i < n - 1; i++) mr[i] = Math.Abs(vals[i + 1] - vals[i]);
            double sigmaWithin = Stats.Average(mr) / D2;
            if (sigmaWithin <= 0) return NoVariation;

            return MinOfSides(xBar, sigmaWithin, lsl, usl);
        }

        /// <summary>=Ppk_Calc(数据, LSL, USL)：整体标准差 Ppk，只取区域第一列。</summary>
        public static object Ppk(object[,] data, object lsl, object usl)
        {
            var vals = Vba.NumericFirstColumn(data);
            if (vals.Length < 2) return InsufficientSample;

            double xBar = Stats.Average(vals);
            double sigmaTotal = Stats.StDevS(vals);
            if (sigmaTotal <= 0) return NoVariation;

            return MinOfSides(xBar, sigmaTotal, lsl, usl);
        }

        // 单侧规格缺失时该侧按 999 处理（与原 VBA 一致）
        private static double MinOfSides(double xBar, double sigma, object lsl, object usl)
        {
            double upper = 999, lower = 999;
            if (!Vba.IsEmpty(usl) && Vba.IsNumeric(usl)) upper = (Vba.CDbl(usl) - xBar) / (3 * sigma);
            if (!Vba.IsEmpty(lsl) && Vba.IsNumeric(lsl)) lower = (xBar - Vba.CDbl(lsl)) / (3 * sigma);
            return Math.Min(upper, lower);
        }

        /// <summary>=Cpk_Rank(Cpk值)：非数值原样返回（例如“样本不足”）。</summary>
        public static object Rank(object cpk)
        {
            if (cpk is CellError) return CellError.Value;
            if (!Vba.IsNumeric(cpk)) return cpk;
            double v = Vba.CDbl(cpk);
            string sep = RankSeparator;
            if (v < 0.67) return "D级" + sep + "严重不足";
            if (v < 1) return "C级" + sep + "能力不足";
            if (v < 1.33) return "B级" + sep + "临界尚可";
            if (v < 1.67) return "A级" + sep + "能力充足";
            return "AA级" + sep + "能力富余";
        }
    }
}
