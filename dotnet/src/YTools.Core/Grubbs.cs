using System;
using System.Collections.Generic;

namespace YTools.Core
{
    /// <summary>Grubbs 异常值检验（双侧）。移植自 其他.bas。</summary>
    public static class Grubbs
    {
        public const string TooFewPoints = "错误：至少需要 3 个数据点";
        public const string NoOutliers = "未检测到异常值";

        /// <summary>
        /// =GrubbsCrit(n, α)：临界值 G = (n-1)/√n · √(t² / (n-2+t²))，t = T.INV.2T(α/n, n-2)。
        /// </summary>
        public static object Critical(int n, double alpha = 0.05)
        {
            if (n <= 2) return "n 必须大于 2";
            return CriticalValue(n, alpha);
        }

        public static double CriticalValue(int n, double alpha)
        {
            double t = Stats.TInv2T(alpha / n, n - 2);
            return (n - 1) / Math.Sqrt(n) * Math.Sqrt(t * t / (n - 2 + t * t));
        }

        /// <summary>=GrubbsG(值, 区域)：单点 G 值，均值与标准差按区域中的全部数值计算。</summary>
        public static object G(double singleValue, double[] rangeNumbers)
        {
            double mean = Stats.Average(rangeNumbers);
            double sd = Stats.StDevS(rangeNumbers);
            if (sd == 0) return "标准差为 0";
            return Math.Abs(singleValue - mean) / sd;
        }

        /// <summary>
        /// =GrubbsOutliers(区域, α)：迭代剔除最极端值，直到不显著，返回剔除出的异常值（一列）。
        /// 与原 VBA 一致：剩余点数 ≤3 时停止检验。
        /// </summary>
        public static object Outliers(object[,] data, double alpha = 0.05)
        {
            var values = new List<double>(Vba.NumericFirstColumn(data));
            if (values.Count < 3) return TooFewPoints;

            var found = DetectOutliers(values, alpha);
            if (found.Count == 0) return NoOutliers;

            var res = new object[found.Count, 1];
            for (int i = 0; i < found.Count; i++) res[i, 0] = found[i];
            return res;
        }

        public static List<double> DetectOutliers(List<double> values, double alpha)
        {
            var data = new List<double>(values);
            var outliers = new List<double>();
            while (data.Count > 3)
            {
                var arr = data.ToArray();
                double mean = Stats.Average(arr);
                double sd = Stats.StDevS(arr);
                if (sd == 0) break;

                double gMax = 0;
                int maxIdx = -1;
                for (int i = 0; i < arr.Length; i++)
                {
                    double g = Math.Abs(arr[i] - mean) / sd;
                    if (g > gMax) { gMax = g; maxIdx = i; }
                }

                if (gMax > CriticalValue(arr.Length, alpha))
                {
                    outliers.Add(arr[maxIdx]);
                    data.RemoveAt(maxIdx);
                }
                else break;
            }
            return outliers;
        }
    }
}
