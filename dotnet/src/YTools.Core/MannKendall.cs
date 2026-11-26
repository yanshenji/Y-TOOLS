using System;
using System.Collections.Generic;

namespace YTools.Core
{
    /// <summary>Mann-Kendall 趋势检验与 Sen 斜率。移植自 模块8。</summary>
    public static class MannKendall
    {
        public enum ResultType { S = 1, Z = 2, P = 3, SenSlope = 4, N = 5 }

        /// <summary>
        /// =MK(区域, 返回类型)：1=S 2=Z 3=P 值 4=Sen 斜率 5=样本数。只取区域第一列；数据少于 4 个返回 #VALUE!。
        /// </summary>
        public static object Compute(object data, int retType)
        {
            double[] arr;
            if (data is object[,] range)
            {
                arr = Vba.NumericFirstColumn(range);
                if (arr.Length == 0) return CellError.Value; // 原 VBA：ReDim Preserve arr(1 To 0) 报错
            }
            else
            {
                arr = new[] { Vba.CDbl(data) };
            }

            int n = arr.Length;
            if (retType == (int)ResultType.N) return (double)n;
            if (n < 4) return CellError.Value;

            switch ((ResultType)retType)
            {
                case ResultType.S: return S(arr);
                case ResultType.Z: return Z(arr);
                case ResultType.P: return P(arr);
                case ResultType.SenSlope: return SenSlope(arr);
                default: return 0.0; // 原 VBA 返回 Empty，单元格显示 0
            }
        }

        public static double S(double[] x)
        {
            double s = 0;
            for (int i = 0; i < x.Length - 1; i++)
                for (int j = i + 1; j < x.Length; j++)
                {
                    if (x[j] > x[i]) s++;
                    else if (x[j] < x[i]) s--;
                }
            return s;
        }

        public static double VarS(double[] x)
        {
            int n = x.Length;
            var sorted = (double[])x.Clone();
            Array.Sort(sorted);
            double tieTerm = 0;
            int k = 0;
            while (k < n - 1)
            {
                if (sorted[k] == sorted[k + 1])
                {
                    int t = 1;
                    while (k + t < n && sorted[k + t] == sorted[k]) t++;
                    tieTerm += (double)t * (t - 1) * (2 * t + 5);
                    k += t;
                }
                else k++;
            }
            return ((double)n * (n - 1) * (2 * n + 5) - tieTerm) / 18;
        }

        public static double Z(double[] x)
        {
            double s = S(x);
            if (s == 0) return 0;
            double sd = Math.Sqrt(VarS(x));
            return s > 0 ? (s - 1) / sd : (s + 1) / sd;
        }

        public static double P(double[] x) => 2 * (1 - Stats.NormSDist(Math.Abs(Z(x))));

        /// <summary>Sen 斜率：所有点对 (xj-xi)/(j-i) 的中位数（按序号等间隔）。</summary>
        public static double SenSlope(double[] x)
        {
            var slopes = new List<double>(x.Length * (x.Length - 1) / 2);
            for (int i = 0; i < x.Length - 1; i++)
                for (int j = i + 1; j < x.Length; j++)
                    slopes.Add((x[j] - x[i]) / (j - i));
            return Stats.Median(slopes.ToArray());
        }
    }
}
