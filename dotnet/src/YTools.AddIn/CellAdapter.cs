using System.Collections.Generic;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>Excel-DNA 参数/返回值 与 YTools.Core 约定值之间的转换。</summary>
    internal static class CellAdapter
    {
        /// <summary>单个值：ExcelEmpty/ExcelMissing → null，ExcelError → CellError。</summary>
        public static object In(object v)
        {
            switch (v)
            {
                case ExcelEmpty _:
                case ExcelMissing _:
                    return null;
                case ExcelError _:
                    return CellError.Value;
                default:
                    return v;
            }
        }

        /// <summary>区域参数统一成二维数组（单个单元格也包成 1×1）。</summary>
        public static object[,] Range(object v)
        {
            if (v is object[,] a)
            {
                var r = new object[a.GetLength(0), a.GetLength(1)];
                for (int i = 0; i < a.GetLength(0); i++)
                    for (int j = 0; j < a.GetLength(1); j++)
                        r[i, j] = In(a[i, j]);
                return r;
            }
            return new[,] { { In(v) } };
        }

        /// <summary>区域里的数值单元格（等同 WorksheetFunction.Average(Range) 的取值规则）。</summary>
        public static double[] Numbers(object v)
        {
            var list = new List<double>();
            if (v is object[,] a)
            {
                foreach (var x in a)
                    if (x is double d) list.Add(d);
            }
            else if (v is double d) list.Add(d);
            else if (v is ExcelError) throw new VbaRuntimeException("区域含错误值");
            return list.ToArray();
        }

        public static bool IsMissing(object v) => v is ExcelMissing || v is ExcelEmpty;

        public static double OptionalDouble(object v, double defaultValue)
            => IsMissing(v) ? defaultValue : Vba.CDbl(In(v));

        public static int OptionalInt(object v, int defaultValue)
            => IsMissing(v) ? defaultValue : (int)System.Math.Round(Vba.CDbl(In(v)), System.MidpointRounding.ToEven);

        public static object Out(object v)
        {
            switch (v)
            {
                case CellError _: return ExcelError.ExcelErrorValue;
                case null: return ExcelEmpty.Value;
                default: return v;
            }
        }
    }
}
