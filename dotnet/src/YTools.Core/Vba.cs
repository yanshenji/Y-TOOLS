using System;
using System.Globalization;

namespace YTools.Core
{
    /// <summary>单元格错误值（#VALUE! 等）的占位，由加载项层映射到 Excel 错误。</summary>
    public sealed class CellError
    {
        public static readonly CellError Value = new CellError("#VALUE!");
        public string Text { get; }
        private CellError(string text) { Text = text; }
        public override string ToString() => Text;
    }

    /// <summary>对应 VBA 运行时错误：在 UDF 中表现为 #VALUE!。</summary>
    public sealed class VbaRuntimeException : Exception
    {
        public VbaRuntimeException(string message) : base(message) { }
    }

    /// <summary>
    /// 复刻 VBA 的类型转换规则，保证迁移后结果与原宏一致。
    /// 单元格值约定：double、string、bool、null（空单元格）、CellError。
    /// </summary>
    public static class Vba
    {
        private const NumberStyles CDblStyles = NumberStyles.Float | NumberStyles.AllowThousands;

        /// <summary>VBA CDbl(字符串)：去首尾空格，允许千分位与科学计数法。</summary>
        public static bool TryParseDouble(string s, out double result)
        {
            result = 0;
            if (s == null) return false;
            var t = s.Trim(' ');
            if (t.Length == 0) return false;
            return double.TryParse(t, CDblStyles, CultureInfo.InvariantCulture, out result);
        }

        /// <summary>VBA IsNumeric(v)。注意：空单元格在 VBA 里也算数值，调用方需要另外排除。</summary>
        public static bool IsNumeric(object v)
        {
            switch (v)
            {
                case null: return true;
                case double _: return true;
                case bool _: return true;
                case string s: return TryParseDouble(s, out _);
                default: return false;
            }
        }

        public static bool IsEmpty(object v) => v == null;

        /// <summary>VBA CDbl(v)。失败时抛出 VbaRuntimeException（类型不匹配）。</summary>
        public static double CDbl(object v)
        {
            if (TryCDbl(v, out var d)) return d;
            throw new VbaRuntimeException("类型不匹配");
        }

        public static bool TryCDbl(object v, out double d)
        {
            switch (v)
            {
                case null: d = 0; return true;
                case double x: d = x; return true;
                case bool b: d = b ? -1 : 0; return true;
                case string s: return TryParseDouble(s, out d);
                default: d = 0; return false;
            }
        }

        /// <summary>VBA CStr(v)，用于把单元格值当作文本处理。</summary>
        public static string CStr(object v)
        {
            switch (v)
            {
                case null: return "";
                case string s: return s;
                case bool b: return b ? "True" : "False";
                case double x: return x.ToString("G15", CultureInfo.InvariantCulture);
                default: throw new VbaRuntimeException("类型不匹配");
            }
        }

        /// <summary>VBA Trim：只去首尾半角空格。</summary>
        public static string Trim(string s) => s.Trim(' ');

        /// <summary>取二维数组第一列的有效数值（VBA：IsNumeric 且非空）。</summary>
        public static double[] NumericFirstColumn(object[,] values)
        {
            int rows = values.GetLength(0);
            var list = new System.Collections.Generic.List<double>(rows);
            for (int i = 0; i < rows; i++)
            {
                var v = values[i, 0];
                if (!IsEmpty(v) && IsNumeric(v)) list.Add(CDbl(v));
            }
            return list.ToArray();
        }
    }
}
