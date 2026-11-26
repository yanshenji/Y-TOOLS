using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>
    /// 工作表函数（UDF）。函数名与原 VBA 保持一致，已有工作簿中的公式无需修改。
    /// 计算逻辑全部在 YTools.Core 中，这里只做参数转换。
    /// </summary>
    public static class Functions
    {
        private const string Category = "工具宏";

        // ───────────── 标准判定（原 函数.bas） ─────────────

        [ExcelFunction(Name = "cmpStd", Category = Category,
            Description = "判断数值是否符合标准表达式，如 ≥10、≤0.5、5~10、5-10")]
        public static bool CmpStd(
            [ExcelArgument(Name = "标准表达式")] string criteriaExpression,
            [ExcelArgument(Name = "数值")] double value)
            => StandardCheck.Compare(criteriaExpression, value);

        [ExcelFunction(Name = "cmpStd2", Category = Category,
            Description = "逐格比对标准区域与数值区域，全部符合才合格")]
        public static object CmpStd2(
            [ExcelArgument(Name = "标准区域")] object stdRange,
            [ExcelArgument(Name = "数值区域")] object valueRange,
            [ExcelArgument(Name = "返回类型", Description = "1=TRUE/FALSE 2=合格/不合格 3=符合/不符合 4=1/2，默认 1")] object returnType)
            => CellAdapter.Out(StandardCheck.CompareRanges(
                CellAdapter.Range(stdRange), CellAdapter.Range(valueRange), CellAdapter.OptionalInt(returnType, 1)));

        // ───────────── 过程能力（原 模块7） ─────────────

        [ExcelFunction(Name = "Cpk_MR", Category = Category, Description = "单值移动极差法 Cpk（取区域第一列）")]
        public static object CpkMR(
            [ExcelArgument(Name = "数据区域")] object data,
            [ExcelArgument(Name = "LSL", Description = "规格下限，可留空")] object lsl,
            [ExcelArgument(Name = "USL", Description = "规格上限，可留空")] object usl)
            => CellAdapter.Out(Capability.CpkMR(CellAdapter.Range(data), CellAdapter.In(lsl), CellAdapter.In(usl)));

        [ExcelFunction(Name = "Ppk_Calc", Category = Category, Description = "整体标准差 Ppk（取区域第一列）")]
        public static object PpkCalc(
            [ExcelArgument(Name = "数据区域")] object data,
            [ExcelArgument(Name = "LSL")] object lsl,
            [ExcelArgument(Name = "USL")] object usl)
            => CellAdapter.Out(Capability.Ppk(CellAdapter.Range(data), CellAdapter.In(lsl), CellAdapter.In(usl)));

        [ExcelFunction(Name = "Cpk_Rank", Category = Category, Description = "Cpk 等级：D/C/B/A/AA")]
        public static object CpkRank([ExcelArgument(Name = "Cpk值")] object cpk)
            => CellAdapter.Out(Capability.Rank(CellAdapter.In(cpk)));

        // ───────────── 趋势检验（原 模块8） ─────────────

        [ExcelFunction(Name = "MK", Category = Category, Description = "Mann-Kendall 趋势检验（取区域第一列）")]
        public static object MK(
            [ExcelArgument(Name = "数据区域")] object data,
            [ExcelArgument(Name = "返回类型", Description = "1=S 2=Z 3=P值 4=Sen斜率 5=样本数")] int retType)
        {
            object arg = data is object[,] ? (object)CellAdapter.Range(data) : CellAdapter.In(data);
            return CellAdapter.Out(MannKendall.Compute(arg, retType));
        }

        // ───────────── Grubbs 异常值（原 其他.bas） ─────────────

        [ExcelFunction(Name = "GrubbsOutliers", Category = Category, Description = "Grubbs 迭代检验，返回检出的异常值（一列）")]
        public static object GrubbsOutliers(
            [ExcelArgument(Name = "数据区域")] object data,
            [ExcelArgument(Name = "alpha", Description = "显著性水平，默认 0.05")] object alpha)
            => CellAdapter.Out(Grubbs.Outliers(CellAdapter.Range(data), CellAdapter.OptionalDouble(alpha, 0.05)));

        [ExcelFunction(Name = "GrubbsG", Category = Category, Description = "单个数据点的 Grubbs G 值")]
        public static object GrubbsG(
            [ExcelArgument(Name = "值")] double singleValue,
            [ExcelArgument(Name = "数据区域")] object data)
            => CellAdapter.Out(Grubbs.G(singleValue, CellAdapter.Numbers(data)));

        [ExcelFunction(Name = "GrubbsCrit", Category = Category, Description = "Grubbs 双侧临界值")]
        public static object GrubbsCrit(
            [ExcelArgument(Name = "n")] int n,
            [ExcelArgument(Name = "alpha", Description = "默认 0.05")] object alpha)
            => CellAdapter.Out(Grubbs.Critical(n, CellAdapter.OptionalDouble(alpha, 0.05)));

        // ───────────── 有效数字修约（原 其他.bas） ─────────────

        [ExcelFunction(Name = "保留两位有效数字_四舍六入五成双", Category = Category,
            Description = "按 GB/T 8170 修约到两位有效数字")]
        public static double SigFig2HalfEven([ExcelArgument(Name = "数值")] double num)
            => SignificantFigures.RoundHalfEven(num);

        [ExcelFunction(Name = "保留两位有效数字_四舍五入", Category = Category,
            Description = "四舍五入到两位有效数字")]
        public static double SigFig2HalfUp([ExcelArgument(Name = "数值")] double num)
            => SignificantFigures.RoundHalfUp(num);

        // ───────────── 诊断 ─────────────

        [ExcelFunction(Name = "工具宏版本", Category = Category, Description = "返回加载项版本，用于确认 .xll 已加载")]
        public static string Version() => Diagnostics.VersionText;
    }
}
