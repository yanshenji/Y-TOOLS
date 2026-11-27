using System;
using System.Collections.Generic;
using System.Linq;

namespace YTools.Core
{
    /// <summary>选区内标注异常值的命令所用的判定（原 其他.GrubbsMarkYellow / FlagOutliersInSheet）。</summary>
    public static class OutlierMarks
    {
        /// <summary>
        /// 单次 Grubbs（GrubbsMarkYellow）：用全部有效数值的均值/标准差算每个点的 G，大于临界值即为异常。
        /// 返回与输入等长的标记；非数值为 null。有效数少于 3 或标准差≈0 时返回 null（不做标记）。
        /// </summary>
        public static bool?[] SinglePass(IList<object> cells, double alpha = 0.05)
        {
            var nums = cells.Where(c => !Vba.IsEmpty(c) && Vba.IsNumeric(c)).Select(Vba.CDbl).ToArray();
            if (nums.Length < 3) return null;
            double mean = Stats.Average(nums), sd = Stats.StDevS(nums);
            if (sd < 1e-12) return null;
            double crit = Grubbs.CriticalValue(nums.Length, alpha);
            return cells.Select(c => !Vba.IsEmpty(c) && Vba.IsNumeric(c)
                ? (bool?)(Math.Abs(Vba.CDbl(c) - mean) / sd > crit)
                : null).ToArray();
        }

        /// <summary>
        /// 迭代 Grubbs（FlagOutliersInSheet）：先跑 GrubbsOutliers 得到异常值，再逐行比对值。
        /// 返回每行的判定文字；非数值行为 null（原 VBA 遇到文本行会报错中断）。
        /// </summary>
        public static string[] Iterative(IList<object> cells, double alpha, out int outlierCount)
        {
            var nums = cells.Where(c => !Vba.IsEmpty(c) && Vba.IsNumeric(c)).Select(Vba.CDbl).ToList();
            var outliers = nums.Count >= 3 ? Grubbs.DetectOutliers(nums, alpha) : new List<double>();
            outlierCount = outliers.Count;
            var set = new HashSet<double>(outliers);
            return cells.Select(c =>
            {
                if (Vba.IsEmpty(c)) return "正常"; // 原 VBA：空单元格按 0 处理
                if (!Vba.IsNumeric(c)) return null;
                return set.Contains(Vba.CDbl(c)) ? "异常值" : "正常";
            }).ToArray();
        }
    }

    /// <summary>标准校验着色（原 函数.cmpStd222）的单元格判定结果。</summary>
    public enum StdMark
    {
        /// <summary>不着色</summary>
        None,
        /// <summary>数值不是数字：红色填充</summary>
        NotNumeric,
        /// <summary>IsNull 规则但有值：橙红填充 RGB(243,84,26)</summary>
        IsNullViolated,
        /// <summary>参考2 项未填：蓝色填充 RGB(37,109,183)</summary>
        Cankao2Missing,
        /// <summary>NotNull 必填项未填：橙色填充 RGB(255,136,16)</summary>
        NotNullMissing,
        /// <summary>普通项目未检测：绿色填充 RGB(169,208,142)</summary>
        Untested,
        /// <summary>超出标准：红色字体</summary>
        OutOfSpec,
    }

    public static class StandardMarks
    {
        /// <summary>单格判定，规则与 cmpStd2 相同，但细分出“未检测/必填未填”等状态用于着色。</summary>
        public static StdMark Classify(object stdCell, object valueCell)
        {
            string std = stdCell is CellError ? "" : Vba.Trim(Vba.CStr(stdCell));
            bool isEmpty = Vba.IsEmpty(valueCell);
            double value = 0;
            if (!isEmpty && !Vba.TryCDbl(valueCell, out value)) return StdMark.NotNumeric;

            if (std.Contains("IsNull"))
                return value != 0 && !isEmpty ? StdMark.IsNullViolated : StdMark.None;

            bool isCankao2 = std.Contains("参考2");
            bool isAny = std.Contains("抽检") || std.Contains("自检");
            bool isValid = (std != "" && std != "/" && std != "-"
                            && !std.Contains("实测") && !std.Contains("监测") && !std.Contains("参考"))
                           || isCankao2;
            if (!isValid) return StdMark.None;

            if (value == 0 || isEmpty)
            {
                if (isAny) return StdMark.None;
                if (isCankao2) return StdMark.Cankao2Missing;
                if (std.Contains("NotNull")) return StdMark.NotNullMissing;
                return StdMark.Untested;
            }

            string cleaned = std;
            foreach (var t in new[] { "NotNull", "抽检", "自检", "参考2", "g/L", "N", "%", "us/cm2", "(", ")", "（", "）" })
                cleaned = cleaned.Replace(t, "");
            cleaned = Vba.Trim(cleaned);
            if (cleaned == "" || cleaned.Length <= 1) return StdMark.None;
            // 原 VBA：参考2 项超标也标红字（与 cmpStd2 不同，cmpStd2 不计入不合格）
            try { return StandardCheck.Compare(cleaned, value) ? StdMark.None : StdMark.OutOfSpec; }
            catch (VbaRuntimeException) { return StdMark.NotNumeric; } // 原 VBA：标准写错时走错误分支，填红色
        }
    }
}
