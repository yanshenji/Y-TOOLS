using System;

namespace YTools.Core
{
    /// <summary>按日期筛选的区间计算（原 模块1）。</summary>
    public static class DateFilters
    {
        /// <summary>
        /// 业务月：上月 28 日（不含）~ 本月 28 日（含），即 29 日到次月 28 日。
        /// monthOffset=0 为本月，-1 为上月。
        /// </summary>
        public static (DateTime StartExclusive, DateTime EndInclusive) BusinessMonth(DateTime today, int monthOffset = 0)
        {
            var anchor = new DateTime(today.Year, today.Month, 1).AddMonths(monthOffset);
            return (anchor.AddMonths(-1).AddDays(27), anchor.AddDays(27));
        }

        /// <summary>最大日期往前 30 天（含两端）。</summary>
        public static (DateTime Start, DateTime End) Last30Days(DateTime maxDate)
            => (maxDate.AddDays(-30), maxDate);

        /// <summary>自动筛选的日期条件用序列号表示，避免日期文本格式随系统区域变化而失效。</summary>
        public static string Criteria(string op, DateTime d)
            => op + d.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}
