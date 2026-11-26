using System;

namespace YTools.Core
{
    /// <summary>标准判定：cmpStd / cmpStd2。移植自 函数.bas，规则与原 VBA 逐条对应。</summary>
    public static class StandardCheck
    {
        public const string ArgumentError = "参数错误";

        /// <summary>
        /// =cmpStd(标准表达式, 数值)。支持 ≥ ≤ ＞ ＜ &gt;= &lt;= &gt; &lt;、a-b、a~b、单值相等。
        /// 无法解析的比较式抛出 VbaRuntimeException（单元格显示 #VALUE!）；无法解析的单值返回 false。
        /// </summary>
        public static bool Compare(string criteriaExpression, double value)
        {
            string expr = Vba.Trim(criteriaExpression ?? "");
            expr = expr.Replace("≥", ">=").Replace("≤", "<=").Replace("＞", ">").Replace("＜", "<");

            if (expr.StartsWith(">=", StringComparison.Ordinal)) return value >= Num(expr.Replace(">=", ""));
            if (expr.StartsWith("<=", StringComparison.Ordinal)) return value <= Num(expr.Replace("<=", ""));
            if (expr.StartsWith(">", StringComparison.Ordinal)) return value > Num(expr.Replace(">", ""));
            if (expr.StartsWith("<", StringComparison.Ordinal)) return value < Num(expr.Replace("<", ""));

            if (expr.IndexOf('-') >= 0 && !expr.StartsWith("-", StringComparison.Ordinal))
                return InRange(expr.Split('-'), value);
            if (expr.IndexOf('~') >= 0)
                return InRange(expr.Split('~'), value);

            return Vba.TryParseDouble(expr, out var c) && value == c;
        }

        private static bool InRange(string[] parts, double value)
        {
            if (parts.Length < 2) throw new VbaRuntimeException("下标越界");
            double lo = Num(Vba.Trim(parts[0])), hi = Num(Vba.Trim(parts[1]));
            return value >= lo && value <= hi;
        }

        private static double Num(string s)
        {
            if (Vba.TryParseDouble(s, out var d)) return d;
            throw new VbaRuntimeException($"无法解析标准值：{s}");
        }

        /// <summary>
        /// =cmpStd2(标准区域, 数值区域, [返回类型])。逐格（按行优先）比对，全部符合才算合格。
        /// 返回类型：1=TRUE/FALSE 2=合格/不合格 3=符合/不符合 4=1/2。
        /// </summary>
        public static object CompareRanges(object[,] stdRange, object[,] valueRange, int returnType = 1)
        {
            var stds = Flatten(stdRange);
            var vals = Flatten(valueRange);
            int count = Math.Min(stds.Length, vals.Length);

            bool isPass = true;
            // 原 VBA 的错误处理在“数值转换失败”后会停留在 On Error Resume Next，
            // 直到下一个非空数值单元格转换成功。此期间发生的错误不会变成“参数错误”。
            bool resumeNext = false;
            string std = "";

            for (int i = 0; i < count; i++)
            {
                try { std = Vba.Trim(Vba.CStr(stds[i])); }
                catch (VbaRuntimeException)
                {
                    if (!resumeNext) return ArgumentError;
                    // Resume Next：std 保持上一格的值
                }

                object raw = vals[i];
                bool isEmptyCell = Vba.IsEmpty(raw);
                double value = 0;
                if (!isEmptyCell)
                {
                    if (!Vba.TryCDbl(raw, out value))
                    {
                        isPass = false;
                        resumeNext = true;
                        continue;
                    }
                    resumeNext = false;
                }

                if (std.Contains("IsNull"))
                {
                    if (value != 0 && !isEmptyCell) isPass = false;
                    continue;
                }

                bool isCankao2 = std.Contains("参考2");
                bool isValid = (std != "" && std != "/" && std != "-"
                                && !std.Contains("实测") && !std.Contains("监测") && !std.Contains("参考"))
                               || isCankao2;
                if (!isValid) continue;
                if (value == 0 || isEmptyCell) continue;

                string cleaned = Clean(std);
                if (cleaned == "" || cleaned.Length <= 1) continue;

                bool ok;
                try { ok = Compare(cleaned, value); }
                catch (VbaRuntimeException)
                {
                    // 走到这里时数值转换一定刚成功过，错误处理已恢复为 GoTo ex
                    return ArgumentError;
                }
                if (!ok && !isCankao2) isPass = false;
            }

            switch (returnType)
            {
                case 2: return isPass ? "合格" : "不合格";
                case 3: return isPass ? "符合" : "不符合";
                case 4: return isPass ? 1.0 : 2.0;
                default: return isPass;
            }
        }

        private static string Clean(string std)
        {
            foreach (var token in new[] { "NotNull", "抽检", "自检", "参考2", "g/L", "N", "%", "us/cm2", "(", ")", "（", "）" })
                std = std.Replace(token, "");
            return Vba.Trim(std);
        }

        /// <summary>按行优先展开，对应 VBA Range.Cells(i) 的顺序。</summary>
        private static object[] Flatten(object[,] a)
        {
            int r = a.GetLength(0), c = a.GetLength(1);
            var res = new object[r * c];
            for (int i = 0; i < r; i++)
                for (int j = 0; j < c; j++)
                    res[i * c + j] = a[i, j];
            return res;
        }
    }
}
