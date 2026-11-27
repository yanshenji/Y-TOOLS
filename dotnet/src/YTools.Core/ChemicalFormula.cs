using System.Collections.Generic;

namespace YTools.Core
{
    public enum ScriptKind { Subscript, Superscript }

    /// <summary>化学式上下标识别（原 其他.ImprovedChemicalFormatter）。</summary>
    public static class ChemicalFormula
    {
        /// <summary>
        /// 返回需要设置上下标的字符位置（从 1 开始，与 Excel Characters 一致）。
        /// 规则：字母后的数字为下标（如 SO4、H2O）；+/- 为上标，紧挨在它前面的电荷数也为上标（如 SO4 2-、Fe3+）。
        /// </summary>
        public static Dictionary<int, ScriptKind> Analyze(string text)
        {
            var result = new Dictionary<int, ScriptKind>();
            if (string.IsNullOrEmpty(text)) return result;
            int len = text.Length;
            char At(int pos) => text[pos - 1]; // 1 基

            // 下标：字母后面的数字，且后一位不是电荷符号
            for (int i = 2; i <= len; i++)
            {
                if (!IsDigit(At(i)) || !IsLetter(At(i - 1))) continue;
                if (i < len && IsCharge(At(i + 1))) continue;
                result[i] = ScriptKind.Subscript;
            }

            // 上标：电荷符号，以及其前面连续的电荷数
            for (int i = 1; i <= len; i++)
            {
                if (!IsCharge(At(i))) continue;
                result[i] = ScriptKind.Superscript;
                int j = i - 1;
                while (j >= 1 && IsDigit(At(j)))
                {
                    bool afterLetter = j > 1 && IsLetter(At(j - 1));
                    if (afterLetter)
                    {
                        // 原 VBA 在这里直接停止，Fe3+ 的 3 既不是上标也不是下标；
                        // 现在：紧挨电荷符号的这一位是电荷数（Fe3+、Ca2+），其余字母后的数字仍是下标
                        if (j == i - 1) result[j] = ScriptKind.Superscript;
                        break;
                    }
                    result[j] = ScriptKind.Superscript;
                    j--;
                }
            }
            return result;
        }

        private static bool IsDigit(char c) => c >= '0' && c <= '9';
        private static bool IsLetter(char c) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        private static bool IsCharge(char c) => c == '+' || c == '-';
    }
}
