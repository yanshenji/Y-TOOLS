using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 其他.GrubbsMarkYellow / FlagOutliersInSheet。</summary>
    public static class OutlierCommands
    {
        private const string Menu = "工具宏·异常值";

        [ExcelCommand(Name = "工具宏_Grubbs标黄", MenuName = Menu, MenuText = "Grubbs 异常值标黄（选区）")]
        public static void MarkYellow() => Cmd.Run("Grubbs 标黄", () =>
        {
            dynamic sel = Xl.RequireRangeSelection();
            if ((int)sel.Areas.Count > 1) throw new UserMessage("请选择连续一块数据区域！");
            sel = Xl.TrimToUsed(sel);
            if (sel == null) return;
            object[,] arr = Xl.Read(sel);
            var cells = arr.Cast<object>().ToList();
            var marks = OutlierMarks.SinglePass(cells);
            if (marks == null) throw new UserMessage("有效样本数必须≥3，且数据不能全部相同。");
            int i = 0, n = 0;
            foreach (dynamic cell in sel.Cells)
            {
                if (marks[i] == true) { cell.Interior.Color = Xl.Rgb(255, 255, 0); n++; }
                else cell.Interior.ColorIndex = Xl.xlColorIndexNone;
                i++;
            }
            Ui.Info($"检出 {n} 个异常值，已标黄。");
        });

        [ExcelCommand(Name = "工具宏_Grubbs标注A列", MenuName = Menu, MenuText = "Grubbs 判定 A 列 → 写入 B 列")]
        public static void FlagColumnA() => Cmd.Run("Grubbs 标注", () =>
        {
            var ws = Xl.RequireActiveSheet();
            int last = Xl.LastRow(ws, 1);
            if (last < 3) throw new UserMessage("A 列数据不足 3 行");
            object[,] arr = Xl.Read(ws.Range["A1:A" + last]);
            var cells = arr.Cast<object>().ToList();
            var flags = OutlierMarks.Iterative(cells, 0.05, out int count);
            var outArr = new object[last, 1];
            for (int r = 0; r < last; r++) outArr[r, 0] = flags[r];
            if (flags[0] == null) outArr[0, 0] = "Grubbs 判定";
            Xl.Write(ws.Range["B1"], outArr);
            Ui.Info("标注完成，共 " + count + " 个异常值");
        });
    }
}
