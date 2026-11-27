using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 模块1（隐藏/显示部分）与 其他.隐藏所有工作表CE列。</summary>
    public static class ViewCommands
    {
        private const string Menu = "工具宏·显示";

        /// <summary>
        /// 隐藏第 2 行起全空的列。合并了原来的 隐藏符合条件的列、隐藏符合条件的列2、根据Ctrl向下查找隐藏列，
        /// 三者判定结果相同。
        /// </summary>
        [ExcelCommand(Name = "工具宏_隐藏空列", MenuName = Menu, MenuText = "隐藏空列（表头以下全空）")]
        public static void HideBlankColumns() => Cmd.Run("隐藏空列", () =>
        {
            var ws = Xl.RequireActiveSheet();
            (int Row, int Col) end = Xl.UsedEnd(ws);
            int lastRow = end.Row, lastCol = end.Col;
            if (lastRow < 2) return;
            object[,] block = Xl.Read(ws.Range[ws.Cells[1, 1], ws.Cells[lastRow, lastCol]]);
            var cols = SheetOps.BlankColumns(block, 1).Select(c => c + 1).ToList();
            Xl.HideColumns(ws, cols);
        });

        /// <summary>从选区第一列到第 1 行最后一列，只看选区所在的行，全空就隐藏该列。</summary>
        [ExcelCommand(Name = "工具宏_隐藏选区空列", MenuName = Menu, MenuText = "隐藏空列（只看选中的行）")]
        public static void HideBlankColumnsInSelection() => Cmd.Run("隐藏选区空列", () =>
        {
            dynamic sel = Xl.RequireRangeSelection();
            var ws = sel.Worksheet;
            int r1 = sel.Row, r2 = r1 + (int)sel.Rows.Count - 1;
            int c1 = sel.Column, c2 = System.Math.Max(Xl.LastCol(ws, 1), c1);
            object[,] block = Xl.Read(ws.Range[ws.Cells[r1, c1], ws.Cells[r2, c2]]);
            var cols = SheetOps.BlankColumns(block).Select(c => c + c1).ToList();
            dynamic app = Xl.App;
            int calc = app.Calculation;
            app.Calculation = Xl.xlCalculationManual;
            try { Xl.HideColumns(ws, cols); }
            finally { app.Calculation = calc; }
        });

        /// <summary>A 列最后一行以内，选中列全为空的行隐藏。</summary>
        [ExcelCommand(Name = "工具宏_隐藏选区空行", MenuName = Menu, MenuText = "隐藏空行（只看选中的列）")]
        public static void HideBlankRowsInSelection() => Cmd.Run("隐藏选区空行", () =>
        {
            dynamic sel = Xl.RequireRangeSelection();
            var ws = sel.Worksheet;
            int c1 = sel.Column, c2 = c1 + (int)sel.Columns.Count - 1;
            int last = Xl.LastRow(ws, 1);
            object[,] block = Xl.Read(ws.Range[ws.Cells[1, c1], ws.Cells[last, c2]]);
            var rows = SheetOps.BlankRows(block).Select(r => r + 1).ToList();
            Xl.HideRows(ws, rows);
        });

        [ExcelCommand(Name = "工具宏_取消隐藏列", MenuName = Menu, MenuText = "取消隐藏所有列")]
        public static void UnhideColumns() => Cmd.Run("取消隐藏列", () =>
            Xl.RequireActiveSheet().Cells.EntireColumn.Hidden = false);

        [ExcelCommand(Name = "工具宏_隐藏所有表CE列", MenuName = Menu, MenuText = "隐藏所有工作表 C~E 列")]
        public static void HideCEAllSheets() => Cmd.Run("隐藏C~E列", () =>
        {
            foreach (dynamic ws in Xl.App.ActiveWorkbook.Worksheets) ws.Columns["C:E"].Hidden = true;
            Ui.Info("已完成所有工作表C到E列的隐藏操作！");
        });
    }
}
