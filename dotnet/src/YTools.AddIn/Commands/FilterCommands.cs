using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 模块1（筛选部分）与 模块4（样品类别轮换筛选）。</summary>
    public static class FilterCommands
    {
        private const string Menu = "工具宏·筛选";
        private const string FilterRange = "$A:$ZZ";
        private static string Num(double d) => d.ToString("R", CultureInfo.InvariantCulture);

        private static void AutoFilter(dynamic ws, int field, string c1 = null, string c2 = null)
        {
            dynamic r = ws.Range[FilterRange];
            if (c1 == null) r.AutoFilter(Field: field);
            else if (c2 == null) r.AutoFilter(Field: field, Criteria1: c1);
            else r.AutoFilter(Field: field, Criteria1: c1, Operator: Xl.xlAnd, Criteria2: c2);
        }

        [ExcelCommand(Name = "工具宏_筛选当前值", MenuName = Menu, MenuText = "筛选为当前单元格值")]
        public static void ByCurrentValue() => Cmd.Run("筛选当前值", () =>
        {
            var ws = Xl.RequireActiveSheet();
            dynamic cell = Xl.App.ActiveCell;
            // 用显示文本做条件：日期、带格式的数字也能正确匹配（原 VBA 用 .Value 拼接，日期常常筛不出）
            AutoFilter(ws, (int)cell.Column, "=" + (string)cell.Text);
        });

        [ExcelCommand(Name = "工具宏_筛选非空", MenuName = Menu, MenuText = "筛选当前列非空值")]
        public static void NonBlank() => Cmd.Run("筛选非空", () =>
            AutoFilter(Xl.RequireActiveSheet(), (int)Xl.App.ActiveCell.Column, "<>"));

        [ExcelCommand(Name = "工具宏_筛选数值", MenuName = Menu, MenuText = "筛选当前列数值（>0）")]
        public static void Numeric() => Cmd.Run("筛选数值", () =>
        {
            var ws = Xl.RequireActiveSheet();
            int col = (int)Xl.App.ActiveCell.Column;
            int last = Xl.LastRow(ws, col);
            double max = Xl.App.WorksheetFunction.Max(ws.Range[ws.Cells[1, col], ws.Cells[last, col]]);
            AutoFilter(ws, col, ">0", "<=" + Num(max));
        });

        [ExcelCommand(Name = "工具宏_筛选近30天", MenuName = Menu, MenuText = "筛选最大日期近30天")]
        public static void Last30Days() => Cmd.Run("筛选近30天", () =>
        {
            var ws = Xl.RequireActiveSheet();
            int last = Xl.LastRow(ws, 1);
            if (last < 2) throw new UserMessage("A 列没有数据。");
            double max = Xl.App.WorksheetFunction.Max(ws.Range["A2:A" + last]);
            if (max <= 0) throw new UserMessage("A 列没有日期。");
            var (s, e) = DateFilters.Last30Days(DateTime.FromOADate(max));
            AutoFilter(ws, 1, DateFilters.Criteria(">=", s), DateFilters.Criteria("<=", e));
        });

        [ExcelCommand(Name = "工具宏_筛选本月", MenuName = Menu, MenuText = "筛选本月（上月29日~本月28日）")]
        public static void ThisMonth() => Cmd.Run("筛选本月", () => Month(0));

        [ExcelCommand(Name = "工具宏_筛选上月", MenuName = Menu, MenuText = "筛选上月")]
        public static void LastMonth() => Cmd.Run("筛选上月", () => Month(-1));

        private static void Month(int offset)
        {
            var (s, e) = DateFilters.BusinessMonth(DateTime.Today, offset);
            AutoFilter(Xl.RequireActiveSheet(), 1, DateFilters.Criteria(">", s), DateFilters.Criteria("<=", e));
        }

        [ExcelCommand(Name = "工具宏_清除当前列筛选", MenuName = Menu, MenuText = "清除当前列筛选")]
        public static void ClearColumn() => Cmd.Run("清除当前列筛选", () =>
            AutoFilter(Xl.RequireActiveSheet(), (int)Xl.App.ActiveCell.Column));

        [ExcelCommand(Name = "工具宏_清除筛选", MenuName = Menu, MenuText = "清除全部筛选")]
        public static void ClearAll() => Cmd.Run("清除筛选", () =>
        {
            var ws = Xl.RequireActiveSheet();
            if ((bool)ws.AutoFilterMode) ws.AutoFilterMode = false;
        });

        [ExcelCommand(Name = "工具宏_保留最后50条", MenuName = Menu, MenuText = "只保留最后50条可见数据")]
        public static void KeepLast50() => Cmd.Run("保留最后50条", () =>
        {
            var ws = Xl.RequireActiveSheet();
            int col = (int)Xl.App.ActiveCell.Column;
            int last = Xl.LastRow(ws, col);
            if (last < 2) throw new UserMessage("当前列没有数据。");
            List<(int Row, int Col, object Value)> cells = Xl.VisibleCells(ws.Range[ws.Cells[2, col], ws.Cells[last, col]]);
            var visible = cells.Select(c => c.Row).Distinct().OrderBy(r => r).ToList();
            var hide = SheetOps.RowsToHideKeepingLast(visible, 50);
            Xl.HideRows(ws, hide);
            Ui.Info(hide.Count > 0 ? $"已隐藏 {hide.Count} 行，只保留了最后 50 条可见数据。" : "可见数据不超过 50 条，无需处理。");
        });

        // ───────── 样品类别轮换筛选（原 模块4）─────────

        private static string _sheetKey;
        private static int _categoryCol;
        private static List<string> _categories;
        private static int _index = -1;

        [ExcelCommand(Name = "工具宏_初始化样品类别", MenuName = Menu, MenuText = "重新读取样品类别")]
        public static void InitCategories() => Cmd.Run("样品类别", () =>
        {
            Init(Xl.RequireActiveSheet());
            Ui.Info($"已找到[样品类别]列，共 {_categories.Count} 个类别。");
        });

        [ExcelCommand(Name = "工具宏_下一个样品类别", MenuName = Menu, MenuText = "下一个样品类别")]
        public static void NextCategory() => Cmd.Run("下一个样品类别", () =>
        {
            var ws = Xl.RequireActiveSheet();
            if (_categories == null || _sheetKey != SheetKey(ws)) Init(ws);
            // 原 VBA 初始化后第一次就跳到第 2 个类别，这里从第 1 个开始
            _index = (_index + 1) % _categories.Count;
            AutoFilter(ws, _categoryCol, _categories[_index]);
            Ui.Info($"已筛选到类别：{_categories[_index]}（{_index + 1}/{_categories.Count}）");
        });

        private static string SheetKey(dynamic ws) => (string)ws.Parent.FullName + "|" + (string)ws.Name;

        private static void Init(dynamic ws)
        {
            _categories = null;
            int lastCol = Xl.LastCol(ws, 1);
            object[,] header = Xl.Read(ws.Range[ws.Cells[1, 1], ws.Cells[1, lastCol]]);
            int col = 0;
            for (int c = 0; c < header.GetLength(1); c++)
                if (header[0, c] is string s && s == "样品类别") { col = c + 1; break; }
            if (col == 0) throw new UserMessage("未找到[样品类别]列！");

            int last = Xl.LastRow(ws, col);
            var list = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // 与 VBA Collection 键一致：不区分大小写
            if (last >= 2)
            {
                object[,] vals = Xl.Read(ws.Range[ws.Cells[2, col], ws.Cells[last, col]]);
                foreach (var v in vals)
                {
                    if (v == null || v is CellError) continue;
                    var s = Vba.CStr(v);
                    if (s.Length > 0 && seen.Add(s)) list.Add(s);
                }
            }
            if (list.Count == 0) throw new UserMessage("[样品类别]列中没有数据！");
            _categories = list;
            _categoryCol = col;
            _sheetKey = SheetKey(ws);
            _index = -1;
        }
    }
}
