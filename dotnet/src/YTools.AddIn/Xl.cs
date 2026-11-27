using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>Excel / WPS COM 对象模型的薄封装（后期绑定，不依赖 Interop 程序集）。</summary>
    internal static class Xl
    {
        // Excel 常量
        public const int xlUp = -4162, xlDown = -4121, xlToLeft = -4159;
        public const int xlCellTypeVisible = 12;
        public const int xlAnd = 1;
        public const int xlTypePDF = 0, xlQualityStandard = 0;
        public const int xlPatternNone = -4142, xlPatternSolid = 1;
        public const int xlColorIndexAutomatic = -4105, xlColorIndexNone = -4142;
        public const int xlCalculationManual = -4135;
        public const int xlWorksheet = -4167;

        public static dynamic App => ExcelDnaUtil.Application;

        public static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

        public static string TypeName(object o) => Microsoft.VisualBasic.Information.TypeName(o);

        /// <summary>当前选区必须是单元格区域，否则提示并中止命令。</summary>
        public static dynamic RequireRangeSelection(string message = "请先选中单元格区域。")
        {
            dynamic sel = App.Selection;
            if (sel == null || TypeName(sel) != "Range") throw new UserMessage(message);
            return sel;
        }

        public static dynamic RequireActiveSheet()
        {
            dynamic ws = App.ActiveSheet;
            if (ws == null || TypeName(ws) != "Worksheet") throw new UserMessage("请先切换到一个工作表。");
            return ws;
        }

        public static int LastRow(dynamic ws, int col) => (int)ws.Cells[ws.Rows.Count, col].End[xlUp].Row;
        public static int LastCol(dynamic ws, int row) => (int)ws.Cells[row, ws.Columns.Count].End[xlToLeft].Column;

        /// <summary>UsedRange 的右下角（绝对行列号）。</summary>
        public static (int Row, int Col) UsedEnd(dynamic ws)
        {
            dynamic u = ws.UsedRange;
            return ((int)u.Row + (int)u.Rows.Count - 1, (int)u.Column + (int)u.Columns.Count - 1);
        }

        /// <summary>选区与已用区域求交，避免选中整列时遍历上百万个空单元格。</summary>
        public static dynamic TrimToUsed(dynamic range)
        {
            dynamic ws = range.Worksheet;
            return App.Intersect(range, ws.UsedRange);
        }

        /// <summary>读取区域值，转成 0 基数组；错误值转 CellError。useValue2=false 时保留日期类型。</summary>
        public static object[,] Read(dynamic range, bool useValue2 = true)
        {
            object raw = useValue2 ? range.Value2 : range.Value;
            if (!(raw is Array arr)) return new[,] { { Norm(raw) } };
            int r0 = arr.GetLowerBound(0), c0 = arr.GetLowerBound(1);
            int rows = arr.GetLength(0), cols = arr.GetLength(1);
            var res = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    res[i, j] = Norm(arr.GetValue(r0 + i, c0 + j));
            return res;
        }

        private static object Norm(object v) => v is int code ? CellError.FromComCode(code) : v;

        /// <summary>把 0 基数组写到以 topLeft 为左上角的区域。</summary>
        public static void Write(dynamic topLeft, object[,] values)
        {
            int rows = values.GetLength(0), cols = values.GetLength(1);
            if (rows == 0 || cols == 0) return;
            var outArr = new object[rows, cols];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < cols; j++)
                    outArr[i, j] = values[i, j] is CellError e ? new ErrorWrapper(e.ComCode) : values[i, j];
            topLeft.Resize[rows, cols].Value = outArr;
        }

        /// <summary>选区中的可见单元格（按区域、按行读取），不存在时返回空列表。</summary>
        public static List<(int Row, int Col, object Value)> VisibleCells(dynamic range)
        {
            var list = new List<(int, int, object)>();
            dynamic vis;
            try { vis = range.SpecialCells(xlCellTypeVisible); }
            catch (COMException) { return list; } // 没有可见单元格
            foreach (dynamic area in vis.Areas)
            {
                int r = area.Row, c = area.Column;
                var vals = Read(area);
                for (int i = 0; i < vals.GetLength(0); i++)
                    for (int j = 0; j < vals.GetLength(1); j++)
                        list.Add((r + i, c + j, vals[i, j]));
            }
            return list;
        }

        /// <summary>把连续的行/列号合并成块，减少 COM 调用次数。</summary>
        public static IEnumerable<(int First, int Last)> Runs(IEnumerable<int> sortedIndexes)
        {
            int first = -1, last = -1;
            foreach (int i in sortedIndexes)
            {
                if (first < 0) { first = last = i; continue; }
                if (i == last + 1) { last = i; continue; }
                yield return (first, last);
                first = last = i;
            }
            if (first >= 0) yield return (first, last);
        }

        public static void HideRows(dynamic ws, IEnumerable<int> rows)
        {
            foreach (var (a, b) in Runs(rows)) ws.Range[ws.Rows[a], ws.Rows[b]].Hidden = true;
        }

        public static void HideColumns(dynamic ws, IEnumerable<int> cols)
        {
            foreach (var (a, b) in Runs(cols)) ws.Range[ws.Columns[a], ws.Columns[b]].Hidden = true;
        }

        /// <summary>按完整路径查找已打开的工作簿。</summary>
        public static dynamic FindOpenWorkbook(string fullPath)
        {
            foreach (dynamic wb in App.Workbooks)
                if (string.Equals((string)wb.FullName, fullPath, StringComparison.OrdinalIgnoreCase)) return wb;
            return null;
        }

        public static IWin32Window Owner
        {
            get
            {
                try { return new WindowWrapper(new IntPtr((int)App.Hwnd)); } catch { return null; }
            }
        }

        private sealed class WindowWrapper : IWin32Window
        {
            public WindowWrapper(IntPtr h) { Handle = h; }
            public IntPtr Handle { get; }
        }
    }

    /// <summary>给用户看的提示（不是程序错误），命令外壳捕获后弹提示框。</summary>
    internal sealed class UserMessage : Exception
    {
        public MessageBoxIcon Icon { get; }
        public UserMessage(string message, MessageBoxIcon icon = MessageBoxIcon.Exclamation) : base(message) { Icon = icon; }
    }
}
