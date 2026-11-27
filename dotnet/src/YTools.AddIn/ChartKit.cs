using System;
using System.Collections.Generic;
using System.Linq;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>图表的公共样式与坐标轴设置（原 模块6 中各画图过程共用的代码）。</summary>
    internal static class ChartKit
    {
        public const int xlLineMarkers = 65, xlXYScatterLines = 74;
        public const int xlCategory = 1, xlValue = 2, xlPrimary = 1;
        public const int xlMarkerStyleCircle = 8, xlMarkerStyleNone = -4142;
        public const int msoLineDash = 4, msoTrue = -1, msoFalse = 0;
        public const int xlLegendPositionBottom = -4107, xlLegendPositionTop = -4160;

        public static readonly int LineColor = Xl.Rgb(91, 155, 213), FillColor = Xl.Rgb(237, 125, 49);

        /// <summary>标准图表尺寸（厘米换算成磅：16.33cm × 6.66cm）。</summary>
        public const double Width = 16.33 * 28.35, Height = 6.66 * 28.35;

        public static dynamic NewChart(dynamic sheet, double left, double top, double w, double h, int type = xlLineMarkers)
        {
            dynamic co = sheet.ChartObjects().Add(left, top, w, h);
            dynamic ch = co.Chart;
            ch.ChartType = type;
            ch.PlotVisibleOnly = true;
            while ((int)ch.SeriesCollection().Count > 0) ch.SeriesCollection(1).Delete();
            return ch;
        }

        public static dynamic AddSeries(dynamic ch, string name, object x, object y)
        {
            dynamic s = ch.SeriesCollection().NewSeries();
            s.Name = name;
            s.XValues = x;
            s.Values = y;
            return s;
        }

        public const string DataSheetName = "工具宏_图表数据";

        /// <summary>
        /// 用数组作为系列数据（与原 VBA 相同）；点数多时数组赋值会因系列公式过长失败，
        /// 此时把数据写到工作簿中隐藏的「工具宏_图表数据」表再引用。
        /// </summary>
        public static dynamic AddSeriesSafe(dynamic ch, string name, object[] x, object[] y)
        {
            try { return AddSeries(ch, name, x, y); }
            catch (System.Runtime.InteropServices.COMException)
            {
                int n = ch.SeriesCollection().Count;
                if (n > 0) { dynamic last = ch.SeriesCollection(n); if ((string)last.Name == name) last.Delete(); }
                dynamic wb = ch.Parent.Parent.Parent; // Chart → ChartObject → Worksheet → Workbook
                dynamic data = null;
                foreach (dynamic s in wb.Worksheets) if ((string)s.Name == DataSheetName) { data = s; break; }
                if (data == null)
                {
                    dynamic active = Xl.App.ActiveSheet;
                    data = wb.Worksheets.Add(After: wb.Worksheets[(int)wb.Worksheets.Count]);
                    data.Name = DataSheetName;
                    data.Visible = 0; // xlSheetHidden
                    try { active.Activate(); } catch { }
                }
                int col = (int)Xl.UsedEnd(data).Col + 1;
                if ((int)data.UsedRange.Count == 1 && data.Cells[1, 1].Value2 == null) col = 1;
                var arr = new object[x.Length + 1, 2];
                arr[0, 0] = name + " X"; arr[0, 1] = name;
                for (int i = 0; i < x.Length; i++) { arr[i + 1, 0] = x[i]; arr[i + 1, 1] = y[i]; }
                Xl.Write(data.Cells[1, col], arr);
                return AddSeries(ch, name, data.Cells[2, col].Resize[x.Length, 1], data.Cells[2, col + 1].Resize[x.Length, 1]);
            }
        }

        public static void StyleMain(dynamic s, int markerSize = 0)
        {
            s.Format.Line.ForeColor.RGB = LineColor;
            s.MarkerStyle = xlMarkerStyleCircle;
            if (markerSize > 0) s.MarkerSize = markerSize;
            s.MarkerForegroundColor = LineColor;
            s.MarkerBackgroundColor = FillColor;
        }

        public static dynamic Axis(dynamic ch, int type)
        {
            try { return ch.Axes(type, xlPrimary); } catch { return null; }
        }

        /// <summary>Y 轴：先按数据±5% 设范围，读取 Excel 自动算出的主刻度，再把范围对齐到刻度（不改刻度单位）。</summary>
        public static void ScaleY(dynamic ch, double dataMin, double dataMax)
        {
            dynamic ax = Axis(ch, xlValue);
            if (ax == null) return;
            var (min, max) = ChartMath.PaddedRange(dataMin, dataMax);
            ax.MinimumScaleIsAuto = false;
            ax.MaximumScaleIsAuto = false;
            ax.MinimumScale = min;
            ax.MaximumScale = max;
            ch.Refresh();
            System.Windows.Forms.Application.DoEvents();
            double unit = ax.MajorUnit;
            (min, max) = ChartMath.AlignToUnit(min, max, unit);
            ax.MinimumScale = min;
            ax.MaximumScale = max;
        }

        /// <summary>标题、字体、网格线、白底蓝边框（原 _final 的样式）。</summary>
        public static void StyleChart(dynamic ch, string title, bool dateAxis)
        {
            ch.HasTitle = true;
            ch.ChartTitle.Text = title;
            ch.ChartTitle.Font.Name = "等线";
            ch.ChartTitle.Font.Size = 14;
            foreach (int t in new[] { xlCategory, xlValue })
            {
                dynamic ax = Axis(ch, t);
                if (ax == null) continue;
                ax.HasTitle = false;
                ax.TickLabels.Font.Name = "等线";
            }
            dynamic y = Axis(ch, xlValue);
            if (y != null)
            {
                y.MajorGridlines.Format.Line.ForeColor.RGB = Xl.Rgb(220, 220, 220);
                y.MajorGridlines.Format.Line.DashStyle = msoLineDash;
            }
            ch.ChartArea.Format.Fill.ForeColor.RGB = Xl.Rgb(255, 255, 255);
            ch.PlotArea.Format.Fill.ForeColor.RGB = Xl.Rgb(255, 255, 255);
            ch.ChartArea.Border.Color = 16711680;
            if (dateAxis)
            {
                dynamic x = Axis(ch, xlCategory);
                try { if (x != null) x.TickLabels.NumberFormat = "m月d日"; } catch { }
            }
        }

        /// <summary>表头行中第一个满足条件的列（1 基），没有返回 0。</summary>
        public static int FindHeader(object[,] headerRow, Func<string, bool> match)
        {
            for (int c = 0; c < headerRow.GetLength(1); c++)
            {
                var v = headerRow[0, c];
                if (v == null || v is CellError) continue;
                if (match(Vba.CStr(v))) return c + 1;
            }
            return 0;
        }

        /// <summary>某列最后一个非空单元格的值（原 VBA：Cells(Rows.Count, col).End(xlUp)）。</summary>
        public static object LastValue(dynamic ws, int col) => col <= 0 ? null : Xl.Read(ws.Cells[ws.Rows.Count, col].End[Xl.xlUp])[0, 0];

        public static IList<bool> HiddenRows(dynamic ws, int firstRow, int lastRow)
        {
            var hidden = new bool[lastRow - firstRow + 1];
            dynamic vis;
            try { vis = ws.Range[ws.Cells[firstRow, 1], ws.Cells[lastRow, 1]].SpecialCells(Xl.xlCellTypeVisible); }
            catch (System.Runtime.InteropServices.COMException) { for (int i = 0; i < hidden.Length; i++) hidden[i] = true; return hidden; }
            for (int i = 0; i < hidden.Length; i++) hidden[i] = true;
            foreach (dynamic a in vis.Areas)
            {
                int r = a.Row, n = a.Rows.Count;
                for (int i = 0; i < n; i++) hidden[r + i - firstRow] = false;
            }
            return hidden;
        }
    }
}
