using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 模块6（画标记折线图_final、画控制图_final、辅助线）与 模块3.画钴杂比标记折线图。</summary>
    public static class ChartCommands
    {
        private const string Menu = "工具宏·图表";
        private const int HelperRow = 100; // 新工作表中图表数据写在第 100 行起（与原控制图一致）

        private sealed class SeriesData
        {
            public int Col;
            public string Name;
            public List<object> X = new List<object>();
            public List<double> Y = new List<double>();
            public bool DateX => X.Count > 0 && X[0] is DateTime;
        }

        private sealed class Context
        {
            public dynamic Ws, Sel;
            public int TitleRow, DateCol, CategoryCol, NameCol;
            public string SheetName;
            public List<int> Cols;
        }

        /// <summary>
        /// 解析选区：表头行、日期列（X 轴）、数据列（排除日期列和隐藏列，可不连续）、样品类别/名称列。
        /// inPlace=true 时沿用旧版“画在原表”的规则：没有自动筛选时表头固定在第 1 行，找不到日期列就用第 1 列。
        /// </summary>
        private static Context Parse(bool inPlace)
        {
            dynamic sel = Xl.RequireRangeSelection("请先选中一个或多个列（可不连续）。");
            dynamic ws = sel.Worksheet;
            var ctx = new Context { Ws = ws, Sel = sel, SheetName = ws.Name };
            ctx.TitleRow = (bool)ws.AutoFilterMode ? (int)ws.AutoFilter.Range.Row : inPlace ? 1 : (int)sel.Rows[1].Row;
            int lastCol = Xl.UsedEnd(ws).Col;
            object[,] header = Xl.Read(ws.Range[ws.Cells[ctx.TitleRow, 1], ws.Cells[ctx.TitleRow, lastCol]]);
            ctx.DateCol = ChartKit.FindHeader(header, h => RuleCharts.Like(h, "*日期*"));
            if (ctx.DateCol == 0)
            {
                if (!inPlace) throw new UserMessage("未找到[取样日期]列，请确认表头。");
                ctx.DateCol = 1;
            }
            ctx.CategoryCol = ChartKit.FindHeader(header, h => string.Equals(h, "样品类别", StringComparison.OrdinalIgnoreCase));
            ctx.NameCol = ChartKit.FindHeader(header, h => string.Equals(h, "样品名称", StringComparison.OrdinalIgnoreCase));
            var cols = new SortedSet<int>();
            var order = new List<int>();
            foreach (dynamic area in sel.Areas)
                foreach (dynamic c in area.Columns)
                {
                    int col = c.Column;
                    if (col != ctx.DateCol && !(bool)c.EntireColumn.Hidden && cols.Add(col)) order.Add(col);
                }
            if (order.Count == 0) throw new UserMessage("没有有效的数据列！");
            ctx.Cols = order;
            return ctx;
        }

        /// <summary>读取一列可画的点（跳过隐藏行与非数值）。</summary>
        private static SeriesData Collect(Context ctx, int col)
        {
            dynamic ws = ctx.Ws;
            int start = ctx.TitleRow + 1;
            int last = Math.Max((int)Xl.LastRow(ws, ctx.DateCol), (int)Xl.LastRow(ws, col));
            var name = Vba.CStr(Xl.Read(ws.Cells[ctx.TitleRow, col])[0, 0] is CellError ? "" : Xl.Read(ws.Cells[ctx.TitleRow, col])[0, 0]);
            var sd = new SeriesData { Col = col, Name = name.Length == 0 ? "Col" + col : name };
            if (last < start) return sd;
            object[,] xs = Xl.Read(ws.Range[ws.Cells[start, ctx.DateCol], ws.Cells[last, ctx.DateCol]], useValue2: false);
            object[,] ys = Xl.Read(ws.Range[ws.Cells[start, col], ws.Cells[last, col]], useValue2: false);
            IList<bool> hidden = ChartKit.HiddenRows(ws, start, last);
            var yl = Enumerable.Range(0, ys.GetLength(0)).Select(i => ys[i, 0]).ToList();
            foreach (int i in ChartMath.PlottableRows(yl, hidden))
            {
                sd.X.Add(xs[i, 0]);
                sd.Y.Add(Vba.CDbl(ys[i, 0]));
            }
            return sd;
        }

        private static string Title(Context ctx, string series, string suffix)
            => ChartMath.Title(ctx.SheetName, ChartKit.LastValue(ctx.Ws, ctx.CategoryCol), ChartKit.LastValue(ctx.Ws, ctx.NameCol), series, suffix);

        private static dynamic NewSheetBefore(dynamic ws, string prefix)
        {
            dynamic s = ws.Parent.Worksheets.Add(Before: ws);
            try { s.Name = (prefix + DateTime.Now.ToString("yyyyMMdd_HHmmss")).Substring(0, Math.Min(31, prefix.Length + 15)); } catch { }
            return s;
        }

        /// <summary>把列数据写到新表第 HelperRow 行起，返回该区域（图表引用区域，比直接赋数组稳定）。</summary>
        private static dynamic WriteColumn(dynamic sheet, int col, IList<object> values)
        {
            var arr = new object[values.Count, 1];
            for (int i = 0; i < values.Count; i++) arr[i, 0] = values[i];
            dynamic top = sheet.Cells[HelperRow, col];
            Xl.Write(top, arr);
            return top.Resize[values.Count, 1];
        }

        // ───────────── 标记折线图 ─────────────

        [ExcelCommand(Name = "工具宏_画折线图", MenuName = Menu, MenuText = "画标记折线图（新工作表）")]
        public static void LineChartNewSheet() => Cmd.Run("画标记折线图", () => LineCharts(inPlace: false));

        [ExcelCommand(Name = "工具宏_画折线图原表", MenuName = Menu, MenuText = "画标记折线图（画在原表）")]
        public static void LineChartInPlace() => Cmd.Run("画标记折线图", () => LineCharts(inPlace: true));

        private static void LineCharts(bool inPlace)
        {
            var ctx = Parse(inPlace);
            dynamic target = inPlace ? ctx.Ws : NewSheetBefore(ctx.Ws, "折线图_");
            double h = inPlace ? ChartKit.Height : ChartKit.Height * 0.9;
            double left0 = inPlace ? (double)ctx.Sel.Left + (double)ctx.Sel.Width : 20;
            double top0 = inPlace ? (double)ctx.Sel.Top + (double)ctx.Sel.Height : 20;
            int made = 0;
            foreach (int col in ctx.Cols)
            {
                var sd = Collect(ctx, col);
                if (sd.Y.Count == 0) continue;
                var (left, top) = ChartMath.GridPosition(made, 2, ChartKit.Width, h, left0, top0);
                dynamic ch = ChartKit.NewChart(target, left, top, ChartKit.Width, h);
                if (inPlace)
                {
                    // 原表：直接引用源数据区域，隐藏行由 PlotVisibleOnly 跳过
                    dynamic ws = ctx.Ws;
                    int start = ctx.TitleRow + 1, last = Math.Max((int)Xl.LastRow(ws, ctx.DateCol), (int)Xl.LastRow(ws, col));
                    ChartKit.StyleMain(ChartKit.AddSeries(ch, sd.Name, ws.Range[ws.Cells[start, ctx.DateCol], ws.Cells[last, ctx.DateCol]],
                                                          ws.Range[ws.Cells[start, col], ws.Cells[last, col]]));
                }
                else
                {
                    // 新表：只写入可画的点，文本、空值不会在折线图里被当成 0
                    dynamic x = WriteColumn(target, 2 * made + 1, sd.X);
                    dynamic y = WriteColumn(target, 2 * made + 2, sd.Y.Cast<object>().ToList());
                    ChartKit.StyleMain(ChartKit.AddSeries(ch, sd.Name, x, y));
                }
                ChartKit.ScaleY(ch, sd.Y.Min(), sd.Y.Max());
                ChartKit.StyleChart(ch, Title(ctx, sd.Name, "趋势图"), sd.DateX);
                ch.HasLegend = false;
                made++;
            }
            Ui.Info("完成！共生成 " + made + " 张图表（两列排列）");
        }

        // ───────────── 控制图 ─────────────

        [ExcelCommand(Name = "工具宏_画控制图", MenuName = Menu, MenuText = "画控制图（I-MR）")]
        public static void ControlChart() => Cmd.Run("画控制图", () =>
        {
            var ctx = Parse(inPlace: false);
            dynamic sheet = NewSheetBefore(ctx.Ws, "控制图_");
            double h = ChartKit.Height * 0.9;
            int made = 0, block = 0;
            foreach (int col in ctx.Cols)
            {
                block++;
                var sd = Collect(ctx, col);
                if (sd.Y.Count < 2) continue;
                var y = sd.Y.ToArray();
                var (cl, ucl, lcl, ooc) = ChartMath.IndividualsLimits(y);
                int b = 1 + (block - 1) * 6, n = y.Length;
                dynamic hCL = WriteColumn(sheet, b, Enumerable.Repeat((object)cl, n).ToList());
                dynamic hUC = WriteColumn(sheet, b + 1, Enumerable.Repeat((object)ucl, n).ToList());
                dynamic hLC = WriteColumn(sheet, b + 2, Enumerable.Repeat((object)lcl, n).ToList());
                dynamic hOOC = WriteColumn(sheet, b + 3, y.Select((v, i) => ooc[i] ? (object)v : null).ToList());
                dynamic hX = WriteColumn(sheet, b + 4, sd.X);
                dynamic hY = WriteColumn(sheet, b + 5, y.Cast<object>().ToList());

                var (left, top) = ChartMath.GridPosition(made, 2, ChartKit.Width, h);
                dynamic ch = ChartKit.NewChart(sheet, left, top, ChartKit.Width, h);
                ChartKit.StyleMain(ChartKit.AddSeries(ch, sd.Name, hX, hY), 6);
                dynamic s;
                s = ChartKit.AddSeries(ch, "CL", hX, hCL); s.Format.Line.ForeColor.RGB = Xl.Rgb(0, 176, 80); s.MarkerStyle = ChartKit.xlMarkerStyleNone;
                foreach (var (nm, rng) in new[] { ("UCL", (object)hUC), ("LCL", (object)hLC) })
                {
                    s = ChartKit.AddSeries(ch, nm, hX, rng);
                    s.Format.Line.ForeColor.RGB = Xl.Rgb(192, 0, 0);
                    s.Format.Line.DashStyle = ChartKit.msoLineDash;
                    s.MarkerStyle = ChartKit.xlMarkerStyleNone;
                }
                s = ChartKit.AddSeries(ch, "超限点", hX, hOOC);
                s.Format.Line.Visible = ChartKit.msoFalse;
                s.MarkerStyle = ChartKit.xlMarkerStyleCircle;
                s.MarkerSize = 9;
                s.MarkerForegroundColor = Xl.Rgb(255, 0, 0);
                s.MarkerBackgroundColor = Xl.Rgb(255, 0, 0);

                ChartKit.ScaleY(ch, Math.Min(y.Min(), Math.Min(lcl, ucl)), Math.Max(y.Max(), Math.Max(lcl, ucl)));
                int oocCount = ooc.Count(f => f);
                var title = Title(ctx, sd.Name, "控制图") + (oocCount > 0 ? "（超限 " + oocCount + " 点）" : "");
                ChartKit.StyleChart(ch, title, sd.DateX);
                ch.HasLegend = true;
                // 原 VBA 用了不存在的常量 xlLegendPositionTopLeft，这里放在顶部
                ch.Legend.Position = ChartKit.xlLegendPositionTop;
                ch.Legend.Font.Name = "等线";
                ch.Legend.Font.Size = 9;
                made++;
            }
            Ui.Info("完成！共生成 " + made + " 张控制图（两列排列）");
        });

        // ───────────── 钴杂比 ─────────────

        [ExcelCommand(Name = "工具宏_画钴杂比", MenuName = Menu, MenuText = "画钴杂比折线图（选中杂质列）")]
        public static void CobaltRatioChart() => Cmd.Run("钴杂比", () =>
        {
            dynamic sel = Xl.RequireRangeSelection("请先选中杂质列中的单元格。");
            dynamic ws = sel.Worksheet;
            int impCol = sel.Column;
            int titleRow = (bool)ws.AutoFilterMode ? (int)ws.AutoFilter.Range.Row : 1;
            int lastCol = Xl.UsedEnd(ws).Col;
            object[,] header = Xl.Read(ws.Range[ws.Cells[titleRow, 1], ws.Cells[titleRow, lastCol]]);
            int coCol = ChartKit.FindHeader(header, h => h.IndexOf("Co", StringComparison.OrdinalIgnoreCase) >= 0);
            if (coCol == 0) throw new UserMessage("未找到标题中包含'Co'的列，请检查数据标题。");

            int start = titleRow + 1, last = Xl.LastRow(ws, impCol);
            if (last < start) throw new UserMessage("杂质列没有数据。");
            object[,] co = Xl.Read(ws.Range[ws.Cells[start, coCol], ws.Cells[last, coCol]]);
            if (co.Cast<object>().All(v => !(v is double d) || d == 0)) throw new UserMessage("Co值全为0，无法绘制钴杂比图表。");
            object[,] imp = Xl.Read(ws.Range[ws.Cells[start, impCol], ws.Cells[last, impCol]]);
            object[,] xs = Xl.Read(ws.Range[ws.Cells[start, 1], ws.Cells[last, 1]], useValue2: false);
            IList<bool> hidden = ChartKit.HiddenRows(ws, start, last);
            var vis = Enumerable.Range(0, hidden.Count).Where(i => !hidden[i]).ToList();
            if (vis.Count == 0) throw new UserMessage("没有找到可见的数据。");
            // 原 VBA 按区域序号去配对 Co 与杂质，筛选后经常错位或报错；这里逐行配对
            var ratio = ChartMath.CobaltRatio(vis.Select(i => co[i, 0]).ToList(), vis.Select(i => imp[i, 0]).ToList());
            var x = vis.Select(i => xs[i, 0]).ToArray();

            string impName = Vba.CStr(header[0, impCol - 1] is CellError ? "" : header[0, impCol - 1]);
            dynamic ch = ChartKit.NewChart(ws, (double)sel.Left + (double)sel.Width, (double)sel.Top + (double)sel.Height, ChartKit.Width, ChartKit.Height);
            ChartKit.StyleMain(ChartKit.AddSeriesSafe(ch, "钴杂比 (Co/" + impName + ")", x, ratio.Cast<object>().ToArray()));
            int catCol = ChartKit.FindHeader(header, h => h == "样品类别"), nameCol = catCol == 0 ? ChartKit.FindHeader(header, h => h == "样品名称") : 0;
            int whoCol = catCol != 0 ? catCol : nameCol;
            string who = whoCol == 0 ? null : Vba.CStr(Xl.Read(ws.Cells[(int)sel.Row, whoCol])[0, 0] ?? "");
            ChartKit.StyleChart(ch, who != null ? $"[{ws.Name} - {who} - 钴杂比趋势图]" : $"[{ws.Name} - 钴杂比趋势图]", xs[vis[0], 0] is DateTime);
            ch.Legend.Position = ChartKit.xlLegendPositionBottom;
        });

        // ───────────── 辅助线 ─────────────

        [ExcelCommand(Name = "工具宏_添加辅助线", MenuName = Menu, MenuText = "给选中图表添加辅助线…")]
        public static void AddHelperLine() => Cmd.Run("添加辅助线", () =>
        {
            dynamic ch = Xl.App.ActiveChart ?? throw new UserMessage("请先选中一个图表区域！");
            if ((int)ch.SeriesCollection().Count == 0) throw new UserMessage("图表中没有数据系列，无法添加辅助线！");
            if (!HelperLineForm.Ask(out string name, out double y)) return;
            AddHelperLineTo(ch, name, y);
            Ui.Info("辅助线已添加到图表！");
        }, freezeScreen: false);

        internal static void AddHelperLineTo(dynamic ch, string name, double y)
        {
            object raw = ch.SeriesCollection(1).XValues;
            var xs = new List<double>();
            if (raw is Array arr)
                foreach (var v in arr)
                    if (v is double d) xs.Add(d);
                    else if (v is DateTime dt) xs.Add(dt.ToOADate());
                    else if (v is string s && Vba.TryParseDouble(s, out var p)) xs.Add(p);
            if (xs.Count == 0) throw new UserMessage("第一条系列的 X 值不是数值或日期，无法添加辅助线。");
            var (hx, hy) = ChartMath.HelperLine(xs, y);
            dynamic s2 = ChartKit.AddSeriesSafe(ch, string.IsNullOrEmpty(name) ? "标准" : name, hx.Cast<object>().ToArray(), hy.Cast<object>().ToArray());
            s2.Format.Line.ForeColor.RGB = Xl.Rgb(255, 0, 0);
            s2.Format.Line.DashStyle = ChartKit.msoLineDash;
            s2.Format.Line.Weight = 1.5;
            s2.MarkerStyle = ChartKit.xlMarkerStyleNone;
        }
    }
}
