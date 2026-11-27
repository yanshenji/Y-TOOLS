using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>
    /// 原 模块5.GenerateCharts_AndExportData_ByRule：按当前工作簿「配置表」中的规则批量筛选、导出数据并画散点折线图。
    /// 配置表：B2 每行图表数，B3 图宽，B4 图高，B5 标题字号；第 7 行起每行一条规则：
    /// A 数据表名 | B 样品类别（标签）| C 样品名称条件 | D 开始日期 | E 结束日期 | F 检测项目（逗号分隔，留空=自动）
    /// </summary>
    public static class RuleChartCommands
    {
        private const int xlPasteAll = -4104, xlAscending = 1, xlYes = 1, xlUpward = -4171, xlHorizontal = -4128;

        [ExcelCommand(Name = "工具宏_按配置表出图", MenuName = "工具宏·图表", MenuText = "按「配置表」批量出图")]
        public static void Run() => Cmd.Run("按配置表出图", () =>
        {
            dynamic app = Xl.App;
            app.DisplayAlerts = false; // 删除空的输出表时不弹确认（命令结束后自动恢复）
            dynamic wb = app.ActiveWorkbook ?? throw new UserMessage("没有打开的工作簿。");
            dynamic cfg = FindSheet(wb, "配置表") ?? throw new UserMessage("未找到配置表！（当前工作簿中需要有名为“配置表”的工作表）");

            int perRow = Math.Max(1, (int)Vba.CDbl(Xl.Read(cfg.Range["B2"])[0, 0]));
            double w = Vba.CDbl(Xl.Read(cfg.Range["B3"])[0, 0]), h = Vba.CDbl(Xl.Read(cfg.Range["B4"])[0, 0]);
            int fontSize = (int)Vba.CDbl(Xl.Read(cfg.Range["B5"])[0, 0]);
            int lastRule = Xl.LastRow(cfg, 1);
            if (lastRule < 7) throw new UserMessage("配置表第 7 行起没有规则。");
            object[,] rows = Xl.Read(cfg.Range[cfg.Cells[7, 1], cfg.Cells[lastRule, 6]]);

            int total = 0;
            var notes = new List<string>();
            for (int r = 0; r < rows.GetLength(0); r++)
            {
                var rule = ChartRule.FromRow(Enumerable.Range(0, 6).Select(c => rows[r, c]).ToArray());
                if (rule == null) continue;
                try
                {
                    int made = RunRule(wb, rule, perRow, w, h, fontSize, out string note);
                    total += made;
                    if (note != null) notes.Add($"第{r + 7}行（{rule.SheetName}）：{note}");
                }
                catch (UserMessage m) { notes.Add($"第{r + 7}行（{rule.SheetName}）：{m.Message}"); }
            }
            cfg.Activate();
            Ui.Info("筛选数据已另存并绘图完成。共生成图表 " + total + " 张。" + (notes.Count > 0 ? "\n\n" + string.Join("\n", notes) : ""));
        });

        private static dynamic FindSheet(dynamic wb, string name)
        {
            foreach (dynamic s in wb.Worksheets) if (string.Equals((string)s.Name, name, StringComparison.OrdinalIgnoreCase)) return s;
            return null;
        }

        private static List<string> SheetNames(dynamic wb)
        {
            var list = new List<string>();
            foreach (dynamic s in wb.Sheets) list.Add((string)s.Name);
            return list;
        }

        private static int RunRule(dynamic wb, ChartRule rule, int perRow, double w, double h, int fontSize, out string note)
        {
            note = null;
            dynamic data = FindSheet(wb, rule.SheetName) ?? throw new UserMessage("找不到该工作表，跳过");
            data.Rows.Hidden = false;
            if ((bool)data.AutoFilterMode) data.AutoFilterMode = false;

            int lastCol = Xl.LastCol(data, 1);
            object[,] header = Xl.Read(data.Range[data.Cells[1, 1], data.Cells[1, lastCol]]);
            int HeaderCol(params string[] names)
            {
                foreach (var n in names) { int c = ChartKit.FindHeader(header, hh => Vba.Trim(hh) == n); if (c > 0) return c; }
                return 0;
            }
            int nameCol = HeaderCol("样品名称", "樣品名稱");
            int dateCol = HeaderCol("送样日期", "送樣日期");
            if (dateCol == 0) dateCol = HeaderCol("取样日期", "取樣日期");
            if (nameCol == 0 || dateCol == 0) throw new UserMessage("第 1 行没有“样品名称”或“送样日期/取样日期”，跳过");

            dynamic region = data.Range["A1"].CurrentRegion;
            int regionRows = region.Rows.Count;
            // 日期筛选（用序列号，不受日期格式影响）
            string D(double d) => d.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            if (rule.StartDate.HasValue && rule.EndDate.HasValue)
                region.AutoFilter(Field: dateCol, Criteria1: ">=" + D(rule.StartDate.Value), Operator: Xl.xlAnd, Criteria2: "<=" + D(rule.EndDate.Value));
            else if (rule.StartDate.HasValue) region.AutoFilter(Field: dateCol, Criteria1: ">=" + D(rule.StartDate.Value));
            else if (rule.EndDate.HasValue) region.AutoFilter(Field: dateCol, Criteria1: "<=" + D(rule.EndDate.Value));

            // 样品名称条件
            var pattern = rule.Pattern ?? "";
            if (pattern.Trim().Length > 0 && pattern != "*")
            {
                if (RuleCharts.HasLogicalOps(pattern))
                {
                    object[,] names = Xl.Read(data.Range[data.Cells[2, nameCol], data.Cells[regionRows, nameCol]]);
                    IList<bool> hidden = ChartKit.HiddenRows(data, 2, regionRows);
                    var hide = new List<int>();
                    for (int i = 0; i < names.GetLength(0); i++)
                        if (!hidden[i] && !RuleCharts.NameMatches(names[i, 0] is CellError ? "" : Vba.CStr(names[i, 0]), pattern)) hide.Add(i + 2);
                    Xl.HideRows(data, hide);
                }
                else region.AutoFilter(Field: nameCol, Criteria1: pattern);
            }

            dynamic visible;
            try { visible = region.SpecialCells(Xl.xlCellTypeVisible); }
            catch (System.Runtime.InteropServices.COMException) { visible = null; }

            var existing = SheetNames(wb);
            string label = rule.Label;
            string chartName = RuleCharts.UniqueSheetName(rule.SheetName + "_" + label + "_图表", existing);
            existing.Add(chartName);
            string dataName = RuleCharts.UniqueSheetName(rule.SheetName + "_" + label + "_数据", existing);
            dynamic outChart = wb.Sheets.Add(After: wb.Sheets[(int)wb.Sheets.Count]);
            outChart.Name = chartName;
            dynamic outData = wb.Sheets.Add(After: wb.Sheets[(int)wb.Sheets.Count]);
            outData.Name = dataName;

            int made = 0;
            try
            {
                if (visible == null) { note = "筛选后没有数据"; return 0; }
                visible.Copy();
                outData.Range["A1"].PasteSpecial(xlPasteAll);
                Xl.App.CutCopyMode = false;
            }
            finally
            {
                // 恢复源表：取消筛选、显示被隐藏的行（原 VBA 会把源表排序并留下隐藏行）
                if ((bool)data.AutoFilterMode) data.AutoFilterMode = false;
                data.Rows.Hidden = false;
            }

            // 在导出的数据表上按日期升序排序（原 VBA 直接排序源表）
            dynamic outRegion = outData.Range["A1"].CurrentRegion;
            int n = (int)outRegion.Rows.Count - 1;
            if (n < 1) { note = "筛选后没有数据"; outChart.Delete(); return 0; }
            outRegion.Sort(Key1: outRegion.Columns[dateCol], Order1: xlAscending, Header: xlYes);
            object[,] outValues = Xl.Read(outRegion);

            // 检测项目：未指定时取第 6 列之后、有数值的列
            var items = rule.ItemList.ToList();
            if (items.Count == 0)
                for (int c = 6; c < outValues.GetLength(1); c++)
                {
                    var hText = outValues[0, c] is CellError ? "" : Vba.Trim(Vba.CStr(outValues[0, c]));
                    if (hText.Length > 0 && Enumerable.Range(1, n).Any(r => outValues[r, c] is double)) items.Add(hText);
                }
            if (items.Count == 0) { note = "没有可画的检测项目"; outChart.Delete(); return 0; }

            dynamic xr = outData.Range[outData.Cells[2, dateCol], outData.Cells[n + 1, dateCol]];
            var dates = Enumerable.Range(1, n).Select(r => outValues[r, dateCol - 1]).OfType<double>().ToList();
            var missing = new List<string>();
            foreach (var item in items)
            {
                int col = 0;
                for (int c = 0; c < outValues.GetLength(1); c++)
                    if (!(outValues[0, c] is CellError) && string.Equals(Vba.CStr(outValues[0, c]), item, StringComparison.OrdinalIgnoreCase)) { col = c + 1; break; }
                if (col == 0) { missing.Add(item); continue; }
                var ys = Enumerable.Range(1, n).Select(r => outValues[r, col - 1]).OfType<double>().ToList();
                if (ys.Count == 0) continue;

                double left = (made % perRow) * (w + 10), top = (made / perRow) * (h + 10);
                dynamic ch = ChartKit.NewChart(outChart, left, top, w, h, ChartKit.xlXYScatterLines);
                ch.HasTitle = true;
                ch.ChartTitle.Text = $"{rule.SheetName} | {label} - {item}\rMax={RuleCharts.SmartNumber(ys.Max())}, Min={RuleCharts.SmartNumber(ys.Min())}, Avg={RuleCharts.SmartNumber(ys.Average())}";
                try { ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = fontSize; } catch { }
                dynamic s = ChartKit.AddSeries(ch, item, xr, outData.Range[outData.Cells[2, col], outData.Cells[n + 1, col]]);
                s.MarkerStyle = ChartKit.xlMarkerStyleCircle;
                s.MarkerSize = 5;
                s.Format.Fill.ForeColor.RGB = Xl.Rgb(255, 140, 0);
                s.Format.Line.Weight = 1.25;
                foreach (dynamic area in new[] { ch.ChartArea, ch.PlotArea })
                {
                    area.Format.Fill.Visible = ChartKit.msoTrue;
                    area.Format.Fill.Solid();
                    area.Format.Fill.ForeColor.RGB = Xl.Rgb(255, 255, 255);
                }
                dynamic yAx = ChartKit.Axis(ch, ChartKit.xlValue);
                if (yAx != null) { yAx.HasMajorGridlines = true; try { yAx.MajorGridlines.Format.Line.Visible = ChartKit.msoTrue; } catch { } }
                dynamic xAx = ChartKit.Axis(ch, ChartKit.xlCategory);
                if (xAx != null && dates.Count > 0)
                {
                    double span = dates.Max() - dates.Min();
                    if (span > 0) xAx.MajorUnit = RuleCharts.DateMajorUnit(span);
                    xAx.TickLabels.NumberFormat = "yyyy-mm-dd";
                    try { xAx.TickLabels.Orientation = xlUpward; } catch { try { xAx.TickLabels.Orientation = xlHorizontal; } catch { } }
                }
                try { ch.Legend.Delete(); } catch { }
                made++;
            }
            if (missing.Count > 0) note = "表头中找不到：" + string.Join("、", missing);
            if (made == 0) outChart.Delete();
            return made;
        }
    }
}
