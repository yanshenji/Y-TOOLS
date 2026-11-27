using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;

namespace YTools.AddIn.Commands
{
    /// <summary>原 其他.bas 的 PDF 导出与批量删除条件格式。</summary>
    public static class ExportCommands
    {
        private const string Menu = "工具宏·导出";

        [ExcelCommand(Name = "工具宏_工作簿导出PDF", MenuName = Menu, MenuText = "当前工作簿导出 PDF")]
        public static void WorkbookToPdf() => Cmd.Run("导出 PDF", () =>
        {
            dynamic wb = Xl.App.ActiveWorkbook ?? throw new UserMessage("没有打开的工作簿。");
            if (string.IsNullOrEmpty((string)wb.Path)) throw new UserMessage("请先保存当前工作簿。");
            string pdf = Path.ChangeExtension((string)wb.FullName, ".pdf");
            // 原 VBA 会连续问两次是否覆盖，这里只问一次
            if (File.Exists(pdf) && !Ui.Confirm("PDF 文件已存在，是否要覆盖？\n" + pdf)) return;
            Export(wb, pdf);
            Ui.Info("已导出：\n" + pdf);
        });

        [ExcelCommand(Name = "工具宏_工作表导出PDF", MenuName = Menu, MenuText = "当前工作表导出 PDF")]
        public static void SheetToPdf() => Cmd.Run("导出 PDF", () =>
        {
            dynamic wb = Xl.App.ActiveWorkbook ?? throw new UserMessage("没有打开的工作簿。");
            if (string.IsNullOrEmpty((string)wb.Path)) throw new UserMessage("请先保存当前工作簿（PDF 保存在工作簿所在文件夹）。");
            dynamic ws = Xl.App.ActiveSheet;
            string name = (string)ws.Name;
            // 工作表还叫 Sheet1 之类的默认名时，用工作簿名
            if (name.StartsWith("Sheet", StringComparison.Ordinal)) name = Path.GetFileNameWithoutExtension((string)wb.Name);
            string pdf = Path.Combine((string)wb.Path, name + ".pdf");
            Export(ws, pdf);
            Ui.Info("PDF已导出至工作表所在路径：\n" + pdf);
        });

        [ExcelCommand(Name = "工具宏_批量导出PDF", MenuName = Menu, MenuText = "文件夹内 Excel 批量导出 PDF…")]
        public static void FolderToPdf() => Cmd.Run("批量导出 PDF", () =>
        {
            var files = PickExcelFiles("选择要批量导出 PDF 的文件夹（PDF 保存在同一文件夹）", out string folder);
            if (files == null) return;
            if (!Ui.Confirm($"将把 {files.Count} 个 Excel 文件导出为 PDF（同名 PDF 会被覆盖），是否继续？")) return;
            var failed = ForEachWorkbook(files, readOnly: true, wb =>
                Export(wb, Path.Combine(folder, Path.GetFileNameWithoutExtension((string)wb.Name) + ".pdf")));
            Report("导出", files.Count, failed);
        });

        [ExcelCommand(Name = "工具宏_批量删除条件格式", MenuName = Menu, MenuText = "文件夹内 Excel 批量删除条件格式…")]
        public static void FolderDeleteConditionalFormats() => Cmd.Run("批量删除条件格式", () =>
        {
            var files = PickExcelFiles("选择要删除条件格式的文件夹", out _);
            if (files == null) return;
            if (!Ui.Confirm($"将删除 {files.Count} 个文件中所有工作表的条件格式并保存，此操作不可撤销。是否继续？")) return;
            var failed = ForEachWorkbook(files, readOnly: false, wb =>
            {
                foreach (dynamic ws in wb.Worksheets)
                    if ((int)ws.Cells.FormatConditions.Count > 0) ws.Cells.FormatConditions.Delete();
                wb.Save();
            });
            Report("删除条件格式", files.Count, failed);
        });

        private static void Export(dynamic target, string pdf)
            => target.ExportAsFixedFormat(Type: Xl.xlTypePDF, Filename: pdf, Quality: Xl.xlQualityStandard,
                                          IncludeDocProperties: true, IgnorePrintAreas: false, OpenAfterPublish: false);

        private static List<string> PickExcelFiles(string description, out string folder)
        {
            folder = Ui.PickFolder(description, Settings.Current.BatchFolder);
            if (folder == null) return null;
            var files = Directory.GetFiles(folder, "*.xls*")
                                 .Where(f => !Path.GetFileName(f).StartsWith("~$")) // 跳过 Excel 临时锁文件
                                 .OrderBy(f => f).ToList();
            if (files.Count == 0) throw new UserMessage("在指定文件夹中没有找到任何 Excel 文件。");
            return files;
        }

        /// <summary>逐个打开处理后关闭；单个文件失败不中断，最后汇总。</summary>
        private static List<string> ForEachWorkbook(List<string> files, bool readOnly, Action<dynamic> work)
        {
            dynamic app = Xl.App;
            app.DisplayAlerts = false;
            var failed = new List<string>();
            foreach (var f in files)
            {
                if (Xl.FindOpenWorkbook(f) != null) { failed.Add(Path.GetFileName(f) + "（已打开，跳过）"); continue; }
                dynamic wb = null;
                try
                {
                    wb = app.Workbooks.Open(Filename: f, ReadOnly: readOnly, UpdateLinks: 0);
                    work(wb);
                }
                catch (Exception ex) { failed.Add(Path.GetFileName(f) + "：" + ex.Message); }
                finally { try { wb?.Close(SaveChanges: false); } catch { } }
            }
            return failed;
        }

        private static void Report(string what, int total, List<string> failed)
        {
            if (failed.Count == 0) Ui.Info($"{total} 个文件已全部{what}完成！");
            else Ui.Warn($"{what}完成 {total - failed.Count}/{total} 个，以下文件失败：\n" + string.Join("\n", failed));
        }
    }
}
