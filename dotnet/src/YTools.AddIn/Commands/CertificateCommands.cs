using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 产品.GenerateQualityCertificate + 产品质量证明书 窗体。</summary>
    public static class CertificateCommands
    {
        private const int xlValues = -4163, xlPart = 2, xlWhole = 1, xlByRows = 1, xlNext = 1, xlOpenXMLWorkbook = 51;

        [ExcelCommand(Name = "工具宏_生成质量证明书", MenuName = "工具宏·产品", MenuText = "生成产品质量证明书…")]
        public static void Generate() => Cmd.Run("质量证明书", () =>
        {
            var cfg = Settings.Current;
            using (var form = new CertificateForm(cfg))
            {
                if (form.ShowDialog(Xl.Owner) != DialogResult.OK) return;
                var (xlsx, pdf) = Create(cfg, form.Batch, form.MetalQuantity, form.PackageWeight, form.Ledger);
                if (Ui.Confirm("质量证书已生成并保存在：\nExcel 文件：" + xlsx + "\nPDF 文件：" + pdf + "\n\n是否打开 PDF？"))
                    Process.Start(pdf);
            }
        }, freezeScreen: false);

        private static (string Xlsx, string Pdf) Create(AppConfig cfg, string batch, double metal, double pkgWeight, DataFile ledgerFile)
        {
            var cc = cfg.Certificate;
            var product = cc.FindProduct(batch) ?? throw new UserMessage("无法识别批号前缀：" + batch);
            var (net, count, _) = CertificateMath.Compute(product, metal, pkgWeight);
            dynamic app = Xl.App;

            // 检测台账：已打开就直接用，否则只读打开（保持打开，连续开证书时更快）
            dynamic ledger = Xl.FindOpenWorkbook(ledgerFile.Path);
            if (ledger == null)
            {
                FileCommands.Open(ledgerFile.Path, true, ledgerFile.Password);
                ledger = Xl.FindOpenWorkbook(ledgerFile.Path) ?? throw new UserMessage("无法打开检测台账：" + ledgerFile.Path);
            }
            dynamic sheet;
            try { sheet = ledger.Worksheets[product.SheetName]; }
            catch (Exception) { throw new UserMessage($"检测台账【{ledgerFile.Name}】中找不到工作表“{product.SheetName}”。"); }

            dynamic dateCell = sheet.Columns[1].Find(What: "日期", LookIn: xlValues, LookAt: xlPart);
            if (dateCell == null) throw new UserMessage("未找到检测数据工作簿中的表头行!（A 列中含“日期”的行）");
            int headerRow = dateCell.Row;
            dynamic batchHeader = sheet.Rows[headerRow].Find(What: "批号", LookIn: xlValues, LookAt: xlPart);
            if (batchHeader == null) throw new UserMessage("未找到检测数据工作簿中的批号列!");
            dynamic found = sheet.Columns[(int)batchHeader.Column].Find(What: batch, LookIn: xlValues, LookAt: xlWhole,
                                                                        SearchOrder: xlByRows, SearchDirection: xlNext, MatchCase: false);
            if (found == null) throw new UserMessage($"在【{ledgerFile.Name}】中未找到批号为 {batch} 的检测数据!");
            int dataRow = found.Row;

            string templatePath = cc.TemplatePath(product);
            if (!File.Exists(templatePath)) throw new UserMessage("无法打开质量证书模板!\n" + templatePath);

            string outDir = cc.OutputFolder;
            Directory.CreateDirectory(outDir);
            string baseName = CertificateMath.FileName(batch, count, net);
            string xlsx = CertificateConfig.JoinPath(outDir, baseName + ".xlsx");
            string pdf = CertificateConfig.JoinPath(outDir, baseName + ".pdf");
            if ((File.Exists(xlsx) || File.Exists(pdf)) && !Ui.Confirm("同名证书已存在，是否覆盖？\n" + baseName)) throw new UserMessage("已取消。", MessageBoxIcon.Information);

            dynamic template = null, cert = null;
            bool alerts = app.DisplayAlerts;
            try
            {
                app.ScreenUpdating = false;
                template = app.Workbooks.Open(Filename: templatePath, ReadOnly: true);
                template.Sheets[1].Copy();
                cert = app.ActiveWorkbook;
                dynamic cs = cert.Sheets[1];
                cs.Cells.Replace(What: "{{批号}}", Replacement: batch, LookAt: xlPart);
                cs.Cells.Replace(What: "{{包数}}", Replacement: count, LookAt: xlPart);
                cs.Cells.Replace(What: "{{净重}}", Replacement: net, LookAt: xlPart);

                int lastCol = Xl.LastCol(sheet, headerRow);
                object[,] headers = Xl.Read(sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow, lastCol]]);
                object[,] values = Xl.Read(sheet.Range[sheet.Cells[dataRow, 1], sheet.Cells[dataRow, lastCol]], useValue2: false);
                for (int c = 0; c < headers.GetLength(1); c++)
                {
                    if (SheetOps.IsBlank(headers[0, c]) || headers[0, c] is CellError) continue;
                    object v = values[0, c];
                    if (v is CellError) continue;
                    cs.Cells.Replace(What: CertificateMath.Placeholder(Vba.CStr(headers[0, c])), Replacement: v ?? "", LookAt: xlPart);
                }

                app.DisplayAlerts = false; // 已确认过覆盖
                cert.SaveAs(Filename: xlsx, FileFormat: xlOpenXMLWorkbook);
                cs.ExportAsFixedFormat(Type: Xl.xlTypePDF, Filename: pdf, Quality: Xl.xlQualityStandard,
                                       IncludeDocProperties: true, IgnorePrintAreas: false, OpenAfterPublish: false);
            }
            finally
            {
                try { cert?.Close(SaveChanges: false); } catch { }
                try { template?.Close(SaveChanges: false); } catch { }
                app.DisplayAlerts = alerts;
                app.ScreenUpdating = true;
            }
            return (xlsx, pdf);
        }
    }
}
