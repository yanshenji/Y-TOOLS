using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>原 Openfiles 模块：打开共享盘数据文件。文件清单在配置文件中维护。</summary>
    public static class FileCommands
    {
        private const string Menu = "工具宏·文件";

        internal static void OpenDataFile(DataFile f) => Cmd.Run("打开 " + f.Name, () => Open(f.Path, f.ReadOnly, f.Password), freezeScreen: false);

        internal static void Open(string path, bool readOnly, string password)
        {
            var already = Xl.FindOpenWorkbook(path);
            if (already != null) { already.Activate(); return; }
            if (!File.Exists(path)) throw new UserMessage("找不到文件，请检查网络盘或配置：\n" + path);
            if (string.IsNullOrEmpty(password))
                Xl.App.Workbooks.Open(Filename: path, ReadOnly: readOnly);
            else
                Xl.App.Workbooks.Open(Filename: path, ReadOnly: readOnly, Password: password);
        }

        [ExcelCommand(Name = "工具宏_打开数据文件", MenuName = Menu, MenuText = "打开数据文件…")]
        public static void PickAndOpen() => Cmd.Run("打开数据文件", () =>
        {
            var f = DataFilePicker.Pick(Settings.Current);
            if (f != null) Open(f.Path, f.ReadOnly, f.Password);
        }, freezeScreen: false);

        [ExcelCommand(Name = "工具宏_打开中控检测结果分类", MenuName = Menu, MenuText = "打开中控检测结果(分类)")]
        public static void OpenFileSet() => Cmd.Run("打开文件组", () =>
        {
            var set = FirstSet();
            var missing = new List<string>();
            foreach (var p in set.FullPaths.Distinct())
            {
                if (!File.Exists(p)) { missing.Add(Path.GetFileName(p)); continue; }
                Open(p, set.ReadOnly, null);
            }
            if (missing.Count > 0) Ui.Warn("以下文件不存在，已跳过：\n" + string.Join("\n", missing));
        }, freezeScreen: false);

        [ExcelCommand(Name = "工具宏_保存关闭中控检测结果分类", MenuName = Menu, MenuText = "保存并关闭中控检测结果(分类)")]
        public static void SaveCloseFileSet() => Cmd.Run("保存并关闭文件组", () =>
        {
            int n = 0;
            foreach (var p in FirstSet().FullPaths.Distinct())
            {
                var wb = Xl.FindOpenWorkbook(p);
                if (wb == null) continue;
                wb.Save();
                wb.Close(SaveChanges: false);
                n++;
            }
            Ui.Info("处理" + n + "个文件完毕");
        });

        private static FileSet FirstSet()
            => Settings.Current.FileSets.FirstOrDefault() ?? throw new UserMessage("配置文件中没有 fileSets。");

        [ExcelCommand(Name = "工具宏_编辑配置", MenuName = Menu, MenuText = "编辑配置文件")]
        public static void EditConfig() => Cmd.Run("编辑配置", Settings.Edit, freezeScreen: false);

        [ExcelCommand(Name = "工具宏_重新加载配置", MenuName = Menu, MenuText = "重新加载配置")]
        public static void ReloadConfig() => Cmd.Run("重新加载配置", () =>
        {
            Settings.Reload();
            YToolsRibbon.Refresh();
            Ui.Info("已加载：" + (Settings.SourcePath ?? "内置默认配置") + "\n数据文件 " + Settings.Current.DataFiles.Count + " 个");
        }, freezeScreen: false);
    }
}
