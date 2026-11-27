using System;
using System.Diagnostics;
using System.IO;
using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>
    /// 配置文件 工具宏.config.json：优先读 .xll 同目录，其次 %APPDATA%\工具宏\，都没有就用内置默认值。
    /// </summary>
    internal static class Settings
    {
        private static AppConfig _current;
        public static string SourcePath { get; private set; }

        public static AppConfig Current => _current ?? (_current = Load());

        private static string XllDir => Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        private static string AppDataDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "工具宏");

        public static AppConfig Load()
        {
            foreach (var dir in new[] { XllDir, AppDataDir })
            {
                var cfg = AppConfig.LoadFrom(dir, out var src);
                if (src != null) { SourcePath = src; return cfg; }
            }
            SourcePath = null;
            return AppConfig.Default();
        }

        public static void Reload() => _current = Load();

        /// <summary>打开配置文件编辑；不存在时先用默认值生成一份。</summary>
        public static void Edit()
        {
            string path = SourcePath;
            if (path == null)
            {
                path = TryWriteDefault(XllDir) ?? TryWriteDefault(AppDataDir)
                       ?? throw new UserMessage("无法创建配置文件（目录没有写入权限）。");
                SourcePath = path;
            }
            Process.Start("notepad.exe", "\"" + path + "\"");
        }

        private static string TryWriteDefault(string dir)
        {
            try
            {
                Directory.CreateDirectory(dir);
                var p = Path.Combine(dir, AppConfig.FileName);
                File.WriteAllText(p, AppConfig.DefaultJson(), new System.Text.UTF8Encoding(false));
                return p;
            }
            catch { return null; }
        }
    }
}
