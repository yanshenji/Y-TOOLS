using System;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace YTools.AddIn
{
    /// <summary>环境信息：用于在 Excel / WPS 中做冒烟测试，确认加载、函数、命令、功能区是否可用。</summary>
    public static class Diagnostics
    {
        public static string VersionText =>
            "工具宏 " + typeof(Diagnostics).Assembly.GetName().Version + (Environment.Is64BitProcess ? " (64 位)" : " (32 位)");

        /// <summary>功能区回调 onLoad 是否执行过（WPS 个人版可能不加载 Excel-DNA 功能区）。</summary>
        internal static bool RibbonLoaded;

        [ExcelCommand(Name = "工具宏_环境信息", MenuName = "工具宏", MenuText = "环境信息")]
        public static void ShowEnvironment()
        {
            var sb = new StringBuilder();
            sb.AppendLine(VersionText);
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                sb.AppendLine("宿主程序：" + app.Name);
                sb.AppendLine("宿主版本：" + app.Version + "  Build " + SafeGet(() => app.Build));
                sb.AppendLine("操作系统：" + app.OperatingSystem);
            }
            catch (Exception ex)
            {
                sb.AppendLine("读取 Application 对象失败：" + ex.Message);
            }
            sb.AppendLine(".NET：" + Environment.Version);
            sb.AppendLine("加载路径：" + ExcelDnaUtil.XllPath);
            sb.AppendLine("功能区：" + (RibbonLoaded ? "已加载" : "未加载（可用“加载项”菜单）"));
            MessageBox.Show(sb.ToString(), "工具宏 - 环境信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string SafeGet(Func<object> f)
        {
            try { return Convert.ToString(f()); } catch { return "?"; }
        }
    }
}
