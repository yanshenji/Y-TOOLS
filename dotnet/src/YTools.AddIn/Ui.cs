using System;
using System.IO;
using System.Windows.Forms;

namespace YTools.AddIn
{
    internal static class Ui
    {
        public const string Title = "工具宏";

        public static void Info(string text) => MessageBox.Show(Xl.Owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        public static void Warn(string text) => MessageBox.Show(Xl.Owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        public static bool Confirm(string text)
            => MessageBox.Show(Xl.Owner, text, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

        public static string PickFolder(string description, string initial)
        {
            using (var d = new FolderBrowserDialog { Description = description, ShowNewFolderButton = false })
            {
                if (!string.IsNullOrEmpty(initial) && Directory.Exists(initial)) d.SelectedPath = initial;
                return d.ShowDialog(Xl.Owner) == DialogResult.OK ? d.SelectedPath : null;
            }
        }

        /// <summary>Excel 原生的选区输入框（Application.InputBox Type:=8），取消返回 null。</summary>
        public static dynamic PickRange(string prompt)
        {
            object r = Xl.App.InputBox(prompt, Title, Type: 8);
            return r is bool ? null : r;
        }

        public static void CopyText(string text) => Clipboard.SetText(text);
    }

    /// <summary>
    /// 所有命令的统一外壳：关屏幕刷新、捕获异常弹框、结束后恢复 Excel 状态。
    /// 原 VBA 中途出错会把 ScreenUpdating/EnableEvents 留在关闭状态，这里保证一定恢复。
    /// </summary>
    internal static class Cmd
    {
        public static void Run(string name, Action action, bool freezeScreen = true)
        {
            dynamic app = Xl.App;
            bool screen = true, events = true, alerts = true;
            try { screen = app.ScreenUpdating; events = app.EnableEvents; alerts = app.DisplayAlerts; } catch { }
            try
            {
                if (freezeScreen) app.ScreenUpdating = false;
                action();
            }
            catch (UserMessage m)
            {
                Restore(app, screen, events, alerts);
                MessageBox.Show(Xl.Owner, m.Message, Ui.Title + " - " + name, MessageBoxButtons.OK, m.Icon);
            }
            catch (Exception ex)
            {
                Restore(app, screen, events, alerts);
                var e = ex is System.Reflection.TargetInvocationException t && t.InnerException != null ? t.InnerException : ex;
                MessageBox.Show(Xl.Owner, "执行出错：" + e.Message, Ui.Title + " - " + name, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Restore(app, screen, events, alerts);
            }
        }

        private static void Restore(dynamic app, bool screen, bool events, bool alerts)
        {
            try { app.ScreenUpdating = screen; app.EnableEvents = events; app.DisplayAlerts = alerts; } catch { }
        }
    }
}
