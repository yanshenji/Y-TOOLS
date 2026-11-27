using System.Drawing;
using System.Windows.Forms;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>辅助线输入窗口（替代原 UserForm 辅助线窗口）。</summary>
    internal sealed class HelperLineForm : Form
    {
        private readonly TextBox _name = new TextBox { Text = "标准", Location = new Point(70, 15), Width = 140 };
        private readonly TextBox _y = new TextBox { Location = new Point(70, 47), Width = 140 };

        private HelperLineForm()
        {
            Text = "添加辅助线";
            Font = new Font("Microsoft YaHei UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(235, 125);
            Controls.Add(new Label { Text = "名称", Location = new Point(15, 18), AutoSize = true });
            Controls.Add(new Label { Text = "Y 值", Location = new Point(15, 50), AutoSize = true });
            Controls.Add(_name);
            Controls.Add(_y);
            var ok = new Button { Text = "确定", Location = new Point(55, 85), Size = new Size(75, 28) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(138, 85), Size = new Size(75, 28) };
            ok.Click += (s, e) =>
            {
                if (!Vba.TryParseDouble(_y.Text, out _))
                {
                    MessageBox.Show(this, "输入的 Y 值无效，请输入一个数字！", Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    _y.SelectAll(); _y.Focus();
                    return;
                }
                DialogResult = DialogResult.OK;
            };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            Shown += (s, e) => _y.Focus();
        }

        public static bool Ask(out string name, out double y)
        {
            using (var f = new HelperLineForm())
            {
                name = null; y = 0;
                if (f.ShowDialog(Xl.Owner) != DialogResult.OK) return false;
                name = f._name.Text.Trim();
                return Vba.TryParseDouble(f._y.Text, out y);
            }
        }
    }
}
