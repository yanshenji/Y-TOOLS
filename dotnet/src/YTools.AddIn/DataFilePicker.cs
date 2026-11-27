using System.Drawing;
using System.Windows.Forms;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>按分组列出配置里的数据文件，双击或回车打开。功能区不可用时（如部分 WPS 版本）的入口。</summary>
    internal sealed class DataFilePicker : Form
    {
        private readonly TreeView _tree = new TreeView { Dock = DockStyle.Fill, Font = new Font("Microsoft YaHei UI", 10f), HideSelection = false };
        private DataFile _chosen;

        private DataFilePicker(AppConfig cfg)
        {
            Text = "打开数据文件";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(420, 520);
            MinimizeBox = MaximizeBox = false;
            var ok = new Button { Text = "打开", Dock = DockStyle.Bottom, Height = 36 };
            Controls.Add(_tree);
            Controls.Add(ok);
            foreach (var g in cfg.Groups)
            {
                var node = _tree.Nodes.Add(g.Key);
                foreach (var f in g) node.Nodes.Add(new TreeNode(f.Name) { Tag = f, ToolTipText = f.Path });
            }
            _tree.ShowNodeToolTips = true;
            _tree.ExpandAll();
            _tree.NodeMouseDoubleClick += (s, e) => Choose(e.Node);
            _tree.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) Choose(_tree.SelectedNode); };
            ok.Click += (s, e) => Choose(_tree.SelectedNode);
        }

        private void Choose(TreeNode n)
        {
            if (n?.Tag is DataFile f) { _chosen = f; DialogResult = DialogResult.OK; Close(); }
        }

        public static DataFile Pick(AppConfig cfg)
        {
            using (var dlg = new DataFilePicker(cfg))
                return dlg.ShowDialog(Xl.Owner) == DialogResult.OK ? dlg._chosen : null;
        }
    }
}
