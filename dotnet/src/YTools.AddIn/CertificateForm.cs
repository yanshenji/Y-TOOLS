using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using YTools.Core;

namespace YTools.AddIn
{
    /// <summary>产品质量证明书输入窗口（替代原 UserForm 产品质量证明书）。</summary>
    internal sealed class CertificateForm : Form
    {
        private readonly TextBox _batch = new TextBox();
        private readonly TextBox _metal = new TextBox();
        private readonly ComboBox _weight = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
        private readonly ComboBox _ledger = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly Label _product = new Label { AutoSize = true, ForeColor = Color.DimGray };
        private readonly Label _preview = new Label { AutoSize = true, ForeColor = Color.DimGray };
        private readonly AppConfig _cfg;

        public string Batch => _batch.Text.Trim();
        public double MetalQuantity { get; private set; }
        public double PackageWeight { get; private set; }
        /// <summary>写入证书与文件名的包数（整数）。</summary>
        public int Packages { get; private set; }
        public DataFile Ledger => _ledger.SelectedItem as DataFile;

        public CertificateForm(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "产品质量证明书";
            Font = new Font("Microsoft YaHei UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(380, 280);

            var grid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, Padding = new Padding(12), AutoSize = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            void Row(string label, Control c) { grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }); c.Dock = DockStyle.Fill; grid.Controls.Add(c); }
            Row("批号", _batch);
            Row("", _product);
            Row("金属量(kg)", _metal);
            Row("每包重量(kg)", _weight);
            Row("", _preview);
            Row("检测台账", _ledger);
            Controls.Add(grid);

            _weight.Items.AddRange(new object[] { "500", "750", "1000" });
            _weight.Text = "1000";
            var ledgers = cfg.DataFiles.Where(f => f.Group == cfg.Certificate.LedgerGroup).ToList();
            _ledger.DisplayMember = nameof(DataFile.Name);
            foreach (var f in ledgers) _ledger.Items.Add(f);
            _ledger.SelectedItem = CertificateMath.PickLedger(ledgers, DateTime.Today.Year);
            _batch.TextChanged += (s, e) => OnBatchChanged();
            _metal.TextChanged += (s, e) => UpdatePreview();
            _weight.TextChanged += (s, e) => UpdatePreview();
            OnBatchChanged();
            UpdatePreview();

            var ok = new Button { Text = "生成", DialogResult = DialogResult.None, Size = new Size(90, 30), Location = new Point(170, 235) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Size = new Size(90, 30), Location = new Point(270, 235) };
            ok.Click += (s, e) => Submit();
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <summary>输入批号后自动识别产品，并把每包重量设为该产品的默认值（原窗体固定默认 1000）。</summary>
        private void OnBatchChanged()
        {
            var p = _cfg.Certificate.FindProduct(Batch);
            if (p == null) { _product.Text = Batch.Length == 0 ? "输入批号后自动识别产品" : "未识别的批号前缀"; UpdatePreview(); return; }
            _product.Text = $"产品：{p.Type}（金属含量 {p.MetalContent.ToString(CultureInfo.InvariantCulture)}）";
            _weight.Text = p.PackageWeight.ToString(CultureInfo.InvariantCulture);
            UpdatePreview();
        }

        /// <summary>实时显示净重与包数；包数不是整数时用橙色提醒。</summary>
        private void UpdatePreview()
        {
            var p = _cfg.Certificate.FindProduct(Batch);
            if (p == null || !Vba.TryParseDouble(_metal.Text, out var m) || m <= 0 || !Vba.TryParseDouble(_weight.Text, out var w) || w <= 0)
            { _preview.Text = ""; return; }
            var (net, count, _) = CertificateMath.Compute(p, m, w);
            bool whole = CertificateMath.TryWholePackages(count, out int n);
            _preview.ForeColor = whole ? Color.DimGray : Color.DarkOrange;
            _preview.Text = whole
                ? $"净重 {net:0.##} kg，{n} 包"
                : $"净重 {net:0.##} kg，包数 {count:0.###} 不是整数";
        }

        private void Submit()
        {
            if (Batch.Length == 0) { Warn("请输入批号。", _batch); return; }
            if (_cfg.Certificate.FindProduct(Batch) == null) { Warn("无法识别批号前缀，请检查批号或配置文件中的 products。", _batch); return; }
            if (!Vba.TryParseDouble(_metal.Text, out var m) || m <= 0) { Warn("金属量请输入大于 0 的数字。", _metal); return; }
            if (!Vba.TryParseDouble(_weight.Text, out var w) || w <= 0) { Warn("每包重量请输入大于 0 的数字。", _weight); return; }
            if (Ledger == null) { Warn("配置文件中没有可用的检测台账（dataFiles 中分组为“" + _cfg.Certificate.LedgerGroup + "”的文件）。", _ledger); return; }
            var (net, count, _) = CertificateMath.Compute(_cfg.Certificate.FindProduct(Batch), m, w);
            if (!CertificateMath.TryWholePackages(count, out int n))
            {
                var msg = $"按金属量 {m:0.###} kg、每包 {w:0.###} kg 计算：\n净重 {net:0.##} kg，包数 {count:0.###}，不是整数。\n\n"
                        + $"是：按 {n} 包生成（最后一包约 {net - (n - 1) * w:0.##} kg）\n否：返回修改";
                if (MessageBox.Show(this, msg, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) { _metal.Focus(); return; }
            }
            MetalQuantity = m;
            PackageWeight = w;
            Packages = n;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void Warn(string text, Control focus)
        {
            MessageBox.Show(this, text, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            focus.Focus();
        }
    }
}
