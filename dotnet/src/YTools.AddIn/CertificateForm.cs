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
        private readonly AppConfig _cfg;

        public string Batch => _batch.Text.Trim();
        public double MetalQuantity { get; private set; }
        public double PackageWeight { get; private set; }
        public DataFile Ledger => _ledger.SelectedItem as DataFile;

        public CertificateForm(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "产品质量证明书";
            Font = new Font("Microsoft YaHei UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(380, 250);

            var grid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, Padding = new Padding(12), AutoSize = true };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            void Row(string label, Control c) { grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }); c.Dock = DockStyle.Fill; grid.Controls.Add(c); }
            Row("批号", _batch);
            Row("", _product);
            Row("金属量(kg)", _metal);
            Row("每包重量(kg)", _weight);
            Row("检测台账", _ledger);
            Controls.Add(grid);

            _weight.Items.AddRange(new object[] { "500", "750", "1000" });
            _weight.Text = "1000";
            var ledgers = cfg.DataFiles.Where(f => f.Group == cfg.Certificate.LedgerGroup).ToList();
            _ledger.DisplayMember = nameof(DataFile.Name);
            foreach (var f in ledgers) _ledger.Items.Add(f);
            _ledger.SelectedItem = CertificateMath.PickLedger(ledgers, DateTime.Today.Year);
            _batch.TextChanged += (s, e) => OnBatchChanged();
            OnBatchChanged();

            var ok = new Button { Text = "生成", DialogResult = DialogResult.None, Size = new Size(90, 30), Location = new Point(170, 205) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Size = new Size(90, 30), Location = new Point(270, 205) };
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
            if (p == null) { _product.Text = Batch.Length == 0 ? "输入批号后自动识别产品" : "未识别的批号前缀"; return; }
            _product.Text = $"产品：{p.Type}（金属含量 {p.MetalContent.ToString(CultureInfo.InvariantCulture)}）";
            _weight.Text = p.PackageWeight.ToString(CultureInfo.InvariantCulture);
        }

        private void Submit()
        {
            if (Batch.Length == 0) { Warn("请输入批号。", _batch); return; }
            if (_cfg.Certificate.FindProduct(Batch) == null) { Warn("无法识别批号前缀，请检查批号或配置文件中的 products。", _batch); return; }
            if (!Vba.TryParseDouble(_metal.Text, out var m) || m <= 0) { Warn("金属量请输入大于 0 的数字。", _metal); return; }
            if (!Vba.TryParseDouble(_weight.Text, out var w) || w <= 0) { Warn("每包重量请输入大于 0 的数字。", _weight); return; }
            if (Ledger == null) { Warn("配置文件中没有可用的检测台账（dataFiles 中分组为“" + _cfg.Certificate.LedgerGroup + "”的文件）。", _ledger); return; }
            MetalQuantity = m;
            PackageWeight = w;
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
