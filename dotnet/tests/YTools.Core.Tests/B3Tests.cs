using System;
using System.Linq;
using Xunit;

namespace YTools.Core.Tests;

public class AnomalyTests
{
    private static object[] Cells(params object[] v) => v;

    [Fact]
    public void 四分位数与Excel一致()
    {
        var x = Enumerable.Range(1, 10).Select(i => (double)i).ToArray();
        Assert.Equal(3.25, Anomaly.QuartileInc(x, 1), 12);
        Assert.Equal(7.75, Anomaly.QuartileInc(x, 3), 12);
        Assert.Equal(5.5, Anomaly.QuartileInc(x, 2), 12);
    }

    [Fact]
    public void 箱线图_标记高值_空格与表头只清色()
    {
        var c = Cells("表头", 10.0, 11.0, 10.5, 10.2, 10.8, 30.0, null);
        var r = Anomaly.Boxplot(c, out double lo, out double hi);
        Assert.Equal(CellMark.Clear, r.Marks[0]);
        Assert.Equal(CellMark.Flag, r.Marks[6]);
        Assert.Equal(CellMark.Clear, r.Marks[7]); // 原 VBA 会把空格标黄
        Assert.Equal(1, r.Flagged);
        Assert.True(lo > 0 && hi < 30);
    }

    [Fact]
    public void 三西格玛_非数值不动()
    {
        var vals = Enumerable.Repeat(10.0, 20).Select((v, i) => (object)(v + (i % 2) * 0.1)).ToList();
        vals.Add(50.0);
        vals.Insert(0, "表头");
        var r = Anomaly.ThreeSigma(vals);
        Assert.Equal(CellMark.Keep, r.Marks[0]);
        Assert.Equal(CellMark.Flag, r.Marks.Last());
        Assert.Equal(1, r.Flagged);
    }

    [Fact]
    public void Grubbs单侧_临界值()
    {
        var c = Cells("h", 2.1, 2.3, 2.2, 2.4, 2.2, 2.3, 2.1, 2.2, 2.3, 5.0);
        var r = Anomaly.GrubbsOneSided(c);
        Assert.Contains("单侧Gcrit=2.1761", r.Info); // 标准单侧临界值表 n=10、α=0.05：2.176（原 VBA 为 1.7984）
        Assert.Equal(CellMark.Flag, r.Marks[10]);
        Assert.Equal(CellMark.Clear, r.Marks[0]);
    }

    // 前 20 点稳定在 10 附近，之后漂移到 11（参考值由独立的 Python 实现算出）
    [Fact]
    public void Grubbs单侧_正常数据不再误标()
    {
        // 原 VBA 临界值约 1.8，这组数据里 10.3 的 G≈1.81 会被误标；按标准临界值（约 2.18）不应标出
        var c = Cells(10.0, 10.2, 9.9, 10.1, 10.0, 9.8, 10.3, 10.1, 10.0, 9.9);
        Assert.Equal(0, Anomaly.GrubbsOneSided(c).Flagged);
    }

    private static readonly object[] Shift = { 10, 10.2, 9.9, 10.1, 10, 9.8, 10.3, 10.1, 10, 9.9, 10.2, 10, 10.1, 9.9, 10, 10.2, 9.8, 10.1, 10, 10.1, 10.8, 10.9, 11, 10.8, 10.9, 11.1 };
    private static object[] ShiftCells() => Shift.Select(v => (object)Convert.ToDouble(v)).ToArray();

    [Fact]
    public void CUSUM_均值漂移()
    {
        var r = Anomaly.Cusum(ShiftCells());
        Assert.Equal(new[] { 24, 25 }, r.Marks.Select((m, i) => (m, i)).Where(x => x.m == CellMark.Flag).Select(x => x.i));
        Assert.Contains("报警点数=2", r.Info);
    }

    [Fact]
    public void EWMA_均值漂移()
    {
        var r = Anomaly.Ewma(ShiftCells());
        Assert.Equal(new[] { 25 }, r.Marks.Select((m, i) => (m, i)).Where(x => x.m == CellMark.Flag).Select(x => x.i));
    }

    [Fact]
    public void 样本不足跳过()
    {
        Assert.True(Anomaly.Cusum(Cells(1.0, 2.0, 3.0, 4.0)).Skipped);
        Assert.True(Anomaly.GrubbsOneSided(Cells(1.0, 2.0)).Skipped);
        Assert.True(Anomaly.ThreeSigma(Cells(5.0, 5.0, 5.0)).Skipped);
    }

    [Fact]
    public void MK汇总()
    {
        var s = Anomaly.MannKendallSummary(Cells("表头", 1.0, 2.0, 3.0, 4.0, 5.0));
        Assert.Contains("n=5", s);
        Assert.Contains("S=10  Z=2.2045  p=0.0275", s);
        Assert.Contains("Sen斜率=1  ▲ 显著上升", s);
        Assert.Contains("无显著趋势", Anomaly.MannKendallSummary(Cells(1.0, 3.0, 2.0, 2.5)));
    }
}

public class CertificateTests
{
    private static readonly CertificateConfig Cfg = AppConfig.Default().Certificate;

    [Theory]
    [InlineData("LHG2611001", "氯化钴")]
    [InlineData("TSG-001", "碳酸钴")]
    [InlineData("ALLOY260101", "2108电钴")]
    [InlineData("HR06A0123", "四钴")]
    public void 按前缀识别产品(string batch, string type) => Assert.Equal(type, Cfg.FindProduct(batch).Type);

    [Fact]
    public void 前缀区分大小写() => Assert.Null(Cfg.FindProduct("lhg001"));

    [Fact]
    public void 净重与包数()
    {
        var p = Cfg.FindProduct("TSG1");
        var (net, count, w) = CertificateMath.Compute(p, 465, 0); // 0 → 用默认 500
        Assert.Equal(1000, net, 9);
        Assert.Equal(2, count, 9);
        Assert.Equal(500, w);
        Assert.Equal("COA - TSG1 - 2 包 - 1000kg", CertificateMath.FileName("TSG1", count, net));
    }

    [Theory]
    [InlineData(2.0, true, 2)]
    [InlineData(1.9999999999, true, 2)]   // 浮点误差
    [InlineData(3.2258064516129, false, 4)]
    [InlineData(0.4, false, 1)]
    public void 包数取整(double count, bool whole, int expected)
    {
        Assert.Equal(whole, CertificateMath.TryWholePackages(count, out int n));
        Assert.Equal(expected, n);
    }

    [Fact]
    public void 模板路径与工作表名()
    {
        var p = Cfg.FindProduct("ALLOY1");
        Assert.EndsWith(@"产品质量证明书模板\2108电钴钴产品质量证明书模板.xlsx", Cfg.TemplatePath(p));
        Assert.Equal("（2108）电钴检测记录表 ", p.SheetName); // 原表名末尾有空格
    }

    [Fact]
    public void 占位符大写() => Assert.Equal("{{CO(%)}}", CertificateMath.Placeholder("Co(%)"));

    [Fact]
    public void 台账按年份选择()
    {
        var g = AppConfig.Default().DataFiles.Where(f => f.Group == "成品数据").ToList();
        Assert.Equal("2026年", CertificateMath.PickLedger(g, 2026).Name);
        Assert.Equal("2024年", CertificateMath.PickLedger(g, 2024).Name);
        Assert.Equal("2026年", CertificateMath.PickLedger(g, 2030).Name);
    }
}

public class ConfigCompatTests
{
    [Fact]
    public void 旧配置缺少证书段时用默认值补齐()
    {
        var dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "yt_cfg_" + Guid.NewGuid().ToString("N"));
        System.IO.Directory.CreateDirectory(dir);
        System.IO.File.WriteAllText(System.IO.Path.Combine(dir, AppConfig.FileName),
            "{\"dataFiles\":[{\"group\":\"成品数据\",\"name\":\"2027年\",\"path\":\"x.xlsx\"}]}");
        var cfg = AppConfig.LoadFrom(dir, out var src);
        Assert.NotNull(src);
        Assert.Single(cfg.DataFiles);                       // 用户写的保留
        Assert.Equal(6, cfg.Certificate.Products.Count);    // 缺的补默认
        Assert.Single(cfg.FileSets);
        System.IO.Directory.Delete(dir, true);
    }
}
