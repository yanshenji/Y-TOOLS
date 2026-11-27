using System.Linq;
using Xunit;

namespace YTools.Core.Tests;

public class ChartMathTests
{
    [Fact]
    public void 可画点_跳过隐藏行文本空值()
    {
        var y = new object[] { 1.0, "未测", null, " ", "2.5", 3.0, 4.0 };
        var hidden = new[] { false, false, false, false, false, true, false };
        Assert.Equal(new[] { 0, 4, 6 }, ChartMath.PlottableRows(y, hidden));
    }

    [Fact]
    public void 单值控制限()
    {
        var (cl, ucl, lcl, ooc) = ChartMath.IndividualsLimits(new[] { 10.0, 12, 11, 13, 12, 30 });
        Assert.Equal(88.0 / 6, cl, 10);
        double mr = (2 + 1 + 2 + 1 + 18) / 5.0;
        Assert.Equal(cl + 3 * mr / 1.128, ucl, 10);
        Assert.Equal(cl - 3 * mr / 1.128, lcl, 10);
        Assert.Equal(new[] { false, false, false, false, false, true }, ooc); // UCL≈27.43
        var (_, _, _, ooc2) = ChartMath.IndividualsLimits(new[] { 10.0, 10.1, 9.9, 10, 10.1, 9.9, 10, 10.1, 9.9, 12 });
        Assert.True(ooc2.Last());
    }

    [Fact]
    public void 坐标范围()
    {
        Assert.Equal((-0.5, 10.5), ChartMath.PaddedRange(0, 10));
        var (a, b) = ChartMath.PaddedRange(5, 5);
        Assert.Equal(4.975, a, 10); Assert.Equal(5.025, b, 10);
        Assert.Equal((-1.0, 11.0), ChartMath.AlignToUnit(-0.5, 10.5, 1));
        Assert.Equal((-4.0, -2.0), ChartMath.AlignToUnit(-3.4, -2.3, 1)); // 原 VBA 得到 (-3, -3)，把数据截掉
        Assert.Equal((2.0, 12.0), ChartMath.AlignToUnit(2.3, 11.2, 2));
    }

    [Fact]
    public void 标题回退顺序()
    {
        Assert.Equal("[中控 - 101 - Co]趋势图", ChartMath.Title("中控", "101", "A", "Co", "趋势图"));
        Assert.Equal("[中控 - A - Co]趋势图", ChartMath.Title("中控", "未知样品", "A", "Co", "趋势图"));
        Assert.Equal("[中控 - Co]控制图", ChartMath.Title("中控", null, "样品名称", "Co", "控制图"));
    }

    [Fact]
    public void 钴杂比()
        => Assert.Equal(new[] { 2.0, 0, 0, 5 }, ChartMath.CobaltRatio(new object[] { 10.0, 5.0, "x", 10.0 }, new object[] { 5.0, 0.0, 1.0, "2" }));

    [Fact]
    public void 辅助线等分()
    {
        var (x, y) = ChartMath.HelperLine(new[] { 3.0, 1, 5 }, 7);
        Assert.Equal(new[] { 1.0, 3, 5 }, x);
        Assert.All(y, v => Assert.Equal(7, v));
        Assert.Equal(new[] { 4.0 }, ChartMath.HelperLine(new[] { 4.0 }, 1).X); // 原 VBA 单点会除以 0
    }

    [Fact]
    public void 网格位置()
    {
        Assert.Equal((20.0, 20.0), ChartMath.GridPosition(0, 2, 100, 50));
        Assert.Equal((135.0, 20.0), ChartMath.GridPosition(1, 2, 100, 50));
        Assert.Equal((20.0, 82.0), ChartMath.GridPosition(2, 2, 100, 50));
    }
}

public class RuleChartTests
{
    [Theory]
    [InlineData("101槽 出口液", "出口", true)]
    [InlineData("101槽 出口液", "102,出口", true)]
    [InlineData("101槽 出口液", "101&出口", true)]
    [InlineData("101槽 出口液", "101&进口", false)]
    [InlineData("ABC", "abc", true)]
    [InlineData("A12", "A##", true)]
    [InlineData("A12", "a?", false)]
    [InlineData("任意", "*", true)]
    [InlineData("任意", "", true)]
    public void 名称条件(string text, string expr, bool expected) => Assert.Equal(expected, RuleCharts.NameMatches(text, expr));

    [Theory]
    [InlineData(12.345, "12.35")]
    [InlineData(1, "1.00")]
    [InlineData(0.012345, "0.012")]
    [InlineData(0.5, "0.5")]
    [InlineData(-0.0456, "-0.046")]
    [InlineData(0, "0")]
    public void 智能数字(double v, string expected) => Assert.Equal(expected, RuleCharts.SmartNumber(v));

    [Fact]
    public void 工作表重名加序号()
    {
        Assert.Equal("A_B", RuleCharts.UniqueSheetName("A/B", new string[0]));
        Assert.Equal("A_B2", RuleCharts.UniqueSheetName("A/B", new[] { "a_b" }));
        Assert.Equal("A_B3", RuleCharts.UniqueSheetName("A/B", new[] { "A_B", "A_B2" }));
        var longName = new string('x', 40);
        Assert.Equal(new string('x', 30) + "2", RuleCharts.UniqueSheetName(longName, new[] { new string('x', 31) }));
    }

    [Fact]
    public void 规则行()
    {
        var r = ChartRule.FromRow(new object[] { " 中控 ", "", "出口", 46000.0, null, "Co, Ni ,," });
        Assert.Equal("中控", r.SheetName);
        Assert.Equal("出口", r.Label);
        Assert.Equal(46000.0, r.StartDate);
        Assert.Null(r.EndDate);
        Assert.Equal(new[] { "Co", "Ni" }, r.ItemList);
        Assert.Null(ChartRule.FromRow(new object[] { "" }));
    }

    [Fact]
    public void 日期主刻度() { Assert.Equal(1, RuleCharts.DateMajorUnit(5)); Assert.Equal(4, RuleCharts.DateMajorUnit(31)); }
}
