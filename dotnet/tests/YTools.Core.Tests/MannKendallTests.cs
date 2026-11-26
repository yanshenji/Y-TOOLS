using Xunit;

namespace YTools.Core.Tests;

public class MannKendallTests
{
    private static readonly object[,] Rising = Col.Of(1.0, 2.0, 3.0, 4.0, 5.0);
    private static readonly object[,] WithTies = Col.Of(3.0, 1.0, 4.0, 1.0, 5.0, 9.0, 2.0, 6.0);

    [Fact]
    public void 单调上升()
    {
        Assert.Equal(10.0, MannKendall.Compute(Rising, 1));
        Assert.Equal(2.2045407685048604, (double)MannKendall.Compute(Rising, 2), 10);
        Assert.Equal(0.0274863361115103, (double)MannKendall.Compute(Rising, 3), 8);
        Assert.Equal(1.0, MannKendall.Compute(Rising, 4));
        Assert.Equal(5.0, MannKendall.Compute(Rising, 5));
    }

    [Fact]
    public void 含相同值的修正()
    {
        Assert.Equal(11.0, MannKendall.Compute(WithTies, 1));
        Assert.Equal(64.33333333333333, MannKendall.VarS(new[] { 3.0, 1, 4, 1, 5, 9, 2, 6 }), 10);
        Assert.Equal(1.2467574523850742, (double)MannKendall.Compute(WithTies, 2), 10);
        Assert.Equal(0.21248644570250086, (double)MannKendall.Compute(WithTies, 3), 8);
        Assert.Equal(0.4642857142857143, (double)MannKendall.Compute(WithTies, 4), 12);
    }

    [Fact]
    public void 少于4个点返回错误_但仍可取样本数()
    {
        var three = Col.Of(1.0, 2.0, "x", 3.0);
        Assert.Same(CellError.Value, MannKendall.Compute(three, 2));
        Assert.Equal(3.0, MannKendall.Compute(three, 5));
    }

    [Fact]
    public void 全无数值返回错误() => Assert.Same(CellError.Value, MannKendall.Compute(Col.Of("a", null), 5));
}
