using Xunit;

namespace YTools.Core.Tests;

public class GrubbsTests
{
    // 双侧 α=0.05 Grubbs 临界值表
    [Theory]
    [InlineData(3, 1.1543)]
    [InlineData(10, 2.2900)]
    [InlineData(20, 2.7082)]
    [InlineData(30, 2.9085)]
    public void 临界值与查表一致(int n, double expected)
        => Assert.Equal(expected, (double)Grubbs.Critical(n, 0.05), 4);

    [Fact]
    public void 临界值_n不大于2() => Assert.Equal("n 必须大于 2", Grubbs.Critical(2));

    [Fact]
    public void 临界值_alpha无效抛错() => Assert.Throws<VbaRuntimeException>(() => Grubbs.Critical(10, 0));

    [Fact]
    public void 检出单个异常值()
    {
        var data = Col.Of(2.1, 2.3, 2.2, 2.4, 2.2, 2.3, 2.1, 2.2, 2.3, 5.0);
        var res = Assert.IsType<object[,]>(Grubbs.Outliers(data));
        Assert.Equal(1, res.GetLength(0));
        Assert.Equal(5.0, res[0, 0]);
    }

    [Fact]
    public void 无异常值()
        => Assert.Equal(Grubbs.NoOutliers, Grubbs.Outliers(Col.Of(1.0, 2.0, 3.0, 4.0, 5.0)));

    [Fact]
    public void 数据点不足()
        => Assert.Equal(Grubbs.TooFewPoints, Grubbs.Outliers(Col.Of(1.0, "a", 2.0)));

    [Fact]
    public void G值()
    {
        var nums = new[] { 2.1, 2.3, 2.2, 2.4, 2.2, 2.3, 2.1, 2.2, 2.3, 5.0 };
        Assert.Equal(2.8296672575826625, (double)Grubbs.G(5.0, nums), 10);
        Assert.Equal("标准差为 0", Grubbs.G(1, new[] { 1.0, 1.0 }));
    }
}
