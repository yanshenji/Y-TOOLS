using Xunit;

namespace YTools.Core.Tests;

public class CapabilityTests
{
    private static readonly object[,] Data = Col.Of(10.0, 12.0, 11.0, 13.0, 12.0);

    [Fact]
    public void CpkMR_双侧取较小值()
        => Assert.Equal(0.9024, (double)Capability.CpkMR(Data, 8.0, 16.0), 10);

    [Fact]
    public void CpkMR_只有上限()
        => Assert.Equal(4.4 / (3 * 1.5 / 1.128), (double)Capability.CpkMR(Data, null, 16.0), 10);

    [Fact]
    public void CpkMR_规格都缺失返回999()
        => Assert.Equal(999.0, Capability.CpkMR(Data, null, "无"));

    [Fact]
    public void CpkMR_跳过文本与空格_接受文本型数字()
        => Assert.Equal(0.9024, (double)Capability.CpkMR(Col.Of(10.0, "未测", null, "12", 11.0, 13.0, 12.0), 8.0, 16.0), 10);

    [Fact]
    public void CpkMR_样本不足()
        => Assert.Equal(Capability.InsufficientSample, Capability.CpkMR(Col.Of(1.0, "x"), 0.0, 2.0));

    [Fact]
    public void CpkMR_无波动()
        => Assert.Equal(Capability.NoVariation, Capability.CpkMR(Col.Of(5.0, 5.0, 5.0), 0.0, 10.0));

    [Fact]
    public void Ppk_样本标准差()
        => Assert.Equal(1.052469623168435, (double)Capability.Ppk(Data, 8.0, 16.0), 10);

    [Fact]
    public void CpkMR_只取第一列()
    {
        var two = new object[,] { { 10.0, 999.0 }, { 12.0, 999.0 }, { 11.0, 0.0 }, { 13.0, 0.0 }, { 12.0, 0.0 } };
        Assert.Equal(0.9024, (double)Capability.CpkMR(two, 8.0, 16.0), 10);
    }

    [Theory]
    [InlineData(0.5, "D级-严重不足")]
    [InlineData(0.67, "C级-能力不足")]
    [InlineData(1.0, "B级-临界尚可")]
    [InlineData(1.33, "A级-能力充足")]
    [InlineData(1.67, "AA级-能力富余")]
    [InlineData(2.5, "AA级-能力富余")]
    public void Rank_分级边界(double cpk, string expected) => Assert.Equal(expected, Capability.Rank(cpk));

    [Fact]
    public void Rank_非数值原样返回() => Assert.Equal("样本不足", Capability.Rank("样本不足"));
}
