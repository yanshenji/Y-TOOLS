using Xunit;

namespace YTools.Core.Tests;

public class StandardCheckTests
{
    [Theory]
    [InlineData("≥10", 10, true)]
    [InlineData(">=10", 9.9, false)]
    [InlineData("≤0.5", 0.5, true)]
    [InlineData("＜5", 5, false)]
    [InlineData("＞5", 5.1, true)]
    [InlineData("5-10", 7, true)]
    [InlineData("5 - 10", 10.1, false)]
    [InlineData("5~10", 5, true)]
    [InlineData("-5", -5, true)]
    [InlineData("3", 3, true)]
    [InlineData("abc", 3, false)]
    public void cmpStd(string expr, double value, bool expected)
        => Assert.Equal(expected, StandardCheck.Compare(expr, value));

    [Fact]
    public void cmpStd_比较式无法解析时报错()
        => Assert.Throws<VbaRuntimeException>(() => StandardCheck.Compare("≥abc", 1));

    private static object Check(object[] std, object[] val, int type = 2)
    {
        var s = new object[1, std.Length];
        var v = new object[1, val.Length];
        for (int i = 0; i < std.Length; i++) s[0, i] = std[i];
        for (int i = 0; i < val.Length; i++) v[0, i] = val[i];
        return StandardCheck.CompareRanges(s, v, type);
    }

    [Fact]
    public void 全部符合()
        => Assert.Equal("合格", Check(new object[] { "≤0.5%", "≥99.8", "5~10g/L" }, new object[] { 0.3, 99.9, 7.0 }));

    [Fact]
    public void 一项超标()
        => Assert.Equal("不合格", Check(new object[] { "≤0.5%", "≥99.8" }, new object[] { 0.6, 99.9 }));

    [Fact]
    public void 实测监测参考与斜杠不判定()
        => Assert.Equal("合格", Check(new object[] { "实测", "监测", "参考≤1", "/", "-" }, new object[] { 9.0, 9.0, 9.0, 9.0, 9.0 }));

    [Fact]
    public void 参考2只解析不判定()
        => Assert.Equal("合格", Check(new object[] { "参考2≤1" }, new object[] { 9.0 }));

    [Fact]
    public void 值为0或空不判定()
        => Assert.Equal("合格", Check(new object[] { "≥5", "≥5" }, new object[] { 0.0, null }));

    [Fact]
    public void IsNull规则()
    {
        Assert.Equal("合格", Check(new object[] { "IsNull" }, new object[] { null }));
        Assert.Equal("不合格", Check(new object[] { "IsNull" }, new object[] { 1.0 }));
    }

    [Fact]
    public void 数值为文本判不合格()
        => Assert.Equal("不合格", Check(new object[] { "≥5" }, new object[] { "未测" }));

    [Fact]
    public void 标准写错返回参数错误()
        => Assert.Equal(StandardCheck.ArgumentError, Check(new object[] { "≥ab" }, new object[] { 1.0 }));

    [Fact]
    public void 单字符标准跳过() // 原 VBA：清理后长度 ≤1 不判定
        => Assert.Equal("合格", Check(new object[] { "5" }, new object[] { 9.0 }));

    [Fact]
    public void 返回类型()
    {
        Assert.Equal(true, Check(new object[] { "≥5" }, new object[] { 6.0 }, 1));
        Assert.Equal("符合", Check(new object[] { "≥5" }, new object[] { 6.0 }, 3));
        Assert.Equal(2.0, Check(new object[] { "≥5" }, new object[] { 4.0 }, 4));
    }

    [Fact]
    public void 长度不一致按较短的比()
        => Assert.Equal("合格", Check(new object[] { "≥5" }, new object[] { 6.0, 1.0 }));
}
