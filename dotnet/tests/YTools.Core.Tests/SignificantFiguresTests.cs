using Xunit;

namespace YTools.Core.Tests;

public class SignificantFiguresTests
{
    // 见 docs/migration/excel-dna-inventory.md §8
    [Theory]
    [InlineData(2.25, 2.3)]
    [InlineData(2.35, 2.4)]
    [InlineData(2.34, 2.3)]
    [InlineData(3.15, 3.2)]
    [InlineData(8.85, 8.9)]
    [InlineData(12.5, 13)]
    [InlineData(0.0125, 0.013)]
    [InlineData(0.335, 0.34)]
    [InlineData(1234, 1200)]
    [InlineData(1250, 1300)]
    [InlineData(1000, 1000)]
    [InlineData(0.001, 0.001)]
    [InlineData(-2.25, -2.3)]
    [InlineData(99.5, 100)]
    [InlineData(0, 0)]
    public void 四舍五入(double input, double expected)
        => Assert.Equal(expected, SignificantFigures.RoundHalfUp(input), 12);

    [Theory]
    [InlineData(2.25, 2.2)]   // 恰为 5，前一位 2 为偶 → 舍
    [InlineData(2.35, 2.4)]   // 恰为 5，前一位 3 为奇 → 进
    [InlineData(2.45, 2.4)]
    [InlineData(2.36, 2.4)]   // 原 VBA 版错误地得到 2.3
    [InlineData(2.34, 2.3)]
    [InlineData(2.251, 2.3)]  // 5 后面还有非零数字 → 进
    [InlineData(2.2500, 2.2)]
    [InlineData(3.15, 3.2)]
    [InlineData(8.85, 8.8)]
    [InlineData(12.5, 12)]
    [InlineData(13.5, 14)]
    [InlineData(0.0125, 0.012)]
    [InlineData(0.335, 0.34)]
    [InlineData(0.0675, 0.068)]
    [InlineData(1250, 1200)]
    [InlineData(1350, 1400)]
    [InlineData(-2.35, -2.4)]
    [InlineData(0, 0)]
    public void 四舍六入五成双_GBT8170(double input, double expected)
        => Assert.Equal(expected, SignificantFigures.RoundHalfEven(input), 12);
}
