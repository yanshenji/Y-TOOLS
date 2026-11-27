using System;
using System.Linq;
using Xunit;

namespace YTools.Core.Tests;

public class ConfigTests
{
    [Fact]
    public void 默认配置包含原Openfiles全部22个文件()
    {
        var cfg = AppConfig.Default();
        Assert.Equal(22, cfg.DataFiles.Count);
        Assert.All(cfg.DataFiles, f => Assert.StartsWith(@"\\10.34.100.50\质量部\", f.Path));
        var cz = cfg.DataFiles.Single(f => f.Group == "中控数据" && f.Name == "2026年");
        Assert.True(cz.ReadOnly);
        Assert.Equal("2026", cz.Password);
        var mn = cfg.DataFiles.Single(f => f.Name.Contains("氯化锰"));
        Assert.False(mn.ReadOnly);
        Assert.Equal("20252025", mn.Password);
    }

    [Fact]
    public void 默认配置的文件组()
    {
        var set = AppConfig.Default().FileSets.Single();
        Assert.Equal(6, set.Files.Count); // 原 VBA 打开时重复打开了 101-105，这里去重
        Assert.All(set.FullPaths, p => Assert.Contains(@"(分类)中控检测结果\", p));
    }

    [Fact]
    public void 分组保持顺序()
        => Assert.Equal("中控质量报表", AppConfig.Default().Groups.First().Key);
}

public class DateFilterTests
{
    [Theory]
    [InlineData("2026-11-27", 0, "2026-10-28", "2026-11-28")]
    [InlineData("2026-11-27", -1, "2026-09-28", "2026-10-28")]
    [InlineData("2026-01-05", 0, "2025-12-28", "2026-01-28")]
    [InlineData("2026-01-05", -1, "2025-11-28", "2025-12-28")]
    [InlineData("2026-03-31", 0, "2026-02-28", "2026-03-28")]
    public void 业务月(string today, int offset, string start, string end)
    {
        var (s, e) = DateFilters.BusinessMonth(DateTime.Parse(today), offset);
        Assert.Equal(DateTime.Parse(start), s);
        Assert.Equal(DateTime.Parse(end), e);
    }

    [Fact]
    public void 条件用日期序列号()
        => Assert.Equal(">=46353", DateFilters.Criteria(">=", new DateTime(2026, 11, 27)));
}

public class ChemicalFormulaTests
{
    private static string Render(string s)
    {
        var marks = ChemicalFormula.Analyze(s);
        return string.Concat(s.Select((c, i) => marks.TryGetValue(i + 1, out var k)
            ? (k == ScriptKind.Subscript ? $"_{c}" : $"^{c}") : c.ToString()));
    }

    [Theory]
    [InlineData("H2O", "H_2O")]
    [InlineData("CoSO4", "CoSO_4")]
    [InlineData("SO42-", "SO_4^2^-")]
    [InlineData("Fe3+", "Fe^3^+")]   // 原 VBA 的 3 不做处理
    [InlineData("Ca2+", "Ca^2^+")]
    [InlineData("NH4+", "NH^4^+")]
    [InlineData("Co3O4", "Co_3O_4")]
    [InlineData("中文", "中文")]
    public void 上下标(string input, string expected) => Assert.Equal(expected, Render(input));
}

public class SheetOpsTests
{
    [Fact]
    public void 空列()
    {
        var b = new object[,] { { "a", "b", "c" }, { null, 1.0, "" }, { "", null, null } };
        Assert.Equal(new[] { 0, 2 }, SheetOps.BlankColumns(b, 1));
        Assert.Empty(SheetOps.BlankColumns(b, 0)); // 从第 0 行查，表头非空
    }

    [Fact]
    public void 空行()
    {
        var b = new object[,] { { "a", null }, { null, "" }, { null, 2.0 } };
        Assert.Equal(new[] { 1 }, SheetOps.BlankRows(b));
    }

    [Fact]
    public void 保留最后50条()
    {
        var rows = Enumerable.Range(2, 60).ToList();
        Assert.Equal(Enumerable.Range(2, 10), SheetOps.RowsToHideKeepingLast(rows));
        Assert.Empty(SheetOps.RowsToHideKeepingLast(rows.Take(50).ToList()));
    }

    [Fact]
    public void 连接()
        => Assert.Equal("A、1.5、B", SheetOps.JoinNonBlank(new object[] { "A", null, 1.5, "", "B" }));

    [Fact]
    public void 统计行()
    {
        Assert.Equal("Co 均值2.0000g/L(1.0000-3.0000g/L)", SheetOps.StatsLine("Co", new object[] { 1.0, "未测", 3.0, 0.0, null, "2" }));
        Assert.Equal("未命名列 无大于0的可见数值", SheetOps.StatsLine("", new object[] { 0.0, "x" }));
    }

    [Fact]
    public void 工作表名()
    {
        Assert.Equal("a_b_窄表", SheetOps.SafeSheetName("a/b?", "窄表"));
        Assert.Equal(31, SheetOps.SafeSheetName(new string('x', 40)).Length);
    }
}

public class NarrowTableTests
{
    [Fact]
    public void 宽表转窄表()
    {
        var wide = new object[,]
        {
            { "样品名称", "Co", "Ni ", "判定" },
            { "A", 1.0, null, "合格" },
            { "B", 2.0, 0.5, "不合格" },
        };
        var t = NarrowTable.Transform(wide);
        Assert.Equal(4, t.GetLength(0));                 // 表头 + 3 条
        Assert.Equal(11, t.GetLength(1));
        Assert.Equal("检测项目", t[0, 9]);
        Assert.Equal("A", t[1, 5]); Assert.Equal("Co", t[1, 9]); Assert.Equal(1.0, t[1, 10]);
        Assert.Equal("合格", t[1, 7]);
        Assert.Null(t[1, 0]);                             // 缺少“送样日期”列：留空
        Assert.Equal("Ni", t[3, 9]); Assert.Equal(0.5, t[3, 10]);
    }
}

public class OutlierMarkTests
{
    private static readonly object[] Data = { "表头", 2.1, 2.3, 2.2, 2.4, 2.2, 2.3, 2.1, 2.2, 2.3, 5.0 };

    [Fact]
    public void 单次检验标黄()
    {
        var m = OutlierMarks.SinglePass(Data);
        Assert.Null(m[0]);
        Assert.True(m[10]);
        Assert.Equal(1, m.Count(x => x == true));
    }

    [Fact]
    public void 迭代检验标注()
    {
        var m = OutlierMarks.Iterative(Data, 0.05, out int n);
        Assert.Equal(1, n);
        Assert.Null(m[0]);
        Assert.Equal("异常值", m[10]);
        Assert.Equal("正常", m[1]);
    }

    [Theory]
    [InlineData("≥5", 6.0, StdMark.None)]
    [InlineData("≥5", 4.0, StdMark.OutOfSpec)]
    [InlineData("≥5", "未测", StdMark.NotNumeric)]
    [InlineData("≥5", null, StdMark.Untested)]
    [InlineData("NotNull≥5", null, StdMark.NotNullMissing)]
    [InlineData("参考2≤1", null, StdMark.Cankao2Missing)]
    [InlineData("参考2≤1", 9.0, StdMark.OutOfSpec)]
    [InlineData("抽检≥5", null, StdMark.None)]
    [InlineData("IsNull", 1.0, StdMark.IsNullViolated)]
    [InlineData("实测", 1.0, StdMark.None)]
    [InlineData("≥ab", 1.0, StdMark.NotNumeric)]
    public void 标准着色(string std, object value, StdMark expected)
        => Assert.Equal(expected, StandardMarks.Classify(std, value));
}
