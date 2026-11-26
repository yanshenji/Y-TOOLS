namespace YTools.Core.Tests;

internal static class Col
{
    /// <summary>把若干单元格值组成一列（模拟 Excel 传入的 object[,]）。</summary>
    public static object[,] Of(params object[] cells)
    {
        var a = new object[cells.Length, 1];
        for (int i = 0; i < cells.Length; i++) a[i, 0] = cells[i];
        return a;
    }
}
