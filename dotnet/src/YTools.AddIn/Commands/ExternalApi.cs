using ExcelDna.Integration;
using YTools.Core;

namespace YTools.AddIn.Commands
{
    /// <summary>
    /// 供其他工作簿的 VBA 宏调用的入口（兼容原 .xlam 中被外部调用的过程）。
    /// 调用方式：Application.Run "cmpStd222", 标准区域, 数值区域
    /// </summary>
    public static class ExternalApi
    {
        /// <summary>
        /// 原 函数.cmpStd222：按标准区域给数值区域着色（规则同「标准校验着色」命令）。
        /// 只能通过 Application.Run 调用；写在单元格公式里会返回 #VALUE!（工作表函数不能修改单元格格式）。
        /// 返回处理的单元格数。与原版不同：不再弹“校验完成！”，避免批量调用时反复弹窗。
        /// </summary>
        [ExcelFunction(Name = "cmpStd222", IsMacroType = true, IsHidden = true,
            Description = "供 VBA 用 Application.Run 调用：按标准区域给数值区域着色")]
        public static object CmpStd222(
            [ExcelArgument(AllowReference = true)] object stdRange,
            [ExcelArgument(AllowReference = true)] object valueRange)
        {
            if (XlCall.Excel(XlCall.xlfCaller) is ExcelReference) return ExcelError.ExcelErrorValue; // 从单元格公式调用
            dynamic std = ToRange(stdRange), val = ToRange(valueRange);
            if (std == null || val == null) return "输入区域不能为空！";

            dynamic app = Xl.App;
            bool screen = app.ScreenUpdating, events = app.EnableEvents;
            try
            {
                app.ScreenUpdating = false;
                return (double)DataCommands.ApplyStandardColors(std, val);
            }
            catch (System.Exception ex)
            {
                return "发生未知错误：" + ex.Message;
            }
            finally
            {
                app.ScreenUpdating = screen;
                app.EnableEvents = events;
            }
        }

        /// <summary>Application.Run 传进来的 Range 在 XLL 里是 ExcelReference，转回 COM 的 Range 对象。</summary>
        private static dynamic ToRange(object arg)
        {
            if (!(arg is ExcelReference r)) return null;
            string full = (string)XlCall.Excel(XlCall.xlSheetNm, r); // 形如 "[Book1.xlsx]Sheet1"
            int close = full.IndexOf(']');
            string book = full.Substring(1, close - 1), sheet = full.Substring(close + 1);
            dynamic ws = Xl.App.Workbooks[book].Worksheets[sheet];
            return ws.Range[ws.Cells[r.RowFirst + 1, r.ColumnFirst + 1], ws.Cells[r.RowLast + 1, r.ColumnLast + 1]];
        }
    }
}
