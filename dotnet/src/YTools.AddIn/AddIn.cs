using ExcelDna.Integration;

namespace YTools.AddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // 未处理的异常统一返回 #VALUE!（与原 VBA UDF 出错时的表现一致）
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => ExcelError.ExcelErrorValue);
        }

        public void AutoClose() { }
    }
}
