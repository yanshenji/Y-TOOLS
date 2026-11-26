using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace YTools.AddIn
{
    /// <summary>“工具宏”功能区选项卡。后续批次的命令按钮都加在这里。</summary>
    [ComVisible(true)]
    public class YToolsRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='yt_tab' label='工具宏'>
        <group id='yt_grp_about' label='关于'>
          <button id='yt_btn_env' label='环境信息' size='large' imageMso='Info' onAction='OnEnvironment' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnLoad(IRibbonUI ribbon) => Diagnostics.RibbonLoaded = true;

        public void OnEnvironment(IRibbonControl control) => Diagnostics.ShowEnvironment();
    }
}
