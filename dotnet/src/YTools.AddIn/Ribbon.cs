using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace YTools.AddIn
{
    /// <summary>
    /// “工具宏”功能区选项卡。按钮的 tag 写命令名（ExcelCommand.Name），统一由 OnCommand 分发。
    /// 每个命令同时注册在「加载项」菜单下（ExcelCommand.MenuName），功能区不显示时（如部分 WPS 版本）可从那里使用。
    /// </summary>
    [ComVisible(true)]
    public class YToolsRibbon : ExcelRibbon
    {
        private static IRibbonUI _ui;

        private static readonly Lazy<Dictionary<string, MethodInfo>> Registry = new Lazy<Dictionary<string, MethodInfo>>(() =>
            typeof(YToolsRibbon).Assembly.GetTypes()
                .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Select(m => (m, a: m.GetCustomAttribute<ExcelCommandAttribute>()))
                .Where(x => x.a?.Name != null)
                .ToDictionary(x => x.a.Name, x => x.m));

        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='yt_tab' label='工具宏'>
        <group id='yt_g_file' label='数据文件'>
          <dynamicMenu id='yt_dm_files' label='打开数据' size='large' imageMso='FileOpen' getContent='GetDataFilesMenu' invalidateContentOnDrop='true' />
          <button id='yt_b_set_open' label='打开分类结果' imageMso='FileOpen' tag='工具宏_打开中控检测结果分类' onAction='OnCommand' screentip='打开 (分类)中控检测结果 下的全部报表' />
          <button id='yt_b_set_close' label='保存并关闭分类' imageMso='FileSave' tag='工具宏_保存关闭中控检测结果分类' onAction='OnCommand' />
          <menu id='yt_m_cfg' label='配置' imageMso='ControlProperties'>
            <button id='yt_b_cfg_edit' label='编辑配置文件' tag='工具宏_编辑配置' onAction='OnCommand' />
            <button id='yt_b_cfg_reload' label='重新加载配置' tag='工具宏_重新加载配置' onAction='OnCommand' />
          </menu>
        </group>
        <group id='yt_g_filter' label='筛选'>
          <button id='yt_b_f_value' label='筛选当前值' size='large' imageMso='FilterBySelection' tag='工具宏_筛选当前值' onAction='OnCommand' />
          <menu id='yt_m_filter' label='更多筛选' imageMso='Filter'>
            <button id='yt_b_f_nonblank' label='当前列非空值' tag='工具宏_筛选非空' onAction='OnCommand' />
            <button id='yt_b_f_num' label='当前列数值（>0）' tag='工具宏_筛选数值' onAction='OnCommand' />
            <menuSeparator id='yt_s_f1' />
            <button id='yt_b_f_month' label='本月（上月29日~本月28日）' tag='工具宏_筛选本月' onAction='OnCommand' />
            <button id='yt_b_f_lmonth' label='上月' tag='工具宏_筛选上月' onAction='OnCommand' />
            <button id='yt_b_f_30' label='最大日期近30天' tag='工具宏_筛选近30天' onAction='OnCommand' />
            <menuSeparator id='yt_s_f2' />
            <button id='yt_b_f_last50' label='只保留最后50条可见数据' tag='工具宏_保留最后50条' onAction='OnCommand' />
            <button id='yt_b_f_initcat' label='重新读取样品类别' tag='工具宏_初始化样品类别' onAction='OnCommand' />
          </menu>
          <button id='yt_b_f_next' label='下一个样品类别' imageMso='RefreshAll' tag='工具宏_下一个样品类别' onAction='OnCommand' />
          <button id='yt_b_f_clearcol' label='清除当前列' imageMso='FilterClearAllFilters' tag='工具宏_清除当前列筛选' onAction='OnCommand' />
          <button id='yt_b_f_clear' label='清除全部' imageMso='FilterToggleFilter' tag='工具宏_清除筛选' onAction='OnCommand' />
        </group>
        <group id='yt_g_view' label='显示'>
          <button id='yt_b_v_blankcol' label='隐藏空列' size='large' imageMso='HideDetail' tag='工具宏_隐藏空列' onAction='OnCommand' screentip='隐藏表头以下全空的列' />
          <menu id='yt_m_view' label='更多' imageMso='ShowDetail'>
            <button id='yt_b_v_selcol' label='隐藏空列（只看选中的行）' tag='工具宏_隐藏选区空列' onAction='OnCommand' />
            <button id='yt_b_v_selrow' label='隐藏空行（只看选中的列）' tag='工具宏_隐藏选区空行' onAction='OnCommand' />
            <button id='yt_b_v_ce' label='隐藏所有工作表 C~E 列' tag='工具宏_隐藏所有表CE列' onAction='OnCommand' />
            <menuSeparator id='yt_s_v1' />
            <button id='yt_b_v_unhide' label='取消隐藏所有列' tag='工具宏_取消隐藏列' onAction='OnCommand' />
          </menu>
        </group>
        <group id='yt_g_data' label='数据处理'>
          <button id='yt_b_d_std' label='标准校验着色' size='large' imageMso='ConditionalFormattingHighlightCellsMenu' tag='工具宏_标准校验着色' onAction='OnCommand' />
          <button id='yt_b_d_stats' label='统计并复制' imageMso='AutoSum' tag='工具宏_统计可见数值' onAction='OnCommand' screentip='选中列的可见数值：均值与范围，复制到剪贴板' />
          <button id='yt_b_d_join' label='连接并复制' imageMso='Copy' tag='工具宏_连接可见单元格' onAction='OnCommand' screentip='用“、”连接可见单元格，复制到剪贴板' />
          <button id='yt_b_d_chem' label='化学式上下标' imageMso='Subscript' tag='工具宏_化学式上下标' onAction='OnCommand' />
          <button id='yt_b_d_narrow' label='转窄表' imageMso='TableInsert' tag='工具宏_检测数据转窄表' onAction='OnCommand' />
        </group>
        <group id='yt_g_anomaly' label='异常检测'>
          <button id='yt_b_a_all' label='五法对比' size='large' imageMso='ConditionalFormattingTopBottomRulesMenu' tag='工具宏_异常_五法对比' onAction='OnCommand' screentip='复制可见数据到新工作表，用五种方法分别标记' />
          <menu id='yt_m_anomaly' label='单项检测' imageMso='ConditionalFormattingHighlightCellsMenu'>
            <button id='yt_b_a_3s' label='3西格玛' tag='工具宏_异常_3西格玛' onAction='OnCommand' />
            <button id='yt_b_a_grubbs1' label='Grubbs 单侧(最大/最小)' tag='工具宏_异常_Grubbs单侧' onAction='OnCommand' />
            <button id='yt_b_a_cusum' label='CUSUM' tag='工具宏_异常_CUSUM' onAction='OnCommand' />
            <button id='yt_b_a_ewma' label='EWMA' tag='工具宏_异常_EWMA' onAction='OnCommand' />
            <button id='yt_b_a_box' label='箱线图(IQR)' tag='工具宏_异常_箱线图' onAction='OnCommand' />
            <menuSeparator id='yt_s_a1' />
            <button id='yt_b_o_yellow' label='Grubbs 迭代标黄（选区）' tag='工具宏_Grubbs标黄' onAction='OnCommand' />
            <button id='yt_b_o_flag' label='Grubbs 判定 A 列 → B 列' tag='工具宏_Grubbs标注A列' onAction='OnCommand' />
          </menu>
          <button id='yt_b_a_mk' label='MK 趋势检验' imageMso='ChartTrendlineMenu' tag='工具宏_MK趋势检验' onAction='OnCommand' />
        </group>
        <group id='yt_g_product' label='产品'>
          <button id='yt_b_p_coa' label='质量证明书' size='large' imageMso='FileSaveAsPdfOrXps' tag='工具宏_生成质量证明书' onAction='OnCommand' screentip='按批号从检测台账取数，生成 Excel 与 PDF 证书' />
        </group>
        <group id='yt_g_export' label='导出'>
          <button id='yt_b_e_sheet' label='工作表转 PDF' size='large' imageMso='FileSaveAsPdfOrXps' tag='工具宏_工作表导出PDF' onAction='OnCommand' />
          <menu id='yt_m_export' label='更多' imageMso='FileSaveAs'>
            <button id='yt_b_e_wb' label='当前工作簿导出 PDF' tag='工具宏_工作簿导出PDF' onAction='OnCommand' />
            <button id='yt_b_e_batch' label='文件夹批量导出 PDF…' tag='工具宏_批量导出PDF' onAction='OnCommand' />
            <menuSeparator id='yt_s_e1' />
            <button id='yt_b_e_cf' label='文件夹批量删除条件格式…' tag='工具宏_批量删除条件格式' onAction='OnCommand' />
          </menu>
        </group>
        <group id='yt_g_about' label='关于'>
          <button id='yt_btn_env' label='环境信息' imageMso='Info' tag='工具宏_环境信息' onAction='OnCommand' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnLoad(IRibbonUI ribbon)
        {
            _ui = ribbon;
            Diagnostics.RibbonLoaded = true;
        }

        public static void Refresh() => _ui?.Invalidate();

        public void OnCommand(IRibbonControl control)
        {
            if (Registry.Value.TryGetValue(control.Tag, out var m)) m.Invoke(null, null);
            else Ui.Warn("未知命令：" + control.Tag);
        }

        /// <summary>“打开数据”下拉菜单：按配置文件的分组生成子菜单。</summary>
        public string GetDataFilesMenu(IRibbonControl control)
        {
            var cfg = Settings.Current;
            var sb = new StringBuilder("<menu xmlns='http://schemas.microsoft.com/office/2009/07/customui'>");
            int i = 0, g = 0;
            foreach (var grp in cfg.Groups)
            {
                sb.Append($"<menu id='yt_dm_g{g++}' label='{Esc(grp.Key)}'>");
                foreach (var f in grp)
                {
                    int idx = cfg.DataFiles.IndexOf(f);
                    sb.Append($"<button id='yt_dm_f{i++}' label='{Esc(f.Name)}' tag='{idx}' screentip='{Esc(f.Path)}' onAction='OnOpenDataFile' />");
                }
                sb.Append("</menu>");
            }
            sb.Append("<menuSeparator id='yt_dm_sep' />");
            sb.Append("<button id='yt_dm_pick' label='打开数据文件…' tag='工具宏_打开数据文件' onAction='OnCommand' />");
            sb.Append("</menu>");
            return sb.ToString();
        }

        public void OnOpenDataFile(IRibbonControl control)
        {
            var files = Settings.Current.DataFiles;
            if (int.TryParse(control.Tag, out int idx) && idx >= 0 && idx < files.Count)
                Commands.FileCommands.OpenDataFile(files[idx]);
        }

        private static string Esc(string s) => SecurityElement.Escape(s ?? "");
    }
}
