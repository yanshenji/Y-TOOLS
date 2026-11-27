using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace YTools.Core
{
    /// <summary>一种产品的证书参数（原 产品.bas 中 GetProductType / GetDefault* / GetTemplatePath / GetDefaultSheetName）。</summary>
    public class ProductSpec
    {
        [JsonProperty("prefix")] public string Prefix { get; set; }
        [JsonProperty("type")] public string Type { get; set; }
        [JsonProperty("metalContent")] public double MetalContent { get; set; }
        [JsonProperty("packageWeight")] public double PackageWeight { get; set; }
        [JsonProperty("template")] public string Template { get; set; }
        [JsonProperty("sheetName")] public string SheetName { get; set; }
    }

    public class CertificateConfig
    {
        [JsonProperty("templateFolder")] public string TemplateFolder { get; set; }
        [JsonProperty("outputFolder")] public string OutputFolder { get; set; }

        /// <summary>检测台账取 dataFiles 里这个分组的文件（默认“成品数据”），按年份选择。</summary>
        [JsonProperty("ledgerGroup")] public string LedgerGroup { get; set; } = "成品数据";

        [JsonProperty("products")] public List<ProductSpec> Products { get; set; } = new List<ProductSpec>();

        /// <summary>按配置顺序匹配批号前缀（区分大小写，与原 VBA 一致）。</summary>
        public ProductSpec FindProduct(string batch)
            => Products.FirstOrDefault(p => !string.IsNullOrEmpty(p.Prefix) && (batch ?? "").StartsWith(p.Prefix, StringComparison.Ordinal));

        public string TemplatePath(ProductSpec p) => JoinPath(TemplateFolder, p.Template);

        /// <summary>Windows 路径拼接（不依赖运行平台的分隔符）。</summary>
        public static string JoinPath(string folder, string name)
        {
            folder = folder ?? "";
            if (folder.Length > 0 && !folder.EndsWith("\\") && !folder.EndsWith("/")) folder += "\\";
            return folder + (name ?? "");
        }
    }

    public static class CertificateMath
    {
        /// <summary>净重 = 金属量 ÷ 金属含量；包数 = 净重 ÷ 每包重量（每包重量为 0 时用产品默认值）。</summary>
        public static (double NetWeight, double PackageCount, double PackageWeight) Compute(ProductSpec p, double metalQuantity, double packageWeight)
        {
            if (p.MetalContent <= 0) throw new ArgumentException($"产品“{p.Type}”的金属含量未设置。");
            if (packageWeight == 0) packageWeight = p.PackageWeight;
            if (packageWeight <= 0) throw new ArgumentException("每包重量必须大于 0。");
            double net = metalQuantity / p.MetalContent;
            return (net, net / packageWeight, packageWeight);
        }

        /// <summary>
        /// 包数取整：与最近整数相差不到 0.001（浮点误差）时视为整数，返回 true；
        /// 否则返回 false，rounded 为向上取整（最后一包不满）。
        /// </summary>
        public static bool TryWholePackages(double packageCount, out int rounded)
        {
            double nearest = Math.Round(packageCount);
            if (Math.Abs(packageCount - nearest) < 1e-3) { rounded = (int)nearest; return true; }
            rounded = (int)Math.Ceiling(packageCount);
            return false;
        }

        /// <summary>“COA - 批号 - 包数 包 - 净重kg”，数字格式与原 VBA 的字符串拼接一致。</summary>
        public static string FileName(string batch, double packageCount, double netWeight)
            => $"COA - {batch} - {Vba.CStr(packageCount)} 包 - {Vba.CStr(netWeight)}kg";

        /// <summary>台账中的表头字段 → 模板占位符 {{字段}}（字段转大写，与原 VBA UCase 一致）。</summary>
        public static string Placeholder(string field) => "{{" + field.ToUpperInvariant() + "}}";

        /// <summary>从“成品数据”分组里选台账：优先当年，其次名称排序最大的一个。</summary>
        public static DataFile PickLedger(IEnumerable<DataFile> group, int year)
        {
            var list = group.ToList();
            return list.FirstOrDefault(f => f.Name == year + "年")
                   ?? list.OrderByDescending(f => f.Name, StringComparer.Ordinal).FirstOrDefault();
        }
    }
}
