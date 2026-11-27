using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace YTools.Core
{
    /// <summary>一个可以从功能区直接打开的数据文件（原 Openfiles 模块的每个“打开 XX”宏）。</summary>
    public class DataFile
    {
        [JsonProperty("group")] public string Group { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("path")] public string Path { get; set; }
        [JsonProperty("readOnly")] public bool ReadOnly { get; set; } = true;
        [JsonProperty("password")] public string Password { get; set; }
    }

    /// <summary>一组需要同时打开、同时保存关闭的文件（原 打开/保存并关闭中控检测结果_分类）。</summary>
    public class FileSet
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("folder")] public string Folder { get; set; }
        [JsonProperty("files")] public List<string> Files { get; set; } = new List<string>();
        [JsonProperty("readOnly")] public bool ReadOnly { get; set; }

        public IEnumerable<string> FullPaths => Files.Select(f => System.IO.Path.Combine(Folder ?? "", f));
    }

    public class AppConfig
    {
        public const string FileName = "工具宏.config.json";

        [JsonProperty("dataFiles")] public List<DataFile> DataFiles { get; set; } = new List<DataFile>();
        [JsonProperty("fileSets")] public List<FileSet> FileSets { get; set; } = new List<FileSet>();

        /// <summary>批量导出 PDF / 批量删除条件格式时，文件夹选择框的默认位置。</summary>
        [JsonProperty("batchFolder")] public string BatchFolder { get; set; }

        /// <summary>按分组保持配置文件中的先后顺序。</summary>
        public IEnumerable<IGrouping<string, DataFile>> Groups => DataFiles.GroupBy(f => f.Group ?? "其他");

        public static AppConfig Parse(string json)
            => JsonConvert.DeserializeObject<AppConfig>(json) ?? new AppConfig();

        public static string DefaultJson()
        {
            using (var s = typeof(AppConfig).Assembly.GetManifestResourceStream("YTools.Core.DefaultConfig.json"))
            using (var r = new StreamReader(s, System.Text.Encoding.UTF8))
                return r.ReadToEnd();
        }

        public static AppConfig Default() => Parse(DefaultJson());

        /// <summary>目录下有配置文件就读它，没有就用内置默认值。</summary>
        public static AppConfig LoadFrom(string directory, out string sourcePath)
        {
            sourcePath = System.IO.Path.Combine(directory ?? "", FileName);
            if (File.Exists(sourcePath)) return Parse(File.ReadAllText(sourcePath, System.Text.Encoding.UTF8));
            sourcePath = null;
            return Default();
        }
    }
}
