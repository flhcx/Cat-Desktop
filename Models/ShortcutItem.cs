using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace DesktopOrganizer.Models
{
    /// <summary>
    /// 快捷方式数据模型，用于存储和序列化快捷方式信息
    /// </summary>
    public class ShortcutItem
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        [JsonPropertyName("name")]
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 目标程序的真实路径
        /// </summary>
        [JsonPropertyName("targetPath")]
        public string TargetPath { get; set; } = string.Empty;

        /// <summary>
        /// 原始文件路径（.lnk 或 .exe 的路径，用于图标提取）
        /// </summary>
        [JsonPropertyName("iconSourcePath")]
        public string IconSourcePath { get; set; } = string.Empty;

        /// <summary>
        /// 启动参数
        /// </summary>
        [JsonPropertyName("arguments")]
        public string Arguments { get; set; } = string.Empty;

        /// <summary>
        /// 工作目录
        /// </summary>
        [JsonPropertyName("workingDirectory")]
        public string WorkingDirectory { get; set; } = string.Empty;
    }

    /// <summary>
    /// 配置根对象，用于 JSON 序列化
    /// </summary>
    public class AppConfig
    {
        [JsonPropertyName("shortcuts")]
        public List<ShortcutItem> Shortcuts { get; set; } = new();

        [JsonPropertyName("windowLeft")]
        public double WindowLeft { get; set; } = 100;

        [JsonPropertyName("windowTop")]
        public double WindowTop { get; set; } = 100;

        [JsonPropertyName("isPanelExpanded")]
        public bool IsPanelExpanded { get; set; } = false;
    }
}
