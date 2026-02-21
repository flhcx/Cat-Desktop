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

        /// <summary>
        /// 分类名称（如 浏览器、开发工具、游戏 等）
        /// </summary>
        [JsonPropertyName("category")]
        public string Category { get; set; } = "未分类";
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

        /// <summary>
        /// 是否记住上次退出的窗口位置（默认开启）
        /// </summary>
        [JsonPropertyName("rememberPosition")]
        public bool RememberPosition { get; set; } = true;

        /// <summary>
        /// 是否开机自启动
        /// </summary>
        [JsonPropertyName("autoStart")]
        public bool AutoStart { get; set; } = false;

        /// <summary>
        /// 自定义悬浮球图片路径（为空则使用默认图片）
        /// </summary>
        [JsonPropertyName("customBallImagePath")]
        public string CustomBallImagePath { get; set; } = string.Empty;

        /// <summary>
        /// 悬浮球大小（像素），范围 36~96，默认 56
        /// </summary>
        [JsonPropertyName("ballSize")]
        public double BallSize { get; set; } = 56;

        /// <summary>
        /// 盒子展开特效名称：Scale / SlideDown / SlideRight / Bounce / Rotate / Fade
        /// </summary>
        [JsonPropertyName("openEffect")]
        public string OpenEffect { get; set; } = "Scale";

        /// <summary>
        /// 是否隐藏桌面图标
        /// </summary>
        [JsonPropertyName("hideDesktopIcons")]
        public bool HideDesktopIcons { get; set; } = false;
    }
}
