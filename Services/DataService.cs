using System;
using System.IO;
using System.Text.Json;
using DesktopOrganizer.Models;

namespace DesktopOrganizer.Services
{
    /// <summary>
    /// 数据持久化服务：使用 JSON 文件保存和加载配置
    /// </summary>
    public static class DataService
    {
        // 配置文件路径：与 exe 同目录下的 config.json
        private static readonly string ConfigPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "config.json");

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        /// <summary>
        /// 保存配置到 JSON 文件
        /// </summary>
        public static void Save(AppConfig config)
        {
            try
            {
                string json = JsonSerializer.Serialize(config, JsonOptions);
                File.WriteAllText(ConfigPath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存配置失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从 JSON 文件加载配置
        /// </summary>
        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                    return new AppConfig();

                string json = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载配置失败: {ex.Message}");
                return new AppConfig();
            }
        }
    }
}
