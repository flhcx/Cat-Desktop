using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DesktopOrganizer.Models;

namespace DesktopOrganizer.Services
{
    /// <summary>
    /// 快捷方式自动分类服务
    /// </summary>
    public static class ShortcutCategorizer
    {
        /// <summary>
        /// 分类规则：关键词 -> 分类名称
        /// 按目标路径和快捷方式名称匹配
        /// </summary>
        private static readonly Dictionary<string, string[]> CategoryRules = new()
        {
            ["🌐 浏览器"] = new[] {
                "chrome", "firefox", "edge", "opera", "brave", "vivaldi",
                "browser", "safari", "360se", "360chrome", "qqbrowser", "sogou"
            },
            ["💻 开发工具"] = new[] {
                "visual studio", "vscode", "code", "idea", "pycharm", "webstorm",
                "android studio", "eclipse", "sublime", "notepad++", "git",
                "postman", "docker", "terminal", "powershell", "cmd",
                "datagrip", "rider", "clion", "goland", "phpstorm",
                "hbuilder", "cursor", "warp", "iterm", "devtools"
            },
            ["🎮 游戏"] = new[] {
                "steam", "epic", "wegame", "origin", "ubisoft", "blizzard",
                "battle.net", "riot", "genshin", "league", "minecraft",
                "game", "games", "pubg", "valorant", "overwatch"
            },
            ["💬 社交通讯"] = new[] {
                "wechat", "qq", "telegram", "discord", "skype", "slack",
                "teams", "zoom", "dingtalk", "feishu", "微信", "钉钉",
                "飞书", "whatsapp", "line", "signal"
            },
            ["🎵 影音娱乐"] = new[] {
                "spotify", "music", "vlc", "potplayer", "foobar", "itunes",
                "bilibili", "网易云", "酷狗", "酷我", "qq音乐", "media",
                "player", "kodi", "mpv", "obs", "premiere", "davinci"
            },
            ["📝 办公效率"] = new[] {
                "word", "excel", "powerpoint", "office", "wps", "onenote",
                "outlook", "notion", "evernote", "typora", "obsidian",
                "adobe", "photoshop", "illustrator", "acrobat", "pdf",
                "xmind", "mindmaster", "todo", "trello"
            },
            ["📁 系统工具"] = new[] {
                "explorer", "control", "regedit", "taskmgr", "settings",
                "设置", "计算器", "calculator", "snipping", "paint",
                "7z", "winrar", "everything", "ccleaner", "dism",
                "disk", "defrag", "system", "driver"
            },
            ["🔒 安全防护"] = new[] {
                "antivirus", "defender", "kaspersky", "norton", "avast",
                "360safe", "火绒", "huorong", "malware", "security"
            },
            ["☁️ 网盘存储"] = new[] {
                "onedrive", "dropbox", "百度网盘", "baidunetdisk",
                "阿里云盘", "天翼云", "坚果云", "mega", "google drive"
            },
            ["📥 下载工具"] = new[] {
                "thunder", "迅雷", "idm", "fdm", "aria2", "motrix",
                "utorrent", "bittorrent", "qbittorrent", "download"
            }
        };

        /// <summary>
        /// 根据名称和路径自动判断分类
        /// </summary>
        public static string Categorize(ShortcutItem item)
        {
            // 合并名称和路径进行匹配
            string searchText = $"{item.Name} {item.TargetPath} {item.IconSourcePath}".ToLowerInvariant();

            foreach (var (category, keywords) in CategoryRules)
            {
                foreach (var keyword in keywords)
                {
                    if (searchText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        return category;
                    }
                }
            }

            return "📦 其他";
        }

        /// <summary>
        /// 对快捷方式列表进行自动归类
        /// </summary>
        public static void CategorizeAll(List<ShortcutItem> shortcuts)
        {
            foreach (var item in shortcuts)
            {
                item.Category = Categorize(item);
            }
        }

        /// <summary>
        /// 扫描桌面目录，获取所有快捷方式和 exe 文件
        /// </summary>
        public static List<string> ScanDesktopFiles()
        {
            var files = new List<string>();
            var extensions = new[] { ".lnk", ".exe", ".url" };

            // 用户桌面
            string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            // 公共桌面
            string publicDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

            foreach (var dir in new[] { userDesktop, publicDesktop })
            {
                if (!Directory.Exists(dir)) continue;
                try
                {
                    foreach (var file in Directory.GetFiles(dir))
                    {
                        string ext = Path.GetExtension(file).ToLowerInvariant();
                        if (extensions.Contains(ext))
                        {
                            files.Add(file);
                        }
                    }
                }
                catch { /* 权限不足时跳过 */ }
            }

            return files;
        }

        /// <summary>
        /// 获取所有分类名称列表（按预定义顺序）
        /// </summary>
        public static List<string> GetCategoryOrder()
        {
            var order = CategoryRules.Keys.ToList();
            order.Add("📦 其他");
            order.Add("未分类");
            return order;
        }
    }
}
