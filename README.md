<div align="center">

# 🐱 Cat Desktop ─ 桌面收纳盒

<img src="tubiao.jpg" width="120" style="border-radius: 50%;" alt="Cat Desktop Icon"/>

**一个轻量、美观的 Windows 桌面快捷方式收纳工具**

[![.NET 7](https://img.shields.io/badge/.NET-7.0-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)
[![WPF](https://img.shields.io/badge/WPF-Desktop-blue?logo=windows)](https://learn.microsoft.com/wpf/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Release](https://img.shields.io/github/v/release/flhcx/Cat-Desktop?include_prereleases)](https://github.com/flhcx/Cat-Desktop/releases)

</div>

---

## ✨ 功能特点

| 功能 | 说明 |
|------|------|
| 🎯 **悬浮球** | 可爱的猫娘图标悬浮球，点击展开/折叠收纳面板 |
| 📦 **快捷方式管理** | 支持拖拽 `.lnk`、`.exe` 文件到面板中统一管理 |
| 🚀 **一键启动** | 左键单击图标即可启动对应程序，自动处理管理员权限 |
| 🖱️ **自由拖拽** | 窗口全区域可拖拽移动，悬浮球折叠状态也能拖动 |
| 🔍 **属性查看** | 右键快捷方式可查看详细属性（目标路径、参数等） |
| 📂 **打开位置** | 右键快捷方式可直接在资源管理器中定位目标文件 |
| 🔄 **开机自启** | 右键悬浮球可切换开机自启动（注册表方式） |
| 💾 **持久保存** | 自动保存快捷方式列表、窗口位置和面板状态 |
| 🎨 **精美 UI** | 无边框透明窗口、毛玻璃效果、丝滑动画 |

## 📸 预览

<div align="center">

> 悬浮球点击展开面板，拖拽文件即可添加快捷方式

</div>

## 📥 快速开始

### 方法一：直接下载（推荐）

1. 前往 [Releases](https://github.com/flhcx/Cat-Desktop/releases) 页面
2. 下载最新版本的 `DesktopOrganizer.exe`
3. 双击运行即可，无需安装任何依赖

### 方法二：源码构建

```bash
# 克隆仓库
git clone https://github.com/flhcx/Cat-Desktop.git
cd Cat-Desktop

# 构建运行
dotnet run

# 发布独立 exe
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o ./publish
```

**构建要求**：[.NET 7 SDK](https://dotnet.microsoft.com/download/dotnet/7.0)

## 🎮 使用说明

| 操作 | 效果 |
|------|------|
| **左键点击**悬浮球 | 展开 / 折叠收纳面板 |
| **按住拖拽**悬浮球 | 移动整个窗口 |
| **右键**悬浮球 | 显示设置菜单（开机自启、退出） |
| **拖拽文件**到面板底部区域 | 添加快捷方式 |
| **左键点击**快捷方式图标 | 启动对应程序 |
| **右键**快捷方式图标 | 查看属性 / 打开位置 / 移除 |

## 🏗️ 项目结构

```
Cat-Desktop/
├── App.xaml / App.xaml.cs        # 应用入口与全局资源
├── MainWindow.xaml               # 主窗口 UI 布局
├── MainWindow.xaml.cs            # 交互逻辑（拖拽、动画、启动等）
├── Models/
│   └── ShortcutItem.cs           # 数据模型（快捷方式 & 配置）
├── Services/
│   ├── ShortcutResolver.cs       # 快捷方式解析 & 图标提取
│   └── DataService.cs            # JSON 配置持久化
├── Assets/
│   └── app.ico                   # 应用图标
├── tubiao.jpg                    # 悬浮球图标
└── DesktopOrganizer.csproj       # 项目配置
```

## ⚙️ 技术栈

- **框架**：WPF (.NET 7)
- **快捷方式解析**：IShellLinkW COM 互操作
- **图标提取**：System.Drawing.Common
- **数据存储**：System.Text.Json (config.json)
- **开机自启**：注册表 `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`

## 📄 开源协议

本项目基于 [MIT License](LICENSE) 开源。

---

<div align="center">

**Made with ❤️ by [flhcx](https://github.com/flhcx)**

</div>
