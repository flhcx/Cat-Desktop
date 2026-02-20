using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using DesktopOrganizer.Models;
using DesktopOrganizer.Services;
using Microsoft.Win32;

namespace DesktopOrganizer
{
    /// <summary>
    /// 主窗口交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        // 当前收纳的快捷方式列表
        private List<ShortcutItem> _shortcuts = new();
        // 面板是否展开
        private bool _isPanelExpanded = false;
        // 应用配置
        private AppConfig _config = new();

        // 悬浮球拖拽状态：区分「点击」和「拖拽移动」
        private bool _isFloatBallDragging = false;
        private Point _floatBallMouseDownPos;

        // 开机自启注册表键名
        private const string AutoStartKeyName = "DesktopOrganizer";
        private const string AutoStartRegistryPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

        public MainWindow()
        {
            InitializeComponent();
            // 悬浮球右键菜单（设置入口）
            SetupFloatButtonContextMenu();
        }

        #region 窗口生命周期

        /// <summary>
        /// 窗口加载时：读取配置并渲染图标
        /// </summary>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 加载配置
            _config = DataService.Load();
            _shortcuts = _config.Shortcuts;

            // 恢复窗口位置
            if (_config.WindowLeft >= 0 && _config.WindowTop >= 0)
            {
                this.Left = _config.WindowLeft;
                this.Top = _config.WindowTop;
            }

            // 渲染已有的快捷方式
            RenderAllShortcuts();

            // 恢复面板状态
            if (_config.IsPanelExpanded)
            {
                ExpandPanel();
            }
        }

        /// <summary>
        /// 窗口关闭时：保存配置
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveConfig();
        }

        #endregion

        #region 悬浮球 & 面板切换

        /// <summary>
        /// 悬浮球鼠标按下：记录位置，准备区分点击/拖拽
        /// </summary>
        private void FloatButton_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isFloatBallDragging = false;
            _floatBallMouseDownPos = e.GetPosition(this);
            // 不在这里调用 DragMove，因为需要先判断是否为拖拽
        }

        /// <summary>
        /// 悬浮球鼠标移动：超过阈值则开始拖拽窗口
        /// </summary>
        private void FloatButton_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isFloatBallDragging)
            {
                var currentPos = e.GetPosition(this);
                var delta = currentPos - _floatBallMouseDownPos;
                // 超过 5 像素视为拖拽
                if (Math.Abs(delta.X) > 5 || Math.Abs(delta.Y) > 5)
                {
                    _isFloatBallDragging = true;
                    this.DragMove();
                }
            }
        }

        /// <summary>
        /// 悬浮球鼠标松开：如果没有拖拽则切换面板
        /// </summary>
        private void FloatButton_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isFloatBallDragging)
            {
                TogglePanel();
            }
            _isFloatBallDragging = false;
        }

        /// <summary>
        /// 切换面板显隐
        /// </summary>
        private void TogglePanel()
        {
            if (_isPanelExpanded)
                CollapsePanel();
            else
                ExpandPanel();
        }

        /// <summary>
        /// 展开面板，带缩放动画
        /// </summary>
        private void ExpandPanel()
        {
            _isPanelExpanded = true;
            PanelBorder.Visibility = Visibility.Visible;
            TitleText.Visibility = Visibility.Visible;
            ControlButtons.Visibility = Visibility.Visible;

            // 缩放展开动画
            var scaleXAnim = new DoubleAnimation(0.8, 1, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            var scaleYAnim = new DoubleAnimation(0.8, 1, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, scaleXAnim);
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, scaleYAnim);

            // 透明度动画
            var opacityAnim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
            PanelBorder.BeginAnimation(OpacityProperty, opacityAnim);
        }

        /// <summary>
        /// 折叠面板，带缩放动画
        /// </summary>
        private void CollapsePanel()
        {
            _isPanelExpanded = false;

            // 缩放收缩动画
            var scaleXAnim = new DoubleAnimation(1, 0.8, TimeSpan.FromMilliseconds(200));
            var scaleYAnim = new DoubleAnimation(1, 0.8, TimeSpan.FromMilliseconds(200));
            var opacityAnim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(200));

            opacityAnim.Completed += (s, e) =>
            {
                PanelBorder.Visibility = Visibility.Collapsed;
                TitleText.Visibility = Visibility.Collapsed;
                ControlButtons.Visibility = Visibility.Collapsed;
            };

            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, scaleXAnim);
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, scaleYAnim);
            PanelBorder.BeginAnimation(OpacityProperty, opacityAnim);
        }

        /// <summary>
        /// 最小化按钮：折叠面板
        /// </summary>
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            CollapsePanel();
        }

        /// <summary>
        /// 关闭按钮：退出应用
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        #endregion

        #region 窗口拖拽移动

        /// <summary>
        /// Grid 区域鼠标按下拖拽：整个窗口区域均可拖拽
        /// 排除快捷方式卡片（Border）、按钮、DropZone 等交互控件
        /// </summary>
        private void DragArea_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // 沿可视树向上检查，排除需要点击的控件
            var source = e.OriginalSource as DependencyObject;
            while (source != null && source != this)
            {
                // 排除按钮、DropZone、快捷方式卡片等
                if (source is Button)
                    return;

                // 排除 DropZone 区域（它有自己的点击处理器：打开文件对话框）
                if (source is Border border && border.Name == "DropZone")
                    return;

                // 排除快捷方式卡片（应用了 ShortcutCardStyle 的 Border，通过 ContextMenu 判断）
                if (source is Border cardBorder && cardBorder.ContextMenu != null)
                    return;

                source = VisualTreeHelper.GetParent(source);
            }

            // 不是交互控件区域，执行窗口拖拽
            this.DragMove();
        }

        #endregion

        #region 拖拽文件进入面板

        /// <summary>
        /// 文件拖入面板区域
        /// </summary>
        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                // 视觉反馈：高亮拖拽区域
                DropZoneBorderBrush.Color = (Color)ColorConverter.ConvertFromString("#90667EEA");
                DropZoneBgBrush.Color = (Color)ColorConverter.ConvertFromString("#25667EEA");
                DropZoneText.Text = "松开以添加";
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        /// <summary>
        /// 文件拖出区域
        /// </summary>
        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            // 恢复默认样式
            DropZoneBorderBrush.Color = (Color)ColorConverter.ConvertFromString("#40667EEA");
            DropZoneBgBrush.Color = (Color)ColorConverter.ConvertFromString("#10667EEA");
            DropZoneText.Text = "拖入 .lnk 或 .exe 文件";
        }

        /// <summary>
        /// 拖拽悬停
        /// </summary>
        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        /// <summary>
        /// 文件放下：解析并添加快捷方式
        /// </summary>
        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            // 恢复拖拽区域样式
            DropZoneBorderBrush.Color = (Color)ColorConverter.ConvertFromString("#40667EEA");
            DropZoneBgBrush.Color = (Color)ColorConverter.ConvertFromString("#10667EEA");
            DropZoneText.Text = "拖入 .lnk 或 .exe 文件";

            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
                return;

            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

            foreach (string file in files)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".lnk" || ext == ".exe")
                {
                    AddShortcutFromFile(file);
                }
            }

            e.Handled = true;
        }

        /// <summary>
        /// 点击拖拽区域：打开文件选择对话框
        /// </summary>
        private void DropZone_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true; // 阻止冒泡到 Grid 拖拽

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择快捷方式或程序",
                Filter = "快捷方式和程序|*.lnk;*.exe|所有文件|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    AddShortcutFromFile(file);
                }
            }
        }

        #endregion

        #region 快捷方式管理

        /// <summary>
        /// 从文件路径添加快捷方式
        /// </summary>
        private void AddShortcutFromFile(string filePath)
        {
            // 检查是否已存在
            if (_shortcuts.Any(s => s.IconSourcePath.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                                 || s.TargetPath.Equals(filePath, StringComparison.OrdinalIgnoreCase)))
            {
                return; // 避免重复添加
            }

            // 解析快捷方式
            var item = ShortcutResolver.ResolveShortcut(filePath);
            _shortcuts.Add(item);

            // 渲染到面板
            AddShortcutCard(item);

            // 保存配置
            SaveConfig();
        }

        /// <summary>
        /// 移除快捷方式
        /// </summary>
        private void RemoveShortcut(ShortcutItem item, Border card)
        {
            _shortcuts.Remove(item);

            // 淡出动画后移除
            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(200));
            fadeOut.Completed += (s, e) =>
            {
                ShortcutsPanel.Children.Remove(card);
            };
            card.BeginAnimation(OpacityProperty, fadeOut);

            SaveConfig();
        }

        #endregion

        #region UI 渲染

        /// <summary>
        /// 渲染所有快捷方式到面板
        /// </summary>
        private void RenderAllShortcuts()
        {
            ShortcutsPanel.Children.Clear();
            foreach (var item in _shortcuts)
            {
                AddShortcutCard(item);
            }
        }

        /// <summary>
        /// 添加一个快捷方式卡片到面板
        /// </summary>
        private void AddShortcutCard(ShortcutItem item)
        {
            var card = new Border
            {
                Style = (Style)FindResource("ShortcutCardStyle"),
                ToolTip = $"{item.Name}\n{item.TargetPath}"
            };

            var stack = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            // 图标
            var iconImage = new Image
            {
                Width = 40,
                Height = 40,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 4)
            };

            // 尝试提取图标
            string iconExtractionPath = !string.IsNullOrEmpty(item.TargetPath) && File.Exists(item.TargetPath)
                ? item.TargetPath
                : item.IconSourcePath;

            var iconSource = ShortcutResolver.ExtractIcon(iconExtractionPath);
            if (iconSource != null)
            {
                iconImage.Source = iconSource;
            }
            else
            {
                // 如果提取失败，显示默认 emoji 图标
                stack.Children.Add(new TextBlock
                {
                    Text = "📄",
                    FontSize = 32,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 4, 0, 4)
                });
                iconImage = null!; // 不使用 Image 控件
            }

            if (iconImage != null)
            {
                stack.Children.Add(iconImage);
            }

            // 名称标签
            var nameLabel = new TextBlock
            {
                Text = TruncateName(item.Name, 8),
                FontSize = 11,
                Foreground = (SolidColorBrush)FindResource("TextPrimaryBrush"),
                HorizontalAlignment = HorizontalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxWidth = 72
            };
            stack.Children.Add(nameLabel);

            card.Child = stack;

            // 左键单击：启动应用
            card.MouseLeftButtonUp += (s, e) =>
            {
                LaunchApp(item);
            };

            // 右键菜单
            var contextMenu = new ContextMenu();

            // 查看属性（弹窗显示）
            var viewPropsItem = new MenuItem { Header = "🔍 查看属性" };
            viewPropsItem.Click += (s, e) =>
            {
                ShowPropertiesDialog(item);
            };
            contextMenu.Items.Add(viewPropsItem);

            // 打开文件所在位置
            var openLocationItem = new MenuItem { Header = "📂 打开文件位置" };
            openLocationItem.Click += (s, e) =>
            {
                OpenFileLocation(item);
            };
            contextMenu.Items.Add(openLocationItem);

            contextMenu.Items.Add(new Separator());

            // 从盒子中移除
            var removeItem = new MenuItem { Header = "❌ 从盒子中移除" };
            removeItem.Click += (s, e) =>
            {
                RemoveShortcut(item, card);
            };
            contextMenu.Items.Add(removeItem);

            card.ContextMenu = contextMenu;

            // 添加到面板
            ShortcutsPanel.Children.Add(card);

            // 添加时的淡入动画
            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(300));
            card.BeginAnimation(OpacityProperty, fadeIn);
        }

        /// <summary>
        /// 截断名称，显示优化
        /// </summary>
        private static string TruncateName(string name, int maxLength)
        {
            if (string.IsNullOrEmpty(name)) return "未知";
            return name.Length > maxLength ? name[..maxLength] + "…" : name;
        }

        #endregion

        #region 应用启动

        /// <summary>
        /// 启动应用程序
        /// </summary>
        private void LaunchApp(ShortcutItem item)
        {
            try
            {
                string path = item.TargetPath;

                if (string.IsNullOrEmpty(path))
                {
                    MessageBox.Show($"快捷方式 \"{item.Name}\" 的目标路径为空。",
                                    "启动失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(path) && !Directory.Exists(path))
                {
                    MessageBox.Show($"目标文件不存在或路径已失效:\n{path}",
                                    "启动失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = path,
                    Arguments = item.Arguments ?? string.Empty,
                    UseShellExecute = true, // 使用系统关联启动
                    WorkingDirectory = !string.IsNullOrEmpty(item.WorkingDirectory)
                        ? item.WorkingDirectory
                        : Path.GetDirectoryName(path) ?? string.Empty
                };

                Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 740)
            {
                // 需要管理员权限
                try
                {
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = item.TargetPath,
                        Verb = "runas",
                        UseShellExecute = true
                    };
                    Process.Start(startInfo);
                }
                catch (Exception innerEx)
                {
                    MessageBox.Show($"以管理员身份启动失败:\n{innerEx.Message}",
                                    "启动失败", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动 \"{item.Name}\" 时发生错误:\n{ex.Message}",
                                "启动失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 打开文件所在位置
        /// </summary>
        private void OpenFileLocation(ShortcutItem item)
        {
            try
            {
                string path = item.TargetPath;
                if (File.Exists(path))
                {
                    Process.Start("explorer.exe", $"/select,\"{path}\"");
                }
                else if (!string.IsNullOrEmpty(path))
                {
                    string? dir = Path.GetDirectoryName(path);
                    if (dir != null && Directory.Exists(dir))
                        Process.Start("explorer.exe", dir);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开文件位置:\n{ex.Message}",
                                "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 属性查看（弹窗）

        /// <summary>
        /// 通过 MessageBox 弹窗显示快捷方式属性
        /// </summary>
        private void ShowPropertiesDialog(ShortcutItem item)
        {
            string info = $"名称：{item.Name}\n\n"
                        + $"目标路径：{(string.IsNullOrEmpty(item.TargetPath) ? "(无)" : item.TargetPath)}\n\n"
                        + $"原始路径：{(string.IsNullOrEmpty(item.IconSourcePath) ? "(无)" : item.IconSourcePath)}\n\n"
                        + $"启动参数：{(string.IsNullOrEmpty(item.Arguments) ? "(无)" : item.Arguments)}\n\n"
                        + $"工作目录：{(string.IsNullOrEmpty(item.WorkingDirectory) ? "(无)" : item.WorkingDirectory)}";

            MessageBox.Show(info, $"属性 - {item.Name}", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion

        #region 开机自启

        /// <summary>
        /// 设置悬浮球右键菜单（包含开机自启开关和退出）
        /// </summary>
        private void SetupFloatButtonContextMenu()
        {
            var contextMenu = new ContextMenu();

            // 开机自启选项
            var autoStartItem = new MenuItem
            {
                Header = "🚀 开机自启",
                IsCheckable = true,
                IsChecked = IsAutoStartEnabled()
            };
            autoStartItem.Click += (s, e) =>
            {
                if (autoStartItem.IsChecked)
                {
                    EnableAutoStart();
                }
                else
                {
                    DisableAutoStart();
                }
            };
            contextMenu.Items.Add(autoStartItem);

            contextMenu.Items.Add(new Separator());

            // 退出应用
            var exitItem = new MenuItem { Header = "✕ 退出" };
            exitItem.Click += (s, e) => { this.Close(); };
            contextMenu.Items.Add(exitItem);

            FloatButton.ContextMenu = contextMenu;
        }

        /// <summary>
        /// 检查是否已设置开机自启
        /// </summary>
        private bool IsAutoStartEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryPath, false);
                if (key == null) return false;
                var val = key.GetValue(AutoStartKeyName);
                return val != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 启用开机自启（写入注册表 HKCU\...\Run）
        /// </summary>
        private void EnableAutoStart()
        {
            try
            {
                // 获取当前可执行文件路径
                string exePath = GetExePath();

                using var key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryPath, true);
                key?.SetValue(AutoStartKeyName, $"\"{exePath}\"");

                MessageBox.Show("已设置开机自启。\n下次开机后将自动运行桌面收纳盒。",
                                "开机自启", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机自启失败:\n{ex.Message}",
                                "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 禁用开机自启（删除注册表键值）
        /// </summary>
        private void DisableAutoStart()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(AutoStartRegistryPath, true);
                key?.DeleteValue(AutoStartKeyName, false);

                MessageBox.Show("已取消开机自启。",
                                "开机自启", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"取消开机自启失败:\n{ex.Message}",
                                "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 获取当前应用的可执行文件路径
        /// </summary>
        private static string GetExePath()
        {
            // 优先使用发布后的 exe 路径
            var processPath = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(processPath))
                return processPath;

            // 回退到入口程序集
            return Process.GetCurrentProcess().MainModule?.FileName
                   ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
        }

        #endregion

        #region 配置保存

        /// <summary>
        /// 保存当前配置到磁盘
        /// </summary>
        private void SaveConfig()
        {
            _config.Shortcuts = _shortcuts;
            _config.WindowLeft = this.Left;
            _config.WindowTop = this.Top;
            _config.IsPanelExpanded = _isPanelExpanded;
            DataService.Save(_config);
        }

        #endregion
    }
}
