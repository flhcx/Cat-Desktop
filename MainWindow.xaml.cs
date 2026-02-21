using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

        // 窗口尺寸：折叠时由悬浮球大小动态计算，展开时固定尺寸
        private double CollapsedWidth => _config.BallSize + 20;
        private double CollapsedHeight => _config.BallSize + 20;
        private const double ExpandedWidth = 420;
        private const double ExpandedHeight = 520;

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

            // 恢复窗口位置（仅在开启记住位置时）
            if (_config.RememberPosition && _config.WindowLeft >= 0 && _config.WindowTop >= 0)
            {
                this.Left = _config.WindowLeft;
                this.Top = _config.WindowTop;
            }

            // 监听窗口位置变化，防止拖出屏幕
            this.LocationChanged += (s, args) => ClampToScreen();

            // 初始化设置面板状态
            InitSettingsPanel();

            // 应用悬浮球大小
            ApplyBallSize(_config.BallSize);

            // 加载自定义悬浮球图片
            if (!string.IsNullOrEmpty(_config.CustomBallImagePath) && System.IO.File.Exists(_config.CustomBallImagePath))
            {
                ApplyBallImage(_config.CustomBallImagePath);
            }

            // 渲染已有的快捷方式
            RenderAllShortcuts();

            // 恢复面板状态
            if (_config.IsPanelExpanded)
            {
                ExpandPanel();
            }
            else
            {
                // 初始状态为折叠，缩小窗口到悬浮球大小
                this.Width = CollapsedWidth;
                this.Height = CollapsedHeight;
            }
        }

        /// <summary>
        /// 窗口关闭时：保存配置
        /// </summary>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveConfig();
        }

        /// <summary>
        /// 限制窗口不超出屏幕工作区域
        /// </summary>
        private void ClampToScreen()
        {
            var screen = SystemParameters.WorkArea;

            // 左边界：窗口左边缘不能超出屏幕左侧
            if (this.Left < screen.Left)
                this.Left = screen.Left;

            // 右边界：窗口右边缘不能超出屏幕右侧
            if (this.Left + this.ActualWidth > screen.Right)
                this.Left = screen.Right - this.ActualWidth;

            // 上边界
            if (this.Top < screen.Top)
                this.Top = screen.Top;

            // 下边界
            if (this.Top + this.ActualHeight > screen.Bottom)
                this.Top = screen.Bottom - this.ActualHeight;
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
            // 防止动画过程中重复触发导致卡死
            if (_isAnimating) return;

            if (_isPanelExpanded)
                CollapsePanel();
            else
                ExpandPanel();
        }

        // 动画锁标志，防止快速连续点击导致竞态
        private bool _isAnimating = false;

        /// <summary>
        /// 展开面板，根据配置的特效播放动画
        /// </summary>
        private void ExpandPanel()
        {
            _isAnimating = true;
            _isPanelExpanded = true;

            // 先恢复窗口尺寸
            this.Width = ExpandedWidth;
            this.Height = ExpandedHeight;

            PanelBorder.Visibility = Visibility.Visible;
            TitleText.Visibility = Visibility.Visible;
            ControlButtons.Visibility = Visibility.Visible;

            // 确保 TranslateTransform 存在
            EnsurePanelTranslateTransform();

            // ★ 关键修复：先清除上一次动画的所有属性锁定
            ResetPanelTransforms();

            string effect = _config.OpenEffect ?? "Scale";
            var duration = TimeSpan.FromMilliseconds(300);

            switch (effect)
            {
                case "SlideDown":
                    PlayExpandSlideDown(duration);
                    break;
                case "SlideRight":
                    PlayExpandSlideRight(duration);
                    break;
                case "Bounce":
                    PlayExpandBounce(duration);
                    break;
                case "Rotate":
                    PlayExpandRotate(duration);
                    break;
                case "Fade":
                    PlayExpandFade(duration);
                    break;
                default: // Scale
                    PlayExpandScale(duration);
                    break;
            }

            // 展开动画完成后解锁
            var timer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = duration + TimeSpan.FromMilliseconds(50)
            };
            timer.Tick += (s, e) => { _isAnimating = false; timer.Stop(); };
            timer.Start();
        }

        /// <summary>
        /// 折叠面板，根据配置的特效播放动画
        /// </summary>
        private void CollapsePanel()
        {
            _isAnimating = true;
            _isPanelExpanded = false;
            EnsurePanelTranslateTransform();

            // ★ 关键修复：先清除之前的动画锁定，再播放折叠动画
            ResetPanelTransforms();

            string effect = _config.OpenEffect ?? "Scale";
            var duration = TimeSpan.FromMilliseconds(200);

            // 所有特效共用的完成回调
            Action onCompleted = () =>
            {
                PanelBorder.Visibility = Visibility.Collapsed;
                TitleText.Visibility = Visibility.Collapsed;
                ControlButtons.Visibility = Visibility.Collapsed;
                this.Width = CollapsedWidth;
                this.Height = CollapsedHeight;
                // 重置变换
                ResetPanelTransforms();
                // 解锁
                _isAnimating = false;
            };

            switch (effect)
            {
                case "SlideDown":
                    PlayCollapseSlideDown(duration, onCompleted);
                    break;
                case "SlideRight":
                    PlayCollapseSlideRight(duration, onCompleted);
                    break;
                case "Bounce":
                    PlayCollapseBounce(duration, onCompleted);
                    break;
                case "Rotate":
                    PlayCollapseRotate(duration, onCompleted);
                    break;
                case "Fade":
                    PlayCollapseFade(duration, onCompleted);
                    break;
                default:
                    PlayCollapseScale(duration, onCompleted);
                    break;
            }
        }

        #region 动画特效实现

        private TranslateTransform? _panelTranslate;
        private RotateTransform? _panelRotate;

        /// <summary>
        /// 确保 PanelBorder 拥有 TransformGroup（包含 Scale、Translate、Rotate）
        /// </summary>
        private void EnsurePanelTranslateTransform()
        {
            if (_panelTranslate != null && _panelRotate != null) return;

            var group = PanelBorder.RenderTransform as TransformGroup;
            if (group == null)
            {
                group = new TransformGroup();
                group.Children.Add(PanelScale);
                _panelTranslate = new TranslateTransform(0, 0);
                group.Children.Add(_panelTranslate);
                _panelRotate = new RotateTransform(0);
                group.Children.Add(_panelRotate);
                PanelBorder.RenderTransform = group;
            }
            else
            {
                // 查找或添加
                _panelTranslate = group.Children.OfType<TranslateTransform>().FirstOrDefault();
                if (_panelTranslate == null)
                {
                    _panelTranslate = new TranslateTransform(0, 0);
                    group.Children.Add(_panelTranslate);
                }
                _panelRotate = group.Children.OfType<RotateTransform>().FirstOrDefault();
                if (_panelRotate == null)
                {
                    _panelRotate = new RotateTransform(0);
                    group.Children.Add(_panelRotate);
                }
            }
        }

        /// <summary>
        /// 重置面板的所有变换为初始状态（必须先清除动画绑定）
        /// </summary>
        private void ResetPanelTransforms()
        {
            // ★ 先清除动画绑定（传入 null 解除 WPF 动画锁定），再设置值
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, null);
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, null);
            PanelScale.ScaleX = 1;
            PanelScale.ScaleY = 1;
            PanelBorder.BeginAnimation(OpacityProperty, null);
            PanelBorder.Opacity = 1;

            if (_panelTranslate != null)
            {
                _panelTranslate.BeginAnimation(TranslateTransform.XProperty, null);
                _panelTranslate.BeginAnimation(TranslateTransform.YProperty, null);
                _panelTranslate.X = 0;
                _panelTranslate.Y = 0;
            }
            if (_panelRotate != null)
            {
                _panelRotate.BeginAnimation(RotateTransform.AngleProperty, null);
                _panelRotate.Angle = 0;
            }
        }

        /// <summary>
        /// 创建带缓动的动画
        /// </summary>
        private DoubleAnimation MakeAnim(double from, double to, TimeSpan dur, IEasingFunction? easing = null)
        {
            var anim = new DoubleAnimation(from, to, dur);
            if (easing != null) anim.EasingFunction = easing;
            return anim;
        }

        // ===== 特效 1：缩放 (Scale) =====
        private void PlayExpandScale(TimeSpan dur)
        {
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(0.8, 1, dur, ease));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(0.8, 1, dur, ease));
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, dur));
        }
        private void PlayCollapseScale(TimeSpan dur, Action onDone)
        {
            var opAnim = MakeAnim(1, 0, dur);
            opAnim.Completed += (s, e) => onDone();
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(1, 0.8, dur));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(1, 0.8, dur));
            PanelBorder.BeginAnimation(OpacityProperty, opAnim);
        }

        // ===== 特效 2：下滑 (SlideDown) =====
        private void PlayExpandSlideDown(TimeSpan dur)
        {
            if (_panelTranslate == null) return;
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
            _panelTranslate.BeginAnimation(TranslateTransform.YProperty, MakeAnim(-60, 0, dur, ease));
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, dur));
        }
        private void PlayCollapseSlideDown(TimeSpan dur, Action onDone)
        {
            if (_panelTranslate == null) { onDone(); return; }
            var anim = MakeAnim(0, -60, dur);
            anim.Completed += (s, e) => onDone();
            _panelTranslate.BeginAnimation(TranslateTransform.YProperty, anim);
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(1, 0, dur));
        }

        // ===== 特效 3：右滑 (SlideRight) =====
        private void PlayExpandSlideRight(TimeSpan dur)
        {
            if (_panelTranslate == null) return;
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
            _panelTranslate.BeginAnimation(TranslateTransform.XProperty, MakeAnim(-80, 0, dur, ease));
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, dur));
        }
        private void PlayCollapseSlideRight(TimeSpan dur, Action onDone)
        {
            if (_panelTranslate == null) { onDone(); return; }
            var anim = MakeAnim(0, -80, dur);
            anim.Completed += (s, e) => onDone();
            _panelTranslate.BeginAnimation(TranslateTransform.XProperty, anim);
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(1, 0, dur));
        }

        // ===== 特效 4：弹性 (Bounce) =====
        private void PlayExpandBounce(TimeSpan dur)
        {
            var ease = new ElasticEase { EasingMode = EasingMode.EaseOut, Oscillations = 1, Springiness = 5 };
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(0.3, 1, TimeSpan.FromMilliseconds(500), ease));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(0.3, 1, TimeSpan.FromMilliseconds(500), ease));
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, dur));
        }
        private void PlayCollapseBounce(TimeSpan dur, Action onDone)
        {
            var ease = new BackEase { EasingMode = EasingMode.EaseIn, Amplitude = 0.4 };
            var opAnim = MakeAnim(1, 0, dur);
            opAnim.Completed += (s, e) => onDone();
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(1, 0.3, dur, ease));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(1, 0.3, dur, ease));
            PanelBorder.BeginAnimation(OpacityProperty, opAnim);
        }

        // ===== 特效 5：旋转 (Rotate) =====
        private void PlayExpandRotate(TimeSpan dur)
        {
            if (_panelRotate == null) return;
            var ease = new CubicEase { EasingMode = EasingMode.EaseOut };
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(0.5, 1, dur, ease));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(0.5, 1, dur, ease));
            _panelRotate.BeginAnimation(RotateTransform.AngleProperty, MakeAnim(-15, 0, dur, ease));
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, dur));
        }
        private void PlayCollapseRotate(TimeSpan dur, Action onDone)
        {
            if (_panelRotate == null) { onDone(); return; }
            var opAnim = MakeAnim(1, 0, dur);
            opAnim.Completed += (s, e) => onDone();
            PanelScale.BeginAnimation(ScaleTransform.ScaleXProperty, MakeAnim(1, 0.5, dur));
            PanelScale.BeginAnimation(ScaleTransform.ScaleYProperty, MakeAnim(1, 0.5, dur));
            _panelRotate.BeginAnimation(RotateTransform.AngleProperty, MakeAnim(0, -15, dur));
            PanelBorder.BeginAnimation(OpacityProperty, opAnim);
        }

        // ===== 特效 6：淡入 (Fade) =====
        private void PlayExpandFade(TimeSpan dur)
        {
            PanelBorder.BeginAnimation(OpacityProperty, MakeAnim(0, 1, TimeSpan.FromMilliseconds(400)));
        }
        private void PlayCollapseFade(TimeSpan dur, Action onDone)
        {
            var anim = MakeAnim(1, 0, TimeSpan.FromMilliseconds(300));
            anim.Completed += (s, e) => onDone();
            PanelBorder.BeginAnimation(OpacityProperty, anim);
        }

        #endregion

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
        /// 渲染所有快捷方式到面板（按分类分组显示）
        /// </summary>
        private void RenderAllShortcuts()
        {
            ShortcutsPanel.Children.Clear();

            // 检查是否有分类数据
            bool hasCategories = _shortcuts.Any(s => s.Category != "未分类" && !string.IsNullOrEmpty(s.Category));

            if (!hasCategories)
            {
                // 无分类时直接显示所有快捷方式
                foreach (var item in _shortcuts)
                {
                    AddShortcutCard(item);
                }
                return;
            }

            // 按分类分组
            var categoryOrder = ShortcutCategorizer.GetCategoryOrder();
            var groups = _shortcuts.GroupBy(s => s.Category ?? "未分类")
                                   .OrderBy(g => {
                                       int idx = categoryOrder.IndexOf(g.Key);
                                       return idx >= 0 ? idx : 999;
                                   });

            foreach (var group in groups)
            {
                // 分类标题
                var header = new TextBlock
                {
                    Text = $"{group.Key} ({group.Count()})",
                    FontSize = 13,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#B0BEC5")),
                    Margin = new Thickness(8, 8, 0, 2)
                };
                // 让 TextBlock 占满整行
                header.SetValue(WrapPanel.WidthProperty, 380.0);
                ShortcutsPanel.Children.Add(header);

                // 分类下的快捷方式卡片
                foreach (var item in group)
                {
                    AddShortcutCard(item);
                }
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

        #region 设置面板

        // 设置面板是否可见
        private bool _isSettingsVisible = false;

        /// <summary>
        /// 初始化设置面板中各开关的状态
        /// </summary>
        private void InitSettingsPanel()
        {
            RememberPositionToggle.IsChecked = _config.RememberPosition;
            AutoStartToggle.IsChecked = _config.AutoStart;

            if (!string.IsNullOrEmpty(_config.CustomBallImagePath))
            {
                BallImagePathText.Text = Path.GetFileName(_config.CustomBallImagePath);
                ResetBallImageBtn.Visibility = Visibility.Visible;
            }

            // 初始化悬浮球大小滑块
            BallSizeSlider.Value = _config.BallSize;
            BallSizeValueText.Text = $"{_config.BallSize} px";

            // 初始化特效选择下拉框
            string effectTag = _config.OpenEffect ?? "Scale";
            foreach (ComboBoxItem item in EffectComboBox.Items)
            {
                if (item.Tag is string tag && tag == effectTag)
                {
                    EffectComboBox.SelectedItem = item;
                    break;
                }
            }

            // 初始化隐藏桌面图标开关
            HideDesktopToggle.IsChecked = _config.HideDesktopIcons;
            if (_config.HideDesktopIcons)
            {
                SetDesktopIconsVisibility(false);
            }
        }

        /// <summary>
        /// 齿轮按钮：切换主页/设置页
        /// </summary>
        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            _isSettingsVisible = !_isSettingsVisible;

            if (_isSettingsVisible)
            {
                // 显示设置，隐藏快捷方式列表和拖拽区域
                ShortcutsScrollViewer.Visibility = Visibility.Collapsed;
                DropZone.Visibility = Visibility.Collapsed;
                SettingsPanel.Visibility = Visibility.Visible;
            }
            else
            {
                // 返回主页
                SettingsPanel.Visibility = Visibility.Collapsed;
                ShortcutsScrollViewer.Visibility = Visibility.Visible;
                DropZone.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// 记住位置开关
        /// </summary>
        private void RememberPositionToggle_Click(object sender, RoutedEventArgs e)
        {
            _config.RememberPosition = RememberPositionToggle.IsChecked == true;
            SaveConfig();
        }

        /// <summary>
        /// 开机自启开关
        /// </summary>
        private void AutoStartToggle_Click(object sender, RoutedEventArgs e)
        {
            bool enabled = AutoStartToggle.IsChecked == true;
            _config.AutoStart = enabled;

            try
            {
                if (enabled)
                    EnableAutoStart();
                else
                    DisableAutoStart();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置开机自启失败:\n{ex.Message}",
                                "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                // 回退 UI 状态
                AutoStartToggle.IsChecked = !enabled;
                _config.AutoStart = !enabled;
            }

            SaveConfig();
        }

        /// <summary>
        /// 更换悬浮球图片按钮
        /// </summary>
        private void ChangeBallImage_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择悬浮球图片",
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|所有文件|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ApplyBallImage(dialog.FileName);
                    _config.CustomBallImagePath = dialog.FileName;
                    BallImagePathText.Text = Path.GetFileName(dialog.FileName);
                    ResetBallImageBtn.Visibility = Visibility.Visible;
                    SaveConfig();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载图片失败:\n{ex.Message}",
                                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// 重置悬浮球图片为默认
        /// </summary>
        private void ResetBallImage_Click(object sender, RoutedEventArgs e)
        {
            _config.CustomBallImagePath = string.Empty;
            BallImagePathText.Text = "默认图片";
            ResetBallImageBtn.Visibility = Visibility.Collapsed;

            // 恢复默认图片 tubiao.jpg（嵌入资源）
            ApplyBallImage(null);
            SaveConfig();
        }

        /// <summary>
        /// 应用悬浮球图片（null = 恢复默认）
        /// </summary>
        private void ApplyBallImage(string? imagePath)
        {
            try
            {
                // 查找悬浮球 Ellipse 的 ImageBrush
                var template = FloatButton.Template;
                var mainCircle = template.FindName("MainCircle", FloatButton) as System.Windows.Shapes.Ellipse;
                if (mainCircle == null) return;

                ImageSource imageSource;
                if (string.IsNullOrEmpty(imagePath))
                {
                    // 恢复默认嵌入资源
                    imageSource = new BitmapImage(new Uri("pack://application:,,,/tubiao.jpg"));
                }
                else
                {
                    // 加载外部文件
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    imageSource = bitmap;
                }

                mainCircle.Fill = new ImageBrush(imageSource) { Stretch = Stretch.UniformToFill };
            }
            catch
            {
                // 图片加载失败时静默处理
            }
        }

        /// <summary>
        /// 应用悬浮球大小（动态修改按钮宽高）
        /// </summary>
        private void ApplyBallSize(double size)
        {
            FloatButton.Width = size;
            FloatButton.Height = size;

            // 同步更新折叠状态下的窗口尺寸
            if (!_isPanelExpanded)
            {
                this.Width = size + 20;
                this.Height = size + 20;
            }
        }

        /// <summary>
        /// 滑动条值改变：调整悬浮球大小
        /// </summary>
        private void BallSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_config == null) return; // 初始化时跳过

            double size = Math.Round(e.NewValue);
            _config.BallSize = size;
            BallSizeValueText.Text = $"{size} px";
            ApplyBallSize(size);
            SaveConfig();
        }

        /// <summary>
        /// 下拉框选择变更：切换展开特效
        /// </summary>
        private void EffectComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (_config == null) return; // 初始化时跳过

            var selected = EffectComboBox.SelectedItem as ComboBoxItem;
            if (selected?.Tag is string effectName)
            {
                _config.OpenEffect = effectName;
                SaveConfig();
            }
        }

        #region 隐藏桌面图标 (Win32 API)

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter,
                                                   string? lpszClass, string? lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // 切换桌面图标显示的 Shell 命令 ID
        private const uint WM_COMMAND = 0x0111;
        private const int TOGGLE_DESKTOP_ICONS_CMD = 0x7402;

        /// <summary>
        /// 获取桌面图标列表窗口句柄 (SHELLDLL_DefView)
        /// </summary>
        private static IntPtr GetDesktopIconsHandle()
        {
            IntPtr progman = FindWindow("Progman", null);
            IntPtr defView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);

            if (defView == IntPtr.Zero)
            {
                // Wallpaper Engine 等软件会把桌面图标移到 WorkerW 子窗口中
                IntPtr workerW = IntPtr.Zero;
                do
                {
                    workerW = FindWindowEx(IntPtr.Zero, workerW, "WorkerW", null);
                    if (workerW != IntPtr.Zero)
                    {
                        defView = FindWindowEx(workerW, IntPtr.Zero, "SHELLDLL_DefView", null);
                        if (defView != IntPtr.Zero) break;
                    }
                } while (workerW != IntPtr.Zero);
            }

            return defView;
        }

        /// <summary>
        /// 切换桌面图标显示/隐藏（等价于右键桌面→查看→显示桌面图标）
        /// 使用 Shell 命令方式，不影响 Wallpaper Engine 等动态壁纸软件
        /// </summary>
        private void ToggleDesktopIcons()
        {
            IntPtr defView = GetDesktopIconsHandle();
            if (defView != IntPtr.Zero)
            {
                SendMessage(defView, WM_COMMAND, (IntPtr)TOGGLE_DESKTOP_ICONS_CMD, IntPtr.Zero);
            }
        }

        /// <summary>
        /// 设置桌面图标可见性（仅在目标状态与当前不同时才切换）
        /// </summary>
        private void SetDesktopIconsVisibility(bool visible)
        {
            // 检测当前是否已有图标显示：找 SHELLDLL_DefView 下的 SysListView32
            IntPtr defView = GetDesktopIconsHandle();
            if (defView == IntPtr.Zero) return;

            IntPtr listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
            if (listView == IntPtr.Zero) return;

            // IsWindowVisible 判断 ListView 当前是否可见
            bool currentlyVisible = NativeIsWindowVisible(listView);

            // 如果当前状态已经是目标状态就不用切换
            if (currentlyVisible != visible)
            {
                ToggleDesktopIcons();
            }
        }

        [DllImport("user32.dll", EntryPoint = "IsWindowVisible")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool NativeIsWindowVisible(IntPtr hWnd);

        #endregion

        /// <summary>
        /// 隐藏桌面图标开关
        /// </summary>
        private void HideDesktopToggle_Click(object sender, RoutedEventArgs e)
        {
            bool hide = HideDesktopToggle.IsChecked == true;
            _config.HideDesktopIcons = hide;
            SetDesktopIconsVisibility(!hide);
            SaveConfig();
        }

        /// <summary>
        /// 一键扫描桌面图标
        /// </summary>
        private void ScanDesktopButton_Click(object sender, RoutedEventArgs e)
        {
            var desktopFiles = ShortcutCategorizer.ScanDesktopFiles();

            if (desktopFiles.Count == 0)
            {
                MessageBox.Show("未在桌面发现快捷方式。", "扫描结果",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 过滤掉已存在的快捷方式（按目标路径或图标路径去重）
            var existingPaths = new HashSet<string>(
                _shortcuts.Select(s => s.TargetPath.ToLowerInvariant())
                .Concat(_shortcuts.Select(s => s.IconSourcePath.ToLowerInvariant())));

            int addedCount = 0;
            foreach (var file in desktopFiles)
            {
                if (existingPaths.Contains(file.ToLowerInvariant())) continue;

                try
                {
                    var item = ShortcutResolver.ResolveShortcut(file);
                    if (existingPaths.Contains(item.TargetPath.ToLowerInvariant())) continue;

                    // 自动分类
                    item.Category = ShortcutCategorizer.Categorize(item);
                    _shortcuts.Add(item);
                    existingPaths.Add(item.TargetPath.ToLowerInvariant());
                    existingPaths.Add(file.ToLowerInvariant());
                    addedCount++;
                }
                catch { /* 解析失败则跳过 */ }
            }

            if (addedCount > 0)
            {
                RenderAllShortcuts();
                SaveConfig();
                MessageBox.Show($"扫描完成！\n新增 {addedCount} 个快捷方式，已自动分类。",
                                "扫描导入", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("桌面上的快捷方式已全部导入过了。",
                                "扫描结果", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// 一键自动归类
        /// </summary>
        private void AutoCategorizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (_shortcuts.Count == 0)
            {
                MessageBox.Show("当前没有快捷方式可以归类。", "自动归类",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            ShortcutCategorizer.CategorizeAll(_shortcuts);
            RenderAllShortcuts();
            SaveConfig();

            // 统计分类结果
            var stats = _shortcuts.GroupBy(s => s.Category)
                                  .Select(g => $"  {g.Key}: {g.Count()} 个")
                                  .ToList();

            MessageBox.Show($"已完成自动归类！\n\n" + string.Join("\n", stats),
                            "归类结果", MessageBoxButton.OK, MessageBoxImage.Information);
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

        #region 窗口边缘拖拽调整大小

        // 边缘检测阈值（像素）
        private const int ResizeGripSize = 6;
        // 最小尺寸
        private const double MinWindowWidth = 300;
        private const double MinWindowHeight = 350;

        // 拖拽调整大小的状态
        private bool _isResizing = false;
        private int _resizeDirection = 0;
        private Point _resizeStartPos; // 屏幕坐标
        private double _resizeStartWidth, _resizeStartHeight, _resizeStartLeft, _resizeStartTop;

        // 方向常量
        private const int DIR_LEFT = 1;
        private const int DIR_RIGHT = 2;
        private const int DIR_TOP = 3;
        private const int DIR_TOPLEFT = 4;
        private const int DIR_TOPRIGHT = 5;
        private const int DIR_BOTTOM = 6;
        private const int DIR_BOTTOMLEFT = 7;
        private const int DIR_BOTTOMRIGHT = 8;

        /// <summary>
        /// 检测鼠标位置相对于窗口边缘的方向（0 = 不在边缘）
        /// </summary>
        private int GetResizeDirection(Point pos)
        {
            if (!_isPanelExpanded) return 0;

            bool left = pos.X < ResizeGripSize;
            bool right = pos.X >= this.ActualWidth - ResizeGripSize;
            bool top = pos.Y < ResizeGripSize;
            bool bottom = pos.Y >= this.ActualHeight - ResizeGripSize;

            if (top && left) return DIR_TOPLEFT;
            if (top && right) return DIR_TOPRIGHT;
            if (bottom && left) return DIR_BOTTOMLEFT;
            if (bottom && right) return DIR_BOTTOMRIGHT;
            if (left) return DIR_LEFT;
            if (right) return DIR_RIGHT;
            if (top) return DIR_TOP;
            if (bottom) return DIR_BOTTOM;

            return 0;
        }

        /// <summary>
        /// 根据方向设置鼠标光标形状
        /// </summary>
        private void UpdateResizeCursor(int dir)
        {
            switch (dir)
            {
                case DIR_LEFT:
                case DIR_RIGHT:
                    this.Cursor = Cursors.SizeWE;
                    break;
                case DIR_TOP:
                case DIR_BOTTOM:
                    this.Cursor = Cursors.SizeNS;
                    break;
                case DIR_TOPLEFT:
                case DIR_BOTTOMRIGHT:
                    this.Cursor = Cursors.SizeNWSE;
                    break;
                case DIR_TOPRIGHT:
                case DIR_BOTTOMLEFT:
                    this.Cursor = Cursors.SizeNESW;
                    break;
                default:
                    this.Cursor = Cursors.Arrow;
                    break;
            }
        }

        /// <summary>
        /// 窗口鼠标移动：更新边缘光标 + 拖拽调整大小
        /// </summary>
        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isResizing && e.LeftButton == MouseButtonState.Pressed)
            {
                // 获取当前屏幕坐标
                Point screenPos = this.PointToScreen(e.GetPosition(this));
                double dx = screenPos.X - _resizeStartPos.X;
                double dy = screenPos.Y - _resizeStartPos.Y;

                // 根据方向调整窗口
                double newW = _resizeStartWidth, newH = _resizeStartHeight;
                double newL = _resizeStartLeft, newT = _resizeStartTop;

                bool hasLeft = _resizeDirection == DIR_LEFT || _resizeDirection == DIR_TOPLEFT || _resizeDirection == DIR_BOTTOMLEFT;
                bool hasRight = _resizeDirection == DIR_RIGHT || _resizeDirection == DIR_TOPRIGHT || _resizeDirection == DIR_BOTTOMRIGHT;
                bool hasTop = _resizeDirection == DIR_TOP || _resizeDirection == DIR_TOPLEFT || _resizeDirection == DIR_TOPRIGHT;
                bool hasBottom = _resizeDirection == DIR_BOTTOM || _resizeDirection == DIR_BOTTOMLEFT || _resizeDirection == DIR_BOTTOMRIGHT;

                if (hasRight) newW = Math.Max(MinWindowWidth, _resizeStartWidth + dx);
                if (hasBottom) newH = Math.Max(MinWindowHeight, _resizeStartHeight + dy);
                if (hasLeft)
                {
                    newW = Math.Max(MinWindowWidth, _resizeStartWidth - dx);
                    newL = _resizeStartLeft + (_resizeStartWidth - newW);
                }
                if (hasTop)
                {
                    newH = Math.Max(MinWindowHeight, _resizeStartHeight - dy);
                    newT = _resizeStartTop + (_resizeStartHeight - newH);
                }

                this.Left = newL;
                this.Top = newT;
                this.Width = newW;
                this.Height = newH;
            }
            else if (e.LeftButton == MouseButtonState.Released)
            {
                _isResizing = false;
                UpdateResizeCursor(GetResizeDirection(e.GetPosition(this)));
            }
        }

        /// <summary>
        /// 窗口鼠标按下：如果在边缘则开始拖拽调整大小
        /// </summary>
        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            int dir = GetResizeDirection(e.GetPosition(this));
            if (dir != 0)
            {
                _isResizing = true;
                _resizeDirection = dir;
                _resizeStartPos = this.PointToScreen(e.GetPosition(this));
                _resizeStartWidth = this.ActualWidth;
                _resizeStartHeight = this.ActualHeight;
                _resizeStartLeft = this.Left;
                _resizeStartTop = this.Top;
                this.CaptureMouse();
                e.Handled = true;
            }
        }

        /// <summary>
        /// 窗口鼠标释放：结束调整大小
        /// </summary>
        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            if (_isResizing)
            {
                _isResizing = false;
                this.ReleaseMouseCapture();
                this.Cursor = Cursors.Arrow;
            }
            base.OnMouseLeftButtonUp(e);
        }

        #endregion
    }
}
