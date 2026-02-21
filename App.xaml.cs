using System;
using System.Threading;
using System.Windows;

namespace DesktopOrganizer
{
    /// <summary>
    /// 应用入口，使用 Mutex 防止多开
    /// </summary>
    public partial class App : Application
    {
        private static Mutex? _mutex;

        protected override void OnStartup(StartupEventArgs e)
        {
            const string mutexName = "DesktopOrganizer_SingleInstance_Mutex";
            _mutex = new Mutex(true, mutexName, out bool isNew);

            if (!isNew)
            {
                // 已有实例在运行，提示后退出
                MessageBox.Show("桌面收纳盒已在运行中！", "提示",
                                MessageBoxButton.OK, MessageBoxImage.Information);
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
            base.OnExit(e);
        }
    }
}
