using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DesktopOrganizer.Models;

namespace DesktopOrganizer.Services
{
    /// <summary>
    /// 快捷方式解析服务：解析 .lnk 文件的真实路径，提取程序图标
    /// </summary>
    public static class ShortcutResolver
    {
        #region COM 接口定义 (IShellLink)

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink { }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile,
                         int cch, IntPtr pfd, uint fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out ushort pwHotkey);
            void SetHotkey(ushort wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath,
                                 int cch, out int piIcon);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
            void Resolve(IntPtr hwnd, uint fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        #endregion

        /// <summary>
        /// 解析 .lnk 快捷方式文件，提取目标路径、参数和工作目录
        /// </summary>
        public static ShortcutItem ResolveShortcut(string lnkPath)
        {
            var item = new ShortcutItem
            {
                IconSourcePath = lnkPath,
                Name = Path.GetFileNameWithoutExtension(lnkPath)
            };

            string extension = Path.GetExtension(lnkPath).ToLowerInvariant();

            if (extension == ".lnk")
            {
                try
                {
                    // 使用 IShellLink COM 接口解析 .lnk 文件
                    var link = (IShellLinkW)new ShellLink();
                    var persistFile = (IPersistFile)link;
                    persistFile.Load(lnkPath, 0);

                    // 获取目标路径
                    var pathBuffer = new StringBuilder(260);
                    link.GetPath(pathBuffer, pathBuffer.Capacity, IntPtr.Zero, 0);
                    item.TargetPath = pathBuffer.ToString();

                    // 获取参数
                    var argsBuffer = new StringBuilder(1024);
                    link.GetArguments(argsBuffer, argsBuffer.Capacity);
                    item.Arguments = argsBuffer.ToString();

                    // 获取工作目录
                    var dirBuffer = new StringBuilder(260);
                    link.GetWorkingDirectory(dirBuffer, dirBuffer.Capacity);
                    item.WorkingDirectory = dirBuffer.ToString();

                    // 如果名称为空，用目标文件名代替
                    if (string.IsNullOrEmpty(item.Name) && !string.IsNullOrEmpty(item.TargetPath))
                    {
                        item.Name = Path.GetFileNameWithoutExtension(item.TargetPath);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"解析 .lnk 失败: {ex.Message}");
                    // 解析失败时使用文件名
                    item.TargetPath = lnkPath;
                }
            }
            else if (extension == ".exe")
            {
                // .exe 文件直接使用路径
                item.TargetPath = lnkPath;
                item.Name = Path.GetFileNameWithoutExtension(lnkPath);
            }
            else
            {
                // 其他文件类型，直接使用路径
                item.TargetPath = lnkPath;
            }

            return item;
        }

        /// <summary>
        /// 从文件路径提取图标，转换为 WPF 可用的 ImageSource
        /// </summary>
        public static ImageSource? ExtractIcon(string filePath)
        {
            try
            {
                // 优先尝试从目标路径提取图标
                string iconPath = filePath;

                if (!File.Exists(iconPath))
                    return null;

                // 使用 System.Drawing.Icon 提取关联图标
                Icon? icon = null;
                try
                {
                    icon = Icon.ExtractAssociatedIcon(iconPath);
                }
                catch
                {
                    // 部分路径可能无法提取图标
                }

                if (icon == null)
                    return null;

                // 将 System.Drawing.Icon 转换为 WPF BitmapSource
                using (icon)
                {
                    var bitmap = icon.ToBitmap();
                    var hBitmap = bitmap.GetHbitmap();
                    try
                    {
                        return Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                    }
                    finally
                    {
                        DeleteObject(hBitmap);
                        bitmap.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"提取图标失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 释放 GDI 位图句柄
        /// </summary>
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
