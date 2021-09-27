using ManagedShell.Common.Helpers;
using System;
using System.Drawing;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.WindowsTray
{
    public class NotificationBalloonInfo
    {
        public string Info { get; internal set; }

        public string Title { get; internal set; }

        public NIIF Flags { get; internal set; }

        public ImageSource Icon { get; internal set; }

        public int Timeout { get; internal set; }

        public NotificationBalloonInfo() { }

        public NotificationBalloonInfo(SafeNotifyIconData nicData)
        {
            Title = nicData.szInfoTitle;
            Info = nicData.szInfo;
            Flags = nicData.dwInfoFlags;
            Timeout = (int)nicData.uVersion;

            if ((NIIF.USER & Flags) != 0)
            {
                if (nicData.hBalloonIcon != 0)
                {
                    SetIconFromHIcon((IntPtr)nicData.hBalloonIcon);
                }
                else if (nicData.hIcon != IntPtr.Zero)
                {
                    SetIconFromHIcon(nicData.hIcon);
                }
            }
            else if ((NIIF.INFO & Flags) != 0)
            {
                Icon = GetSystemIcon(SystemIcons.Information.Handle);
            }
            else if ((NIIF.WARNING & Flags) != 0)
            {
                Icon = GetSystemIcon(SystemIcons.Warning.Handle);
            }
            else if ((NIIF.ERROR & Flags) != 0)
            {
                Icon = GetSystemIcon(SystemIcons.Error.Handle);
            }
        }

        private void SetIconFromHIcon(IntPtr hIcon)
        {
            if (hIcon == IntPtr.Zero)
            {
                if (Icon == null)
                {
                    // Use default only if we don't have a valid icon already
                    Icon = IconImageConverter.GetDefaultIcon();
                }

                return;
            }

            ImageSource icon = IconImageConverter.GetImageFromHIcon(hIcon, false);

            if (icon != null)
            {
                Icon = icon;
            }
            else if (icon == null && Icon == null)
            {
                // Use default only if we don't have a valid icon already
                Icon = IconImageConverter.GetDefaultIcon();
            }
        }

        private BitmapSource GetSystemIcon(IntPtr hIcon)
        {
            BitmapSource bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(hIcon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            bs.Freeze();

            return bs;
        }
    }
}
