using ManagedShell.Common.Logging;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ManagedShell.Common.Helpers;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.AppBar
{
    public class AppBarManager : IDisposable
    {
        private static object appBarLock = new object();
        
        private readonly ExplorerHelper _explorerHelper;
        private int uCallBack;
        
        public List<AppBarWindow> AppBars { get; } = new List<AppBarWindow>();
        public EventHandler<AppBarEventArgs> AppBarEvent;

        public AppBarManager(ExplorerHelper explorerHelper)
        {
            _explorerHelper = explorerHelper;
        }

        public void SignalGracefulShutdown()
        {
            foreach (AppBarWindow window in AppBars)
            {
                window.AllowClose = true;
            }
        }

        public void NotifyAppBarEvent(AppBarWindow sender, AppBarEventReason reason)
        {
            AppBarEventArgs args = new AppBarEventArgs { Reason = reason };
            AppBarEvent?.Invoke(sender, args);
        }

        #region AppBar message helpers
        public int RegisterBar(AppBarWindow abWindow, double width, double height, AppBarEdge edge = AppBarEdge.Top)
        {
            lock (appBarLock)
            {
                APPBARDATA abd = new APPBARDATA();
                abd.cbSize = Marshal.SizeOf(typeof(APPBARDATA));
                abd.hWnd = abWindow.Handle;

                if (!AppBars.Contains(abWindow))
                {
                    if (!EnvironmentHelper.IsAppRunningAsShell)
                    {
                        uCallBack = RegisterWindowMessage("AppBarMessage");
                        abd.uCallbackMessage = uCallBack;
                        
                        SHAppBarMessage((int) ABMsg.ABM_NEW, ref abd);
                    }
                    
                    AppBars.Add(abWindow);
                    
                    ShellLogger.Debug($"AppBarManager: Created AppBar for handle {abWindow.Handle}");

                    if (!EnvironmentHelper.IsAppRunningAsShell)
                    {
                        ABSetPos(abWindow, width, height, edge, true);
                    }
                    else
                    {
                        SetWorkArea(abWindow.Screen);
                    }
                }
                else
                {
                    if (!EnvironmentHelper.IsAppRunningAsShell)
                    {
                        SHAppBarMessage((int) ABMsg.ABM_REMOVE, ref abd);
                    }

                    AppBars.Remove(abWindow);
                    ShellLogger.Debug($"AppBarManager: Removed AppBar for handle {abWindow.Handle}");

                    if (EnvironmentHelper.IsAppRunningAsShell)
                    {
                        SetWorkArea(abWindow.Screen);
                    }

                    return 0;
                }
            }

            return uCallBack;
        }

        public void AppBarActivate(IntPtr hwnd)
        {
            APPBARDATA abd = new APPBARDATA
            {
                cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                hWnd = hwnd,
                lParam = (IntPtr)Convert.ToInt32(true)
            };
            
            SHAppBarMessage((int)ABMsg.ABM_ACTIVATE, ref abd);

            // apparently the TaskBars like to pop up when AppBars change
            if (_explorerHelper.HideExplorerTaskbar)
            {
                _explorerHelper.SetSecondaryTaskbarVisibility((int)SetWindowPosFlags.SWP_HIDEWINDOW);
            }
        }

        public void AppBarWindowPosChanged(IntPtr hwnd)
        {
            APPBARDATA abd = new APPBARDATA
            {
                cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                hWnd = hwnd
            };
            
            SHAppBarMessage((int)ABMsg.ABM_WINDOWPOSCHANGED, ref abd);

            // apparently the TaskBars like to pop up when AppBars change
            if (_explorerHelper.HideExplorerTaskbar)
            {
                _explorerHelper.SetSecondaryTaskbarVisibility((int)SetWindowPosFlags.SWP_HIDEWINDOW);
            }
        }

        public void ABSetPos(AppBarWindow abWindow, double width, double height, AppBarEdge edge, bool isCreate = false)
        {
            lock (appBarLock)
            {
                APPBARDATA abd = new APPBARDATA
                {
                    cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                    hWnd = abWindow.Handle,
                    uEdge = (int)edge
                };

                int sWidth = (int)width;
                int sHeight = (int)height;

                int top = 0;
                int left = 0;
                int right = ScreenHelper.PrimaryMonitorDeviceSize.Width;
                int bottom = ScreenHelper.PrimaryMonitorDeviceSize.Height;

                int edgeOffset = 0;

                if (abWindow.Screen != null)
                {
                    top = abWindow.Screen.Bounds.Y;
                    left = abWindow.Screen.Bounds.X;
                    right = abWindow.Screen.Bounds.Right;
                    bottom = abWindow.Screen.Bounds.Bottom;
                }

                if (!abWindow.RequiresScreenEdge)
                {
                    edgeOffset = Convert.ToInt32(GetAppBarEdgeWindowsHeight((AppBarEdge)abd.uEdge, abWindow.Screen));
                }

                if (abd.uEdge == (int)AppBarEdge.Left || abd.uEdge == (int)AppBarEdge.Right)
                {
                    abd.rc.Top = top;
                    abd.rc.Bottom = bottom;
                    if (abd.uEdge == (int)AppBarEdge.Left)
                    {
                        abd.rc.Left = left + edgeOffset;
                        abd.rc.Right = abd.rc.Left + sWidth;
                    }
                    else
                    {
                        abd.rc.Right = right - edgeOffset;
                        abd.rc.Left = abd.rc.Right - sWidth;
                    }
                }
                else
                {
                    abd.rc.Left = left;
                    abd.rc.Right = right;
                    if (abd.uEdge == (int)AppBarEdge.Top)
                    {
                        abd.rc.Top = top + edgeOffset;
                        abd.rc.Bottom = abd.rc.Top + sHeight;
                    }
                    else
                    {
                        abd.rc.Bottom = bottom - edgeOffset;
                        abd.rc.Top = abd.rc.Bottom - sHeight;
                    }
                }

                _explorerHelper.SuspendTrayService();
                SHAppBarMessage((int)ABMsg.ABM_QUERYPOS, ref abd);
                _explorerHelper.ResumeTrayService();

                // system doesn't adjust all edges for us, do some adjustments
                switch (abd.uEdge)
                {
                    case (int)AppBarEdge.Left:
                        abd.rc.Right = abd.rc.Left + sWidth;
                        break;
                    case (int)AppBarEdge.Right:
                        abd.rc.Left = abd.rc.Right - sWidth;
                        break;
                    case (int)AppBarEdge.Top:
                        abd.rc.Bottom = abd.rc.Top + sHeight;
                        break;
                    case (int)AppBarEdge.Bottom:
                        abd.rc.Top = abd.rc.Bottom - sHeight;
                        break;
                }

                _explorerHelper.SuspendTrayService();
                SHAppBarMessage((int)ABMsg.ABM_SETPOS, ref abd);
                _explorerHelper.ResumeTrayService();

                // check if new coords
                bool isSameCoords = false;
                if (!isCreate)
                {
                    bool topUnchanged = abd.rc.Top == (abWindow.Top * abWindow.DpiScale);
                    bool leftUnchanged = abd.rc.Left == (abWindow.Left * abWindow.DpiScale);
                    bool bottomUnchanged = abd.rc.Bottom == (abWindow.Top * abWindow.DpiScale) + sHeight;
                    bool rightUnchanged = abd.rc.Right == (abWindow.Left * abWindow.DpiScale) + sWidth;

                    isSameCoords = topUnchanged
                                   && leftUnchanged
                                   && bottomUnchanged
                                   && rightUnchanged;
                }

                if (!isSameCoords)
                {
                    ShellLogger.Debug($"AppBarManager: {abWindow.Name} changing position (TxLxBxR) to {abd.rc.Top}x{abd.rc.Left}x{abd.rc.Bottom}x{abd.rc.Right} from {abWindow.Top * abWindow.DpiScale}x{abWindow.Left * abWindow.DpiScale}x{(abWindow.Top * abWindow.DpiScale) + sHeight}x{ (abWindow.Left * abWindow.DpiScale) + sWidth}");
                    abWindow.SetAppBarPosition(abd.rc);
                }

                abWindow.AfterAppBarPos(isSameCoords, abd.rc);

                if (abd.rc.Bottom - abd.rc.Top < sHeight)
                {
                    ABSetPos(abWindow, width, height, edge);
                }
            }
        }
        #endregion

        #region Work area
        public double GetAppBarEdgeWindowsHeight(AppBarEdge edge, AppBarScreen screen)
        {
            double edgeHeight = 0;
            double dpiScale = 1;
            Rect workAreaRect = GetWorkArea(ref dpiScale, screen, true, true);

            switch (edge)
            {
                case AppBarEdge.Top:
                    edgeHeight += workAreaRect.Top / dpiScale;
                    break;
                case AppBarEdge.Bottom:
                    edgeHeight += (screen.Bounds.Bottom - workAreaRect.Bottom) / dpiScale;
                    break;
                case AppBarEdge.Left:
                    edgeHeight += workAreaRect.Left / dpiScale;
                    break;
                case AppBarEdge.Right:
                    edgeHeight += (screen.Bounds.Right - workAreaRect.Right) / dpiScale;
                    break;
            }

            return edgeHeight;
        }

        public Rect GetWorkArea(ref double dpiScale, AppBarScreen screen, bool edgeBarsOnly, bool enabledBarsOnly)
        {
            double topEdgeWindowHeight = 0;
            double bottomEdgeWindowHeight = 0;
            double leftEdgeWindowWidth = 0;
            double rightEdgeWindowWidth = 0;
            Rect rc;

            // get appropriate windows for this display
            foreach (var window in AppBars)
            {
                if (window.Screen.DeviceName == screen.DeviceName)
                {
                    if ((window.EnableAppBar || !enabledBarsOnly) && (window.RequiresScreenEdge || !edgeBarsOnly))
                    {
                        if (window.AppBarEdge == AppBarEdge.Top)
                        {
                            topEdgeWindowHeight += window.ActualHeight;
                        }
                        else if (window.AppBarEdge == AppBarEdge.Bottom)
                        {
                            bottomEdgeWindowHeight += window.ActualHeight;
                        }
                        else if (window.AppBarEdge == AppBarEdge.Left)
                        {
                            leftEdgeWindowWidth += window.ActualWidth;
                        }
                        else if (window.AppBarEdge == AppBarEdge.Right)
                        {
                            rightEdgeWindowWidth += window.ActualWidth;
                        }
                    }

                    dpiScale = window.DpiScale;
                }
            }

            rc.Top = screen.Bounds.Top + (int)(topEdgeWindowHeight * dpiScale);
            rc.Bottom = screen.Bounds.Bottom - (int)(bottomEdgeWindowHeight * dpiScale);
            rc.Left = screen.Bounds.Left + (int)(leftEdgeWindowWidth * dpiScale);
            rc.Right = screen.Bounds.Right - (int)(rightEdgeWindowWidth * dpiScale);

            return rc;
        }

        public void SetWorkArea(AppBarScreen screen)
        {
            double dpiScale = 1;
            Rect rc = GetWorkArea(ref dpiScale, screen, false, true);

            SystemParametersInfo((int)SPI.SETWORKAREA, 1, ref rc, (uint)(SPIF.UPDATEINIFILE | SPIF.SENDWININICHANGE));
        }

        public static void ResetWorkArea()
        {
            if (EnvironmentHelper.IsAppRunningAsShell)
            {
                // TODO this is wrong for multi-display
                // set work area back to full screen size. we can't assume what pieces of the old work area may or may not be still used
                Rect oldWorkArea;
                oldWorkArea.Left = SystemInformation.VirtualScreen.Left;
                oldWorkArea.Top = SystemInformation.VirtualScreen.Top;
                oldWorkArea.Right = SystemInformation.VirtualScreen.Right;
                oldWorkArea.Bottom = SystemInformation.VirtualScreen.Bottom;

                SystemParametersInfo((int)SPI.SETWORKAREA, 1, ref oldWorkArea,
                    (uint)(SPIF.UPDATEINIFILE | SPIF.SENDWININICHANGE));
            }
        }
        #endregion

        public void Dispose()
        {
            ResetWorkArea();
        }
    }
}