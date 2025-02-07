using ManagedShell.Common.Logging;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ManagedShell.Common.Helpers;
using static ManagedShell.Interop.NativeMethods;
using ManagedShell.WindowsTray;

namespace ManagedShell.AppBar
{
    public class AppBarManager : IDisposable
    {
        private static object appBarLock = new object();
        
        private readonly ExplorerHelper _explorerHelper;
        private AppBarMessageDelegate _appBarMessageDelegate;
        private int uCallBack;

        private int retryNum;
        private Rect retryRect;
        private DateTime retryTimestamp;
        private int maxRetryNum = 20;
        private TimeSpan maxRetryTimespan = TimeSpan.FromSeconds(10);

        public List<AppBarWindow> AppBars { get; } = new List<AppBarWindow>();
        public List<AppBarWindow> AutoHideBars { get; } = new List<AppBarWindow>();
        public EventHandler<AppBarEventArgs> AppBarEvent;

        public AppBarManager(ExplorerHelper explorerHelper)
        {
            _appBarMessageDelegate = appBarMessageDelegate;
            _explorerHelper = explorerHelper;

            _explorerHelper._notificationArea?.SetAppBarMessageCallback(_appBarMessageDelegate);
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

        private IntPtr appBarMessageDelegate(APPBARMSGDATAV3 amd, ref bool handled)
        {
            // only handle certain messages, send other AppBar messages to default handler
            switch ((ABMsg)amd.dwMessage)
            {
                case ABMsg.ABM_GETTASKBARPOS:
                    return appBarMessage_GetTaskbarPos(amd, ref handled);
                case ABMsg.ABM_QUERYPOS:
                case ABMsg.ABM_SETPOS:
                    return appBarMessage_QuerySetPos(amd, ref handled);
                case ABMsg.ABM_GETSTATE:
                    return appBarMessage_GetState(amd, ref handled);
                case ABMsg.ABM_GETAUTOHIDEBAR:
                case ABMsg.ABM_GETAUTOHIDEBAREX:
                    return appBarMessage_GetAutoHideBar(amd, ref handled);
                case ABMsg.ABM_ACTIVATE:
                case ABMsg.ABM_WINDOWPOSCHANGED:
                    handled = true;
                    return (IntPtr)1;
            }
            return IntPtr.Zero;
        }

        #region AppBar message handlers
        private IntPtr appBarMessage_GetTaskbarPos(APPBARMSGDATAV3 amd, ref bool handled)
        {
            IntPtr hShared = SHLockShared((IntPtr)amd.hSharedMemory, (uint)amd.dwSourceProcessId);
            APPBARDATAV2 abd = (APPBARDATAV2)Marshal.PtrToStructure(hShared, typeof(APPBARDATAV2));

            if (_explorerHelper._notificationArea != null)
            {
                _explorerHelper._notificationArea.FillTrayHostSizeData(ref abd);
            }

            Marshal.StructureToPtr(abd, hShared, false);
            SHUnlockShared(hShared);
            handled = true;
            return (IntPtr)1;
        }

        private IntPtr appBarMessage_QuerySetPos(APPBARMSGDATAV3 amd, ref bool handled)
        {
            // These two messages use shared memory, and forwarding over the message as-is doesn't
            // seem to allow Explorer to access the shared memory. Here we grab the existing (old)
            // shared memory and allocate it into new shared memory, then update AppBarMessageData
            // and forward it on to Explorer.

            if (EnvironmentHelper.IsAppRunningAsShell)
            {
                // some day we will manage AppBars if we are shell, but today is not that day.
                return IntPtr.Zero;
            }

            // Get Explorer tray handle and PID
            IntPtr ignoreHwnd = IntPtr.Zero;
            IntPtr explorerTray;

            if (_explorerHelper._notificationArea != null)
            {
                ignoreHwnd = _explorerHelper._notificationArea.Handle;
            }
            explorerTray = WindowHelper.FindWindowsTray(ignoreHwnd);

            GetWindowThreadProcessId(explorerTray, out uint explorerPid);

            // recreate shared memory so that Explorer gets access to it
            IntPtr hSharedOld = SHLockShared((IntPtr)amd.hSharedMemory, (uint)amd.dwSourceProcessId);
            IntPtr hSharedNew = SHAllocShared(IntPtr.Zero, (uint)Marshal.SizeOf(typeof(APPBARDATAV2)), explorerPid);

            // Copy the data from the old shared memory into the new
            IntPtr hSharedData = SHLockShared(hSharedNew, explorerPid);
            if (hSharedData == IntPtr.Zero)
            {
                // Failed, bail out bail out!
                SHFreeShared(hSharedNew, explorerPid);
                return IntPtr.Zero;
            }

            APPBARDATAV2 abdOld = (APPBARDATAV2)Marshal.PtrToStructure(hSharedOld, typeof(APPBARDATAV2));
            Marshal.StructureToPtr(abdOld, hSharedData, false);
            SHUnlockShared(hSharedData);

            // Update AppBarMessageData with the new shared memory handle and PID
            amd.hSharedMemory = (long)hSharedNew;
            amd.dwSourceProcessId = (int)explorerPid;

            // Prepare structs to send onward
            IntPtr hAmd = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(APPBARMSGDATAV3)));
            Marshal.StructureToPtr(amd, hAmd, false);

            COPYDATASTRUCT copyData = new COPYDATASTRUCT
            {
                cbData = Marshal.SizeOf(typeof(APPBARMSGDATAV3)),
                dwData = (IntPtr)0,
                lpData = hAmd
            };
            IntPtr hCopyData = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(COPYDATASTRUCT)));
            Marshal.StructureToPtr(copyData, hCopyData, false);

            IntPtr result = SendMessage(explorerTray, (int)WM.COPYDATA, (IntPtr)amd.abd.hWnd, hCopyData);
            handled = true;

            // It's possible that Explorer modified the data we sent, so read the data back out.
            IntPtr hSharedFromExplorer = SHLockShared(hSharedNew, explorerPid);
            if (hSharedFromExplorer != IntPtr.Zero)
            {
                APPBARDATAV2 abdNew = (APPBARDATAV2)Marshal.PtrToStructure(hSharedFromExplorer, typeof(APPBARDATAV2));
                SHUnlockShared(hSharedFromExplorer);

                Marshal.StructureToPtr(abdNew, hSharedOld, false);
                SHUnlockShared(hSharedOld);
            }

            SHFreeShared(hSharedNew, explorerPid);

            return result;
        }

        private IntPtr appBarMessage_GetState(APPBARMSGDATAV3 amd, ref bool handled)
        {
            if (_explorerHelper._notificationArea != null && amd.abd.hWnd == (uint)WindowHelper.FindWindowsTray(_explorerHelper._notificationArea.Handle))
            {
                // If the Explorer AppBar is being queried specifically, forward the request on to it.
                // ExplorerHelper queries this to get the pre-existing taskbar state before hiding.
                return IntPtr.Zero;
            }

            handled = true;

            if (AutoHideBars.Count > 0)
            {
                return (IntPtr)ABState.AutoHide;
            }

            return (IntPtr)ABState.Default;
        }

        private IntPtr appBarMessage_GetAutoHideBar(APPBARMSGDATAV3 amd, ref bool handled)
        {
            var autoHideBar = AutoHideBars.Find(bar => (int)bar.AppBarEdge == amd.abd.uEdge && bar.AppBarMode == AppBarMode.AutoHide);

            handled = true;

            if (autoHideBar != null)
            {
                // Return the notification area hwnd instead of the AppBar's. Why?
                // Some apps (Firefox) check the class of the AppBar matches Shell_TrayWnd.
                // However, an AppBarWindow should be providing its coordinates to a
                // NotificationArea to position it appropriately anyway.
                return _explorerHelper?._notificationArea?.Handle ?? autoHideBar.Handle;
            }

            return IntPtr.Zero;
        }
        #endregion

        #region AppBar message helpers
        public void RegisterAutoHideBar(AppBarWindow window)
        {
            if (AutoHideBars.Contains(window))
            {
                return;
            }

            AutoHideBars.Add(window);
        }

        public void UnregisterAutoHideBar(AppBarWindow window)
        {
            if (!AutoHideBars.Contains(window))
            {
                return;
            }

            AutoHideBars.Remove(window);
        }

        public int RegisterBar(AppBarWindow abWindow)
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
                        ABSetPos(abWindow, true);
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

        public void ABSetPos(AppBarWindow abWindow, bool isCreate = false)
        {
            lock (appBarLock)
            {
                Rect desiredRect = abWindow.GetDesiredRect();
                APPBARDATA abd = new APPBARDATA
                {
                    cbSize = Marshal.SizeOf(typeof(APPBARDATA)),
                    hWnd = abWindow.Handle,
                    uEdge = (int)abWindow.AppBarEdge,
                    rc = desiredRect
                };

                SHAppBarMessage((int)ABMsg.ABM_QUERYPOS, ref abd);

                // system doesn't adjust all edges for us, do some adjustments
                switch (abd.uEdge)
                {
                    case (int)AppBarEdge.Left:
                        abd.rc.Right = abd.rc.Left + desiredRect.Width;
                        break;
                    case (int)AppBarEdge.Right:
                        abd.rc.Left = abd.rc.Right - desiredRect.Width;
                        break;
                    case (int)AppBarEdge.Top:
                        abd.rc.Bottom = abd.rc.Top + desiredRect.Height;
                        break;
                    case (int)AppBarEdge.Bottom:
                        abd.rc.Top = abd.rc.Bottom - desiredRect.Height;
                        break;
                }

                SHAppBarMessage((int)ABMsg.ABM_SETPOS, ref abd);

                // check if new coords
                bool isSameCoords = false;
                Rect currentRect;
                GetWindowRect(abWindow.Handle, out currentRect);
                if (!isCreate)
                {
                    bool topUnchanged = abd.rc.Top == currentRect.Top;
                    bool leftUnchanged = abd.rc.Left == currentRect.Left;
                    bool bottomUnchanged = abd.rc.Bottom == currentRect.Bottom;
                    bool rightUnchanged = abd.rc.Right == currentRect.Right;

                    isSameCoords = topUnchanged
                                   && leftUnchanged
                                   && bottomUnchanged
                                   && rightUnchanged;
                }

                if (!isSameCoords)
                {
                    ShellLogger.Debug($"AppBarManager: {abWindow.Name} changing position (TxLxBxR) to {abd.rc.Top}x{abd.rc.Left}x{abd.rc.Bottom}x{abd.rc.Right} from {currentRect.Top}x{currentRect.Left}x{currentRect.Bottom}x{currentRect.Right}");
                    abWindow.SetAppBarPosition(abd.rc);
                }

                abWindow.AfterAppBarPos(isSameCoords, abd.rc);

                if ((((abd.uEdge == (int)AppBarEdge.Top || abd.uEdge == (int)AppBarEdge.Bottom) && abd.rc.Bottom - abd.rc.Top < desiredRect.Height) ||
                    ((abd.uEdge == (int)AppBarEdge.Left || abd.uEdge == (int)AppBarEdge.Right) && abd.rc.Right - abd.rc.Left < desiredRect.Width)) && allowRetry(abd.rc))
                {
                    // The system did not respect the coordinates we selected, resulting in an unexpected window size.
                    ABSetPos(abWindow);
                }
            }
        }

        private bool allowRetry(Rect rect)
        {
            // The system did not respect the coordinates we selected. This may or may not need remediation, so keep track of attempts to prevent infinite looping.

            if (rect.Top == retryRect.Top && rect.Left == retryRect.Left && rect.Right == retryRect.Right && rect.Bottom == retryRect.Bottom)
            {
                // Repeat rect
                if (DateTime.Now.Subtract(retryTimestamp) < maxRetryTimespan)
                {
                    // within retry span
                    if (retryNum >= maxRetryNum)
                    {
                        // hit max retries
                        ShellLogger.Debug("AppBarManager: Max retries limit reached");
                        return false;
                    }
                    else
                    {
                        // allow retry
                        ShellLogger.Debug("AppBarManager: Allowing retry of ABSetPos");
                        retryNum++;
                        return true;
                    }
                }
            }

            // Reset
            ShellLogger.Debug("AppBarManager: Resetting retry of ABSetPos");
            retryNum = 0;
            retryRect = rect;
            retryTimestamp = DateTime.Now;

            return true;
        }
        #endregion

        #region Work area
        public int GetAppBarEdgeWindowsHeight(AppBarEdge edge, AppBarScreen screen)
        {
            int edgeHeight = 0;
            Rect workAreaRect = GetWorkArea(screen, true, true);

            switch (edge)
            {
                case AppBarEdge.Top:
                    edgeHeight += workAreaRect.Top - screen.Bounds.Top;
                    break;
                case AppBarEdge.Bottom:
                    edgeHeight += screen.Bounds.Bottom - workAreaRect.Bottom;
                    break;
                case AppBarEdge.Left:
                    edgeHeight += workAreaRect.Left - screen.Bounds.Left;
                    break;
                case AppBarEdge.Right:
                    edgeHeight += screen.Bounds.Right - workAreaRect.Right;
                    break;
            }

            return edgeHeight;
        }

        public Rect GetWorkArea(AppBarScreen screen, bool edgeBarsOnly, bool enabledBarsOnly)
        {
            int topEdgeWindowHeight = 0;
            int bottomEdgeWindowHeight = 0;
            int leftEdgeWindowWidth = 0;
            int rightEdgeWindowWidth = 0;
            Rect rc;

            // get appropriate windows for this display
            foreach (var window in AppBars)
            {
                if (window.Screen.DeviceName == screen.DeviceName && 
                    (window.AppBarMode == AppBarMode.Normal || !enabledBarsOnly) && 
                    (window.RequiresScreenEdge || !edgeBarsOnly))
                {
                    Rect windowRect;
                    GetWindowRect(window.Handle, out windowRect);
                    switch (window.AppBarEdge)
                    {
                        case AppBarEdge.Left:
                            leftEdgeWindowWidth += windowRect.Width;
                            break;
                        case AppBarEdge.Right:
                            rightEdgeWindowWidth += windowRect.Width;
                            break;
                        case AppBarEdge.Bottom:
                            bottomEdgeWindowHeight += windowRect.Height;
                            break;
                        case AppBarEdge.Top:
                            topEdgeWindowHeight += windowRect.Height;
                            break;
                    }
                }
            }

            rc.Top = screen.Bounds.Top + topEdgeWindowHeight;
            rc.Bottom = screen.Bounds.Bottom - bottomEdgeWindowHeight;
            rc.Left = screen.Bounds.Left + leftEdgeWindowWidth;
            rc.Right = screen.Bounds.Right - rightEdgeWindowWidth;

            return rc;
        }

        public void SetWorkArea(AppBarScreen screen)
        {
            Rect rc = GetWorkArea(screen, false, true);

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