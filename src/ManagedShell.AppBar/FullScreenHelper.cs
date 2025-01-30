using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.WindowsTasks;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Threading;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.AppBar
{
    public sealed class FullScreenHelper : IDisposable
    {
        private readonly DispatcherTimer _fullscreenCheck;
        private readonly TasksService _tasksService;

        public ObservableCollection<FullScreenApp> FullScreenApps = new ObservableCollection<FullScreenApp>();

        public FullScreenHelper(TasksService tasksService)
        {
            _tasksService = tasksService;

            if (_tasksService != null && EnvironmentHelper.IsWindows8OrBetter)
            {
                // On Windows 8 and newer, TasksService will tell us when windows enter and exit full screen
                _tasksService.FullScreenEntered += TasksService_Event;
                _tasksService.FullScreenLeft += TasksService_Event;
                _tasksService.MonitorChanged += TasksService_Event;
                _tasksService.DesktopActivated += TasksService_Event;
                _tasksService.WindowActivated += TasksService_Event;
                return;
            }

            _fullscreenCheck = new DispatcherTimer(DispatcherPriority.Background, System.Windows.Application.Current.Dispatcher)
            {
                Interval = new TimeSpan(0, 0, 0, 0, 100)
            };

            _fullscreenCheck.Tick += FullscreenCheck_Tick;
            _fullscreenCheck.Start();
        }

        private void TasksService_Event(object sender, EventArgs e)
        {
            updateFullScreenWindows();
        }

        private void FullscreenCheck_Tick(object sender, EventArgs e)
        {
            updateFullScreenWindows();
        }

        private void updateFullScreenWindows()
        {
            IntPtr hWnd = GetForegroundWindow();

            List<FullScreenApp> removeApps = new List<FullScreenApp>();
            bool skipAdd = false;

            // first check if this window is already in our list. if so, remove it if necessary
            foreach (FullScreenApp app in FullScreenApps)
            {
                FullScreenApp appCurrentState = getFullScreenApp(app.hWnd);

                if (app.hWnd == hWnd && appCurrentState != null && app.screen.DeviceName == appCurrentState.screen.DeviceName)
                {
                    // this window, still same screen, do nothing
                    skipAdd = true;
                    continue;
                }

                if (appCurrentState != null && app.hWnd != hWnd &&
                    app.screen.DeviceName == appCurrentState.screen.DeviceName &&
                    Screen.FromHandle(hWnd).DeviceName != appCurrentState.screen.DeviceName)
                {
                    // if the full-screen window is no longer foreground, keep it
                    // as long as the foreground window is on a different screen.
                    continue;
                }

                removeApps.Add(app);
            }

            // remove any changed windows we found
            if (removeApps.Count > 0)
            {
                foreach (FullScreenApp existingApp in removeApps)
                {
                    ShellLogger.Debug($"FullScreenHelper: Removing full screen app {existingApp.hWnd} ({existingApp.title})");
                    FullScreenApps.Remove(existingApp);
                }
            }

            // check if this is a new full screen app
            if (!skipAdd)
            {
                FullScreenApp appNew = getFullScreenApp(hWnd);
                if (appNew != null)
                {
                    ShellLogger.Debug($"FullScreenHelper: Adding full screen app {appNew.hWnd} ({appNew.title})");
                    FullScreenApps.Add(appNew);
                }
            }
        }

        private FullScreenApp getFullScreenApp(IntPtr hWnd)
        {
            int style = GetWindowLong(hWnd, GWL_STYLE);
            Rect rect;

            if ((((int)WindowStyles.WS_CAPTION | (int)WindowStyles.WS_THICKFRAME) & style) == ((int)WindowStyles.WS_CAPTION | (int)WindowStyles.WS_THICKFRAME))
            {
                GetClientRect(hWnd, out rect);
                MapWindowPoints(hWnd, IntPtr.Zero, ref rect, 2);
            }
            else
            {
                GetWindowRect(hWnd, out rect);
            }

            var allScreens = Screen.AllScreens.Select(ScreenInfo.Create).ToList();
            if (allScreens.Count > 1) allScreens.Add(ScreenInfo.CreateVirtualScreen());

            // check if this is a fullscreen app
            foreach (var screen in allScreens)
            {
                if (rect.Top == screen.Bounds.Top && rect.Left == screen.Bounds.Left &&
                    rect.Bottom == screen.Bounds.Bottom && rect.Right == screen.Bounds.Right)
                {
                    // make sure this is not us
                    GetWindowThreadProcessId(hWnd, out uint hwndProcId);
                    if (hwndProcId == GetCurrentProcessId())
                    {
                        return null;
                    }

                    // make sure this is fullscreen-able
                    if (!IsWindow(hWnd) || !IsWindowVisible(hWnd) || IsIconic(hWnd))
                    {
                        return null;
                    }

                    // Make sure this isn't explicitly marked as being non-rude
                    IntPtr isNonRudeHwnd = GetProp(hWnd, "NonRudeHWND");
                    if (isNonRudeHwnd != IntPtr.Zero)
                    {
                        return null;
                    }

                    // make sure this is not a cloaked window
                    if (EnvironmentHelper.IsWindows8OrBetter)
                    {
                        int cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(uint));
                        DwmGetWindowAttribute(hWnd, DWMWINDOWATTRIBUTE.DWMWA_CLOAKED, out uint cloaked, cbSize);
                        if (cloaked > 0)
                        {
                            return null;
                        }
                    }

                    ApplicationWindow win = new ApplicationWindow(null, hWnd);
                    if (!EnvironmentHelper.IsWindows8OrBetter)
                    {
                        // make sure this is not the shell desktop
                        // In Windows 8 and newer, the NonRudeHWND property is set and this is not needed
                        if (win.ClassName == "Progman" || win.ClassName == "WorkerW")
                        {
                            return null;
                        }
                    }

                    // make sure this is not a transparent window
                    int styles = win.ExtendedWindowStyles;
                    if ((styles & (int)ExtendedWindowStyles.WS_EX_LAYERED) != 0 && ((styles & (int)ExtendedWindowStyles.WS_EX_TRANSPARENT) != 0 || (styles & (int)ExtendedWindowStyles.WS_EX_NOACTIVATE) != 0))
                    {
                        return null;
                    }

                    // this is a full screen app on this screen
                    return new FullScreenApp { hWnd = hWnd, screen = screen, rect = rect, title = win.Title };
                }
            }

            return null;
        }

        private void ResetScreenCache()
        {
            // use reflection to empty screens cache
            const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic;
            var fi = typeof(Screen).GetField("screens", flags) ?? typeof(Screen).GetField("s_screens", flags);

            if (fi == null)
            {
                ShellLogger.Warning("FullScreenHelper: Unable to reset screens cache");
                return;
            }

            fi.SetValue(null, null);
        }

        public void NotifyScreensChanged()
        {
            ResetScreenCache();
        }

        public void Dispose()
        {
            _fullscreenCheck?.Stop();

            if (_tasksService != null && EnvironmentHelper.IsWindows8OrBetter)
            {
                _tasksService.FullScreenEntered -= TasksService_Event;
                _tasksService.FullScreenLeft -= TasksService_Event;
                _tasksService.MonitorChanged -= TasksService_Event;
                _tasksService.DesktopActivated -= TasksService_Event;
                _tasksService.WindowActivated -= TasksService_Event;
                return;
            }
        }
    }
}
