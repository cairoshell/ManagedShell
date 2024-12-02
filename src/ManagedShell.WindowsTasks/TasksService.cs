using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.Interop;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using ManagedShell.Common.Enums;
using ManagedShell.Common.SupportingClasses;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.WindowsTasks
{
    public class TasksService : DependencyObject, IDisposable
    {
        public static readonly IconSize DEFAULT_ICON_SIZE = IconSize.Small;

        public event EventHandler<WindowActivatedEventArgs> WindowActivated;
        public event EventHandler<EventArgs> DesktopActivated;

        private NativeWindowEx _HookWin;
        private object _windowsLock = new object();
        internal bool IsInitialized;
        private IconSize _taskIconSize;

        private static int WM_SHELLHOOKMESSAGE = -1;
        private static int WM_TASKBARCREATEDMESSAGE = -1;
        private static int TASKBARBUTTONCREATEDMESSAGE = -1;
        private static IntPtr cloakEventHook = IntPtr.Zero;
        private WinEventProc cloakEventProc;
        private static IntPtr moveEventHook = IntPtr.Zero;
        private WinEventProc moveEventProc;

        internal ITaskCategoryProvider TaskCategoryProvider;
        private TaskCategoryChangeDelegate CategoryChangeDelegate;

        public IconSize TaskIconSize
        {
            get { return _taskIconSize; }
            set
            {
                if (value == _taskIconSize)
                {
                    return;
                }

                _taskIconSize = value;

                if (!IsInitialized)
                {
                    return;
                }

                foreach (var window in Windows)
                {
                    if (!window.ShowInTaskbar)
                    {
                        return;
                    }

                    window.UpdateProperties();
                }
            }
        }

        public TasksService() : this(DEFAULT_ICON_SIZE)
        {
        }
        
        public TasksService(IconSize iconSize)
        {
            TaskIconSize = iconSize;
        }

        internal void Initialize(bool withMultiMonTracking)
        {
            if (IsInitialized)
            {
                return;
            }

            try
            {
                ShellLogger.Debug("TasksService: Starting");

                // create window to receive task events
                _HookWin = new NativeWindowEx();
                _HookWin.CreateHandle(new CreateParams());

                // prevent other shells from working properly
                SetTaskmanWindow(_HookWin.Handle);

                // register to receive task events
                RegisterShellHookWindow(_HookWin.Handle);
                WM_SHELLHOOKMESSAGE = RegisterWindowMessage("SHELLHOOK");
                WM_TASKBARCREATEDMESSAGE = RegisterWindowMessage("TaskbarCreated");
                TASKBARBUTTONCREATEDMESSAGE = RegisterWindowMessage("TaskbarButtonCreated");
                _HookWin.MessageReceived += ShellWinProc;

                if (EnvironmentHelper.IsWindows8OrBetter)
                {
                    // set event hook for cloak/uncloak events
                    cloakEventProc = CloakEventCallback;

                    if (cloakEventHook == IntPtr.Zero)
                    {
                        cloakEventHook = SetWinEventHook(
                            EVENT_OBJECT_CLOAKED,
                            EVENT_OBJECT_UNCLOAKED,
                            IntPtr.Zero,
                            cloakEventProc,
                            0,
                            0,
                            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
                    }
                }

                if (withMultiMonTracking)
                {
                    // set event hook for move events
                    moveEventProc = MoveEventCallback;

                    if (moveEventHook == IntPtr.Zero)
                    {
                        moveEventHook = SetWinEventHook(
                            EVENT_OBJECT_LOCATIONCHANGE,
                            EVENT_OBJECT_LOCATIONCHANGE,
                            IntPtr.Zero,
                            moveEventProc,
                            0,
                            0,
                            WINEVENT_OUTOFCONTEXT);
                    }
                }

                // set window for ITaskbarList
                setTaskbarListHwnd(_HookWin.Handle);

                // adjust minimize animation
                SetMinimizedMetrics();

                // enumerate windows already opened and set active window
                getInitialWindows();

                IsInitialized = true;
            }
            catch (Exception ex)
            {
                ShellLogger.Info("TasksService: Unable to start: " + ex.Message);
            }
        }

        internal void SetTaskCategoryProvider(ITaskCategoryProvider provider)
        {
            TaskCategoryProvider = provider;

            if (CategoryChangeDelegate == null)
            {
                CategoryChangeDelegate = CategoriesChanged;
            }

            TaskCategoryProvider.SetCategoryChangeDelegate(CategoryChangeDelegate);
        }

        private void getInitialWindows()
        {
            EnumWindows((hwnd, lParam) =>
            {
                ApplicationWindow win = new ApplicationWindow(this, hwnd);

                if (win.CanAddToTaskbar && win.ShowInTaskbar && !Windows.Contains(win))
                {
                    Windows.Add(win);

                    sendTaskbarButtonCreatedMessage(win.Handle);
                }

                return true;
            }, 0);

            IntPtr hWndForeground = GetForegroundWindow();
            if (Windows.Any(i => i.Handle == hWndForeground && i.ShowInTaskbar))
            {
                ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWndForeground);
                win.State = ApplicationWindow.WindowState.Active;
                win.SetShowInTaskbar();
            }
        }

        public void Dispose()
        {
            if (IsInitialized)
            {
                ShellLogger.Debug("TasksService: Deregistering hooks");
                DeregisterShellHookWindow(_HookWin.Handle);
                if (cloakEventHook != IntPtr.Zero) UnhookWinEvent(cloakEventHook);
                if (moveEventHook != IntPtr.Zero) UnhookWinEvent(moveEventHook);
                _HookWin.DestroyHandle();
                setTaskbarListHwnd(IntPtr.Zero);
            }

            TaskCategoryProvider?.Dispose();
        }

        private void CategoriesChanged()
        {
            foreach (ApplicationWindow window in Windows)
            {
                if (window.ShowInTaskbar)
                {
                    window.Category = TaskCategoryProvider?.GetCategory(window);
                }
            }
        }

        private void SetMinimizedMetrics()
        {
            MinimizedMetrics mm = new MinimizedMetrics
            {
                cbSize = (uint)Marshal.SizeOf(typeof(MinimizedMetrics))
            };

            IntPtr mmPtr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(MinimizedMetrics)));

            try
            {
                Marshal.StructureToPtr(mm, mmPtr, true);
                SystemParametersInfo(SPI.GETMINIMIZEDMETRICS, mm.cbSize, mmPtr, SPIF.None);
                mm.iWidth = 140;
                mm.iArrange |= MinimizedMetricsArrangement.Hide;
                Marshal.StructureToPtr(mm, mmPtr, true);
                SystemParametersInfo(SPI.SETMINIMIZEDMETRICS, mm.cbSize, mmPtr, SPIF.None);
            }
            finally
            {
                Marshal.DestroyStructure(mmPtr, typeof(MinimizedMetrics));
                Marshal.FreeHGlobal(mmPtr);
            }
        }

        public void CloseWindow(ApplicationWindow window)
        {
            if (window.DoClose() != IntPtr.Zero)
            {
                ShellLogger.Debug($"TasksService: Removing window {window.Title} from collection due to no response");
                window.Dispose();
                Windows.Remove(window);
            }
        }

        private void sendTaskbarButtonCreatedMessage(IntPtr hWnd)
        {
            // Server Core doesn't support ITaskbarList, so sending this message on that OS could cause some assuming apps to crash
            if (!EnvironmentHelper.IsServerCore) SendNotifyMessage(hWnd, (uint)TASKBARBUTTONCREATEDMESSAGE, UIntPtr.Zero, IntPtr.Zero);
        }

        private ApplicationWindow addWindow(IntPtr hWnd, ApplicationWindow.WindowState initialState = ApplicationWindow.WindowState.Inactive, bool sanityCheck = false)
        {
            ApplicationWindow win = new ApplicationWindow(this, hWnd);

            // set window state if a non-default value is provided
            if (initialState != ApplicationWindow.WindowState.Inactive) win.State = initialState;

            // add window unless we need to validate it is eligible to show in taskbar
            if (!sanityCheck || win.CanAddToTaskbar) Windows.Add(win);

            // Only send TaskbarButtonCreated if we are shell, and if OS is not Server Core
            // This is because if Explorer is running, it will send the message, so we don't need to
            if (EnvironmentHelper.IsAppRunningAsShell) sendTaskbarButtonCreatedMessage(win.Handle);

            return win;
        }

        private void removeWindow(IntPtr hWnd)
        {
            if (Windows.Any(i => i.Handle == hWnd))
            {
                do
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    win.Dispose();
                    Windows.Remove(win);
                }
                while (Windows.Any(i => i.Handle == hWnd));
            }
        }

        private void redrawWindow(ApplicationWindow win)
        {
            win.UpdateProperties();

            foreach (ApplicationWindow wind in Windows)
            {
                if (wind.WinFileName == win.WinFileName && wind.Handle != win.Handle)
                {
                    wind.UpdateProperties();
                }
            }
        }

        private void ShellWinProc(Message msg)
        {
            if (msg.Msg == WM_SHELLHOOKMESSAGE)
            {
                try
                {
                    lock (_windowsLock)
                    {
                        switch ((HSHELL)msg.WParam.ToInt32())
                        {
                            case HSHELL.WINDOWCREATED:
                                ShellLogger.Debug("TasksService: Created: " + msg.LParam);
                                if (!Windows.Any(i => i.Handle == msg.LParam))
                                {
                                    addWindow(msg.LParam);
                                }
                                else
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msg.LParam);
                                    win.UpdateProperties();
                                }
                                break;

                            case HSHELL.WINDOWDESTROYED:
                                ShellLogger.Debug("TasksService: Destroyed: " + msg.LParam);
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.WINDOWREPLACING:
                                ShellLogger.Debug("TasksService: Replacing: " + msg.LParam);
                                if (Windows.Any(i => i.Handle == msg.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msg.LParam);
                                    win.State = ApplicationWindow.WindowState.Inactive;
                                    win.SetShowInTaskbar();
                                }
                                else
                                {
                                    addWindow(msg.LParam);
                                }
                                break;
                            case HSHELL.WINDOWREPLACED:
                                ShellLogger.Debug("TasksService: Replaced: " + msg.LParam);
                                // TODO: If a window gets replaced, we lose app-level state such as overlay icons.
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.WINDOWACTIVATED:
                            case HSHELL.RUDEAPPACTIVATED:
                                ShellLogger.Debug("TasksService: Activated: " + msg.LParam);

                                foreach (var aWin in Windows.Where(w => w.State == ApplicationWindow.WindowState.Active))
                                {
                                    aWin.State = ApplicationWindow.WindowState.Inactive;
                                }

                                if (msg.LParam != IntPtr.Zero)
                                {
                                    ApplicationWindow win = null;

                                    if (Windows.Any(i => i.Handle == msg.LParam))
                                    {
                                        win = Windows.First(wnd => wnd.Handle == msg.LParam);
                                        win.State = ApplicationWindow.WindowState.Active;
                                        win.SetShowInTaskbar();
                                    }
                                    else
                                    {
                                        win = addWindow(msg.LParam, ApplicationWindow.WindowState.Active);
                                    }

                                    if (win != null)
                                    {
                                        foreach (ApplicationWindow wind in Windows)
                                        {
                                            if (wind.WinFileName == win.WinFileName && wind.Handle != win.Handle)
                                                wind.SetShowInTaskbar();
                                        }

                                        WindowActivatedEventArgs args = new WindowActivatedEventArgs
                                        {
                                            Window = win
                                        };

                                        WindowActivated?.Invoke(this, args);
                                    }
                                }
                                else
                                {
                                    DesktopActivated?.Invoke(this, new EventArgs());
                                }
                                break;

                            case HSHELL.FLASH:
                                ShellLogger.Debug("TasksService: Flashing window: " + msg.LParam);
                                if (Windows.Any(i => i.Handle == msg.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msg.LParam);
                                    
                                    if (win.State != ApplicationWindow.WindowState.Active)
                                    {
                                        win.State = ApplicationWindow.WindowState.Flashing;
                                    }

                                    redrawWindow(win);
                                }
                                else
                                {
                                    addWindow(msg.LParam, ApplicationWindow.WindowState.Flashing, true);
                                }
                                break;

                            case HSHELL.ACTIVATESHELLWINDOW:
                                ShellLogger.Debug("TasksService: Activate shell window called.");
                                break;

                            case HSHELL.ENDTASK:
                                ShellLogger.Debug("TasksService: EndTask called: " + msg.LParam);
                                removeWindow(msg.LParam);
                                break;

                            case HSHELL.GETMINRECT:
                                ShellLogger.Debug("TasksService: GetMinRect called: " + msg.LParam);
                                SHELLHOOKINFO winHandle = (SHELLHOOKINFO)Marshal.PtrToStructure(msg.LParam, typeof(SHELLHOOKINFO));
                                winHandle.rc = new NativeMethods.Rect { Bottom = 100, Left = 0, Right = 100, Top = 0 };
                                Marshal.StructureToPtr(winHandle, msg.LParam, true);
                                msg.Result = winHandle.hwnd;
                                return; // return here so the result isnt reset to DefWindowProc

                            case HSHELL.REDRAW:
                                ShellLogger.Debug("TasksService: Redraw called: " + msg.LParam);
                                if (Windows.Any(i => i.Handle == msg.LParam))
                                {
                                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == msg.LParam);

                                    if (win.State == ApplicationWindow.WindowState.Flashing)
                                    {
                                        win.State = ApplicationWindow.WindowState.Inactive;
                                    }

                                    redrawWindow(win);
                                }
                                else
                                {
                                    addWindow(msg.LParam, ApplicationWindow.WindowState.Inactive, true);
                                }
                                break;

                            // TaskMan needs to return true if we provide our own task manager to prevent explorers.
                            // case HSHELL.TASKMAN:
                            //     SingletonLogger.Instance.Info("TaskMan Message received.");
                            //     break;

                            default:
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ShellLogger.Error("TasksService: Error in ShellWinProc. ", ex);
                    Debugger.Break();
                }
            }
            else if (msg.Msg == WM_TASKBARCREATEDMESSAGE)
            {
                ShellLogger.Debug("TasksService: TaskbarCreated received, setting ITaskbarList window");
                setTaskbarListHwnd(_HookWin.Handle);
            }
            else
            {
                // Handle ITaskbarList functions, most not implemented yet

                ApplicationWindow win = null;

                switch (msg.Msg)
                {
                    case (int)WM.USER + 50:
                        // ActivateTab
                        // Also sends WM_SHELLHOOK message
                        ShellLogger.Debug("TasksService: ITaskbarList: ActivateTab HWND:" + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 60:
                        // MarkFullscreenWindow
                        ShellLogger.Debug("TasksService: ITaskbarList: MarkFullscreenWindow HWND:" + msg.LParam + " Entering? " + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 64:
                        // SetProgressValue
                        ShellLogger.Debug("TasksService: ITaskbarList: SetProgressValue HWND:" + msg.WParam + " Progress: " + msg.LParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msg.WParam);
                            win.ProgressValue = (int)msg.LParam;
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 65:
                        // SetProgressState
                        ShellLogger.Debug("TasksService: ITaskbarList: SetProgressState HWND:" + msg.WParam + " Flags: " + msg.LParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msg.WParam);
                            win.ProgressState = (TBPFLAG)msg.LParam;
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 67:
                        // RegisterTab
                        ShellLogger.Debug("TasksService: ITaskbarList: RegisterTab MDI HWND:" + msg.LParam + " Tab HWND: " + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 68:
                        // UnregisterTab
                        ShellLogger.Debug("TasksService: ITaskbarList: UnregisterTab Tab HWND: " + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 71:
                        // SetTabOrder
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabOrder HWND:" + msg.WParam + " Before HWND: " + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 72:
                        // SetTabActive
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabActive HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 75:
                        // Unknown
                        ShellLogger.Debug("TasksService: ITaskbarList: Unknown HWND:" + msg.WParam + " LParam: " + msg.LParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 76:
                        // ThumbBarAddButtons
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarAddButtons HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 77:
                        // ThumbBarUpdateButtons
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarUpdateButtons HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 78:
                        // ThumbBarSetImageList
                        ShellLogger.Debug("TasksService: ITaskbarList: ThumbBarSetImageList HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 79:
                        // SetOverlayIcon - Icon
                        ShellLogger.Debug("TasksService: ITaskbarList: SetOverlayIcon - Icon HWND:" + msg.WParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msg.WParam);
                            win.SetOverlayIcon(msg.LParam);
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 80:
                        // SetThumbnailTooltip
                        ShellLogger.Debug("TasksService: ITaskbarList: SetThumbnailTooltip HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 81:
                        // SetThumbnailClip
                        ShellLogger.Debug("TasksService: ITaskbarList: SetThumbnailClip HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 85:
                        // SetOverlayIcon - Description
                        ShellLogger.Debug("TasksService: ITaskbarList: SetOverlayIcon - Description HWND:" + msg.WParam);

                        win = new ApplicationWindow(this, msg.WParam);
                        if (Windows.Contains(win))
                        {
                            win = Windows.First(wnd => wnd.Handle == msg.WParam);
                            win.SetOverlayIconDescription(msg.LParam);
                        }

                        msg.Result = IntPtr.Zero;
                        return;
                    case (int)WM.USER + 87:
                        // SetTabProperties
                        ShellLogger.Debug("TasksService: ITaskbarList: SetTabProperties HWND:" + msg.WParam);
                        msg.Result = IntPtr.Zero;
                        return;
                    default:
                        ShellLogger.Debug($"TasksService: Unknown ITaskbarList Msg: {msg.Msg} LParam: {msg.LParam} WParam: {msg.WParam}");
                        break;
                }
            }

            msg.Result = DefWindowProc(msg.HWnd, msg.Msg, msg.WParam, msg.LParam);
        }

        private void MoveEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hWnd != IntPtr.Zero && idObject == 0 && idChild == 0)
            {
                if (Windows.Any(i => i.Handle == hWnd))
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    win.SetMonitor();
                }
            }
        }

        private void CloakEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (hWnd != IntPtr.Zero && idObject == 0 && idChild == 0)
            {
                if (Windows.Any(i => i.Handle == hWnd))
                {
                    ApplicationWindow win = Windows.First(wnd => wnd.Handle == hWnd);
                    ShellLogger.Debug($"TasksService: {(eventType == EVENT_OBJECT_CLOAKED ? "Cloak" : "Uncloak")} event received for {win.Title}");
                    win.SetShowInTaskbar();
                }
            }
        }

        // set property on hook window that should receive ITaskbarList messages
        private void setTaskbarListHwnd(IntPtr hwndHook)
        {
            bool resetProp = true;

            // get the topmost tray
            IntPtr taskbarHwnd = WindowHelper.FindWindowsTray(IntPtr.Zero);
            
            if (taskbarHwnd == IntPtr.Zero)
            {
                return;
            }

            // if our tray is running, there may also be a second tray running
            IntPtr systemTaskbarHwnd = WindowHelper.FindWindowsTray(taskbarHwnd);

            if (hwndHook == IntPtr.Zero)
            {
                // no target hwnd provided
                // Try to find and use the handle of the Explorer hook window
                resetProp = false;
                hwndHook = getChildHwndByClass(systemTaskbarHwnd == IntPtr.Zero ? taskbarHwnd : systemTaskbarHwnd, "MSTaskSwWClass");
            }

            if (hwndHook == IntPtr.Zero)
            {
                // if still no hwnd to hook, we can't do anything
                return;
            }

            ShellLogger.Debug("TasksService: Adding TaskbandHWND prop to hwnd: " + taskbarHwnd);
            SetProp(taskbarHwnd, "TaskbandHWND", hwndHook);

            // Remove the property from the Explorer taskbar, if it is not the only tray
            if (resetProp && systemTaskbarHwnd != IntPtr.Zero)
            {
                ShellLogger.Debug("TasksService: Removing TaskbandHWND prop from hwnd: " + systemTaskbarHwnd);
                RemoveProp(systemTaskbarHwnd, "TaskbandHWND");
            }
        }

        private IntPtr getChildHwndByClass(IntPtr parentHwnd, string wndClass)
        {
            IntPtr childHwnd = IntPtr.Zero;
            EnumChildWindows(parentHwnd, (hwnd, lParam) =>
            {
                StringBuilder cName = new StringBuilder(256);
                GetClassName(hwnd, cName, cName.Capacity);
                if (cName.ToString() == wndClass)
                {
                    childHwnd = hwnd;
                    return false;
                }

                return true;
            }, 0);

            return childHwnd;
        }

        internal ObservableCollection<ApplicationWindow> Windows
        {
            get
            {
                return base.GetValue(windowsProperty) as ObservableCollection<ApplicationWindow>;
            }
            set
            {
                SetValue(windowsProperty, value);
            }
        }

        private DependencyProperty windowsProperty = DependencyProperty.Register("Windows",
            typeof(ObservableCollection<ApplicationWindow>), typeof(TasksService),
            new PropertyMetadata(new ObservableCollection<ApplicationWindow>()));
    }
}
