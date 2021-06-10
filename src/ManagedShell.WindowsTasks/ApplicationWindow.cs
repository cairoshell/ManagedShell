using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.Interop;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using ManagedShell.Common.Enums;

namespace ManagedShell.WindowsTasks
{
    [DebuggerDisplay("Title: {Title}, Handle: {Handle}")]
    public class ApplicationWindow : IEquatable<ApplicationWindow>, INotifyPropertyChanged, IDisposable
    {
        const int TITLE_LENGTH = 1024;
        private readonly TasksService _tasksService;
        StringBuilder titleBuilder = new StringBuilder(TITLE_LENGTH);

        public ApplicationWindow(TasksService tasksService, IntPtr handle)
        {
            _tasksService = tasksService;
            Handle = handle;
            State = WindowState.Inactive;
        }

        public void Dispose()
        {
            // no longer required
        }

        public IntPtr Handle
        {
            get;
            set;
        }

        private string _appUserModelId = null;

        public string AppUserModelID
        {
            get
            {
                if (string.IsNullOrEmpty(_appUserModelId))
                {
                    _appUserModelId = ShellHelper.GetAppUserModelIdPropertyForHandle(Handle);
                }

                return _appUserModelId;
            }
        }

        private bool? _isUWP = null;

        public bool IsUWP
        {
            get
            {
                if (_isUWP == null)
                {
                    _isUWP = WinFileName.ToLower().Contains("applicationframehost.exe");
                }

                return (bool)_isUWP;
            }
        }

        private string _winFileName = "";

        public string WinFileName
        {
            get
            {
                if (string.IsNullOrEmpty(_winFileName))
                {
                    _winFileName = ShellHelper.GetPathForHandle(Handle);
                }

                return _winFileName;
            }
        }

        private uint? _procId;

        public uint? ProcId => _procId = _procId ?? ShellHelper.GetProcIdForHandle(Handle);

        private string _category;

        public string Category
        {
            get
            {
                return _category;
            }
            set
            {
                if (value != _category)
                {
                    _category = value;
                    OnPropertyChanged("Category");
                }
            }
        }

        private string _title;

        public string Title
        {
            get
            {
                if (_title == null)
                {
                    setTitle();
                }

                return _title;
            }
        }

        private void setTitle()
        {
            string title = "";
            try
            {
                titleBuilder.Clear();
                NativeMethods.GetWindowText(Handle, titleBuilder, TITLE_LENGTH + 1);

                title = titleBuilder.ToString();
            }
            catch { }

            if (_title != title)
            {
                _title = title;
                OnPropertyChanged("Title");
            }
        }

        private bool _iconLoading;
        private ImageSource _icon;
        private IntPtr _hIcon = IntPtr.Zero;

        public ImageSource Icon
        {
            get
            {
                if (_icon == null)
                    setIcon();

                return _icon;
            }
            set
            {
                _icon = value;
                OnPropertyChanged("Icon");
            }
        }
        
        private ImageSource _overlayIcon;

        public ImageSource OverlayIcon
        {
            get
            {
                return _overlayIcon;
            }
            private set
            {
                _overlayIcon = value;
                OnPropertyChanged("OverlayIcon");
            }
        }

        private string _overlayIconDescription;

        public string OverlayIconDescription
        {
            get
            {
                return _overlayIconDescription;
            }
            private set
            {
                _overlayIconDescription = value;
                OnPropertyChanged("OverlayIconDescription");
            }
        }

        private NativeMethods.TBPFLAG _progressState;

        public NativeMethods.TBPFLAG ProgressState
        {
            get
            {
                return _progressState;
            }

            set
            {
                _progressState = value;

                if (value == NativeMethods.TBPFLAG.TBPF_NOPROGRESS)
                {
                    ProgressValue = 0;
                }

                OnPropertyChanged("ProgressState");
            }
        }

        private int _progressValue;

        public int ProgressValue
        {
            get
            {
                return _progressValue;
            }

            set
            {
                _progressValue = value;
                OnPropertyChanged("ProgressValue");
            }
        }

        private WindowState _state;

        public WindowState State
        {
            get
            {
                return _state;
            }

            set
            {
                _state = value;
                OnPropertyChanged("State");
            }
        }

        public bool IsMinimized
        {
            get { return NativeMethods.IsIconic(Handle); }
        }

        public NativeMethods.WindowShowStyle ShowStyle
        {
            get { return GetWindowShowStyle(Handle); }
        }

        public int WindowStyles
        {
            get
            {
                return NativeMethods.GetWindowLong(Handle, NativeMethods.GWL_STYLE);
            }
        }

        public int ExtendedWindowStyles
        {
            get
            {
                return NativeMethods.GetWindowLong(Handle, NativeMethods.GWL_EXSTYLE);
            }
        }

        public bool CanAddToTaskbar
        {
            get
            {
                int extendedWindowStyles = ExtendedWindowStyles;
                bool isWindow = NativeMethods.IsWindow(Handle);
                bool isVisible = NativeMethods.IsWindowVisible(Handle);
                bool isToolWindow = (extendedWindowStyles & (int)NativeMethods.ExtendedWindowStyles.WS_EX_TOOLWINDOW) != 0;
                bool isAppWindow = (extendedWindowStyles & (int)NativeMethods.ExtendedWindowStyles.WS_EX_APPWINDOW) != 0;
                bool isNoActivate = (extendedWindowStyles & (int)NativeMethods.ExtendedWindowStyles.WS_EX_NOACTIVATE) != 0;
                IntPtr ownerWin = NativeMethods.GetWindow(Handle, NativeMethods.GetWindow_Cmd.GW_OWNER);

                return isWindow && isVisible && (ownerWin == IntPtr.Zero || isAppWindow) && (!isNoActivate || isAppWindow) && !isToolWindow;
            }
        }

        private bool? _showInTaskbar;

        // True if this window should be shown in the taskbar
        public bool ShowInTaskbar
        {
            get
            {
                if (_showInTaskbar == null)
                {
                    SetShowInTaskbar();
                }

                return (bool)_showInTaskbar;
            }
        }

        public void SetShowInTaskbar()
        {
            bool showInTaskbar = getShowInTaskbar();

            if (_showInTaskbar != showInTaskbar)
            {
                _showInTaskbar = showInTaskbar;

                // If we are becoming visible in the taskbar, get the category if it hasn't been set yet
                if (_showInTaskbar == true && Category == null && _tasksService.TaskCategoryProvider != null)
                {
                    // This is still a bit hacky because we need to call this synchronous as we access the list of Windows in TryGetCategoryDisplayNameByProcId
                    // and also mutate that same list in getInitialWindows so we can not evaluate categories in parallel
                    var task = Task.Run(async ()
                        => await _tasksService.TaskCategoryProvider.GetCategoryAsync(this)
                            .ConfigureAwait(false));
                    Category = task.GetAwaiter().GetResult();
                }

                OnPropertyChanged("ShowInTaskbar");
            }
        }

        private bool getShowInTaskbar()
        {
            // EnumWindows and ShellHook return UWP app windows that are 'cloaked', which should not be visible in the taskbar.
            if (EnvironmentHelper.IsWindows8OrBetter)
            {
                int cbSize = Marshal.SizeOf(typeof(uint));
                NativeMethods.DwmGetWindowAttribute(Handle, NativeMethods.DWMWINDOWATTRIBUTE.DWMWA_CLOAKED, out var cloaked, cbSize);

                if (cloaked > 0)
                {
                    ShellLogger.Debug($"ApplicationWindow: Cloaked ({cloaked}) window ({Title}) hidden from taskbar");
                    return false;
                }

                // UWP shell windows that are not cloaked should be hidden from the taskbar, too.
                StringBuilder cName = new StringBuilder(256);
                NativeMethods.GetClassName(Handle, cName, cName.Capacity);
                if (cName.ToString() == "ApplicationFrameWindow" || cName.ToString() == "Windows.UI.Core.CoreWindow" || cName.ToString() == "Shell_CharmWindow" || cName.ToString() == "ImmersiveLauncher")
                {
                    if ((ExtendedWindowStyles & (int)NativeMethods.ExtendedWindowStyles.WS_EX_WINDOWEDGE) == 0)
                    {
                        ShellLogger.Debug($"ApplicationWindow: Hiding UWP non-window {Title}");
                        return false;
                    }
                }
            }

            return CanAddToTaskbar;
        }

        internal void Uncloak()
        {
            ShellLogger.Debug($"ApplicationWindow: Uncloak event received for {Title}");

            SetShowInTaskbar();
        }

        private void setIcon()
        {
            if (!_iconLoading && ShowInTaskbar)
            {
                _iconLoading = true;

                Task.Factory.StartNew(() =>
                {
                    if (IsUWP && !string.IsNullOrEmpty(AppUserModelID))
                    {
                        // UWP apps
                        try
                        {
                            var storeApp = UWPInterop.StoreAppHelper.AppList.GetAppByAumid(AppUserModelID);

                            if (storeApp != null)
                            {
                                Icon = storeApp.GetIconImageSource(_tasksService.TaskIconSize);
                            }
                            else
                            {
                                Icon = IconImageConverter.GetDefaultIcon();
                            }
                        }
                        catch
                        {
                            if (_icon == null) Icon = IconImageConverter.GetDefaultIcon();
                        }
                    }
                    else
                    {
                        // non-UWP apps
                        IntPtr hIco = default;
                        uint WM_GETICON = (uint)NativeMethods.WM.GETICON;
                        uint WM_QUERYDRAGICON = (uint)NativeMethods.WM.QUERYDRAGICON;
                        int GCL_HICON = -14;
                        int GCL_HICONSM = -34;
                        IconSize sizeSetting = _tasksService.TaskIconSize;

                        if (sizeSetting == IconSize.Small)
                        {
                            NativeMethods.SendMessageTimeout(Handle, WM_GETICON, 2, 0, 2, 1000, ref hIco);
                            if (hIco == IntPtr.Zero)
                                NativeMethods.SendMessageTimeout(Handle, WM_GETICON, 0, 0, 2, 1000, ref hIco);
                        }
                        else
                        {
                            NativeMethods.SendMessageTimeout(Handle, WM_GETICON, 1, 0, 2, 1000, ref hIco);
                        }

                        if (hIco == IntPtr.Zero && sizeSetting == IconSize.Small)
                        {
                            if (!Environment.Is64BitProcess)
                                hIco = NativeMethods.GetClassLong(Handle, GCL_HICONSM);
                            else
                                hIco = NativeMethods.GetClassLongPtr(Handle, GCL_HICONSM);
                        }

                        if (hIco == IntPtr.Zero)
                        {
                            if (!Environment.Is64BitProcess)
                                hIco = NativeMethods.GetClassLong(Handle, GCL_HICON);
                            else
                                hIco = NativeMethods.GetClassLongPtr(Handle, GCL_HICON);
                        }

                        if (hIco == IntPtr.Zero)
                        {
                            NativeMethods.SendMessageTimeout(Handle, WM_QUERYDRAGICON, 0, 0, 0, 1000, ref hIco);
                        }

                        if (hIco == IntPtr.Zero && _icon == null)
                        {
                            // last resort: find icon by executable. if we already have an icon from a previous fetch, then just skip this
                            if (ShellHelper.Exists(WinFileName))
                            {
                                IconSize size = IconSize.Small;
                                if (sizeSetting != size)
                                    size = IconSize.Large;

                                hIco = IconHelper.GetIconByFilename(WinFileName, size);
                            }
                        }

                        if (hIco != IntPtr.Zero)
                        {
                            if (_hIcon != hIco)
                            {
                                _hIcon = hIco;
                                bool returnDefault = (_icon == null); // only return a default icon if we don't already have one. otherwise let's use what we have.
                                ImageSource icon = IconImageConverter.GetImageFromHIcon(hIco, returnDefault);
                                if (icon != null)
                                {
                                    icon.Freeze();
                                    Icon = icon;
                                }
                            }
                            else
                            {
                                NativeMethods.DestroyIcon(hIco);
                            }
                        }
                    }

                    _iconLoading = false;
                }, CancellationToken.None, TaskCreationOptions.None, IconHelper.IconScheduler);
            }
        }

        public void SetOverlayIcon(IntPtr hIcon)
        {
            if (hIcon == IntPtr.Zero)
            {
                OverlayIcon = null;
                return;
            }

            ImageSource icon = IconImageConverter.GetImageFromHIcon(hIcon, false);
            if (icon != null)
            {
                icon.Freeze();
                OverlayIcon = icon;
            }
        }

        public void SetOverlayIconDescription(IntPtr lParam)
        {
            try
            {
                if (ProcId is uint procId)
                {
                    IntPtr hShared = NativeMethods.SHLockShared(lParam, procId);
                    if (hShared == default) return;
                    string str = Marshal.PtrToStringAuto(hShared);
                    NativeMethods.SHUnlockShared(hShared);

                    OverlayIconDescription = str;
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ApplicationWindow: Unable to get overlay icon description from process {Title}: {e.Message}");
            }
        }

        internal void UpdateProperties()
        {
            setTitle();
            SetShowInTaskbar();
            setIcon();
        }

        public void BringToFront()
        {
            // call restore if window is minimized
            if (IsMinimized)
            {
                Restore();
            }
            else
            {
                NativeMethods.ShowWindow(Handle, NativeMethods.WindowShowStyle.Show);
                NativeMethods.SetForegroundWindow(Handle);

                if (State == WindowState.Flashing) State = WindowState.Active; // some stubborn windows (Outlook) start flashing while already active, this lets us stop
            }
        }

        public void Minimize()
        {
            if ((WindowStyles & (int)NativeMethods.WindowStyles.WS_MINIMIZEBOX) != 0)
            {
                IntPtr retval = IntPtr.Zero;
                NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_MINIMIZE, 0, 2, 200, ref retval);
            }
        }

        public void Restore()
        {
            IntPtr retval = IntPtr.Zero;
            NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_RESTORE, 0, 2, 200, ref retval);

            NativeMethods.SetForegroundWindow(Handle);
        }

        public void Maximize()
        {
            bool maximizeResult = NativeMethods.ShowWindow(Handle, NativeMethods.WindowShowStyle.Maximize);
            if (!maximizeResult)
            {
                // we don't have a fallback for elevated windows here since our only hope, SC_MAXIMIZE, doesn't seem to work for them. fall back to restore.
                IntPtr retval = IntPtr.Zero;
                NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_RESTORE, 0, 2, 200, ref retval);
            }
            NativeMethods.SetForegroundWindow(Handle);
        }

        internal IntPtr DoClose()
        {
            IntPtr retval = IntPtr.Zero;
            NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_CLOSE, 0, 2, 200, ref retval);

            return retval;
        }

        public void Close()
        {
            _tasksService.CloseWindow(this);
        }

        public void Move()
        {
            // move window via arrow keys; must be active window to control
            BringToFront();
            IntPtr retval = IntPtr.Zero;
            NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_MOVE, 0, 2, 200, ref retval);
        }

        public void Size()
        {
            // size window via arrow keys; must be active window to control
            BringToFront();
            IntPtr retval = IntPtr.Zero;
            NativeMethods.SendMessageTimeout(Handle, (int)NativeMethods.WM.SYSCOMMAND, NativeMethods.SC_SIZE, 0, 2, 200, ref retval);
        }

        /// <summary>
        /// Returns whether a window is normal (1), minimized (2), or maximized (3).
        /// </summary>
        /// <param name="hWnd">The handle of the window.</param>
        private NativeMethods.WindowShowStyle GetWindowShowStyle(IntPtr hWnd)
        {
            NativeMethods.WINDOWPLACEMENT placement = new NativeMethods.WINDOWPLACEMENT();
            NativeMethods.GetWindowPlacement(hWnd, ref placement);
            return placement.showCmd;
        }

        #region IEquatable<Window> Members

        public bool Equals(ApplicationWindow other)
        {
            return Handle.Equals(other.Handle);
        }

        #endregion

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged(string PropertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(PropertyName));
            }
        }

        #endregion

        public enum WindowState
        {
            Active,
            Inactive,
            Hidden,
            Flashing,
            Unknown = 999
        }
    }

}
