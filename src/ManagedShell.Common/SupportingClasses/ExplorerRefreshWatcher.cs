using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ManagedShell.Common.Logging;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.Common.SupportingClasses
{
    /// <summary>
    /// Once another process registers itself as the OS shell window (via SetShellWindow, as
    /// ShellWindow does), Explorer's own folder view windows stop receiving the shell-level
    /// change notifications they rely on for auto-refresh, since that delivery path is
    /// restricted to windows belonging to whichever process the OS currently trusts as "the
    /// shell". This watcher registers for the same notifications at interrupt level, which is
    /// not subject to that restriction, and manually refreshes any open Explorer windows
    /// viewing an affected folder.
    /// </summary>
    public class ExplorerRefreshWatcher : IDisposable
    {
        private const SHCNE WatchedEvents = SHCNE.CREATE | SHCNE.DELETE | SHCNE.MKDIR | SHCNE.RMDIR |
                                             SHCNE.RENAMEITEM | SHCNE.RENAMEFOLDER | SHCNE.UPDATEDIR |
                                             SHCNE.UPDATEITEM | SHCNE.ATTRIBUTES | SHCNE.MEDIAINSERTED |
                                             SHCNE.MEDIAREMOVED | SHCNE.DRIVEADD | SHCNE.DRIVEREMOVED;

        // Events whose item pidl(s) refer to the directory whose own listing needs refreshing.
        // Everything else (create/delete/rename an item, make/remove a subfolder, rename a
        // subfolder) targets an item *inside* a directory, so it's that item's parent whose
        // window needs refreshing instead.
        private const SHCNE DirectoryTargetEvents = SHCNE.UPDATEDIR | SHCNE.MEDIAINSERTED |
                                                      SHCNE.MEDIAREMOVED | SHCNE.DRIVEADD | SHCNE.DRIVEREMOVED;

        private readonly NativeWindowEx _window;
        private readonly int _notifyMessage;
        private IntPtr _registration;
        private dynamic _shellApp;

        public ExplorerRefreshWatcher(NativeWindowEx window)
        {
            _window = window;
            _notifyMessage = RegisterWindowMessage("ManagedShell_ExplorerRefreshWatcher");
            _window.MessageReceived += WndProc;

            SHChangeNotifyEntry entry = new SHChangeNotifyEntry
            {
                pIdl = IntPtr.Zero,
                Recursively = true
            };

            _registration = SHChangeNotifyRegister(_window.Handle, SHCNRF.InterruptLevel | SHCNRF.NewDelivery,
                WatchedEvents, (uint)_notifyMessage, 1, ref entry);

            if (_registration == IntPtr.Zero)
            {
                ShellLogger.Warning("ExplorerRefreshWatcher: Failed to register for shell change notifications");
            }
        }

        private void WndProc(ref Message msg, ref bool handled)
        {
            if (msg.Msg != _notifyMessage)
            {
                return;
            }

            IntPtr lockHandle = SHChangeNotification_Lock(msg.WParam, 0, out IntPtr pidlArray, out uint eventId);

            if (lockHandle == IntPtr.Zero)
            {
                return;
            }

            try
            {
                bool isDirectoryTarget = (WatchedEvents & (SHCNE)eventId & DirectoryTargetEvents) != 0;

                string path1 = GetPathFromPidl(Marshal.ReadIntPtr(pidlArray));
                string path2 = GetPathFromPidl(Marshal.ReadIntPtr(pidlArray, IntPtr.Size));

                string target1 = ResolveRefreshTarget(path1, isDirectoryTarget);
                string target2 = ResolveRefreshTarget(path2, isDirectoryTarget);

                if (!string.IsNullOrEmpty(target1) || !string.IsNullOrEmpty(target2))
                {
                    RefreshExplorerWindows(target1, target2);
                }
            }
            catch (Exception ex)
            {
                ShellLogger.Warning("ExplorerRefreshWatcher: Error handling shell change notification", ex);
            }
            finally
            {
                SHChangeNotification_Unlock(lockHandle);
            }
        }

        private static string GetPathFromPidl(IntPtr pidl)
        {
            if (pidl == IntPtr.Zero)
            {
                return null;
            }

            StringBuilder path = new StringBuilder(260);
            return SHGetPathFromIDList(pidl, path) ? path.ToString() : null;
        }

        // For events where the pidl refers to an item inside a directory (create/delete/rename
        // a file, make/remove/rename a subfolder), the Explorer window that needs refreshing is
        // the one browsing that item's parent, not the item's own path.
        private static string ResolveRefreshTarget(string path, bool isDirectoryTarget)
        {
            if (string.IsNullOrEmpty(path))
            {
                return null;
            }

            if (isDirectoryTarget)
            {
                return path;
            }

            try
            {
                return System.IO.Path.GetDirectoryName(path);
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        private void RefreshExplorerWindows(string target1, string target2)
        {
            try
            {
                _shellApp ??= Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                dynamic windows = _shellApp.Windows();

                foreach (dynamic window in windows)
                {
                    try
                    {
                        string windowPath = window?.Document?.Folder?.Self?.Path as string;

                        if (!string.IsNullOrEmpty(windowPath) && (PathsMatch(windowPath, target1) || PathsMatch(windowPath, target2)))
                        {
                            window.Refresh();
                        }
                    }
                    catch (COMException)
                    {
                        // window doesn't expose a Document/Folder (e.g. an IE window), skip it
                    }
                    finally
                    {
                        if (window != null)
                        {
                            Marshal.ReleaseComObject(window);
                        }
                    }
                }

                Marshal.ReleaseComObject(windows);
            }
            catch (Exception ex)
            {
                ShellLogger.Warning("ExplorerRefreshWatcher: Unable to refresh Explorer windows", ex);
            }
        }

        private static bool PathsMatch(string windowPath, string target)
        {
            return !string.IsNullOrEmpty(target) &&
                   string.Equals(windowPath.TrimEnd('\\'), target.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase);
        }

        public void Dispose()
        {
            _window.MessageReceived -= WndProc;

            if (_registration != IntPtr.Zero)
            {
                SHChangeNotifyDeregister(_registration);
                _registration = IntPtr.Zero;
            }

            if (_shellApp != null)
            {
                Marshal.ReleaseComObject(_shellApp);
                _shellApp = null;
            }
        }
    }
}
