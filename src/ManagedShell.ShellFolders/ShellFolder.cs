using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ManagedShell.Common.Common;
using ManagedShell.Common.Logging;
using ManagedShell.Interop;
using ManagedShell.ShellFolders.Enums;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public class ShellFolder : ShellItem, IDisposable
    {
        private readonly IntPtr _hwndInput;
        private readonly bool _loadAsync;
        private readonly ChangeWatcher _changeWatcher;

        private bool _isDisposed;
        private IShellFolder _shellFolder;

        private ThreadSafeObservableCollection<ShellFile> _files;

        public ThreadSafeObservableCollection<ShellFile> Files
        {
            get
            {
                if (_files != null)
                {
                    return _files;
                }

                _files = new ThreadSafeObservableCollection<ShellFile>();
                Initialize();

                return _files;
            }
        }
        
        public ShellFolder(string parsingName, IntPtr hwndInput, bool loadAsync = false) : base(parsingName)
        {
            _hwndInput = hwndInput;
            _loadAsync = loadAsync;

            if (_shellItem == null && parsingName.StartsWith("{"))
            {
                parsingName = "::" + parsingName;
                _shellItem = GetShellItem(parsingName);
            }

            if (_shellItem == null && !parsingName.ToLower().StartsWith("shell:"))
            {
                parsingName = "shell:" + parsingName;
                _shellItem = GetShellItem(parsingName);
            }

            if (_shellItem != null && IsFileSystem)
            {
                _changeWatcher = new ChangeWatcher(Path, ChangedEventHandler, CreatedEventHandler, DeletedEventHandler, RenamedEventHandler);
            }
        }

        private void Initialize()
        {
            _changeWatcher?.StartWatching();
            
            if (_loadAsync)
            {
                // Enumerate the directory on a new thread so that we don't block the UI during a potentially long operation
                // Because files is an ObservableCollection, we don't need to do anything special for the UI to update
                Task.Factory.StartNew(Enumerate, CancellationToken.None, TaskCreationOptions.None, Interop.ShellItemScheduler);
            }
            else
            {
                Enumerate();
            }
        }

        private void Enumerate()
        {
            IntPtr hEnum = IntPtr.Zero;

            if (_shellFolder == null && AbsolutePidl != IntPtr.Zero)
            {
                _shellFolder = GetShellFolder(AbsolutePidl);
            }

            if (_shellFolder == null)
            {
                return;
            }

            Files.Clear();

            try
            {
                if (_shellFolder?.EnumObjects(_hwndInput, SHCONTF.FOLDERS | SHCONTF.NONFOLDERS,
                    out hEnum) == NativeMethods.S_OK)
                {
                    IEnumIDList enumIdList =
                        (IEnumIDList)Marshal.GetTypedObjectForIUnknown(hEnum, typeof(IEnumIDList));

                    while (enumIdList.Next(1, out var pidlChild, out var numFetched) == NativeMethods.S_OK && numFetched == 1)
                    {
                        if (_isDisposed)
                        {
                            break;
                        }

                        AddFile(pidlChild);
                    }

                    Marshal.FinalReleaseComObject(enumIdList);
                }
                else
                {
                    ShellLogger.Error($"ShellFolder: Unable to enumerate IShellFolder");
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellFolder: Unable to enumerate IShellFolder: {e.Message}");
            }
        }

        private void ChangedEventHandler(object sender, FileSystemEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                ShellLogger.Info($"{e.ChangeType}: {e.Name} ({e.FullPath})");

                bool exists = false;

                foreach (var file in Files)
                {
                    if (_isDisposed)
                    {
                        break;
                    }

                    if (file.Path == e.FullPath)
                    {
                        exists = true;
                        file.Refresh();

                        break;
                    }
                }

                if (!exists)
                {
                    AddFile(e.FullPath);
                }
            }, CancellationToken.None, TaskCreationOptions.None, Interop.ShellItemScheduler);
        }

        private void CreatedEventHandler(object sender, FileSystemEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                ShellLogger.Info($"{e.ChangeType}: {e.Name} ({e.FullPath})");

                if (!FileExists(e.FullPath))
                {
                    AddFile(e.FullPath);
                }
            }, CancellationToken.None, TaskCreationOptions.None, Interop.ShellItemScheduler);
        }

        private void DeletedEventHandler(object sender, FileSystemEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                ShellLogger.Info($"{e.ChangeType}: {e.Name} ({e.FullPath})");

                RemoveFile(e.FullPath);
            }, CancellationToken.None, TaskCreationOptions.None, Interop.ShellItemScheduler);
        }

        private void RenamedEventHandler(object sender, RenamedEventArgs e)
        {
            Task.Factory.StartNew(() =>
            {
                ShellLogger.Info($"{e.ChangeType}: From {e.OldName} ({e.OldFullPath}) to {e.Name} ({e.FullPath})");

                int existing = RemoveFile(e.OldFullPath);

                if (!FileExists(e.FullPath))
                {
                    AddFile(e.FullPath, existing);
                }
            }, CancellationToken.None, TaskCreationOptions.None, Interop.ShellItemScheduler);
        }

        private IShellFolder GetShellFolder()
        {
            Interop.SHGetDesktopFolder(out IntPtr desktopFolderPtr);
            return (IShellFolder)Marshal.GetTypedObjectForIUnknown(desktopFolderPtr, typeof(IShellFolder));
        }

        private IShellFolder GetShellFolder(IntPtr folderPidl)
        {
            IShellFolder desktop = GetShellFolder();
            Guid guid = typeof(IShellFolder).GUID;

            if (desktop.BindToObject(folderPidl, IntPtr.Zero, ref guid, out IntPtr folderPtr) == NativeMethods.S_OK)
            {
                Marshal.FinalReleaseComObject(desktop);
                return (IShellFolder)Marshal.GetTypedObjectForIUnknown(folderPtr, typeof(IShellFolder));
            }

            ShellLogger.Error($"ShellFolder: Unable to bind IShellFolder for {folderPidl}");
            return null;
        }

        #region Helpers
        private bool AddFile(string parsingName, int position = -1)
        {
            ShellFile file = new ShellFile(this, parsingName);

            return AddFile(file, position);
        }

        private bool AddFile(IntPtr relPidl, int position = -1)
        {
            ShellFile file = new ShellFile(this, _shellFolder, relPidl, _loadAsync);

            return AddFile(file, position);
        }

        private bool AddFile(ShellFile file, int position = -1)
        {
            if (file.Loaded)
            {
                if (position >= 0)
                {
                    Files.Insert(position, file);
                }
                else
                {
                    Files.Add(file);
                }

                return true;
            }

            return false;
        }

        private int RemoveFile(string parsingName)
        {
            for (int i = 0; i < Files.Count; i++)
            {
                if (Files[i].Path == parsingName)
                {
                    ShellFile file = Files[i];
                    Files.RemoveAt(i);
                    file.Dispose();

                    return i;
                }

                if (_isDisposed)
                {
                    break;
                }
            }

            return -1;
        }

        private bool FileExists(string parsingName)
        {
            bool exists = false;

            foreach (var file in Files)
            {
                if (file.Path == parsingName)
                {
                    exists = true;
                    break;
                }

                if (_isDisposed)
                {
                    break;
                }
            }

            return exists;
        }
        #endregion

        public new void Dispose()
        {
            _isDisposed = true;
            _changeWatcher?.Dispose();
            
            try
            {
                if (_files != null)
                {
                    foreach (var file in Files)
                    {
                        file.Dispose();
                    }

                    Files.Clear();
                }
            }
            catch (Exception e)
            {
                ShellLogger.Warning($"ShellFolder: Unable to dispose files: {e.Message}");
            }

            if (_shellFolder != null)
            {
                Marshal.ReleaseComObject(_shellFolder);
                _shellFolder = null;
            }

            if (_parentAbsolutePidl != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_parentAbsolutePidl);
            }
            
            base.Dispose();
        }
    }
}
