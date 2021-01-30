﻿using System;
using System.Runtime.InteropServices;
using System.Threading;
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
        
        public ShellFolder(string parsingName, IntPtr hwndInput, bool loadAsync) : base(parsingName)
        {
            _hwndInput = hwndInput;
            _loadAsync = loadAsync;
        }

        private void Initialize()
        {
            if (_loadAsync)
            {
                // Enumerate the directory on a new thread so that we don't block the UI during a potentially long operation
                // Because files is an ObservableCollection, we don't need to do anything special for the UI to update
                var thread = new Thread(() =>
                {
                    Enumerate(_hwndInput);
                });
                thread.IsBackground = true;
                thread.Start();
            }
            else
            {
                Enumerate(_hwndInput);
            }
        }

        private void Enumerate(IntPtr hwndInput)
        {
            IntPtr hEnum = IntPtr.Zero;

            if (_shellFolder == null)
            {
                _shellFolder = GetShellFolder(AbsolutePidl);
            }

            Files.Clear();

            if (_shellFolder?.EnumObjects(hwndInput, SHCONTF.FOLDERS | SHCONTF.NONFOLDERS,
                out hEnum) == NativeMethods.S_OK)
            {
                IEnumIDList enumIdList =
                    (IEnumIDList)Marshal.GetTypedObjectForIUnknown(hEnum, typeof(IEnumIDList));

                while (enumIdList.Next(1, out var pidlChild, out var numFetched) == NativeMethods.S_OK && numFetched == 1)
                {
                    Files.Add(new ShellFile(AbsolutePidl, _shellFolder, pidlChild));
                }

                Marshal.FinalReleaseComObject(enumIdList);
            }
            else
            {
                ShellLogger.Error($"ShellFolder: Unable to enumerate IShellFolder");
            }
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

        public new void Dispose()
        {
            base.Dispose();
            
            if (_shellFolder != null)
            {
                Marshal.FinalReleaseComObject(_shellFolder);
            }
        }
    }
}