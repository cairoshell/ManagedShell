using System;
using System.IO;
using ManagedShell.Common.Logging;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public class ShellFile : ShellItem
    {
        private long _fileSize;
        public long FileSize
        {
            get
            {
                if (_fileSize == 0)
                {
                    _fileSize = GetFileSize();
                }

                return _fileSize;
            }
        }
        
        public ShellFile(string parsingName) : base(parsingName)
        {
            ShellLogger.Info($"Found {DisplayName} : {FileName} : {Path} | {Attributes}");
        }

        public ShellFile(IntPtr absolutePidl, IShellFolder parentShellFolder, IntPtr relativePidl) : base(absolutePidl, parentShellFolder, relativePidl)
        {
            ShellLogger.Info($"Found {DisplayName} : {FileName} : {Path} | {Attributes}");
        }

        private long GetFileSize()
        {
            // TODO: Replace this using properties via IShellItem2
            using (var file = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return file.Length;
            }
        }
    }
}
