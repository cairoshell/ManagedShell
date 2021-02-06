using System;
using System.IO;
using ManagedShell.Common.Logging;

namespace ManagedShell.ShellFolders
{
    class ChangeWatcher : IDisposable
    {
        private readonly FileSystemEventHandler _changedEventHandler;
        private readonly FileSystemEventHandler _createdEventHandler;
        private readonly FileSystemEventHandler _deletedEventHandler;
        private readonly RenamedEventHandler _renamedEventHandler;
        private FileSystemWatcher _watcher;

        public ChangeWatcher(string path, FileSystemEventHandler changedEventHandler, FileSystemEventHandler createdEventHandler, FileSystemEventHandler deletedEventHandler, RenamedEventHandler renamedEventHandler)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            _changedEventHandler = changedEventHandler;
            _createdEventHandler = createdEventHandler;
            _deletedEventHandler = deletedEventHandler;
            _renamedEventHandler = renamedEventHandler;

            try
            {
                _watcher = new FileSystemWatcher(path)
                {
                    NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size
                };

                _watcher.Changed += _changedEventHandler;
                _watcher.Created += _createdEventHandler;
                _watcher.Deleted += _deletedEventHandler;
                _watcher.Renamed += _renamedEventHandler;
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ChangeWatcher: Unable to instantiate watcher: {e.Message}");
            }
        }

        public void StartWatching()
        {
            if (_watcher == null)
            {
                ShellLogger.Error("ChangeWatcher: Unable to start watching directory.");
                return;
            }

            try
            {
                _watcher.EnableRaisingEvents = true;
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ChangeWatcher: Unable to start watching directory: {e.Message}");
            }
        }

        public void Dispose()
        {
            if (_watcher == null)
            {
                return;
            }

            _watcher.Changed -= _changedEventHandler;
            _watcher.Created -= _createdEventHandler;
            _watcher.Deleted -= _deletedEventHandler;
            _watcher.Renamed -= _renamedEventHandler;

            _watcher.Dispose();
            _watcher = null;
        }
    }
}
