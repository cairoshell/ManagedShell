using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using ManagedShell.Common.Enums;
using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.Interop;
using ManagedShell.ShellFolders.Enums;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public class ShellItem : INotifyPropertyChanged, IDisposable
    {
        private IShellItemImageFactory _imageFactory;
        protected readonly IShellItem _shellItem;

        private bool _smallIconLoading;
        private bool _largeIconLoading;
        private bool _extraLargeIconLoading;

        #region Properties
        private bool? _isFolder;

        public bool IsFolder
        {
            get
            {
                if (_isFolder == null)
                {
                    _isFolder = ((Attributes & SFGAO.FOLDER) != 0);
                }
                
                return (bool)_isFolder;
            }
        }

        private ShellItem _parentItem;

        public ShellItem ParentItem
        {
            get
            {
                if (_parentItem == null)
                {
                    _parentItem = new ShellItem(GetParentShellItem());
                }

                return _parentItem;
            }
        }

        private IntPtr _absolutePidl;
        
        public IntPtr AbsolutePidl
        {
            get
            {
                if (_absolutePidl == IntPtr.Zero)
                {
                    _absolutePidl = GetAbsolutePidl();
                }

                return _absolutePidl;
            }
            protected set
            {
                _absolutePidl = value;
            }
        }

        private string _path;
        
        public string Path
        {
            get
            {
                if (_path == null)
                {
                    _path = GetDisplayName(SIGDN.DESKTOPABSOLUTEPARSING);
                }

                return _path;
            }
        }

        private string _fileName;
        
        public string FileName
        {
            get
            {
                if (_fileName == null)
                {
                    _fileName = GetDisplayName(SIGDN.PARENTRELATIVEPARSING);
                }

                return _fileName;
            }
        }

        private string _displayName;
        
        public string DisplayName
        {
            get
            {
                if (_displayName == null)
                {
                    _displayName = GetDisplayName(SIGDN.NORMALDISPLAY);
                }

                return _displayName;
            }
        }

        private SFGAO _attributes = 0;
        
        public SFGAO Attributes
        {
            get
            {
                if (_attributes == 0)
                {
                    _attributes = GetAttributes();
                }

                return _attributes;
            }
        }

        private ImageSource _smallIcon;

        public ImageSource SmallIcon
        {
            get
            {
                if (_smallIcon == null && !_smallIconLoading)
                {
                    _smallIconLoading = true;

                    Task.Factory.StartNew(() =>
                    {
                        SmallIcon = GetDisplayIcon(IconSize.Small);
                        SmallIcon?.Freeze();
                        _smallIconLoading = false;
                    }, CancellationToken.None, TaskCreationOptions.None, IconHelper.IconScheduler);
                }

                return _smallIcon;
            }
            private set
            {
                _smallIcon = value;
                OnPropertyChanged();
            }
        }

        private ImageSource _largeIcon;

        public ImageSource LargeIcon
        {
            get
            {
                if (_largeIcon == null && !_largeIconLoading)
                {
                    _largeIconLoading = true;

                    Task.Factory.StartNew(() =>
                    {
                        LargeIcon = GetDisplayIcon(IconSize.Large);
                        LargeIcon?.Freeze();
                        _largeIconLoading = false;
                    }, CancellationToken.None, TaskCreationOptions.None, IconHelper.IconScheduler);
                }

                return _largeIcon;
            }
            private set
            {
                _largeIcon = value;
                OnPropertyChanged();
            }
        }

        private ImageSource _extraLargeIcon;

        public ImageSource ExtraLargeIcon
        {
            get
            {
                if (_extraLargeIcon == null && !_extraLargeIconLoading)
                {
                    _extraLargeIconLoading = true;

                    Task.Factory.StartNew(() =>
                    {
                        ExtraLargeIcon = GetDisplayIcon(IconSize.ExtraLarge);
                        ExtraLargeIcon?.Freeze();
                        _extraLargeIconLoading = false;
                    }, CancellationToken.None, TaskCreationOptions.None, IconHelper.IconScheduler);
                }

                return _extraLargeIcon;
            }
            private set
            {
                _extraLargeIcon = value;
                OnPropertyChanged();
            }
        }
        #endregion

        public ShellItem(IShellItem shellItem)
        {
            _shellItem = shellItem;
        }

        public ShellItem(string parsingName)
        {
            _shellItem = GetShellItem(parsingName);
        }

        public ShellItem(IntPtr parentPidl, IShellFolder parentShellFolder, IntPtr relativePidl)
        {
            _shellItem = GetShellItem(parentPidl, parentShellFolder, relativePidl);
        }

        #region Retrieve interfaces
        private IShellItem GetParentShellItem()
        {
            IShellItem parent = null;

            if (_shellItem?.GetParent(out parent) != NativeMethods.S_OK)
            {
                parent = null;
            }

            return parent;
        }

        private IShellItem GetShellItem(string parsingName)
        {
            try
            {
                Interop.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, typeof(IShellItem).GUID, out IShellItem ppv);
                return ppv;
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellItem: Unable to get shell item for {parsingName}: {e.Message}");
                return null;
            }
        }

        private IShellItem GetShellItem(IntPtr parentPidl, IShellFolder parentShellFolder, IntPtr relativePidl)
        {
            try
            {
                Interop.SHCreateItemWithParent(parentPidl, parentShellFolder, relativePidl, typeof(IShellItem).GUID, out IShellItem ppv);
                return ppv;
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellItem: Unable to get shell item for {relativePidl}: {e.Message}");
                return null;
            }
        }

        private IShellItemImageFactory GetImageFactory(IntPtr absolutePidl)
        {
            try
            {
                Interop.SHCreateItemFromIDList(absolutePidl, typeof(IShellItemImageFactory).GUID, out IShellItemImageFactory ppv);
                return ppv;
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellItem: Unable to get shell item image factory for {absolutePidl}: {e.Message}");
                return null;
            }
        }
        #endregion

        #region Retrieve properties
        private IntPtr GetAbsolutePidl()
        {
            IntPtr pidl = IntPtr.Zero;

            Interop.SHGetIDListFromObject(_shellItem, out pidl);

            return pidl;
        }

        private SFGAO GetAttributes()
        {
            SFGAO attrs = 0;

            if (_shellItem?.GetAttributes(SFGAO.FILESYSTEM | SFGAO.FOLDER | SFGAO.HIDDEN, out attrs) != NativeMethods.S_OK)
            {
                attrs = 0;
            }

            return attrs;
        }

        private string GetDisplayName(SIGDN purpose)
        {
            IntPtr hString = IntPtr.Zero;
            string name = string.Empty;

            try
            {
                if (_shellItem?.GetDisplayName(purpose, out hString) == NativeMethods.S_OK)
                {
                    if (hString != IntPtr.Zero)
                    {
                        name = Marshal.PtrToStringAuto(hString);
                    }
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellItem: Unable to get {purpose} display name: {e.Message}");
            }
            finally
            {
                Marshal.FreeCoTaskMem(hString);
            }

            return name;
        }

        private ImageSource GetDisplayIcon(IconSize size)
        {
            if (_imageFactory == null)
            {
                _imageFactory = GetImageFactory(AbsolutePidl);
            }

            int iconPoints = IconHelper.GetSize(size);
            SIZE imageSize = new SIZE {cx = iconPoints, cy = iconPoints};
            
            IntPtr hBitmap = IntPtr.Zero;
            try
            {
                SIIGBF flags = 0;

                if (size == IconSize.Small)
                {
                    // for 16pt icons, thumbnails are too small
                    flags = SIIGBF.ICONONLY;
                }
                
                if (_imageFactory?.GetImage(imageSize, flags, out hBitmap) == NativeMethods.S_OK)
                {
                    if (hBitmap != IntPtr.Zero)
                    {
                        return IconImageConverter.GetImageFromHBitmap(hBitmap);
                    }
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellItem: Unable to get icon: {e.Message}");
            }

            return null;
        }
        #endregion

        #region IDisposable
        public void Dispose()
        {
            if (_shellItem != null)
            {
                Marshal.FinalReleaseComObject(_shellItem);
            }

            if (_imageFactory != null)
            {
                Marshal.FinalReleaseComObject(_imageFactory);
            }

            if (_absolutePidl != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_absolutePidl);
            }
        }
        #endregion

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
