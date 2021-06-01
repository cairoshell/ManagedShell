using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using ManagedShell.Interop;

namespace ManagedShell.WindowsTasks
{
    public class TaskThumbButton : INotifyPropertyChanged
    {
        private uint _id;
        public uint Id
        {
            get => _id;
            internal set
            {
                _id = value;
                OnPropertyChanged();
            }
        }

        private ImageSource _icon;
        public ImageSource Icon
        {
            get => _icon;
            internal set
            {
                _icon = value;
                OnPropertyChanged();
            }
        }

        private string _title;
        public string Title
        {
            get => _title;
            internal set
            {
                _title = value;
                OnPropertyChanged();
            }
        }

        private NativeMethods.THUMBBUTTONFLAGS _flags;
        public NativeMethods.THUMBBUTTONFLAGS Flags
        {
            get => _flags;
            internal set
            {
                _flags = value;
                OnPropertyChanged();
            }
        }

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }
}
