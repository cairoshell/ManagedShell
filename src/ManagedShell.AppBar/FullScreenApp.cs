using System;
using System.Windows.Forms;
using ManagedShell.Interop;

namespace ManagedShell.AppBar
{
    public class FullScreenApp
    {
        public IntPtr hWnd;
        public string screenDeviceName;
        public NativeMethods.Rect rect;
    }
}