using System;
using System.Runtime.InteropServices;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public class SubMenu
    {
        internal IContextMenu iContextMenu;
        internal IContextMenu2 iContextMenu2;
        internal IContextMenu3 iContextMenu3;
        
        internal IntPtr iContextMenuPtr;
        internal IntPtr iContextMenu2Ptr;
        internal IntPtr iContextMenu3Ptr;
        
        internal IntPtr subMenuPtr;
        
        internal void FreeResources()
        {
            if (iContextMenu != null)
            {
                Marshal.FinalReleaseComObject(iContextMenu);
                iContextMenu = null;
            }

            if (iContextMenu2 != null)
            {
                Marshal.FinalReleaseComObject(iContextMenu2);
                iContextMenu2 = null;
            }

            if (iContextMenu3 != null)
            {
                Marshal.FinalReleaseComObject(iContextMenu3);
                iContextMenu3 = null;
            }

            if (iContextMenuPtr != IntPtr.Zero)
                Marshal.Release(iContextMenuPtr);

            if (iContextMenu2Ptr != IntPtr.Zero)
                Marshal.Release(iContextMenu2Ptr);

            if (iContextMenu3Ptr != IntPtr.Zero)
                Marshal.Release(iContextMenu3Ptr);

            subMenuPtr = IntPtr.Zero;
        }
    }
}
