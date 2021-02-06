using System;
using System.Runtime.InteropServices;
using ManagedShell.Common.Logging;
using ManagedShell.Interop;
using ManagedShell.ShellFolders.Enums;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public class ShellNewMenuCommand : ShellMenuCommand
    {
        public NativeContextMenu AddSubMenu(ShellFolder folder, int position, ref IntPtr parentNativeMenuPtr)
        {
            NativeContextMenu menu = new NativeContextMenu();
            
            if (GetNewContextMenu(folder, ref menu))
            {
                menu.iContextMenu.QueryContextMenu(
                    parentNativeMenuPtr,
                    (uint)position,
                    Interop.CMD_FIRST,
                    Interop.CMD_LAST,
                    CMF.NORMAL);

                menu.nativeMenuPtr = Interop.GetSubMenu(parentNativeMenuPtr, position);

                if (Marshal.QueryInterface(menu.iContextMenuPtr, ref Interop.IID_IContextMenu2,
                    out menu.iContextMenu2Ptr) == NativeMethods.S_OK)
                {
                    if (menu.iContextMenu2Ptr != IntPtr.Zero)
                    {
                        try
                        {
                            menu.iContextMenu2 =
                                (IContextMenu2)Marshal.GetTypedObjectForIUnknown(menu.iContextMenu2Ptr, typeof(IContextMenu2));
                        }
                        catch (Exception e)
                        {
                            ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu2 interface: {e.Message}");
                        }
                    }
                }

                if (Marshal.QueryInterface(menu.iContextMenuPtr, ref Interop.IID_IContextMenu3,
                    out menu.iContextMenu3Ptr) == NativeMethods.S_OK)
                {
                    if (menu.iContextMenu3Ptr != IntPtr.Zero)
                    {
                        try
                        {
                            menu.iContextMenu3 =
                                (IContextMenu3)Marshal.GetTypedObjectForIUnknown(menu.iContextMenu3Ptr, typeof(IContextMenu3));
                        }
                        catch (Exception e)
                        {
                            ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu3 interface: {e.Message}");
                        }
                    }
                }
            }

            return menu;
        }

        private bool GetNewContextMenu(ShellFolder folder, ref NativeContextMenu menu)
        {
            if (Interop.CoCreateInstance(
                ref Interop.CLSID_NewMenu,
                IntPtr.Zero,
                CLSCTX.INPROC_SERVER,
                ref Interop.IID_IContextMenu,
                out menu.iContextMenuPtr) == NativeMethods.S_OK)
            {
                menu.iContextMenu = Marshal.GetTypedObjectForIUnknown(menu.iContextMenuPtr, typeof(IContextMenu)) as IContextMenu;

                if (Marshal.QueryInterface(
                    menu.iContextMenuPtr,
                    ref Interop.IID_IShellExtInit,
                    out var iShellExtInitPtr) == NativeMethods.S_OK)
                {
                    IShellExtInit iShellExtInit = Marshal.GetTypedObjectForIUnknown(
                        iShellExtInitPtr, typeof(IShellExtInit)) as IShellExtInit;

                    iShellExtInit?.Initialize(folder.AbsolutePidl, IntPtr.Zero, 0);

                    if (iShellExtInit != null)
                    {
                        Marshal.ReleaseComObject(iShellExtInit);
                    }

                    if (iShellExtInitPtr != IntPtr.Zero)
                    {
                        Marshal.Release(iShellExtInitPtr);
                    }

                    return true;
                }
            }
            
            return false;
        }
    }
}
