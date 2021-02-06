using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.ShellFolders.Interfaces;
using ManagedShell.ShellFolders.Enums;
using NativeMethods = ManagedShell.Interop.NativeMethods;

namespace ManagedShell.ShellFolders
{
    public class ShellItemContextMenu : ShellContextMenu
    {
        public delegate void ItemSelectAction(string item, ShellItem[] items);
        private readonly ItemSelectAction itemSelected;

        public ShellItemContextMenu(ShellItem[] files, IntPtr hwndOwner, ItemSelectAction itemSelected, bool isInteractive) : this(files, hwndOwner, itemSelected, isInteractive, new ShellMenuCommandBuilder(), new ShellMenuCommandBuilder())
        { }

        public ShellItemContextMenu(ShellItem[] files, IntPtr hwndOwner, ItemSelectAction itemSelected, bool isInteractive, ShellMenuCommandBuilder preBuilder, ShellMenuCommandBuilder postBuilder)
        {
            if (files == null || files.Length < 1)
            {
                return;
            }

            lock (IconHelper.ComLock)
            {
                x = Cursor.Position.X;
                y = Cursor.Position.Y;

                this.itemSelected = itemSelected;

                SetupContextMenu(files, hwndOwner, isInteractive, preBuilder, postBuilder);
            }
        }

        private uint ConfigureMenuItems(bool allFolders, IntPtr contextMenu, ShellMenuCommandBuilder builder)
        {
            uint numAdded = 0;

            foreach (var command in builder.Commands)
            {
                if (allFolders || !command.FoldersOnly)
                {
                    Interop.AppendMenu(contextMenu, command.Flags, command.UID, command.Label);
                }

                numAdded++;
            }

            return numAdded;
        }

        private void SetupContextMenu(ShellItem[] files, IntPtr hwndOwner, bool isInteractive, ShellMenuCommandBuilder preBuilder, ShellMenuCommandBuilder postBuilder)
        {
            try
            {
                if (GetIContextMenu(files, hwndOwner, out iContextMenuPtr, out iContextMenu))
                {
                    // get some properties about our file(s)
                    bool allFolders = ItemsAllFolders(files);

                    CMF flags = CMF.EXPLORE |
                        CMF.ITEMMENU |
                        CMF.CANRENAME |
                        ((Control.ModifierKeys & Keys.Shift) != 0 ? CMF.EXTENDEDVERBS : 0);

                    if (!isInteractive)
                    {
                        flags |= CMF.DEFAULTONLY;
                    }

                    nativeMenuPtr = Interop.CreatePopupMenu();

                    uint numPrepended = ConfigureMenuItems(allFolders, nativeMenuPtr, preBuilder);

                    iContextMenu.QueryContextMenu(
                        nativeMenuPtr,
                        numPrepended,
                        Interop.CMD_FIRST,
                        Interop.CMD_LAST,
                        flags);

                    ConfigureMenuItems(allFolders, nativeMenuPtr, postBuilder);

                    if (isInteractive)
                    {
                        if (Marshal.QueryInterface(iContextMenuPtr, ref Interop.IID_IContextMenu2,
                            out iContextMenu2Ptr) == NativeMethods.S_OK)
                        {
                            if (iContextMenu2Ptr != IntPtr.Zero)
                            {
                                try
                                {
                                    iContextMenu2 =
                                        (IContextMenu2)Marshal.GetTypedObjectForIUnknown(iContextMenu2Ptr, typeof(IContextMenu2));
                                }
                                catch (Exception e)
                                {
                                    ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu2 interface: {e.Message}");
                                }
                            }
                        }

                        if (Marshal.QueryInterface(iContextMenuPtr, ref Interop.IID_IContextMenu3,
                            out iContextMenu3Ptr) == NativeMethods.S_OK)
                        {
                            if (iContextMenu3Ptr != IntPtr.Zero)
                            {
                                try
                                {
                                    iContextMenu3 =
                                        (IContextMenu3)Marshal.GetTypedObjectForIUnknown(iContextMenu3Ptr, typeof(IContextMenu3));
                                }
                                catch (Exception e)
                                {
                                    ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu3 interface: {e.Message}");
                                }
                            }
                        }

                        ShowMenu(files, allFolders);
                    }
                    else
                    {
                        uint selected = Interop.GetMenuDefaultItem(nativeMenuPtr, 0, 0);

                        HandleMenuCommand(files, selected, allFolders);
                    }
                }
                else
                {
                    ShellLogger.Debug("Error retrieving IContextMenu");
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellContextMenu: Error building context menu: {e.Message}");
            }
            finally
            {
                FreeResources();
                
                foreach (var subMenu in ShellNewMenus)
                {
                    subMenu.FreeResources();
                }
            }
        }

        private void ShowMenu(ShellItem[] files, bool allFolders)
        {
            CreateHandle(new CreateParams());

            uint selected = Interop.TrackPopupMenuEx(
                nativeMenuPtr,
                TPM.RETURNCMD,
                x,
                y,
                Handle,
                IntPtr.Zero);

            HandleMenuCommand(files, selected, allFolders);

            DestroyHandle();
        }

        private void HandleMenuCommand(ShellItem[] files, uint selected, bool allFolders)
        {
            if (selected >= Interop.CMD_FIRST && selected < uint.MaxValue)
            {
                string command = GetCommandString(iContextMenu, selected - Interop.CMD_FIRST, true);

                if (command == "open" && allFolders)
                {
                    // suppress running system code
                    command = "openFolder";
                }
                else
                {
                    InvokeCommand(
                        iContextMenu,
                        selected - Interop.CMD_FIRST,
                        new Point(x, y));
                }

                if (string.IsNullOrEmpty(command))
                {
                    // custom commands only have an ID, so pass that instead
                    command = selected.ToString();
                }

                itemSelected?.Invoke(command, files);
            }
        }

        #region Helpers
        private bool ItemsAllFolders(ShellItem[] items)
        {
            bool allFolders = true;
            foreach (var item in items)
            {
                if (!item.IsFolder || !item.IsFileSystem)
                {
                    // If the item is a folder, but not a filesystem object, don't treat it as a folder.
                    allFolders = false;
                    break;
                }
            }

            return allFolders;
        }
        
        protected bool GetIContextMenu(ShellItem[] items, IntPtr hwndOwner, out IntPtr icontextMenuPtr, out IContextMenu iContextMenu)
        {
            if (items.Length < 1)
            {
                icontextMenuPtr = IntPtr.Zero;
                iContextMenu = null;

                return false;
            }

            IShellFolder parent = items[0].ParentFolder;
            IntPtr[] pidls = new IntPtr[items.Length];

            for (int i = 0; i < items.Length; i++)
            {
                pidls[i] = items[i].RelativePidl;
            }

            if (parent.GetUIObjectOf(
                hwndOwner,
                (uint)pidls.Length,
                pidls,
                ref Interop.IID_IContextMenu,
                IntPtr.Zero,
                out icontextMenuPtr) == NativeMethods.S_OK)
            {
                iContextMenu =
                    (IContextMenu)Marshal.GetTypedObjectForIUnknown(
                        icontextMenuPtr, typeof(IContextMenu));

                return true;
            }

            icontextMenuPtr = IntPtr.Zero;
            iContextMenu = null;

            return false;
        }
        #endregion
    }
}
