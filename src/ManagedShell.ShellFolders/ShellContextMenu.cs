using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.ShellFolders.Interfaces;
using ManagedShell.ShellFolders.Enums;
using ManagedShell.ShellFolders.Structs;
using NativeMethods = ManagedShell.Interop.NativeMethods;

namespace ManagedShell.ShellFolders
{
    public class ShellContextMenu : NativeWindow
    {
        // Properties
        public delegate void ItemSelectAction(string item, ShellItem[] items);
        public delegate void FolderItemSelectAction(uint uid, string path);
        private ItemSelectAction itemSelected;
        private FolderItemSelectAction folderItemSelected;

        private IContextMenu iContextMenu;
        private IContextMenu2 iContextMenu2;
        private IContextMenu3 iContextMenu3;

        private List<SubMenu> SubMenus = new List<SubMenu>();
        
        private readonly int x;
        private readonly int y;
        
        public ShellContextMenu(ShellItem[] files, IntPtr hwndOwner, ItemSelectAction itemSelected, bool isInteractive) : this(files, hwndOwner, itemSelected, isInteractive, new ShellCommandBuilder(), new ShellCommandBuilder())
        { }

        public ShellContextMenu(ShellItem[] files, IntPtr hwndOwner, ItemSelectAction itemSelected, bool isInteractive, ShellCommandBuilder preBuilder, ShellCommandBuilder postBuilder)
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

        public ShellContextMenu(ShellFolder folder, FolderItemSelectAction folderItemSelected, ShellCommandBuilder builder)
        {
            if (folder == null)
            {
                return;
            }

            lock (IconHelper.ComLock)
            {
                x = Cursor.Position.X;
                y = Cursor.Position.Y;

                this.folderItemSelected = folderItemSelected;
                
                SetupContextMenu(folder, builder);
            }
        }

        private uint ConfigureMenuItems(bool allFolders, IntPtr contextMenu, ShellCommandBuilder builder)
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

        private void ConfigureMenuItems(ShellFolder folder, IntPtr contextMenu, ShellCommandBuilder builder)
        {
            int numAdded = 0;
            SubMenus.Clear();

            foreach (var command in builder.Commands)
            {
                if (command is ShellNewCommand shellNewCommand)
                {
                    SubMenus.Add(shellNewCommand.AddSubMenu(folder, numAdded, ref contextMenu));
                }
                else
                {
                    Interop.AppendMenu(contextMenu, command.Flags, command.UID, command.Label);
                }

                numAdded++;
            }
        }

        private void SetupContextMenu(ShellItem[] files, IntPtr hwndOwner, bool isInteractive, ShellCommandBuilder preBuilder, ShellCommandBuilder postBuilder)
        {
            IntPtr contextMenu = IntPtr.Zero,
                iContextMenuPtr = IntPtr.Zero,
                iContextMenuPtr2 = IntPtr.Zero,
                iContextMenuPtr3 = IntPtr.Zero;

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

                    contextMenu = Interop.CreatePopupMenu();

                    uint numPrepended = ConfigureMenuItems(allFolders, contextMenu, preBuilder);

                    iContextMenu.QueryContextMenu(
                        contextMenu,
                        numPrepended,
                        Interop.CMD_FIRST,
                        Interop.CMD_LAST,
                        flags);

                    ConfigureMenuItems(allFolders, contextMenu, postBuilder);

                    if (isInteractive)
                    {
                        if (Marshal.QueryInterface(iContextMenuPtr, ref Interop.IID_IContextMenu2,
                            out iContextMenuPtr2) == NativeMethods.S_OK)
                        {
                            if (iContextMenuPtr2 != IntPtr.Zero)
                            {
                                try
                                {
                                    iContextMenu2 =
                                        (IContextMenu2)Marshal.GetTypedObjectForIUnknown(iContextMenuPtr2, typeof(IContextMenu2));
                                }
                                catch (Exception e)
                                {
                                    ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu2 interface: {e.Message}");
                                }
                            }
                        }

                        if (Marshal.QueryInterface(iContextMenuPtr, ref Interop.IID_IContextMenu3,
                            out iContextMenuPtr3) == NativeMethods.S_OK)
                        {
                            if (iContextMenuPtr3 != IntPtr.Zero)
                            {
                                try
                                {
                                    iContextMenu3 =
                                        (IContextMenu3)Marshal.GetTypedObjectForIUnknown(iContextMenuPtr3, typeof(IContextMenu3));
                                }
                                catch (Exception e)
                                {
                                    ShellLogger.Error($"ShellContextMenu: Error retrieving IContextMenu3 interface: {e.Message}");
                                }
                            }
                        }

                        ShowMenu(files, contextMenu, allFolders);
                    }
                    else
                    {
                        uint selected = Interop.GetMenuDefaultItem(contextMenu, 0, 0);

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

                if (contextMenu != IntPtr.Zero)
                    Interop.DestroyMenu(contextMenu);

                if (iContextMenuPtr != IntPtr.Zero)
                    Marshal.Release(iContextMenuPtr);

                if (iContextMenuPtr2 != IntPtr.Zero)
                    Marshal.Release(iContextMenuPtr2);

                if (iContextMenuPtr3 != IntPtr.Zero)
                    Marshal.Release(iContextMenuPtr3);

                foreach (var subMenu in SubMenus)
                {
                    subMenu.FreeResources();
                }
            }
        }

        private void SetupContextMenu(ShellFolder folder, ShellCommandBuilder builder)
        {
            IntPtr contextMenu = IntPtr.Zero;

            try
            {
                contextMenu = Interop.CreatePopupMenu();
                
                ConfigureMenuItems(folder, contextMenu, builder);
                
                ShowMenu(folder, contextMenu);
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellContextMenu: Error building folder context menu: {e.Message}");
            }
            finally
            {
                if (contextMenu != IntPtr.Zero)
                    Interop.DestroyMenu(contextMenu);

                foreach (var subMenu in SubMenus)
                {
                    subMenu.FreeResources();
                }
            }
        }

        private void ShowMenu(ShellItem[] files, IntPtr contextMenu, bool allFolders)
        {
            CreateHandle(new CreateParams());
            
            uint selected = Interop.TrackPopupMenuEx(
                contextMenu,
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

        private void ShowMenu(ShellFolder folder, IntPtr contextMenu)
        {
            CreateHandle(new CreateParams());
            
            uint selected = Interop.TrackPopupMenuEx(
                contextMenu,
                TPM.RETURNCMD,
                x,
                y,
                Handle,
                IntPtr.Zero);

            if (selected >= Interop.CMD_FIRST)
            {
                if (selected <= Interop.CMD_LAST)
                {
                    // custom commands are greater than CMD_LAST, so this must be a sub menu item
                    foreach (var subMenu in SubMenus)
                    {
                        if (subMenu.iContextMenu != null)
                        {
                            InvokeCommand(
                                subMenu.iContextMenu,
                                selected - Interop.CMD_FIRST,
                                new Point(x, y));
                        }
                    }
                }

                folderItemSelected?.Invoke(selected, folder.Path);
            }
            
            DestroyHandle();
        }

        #region Helpers
        private bool ItemsAllFolders(ShellItem[] items)
        {
            bool allFolders = true;
            foreach (var item in items)
            {
                if (!item.IsFolder || (item.Attributes & SFGAO.FILESYSTEM) == 0)
                {
                    // If the item is a folder, but not a filesystem object, don't treat it as a folder.
                    allFolders = false;
                    break;
                }
            }

            return allFolders;
        }
        
        private string GetCommandString(IContextMenu iContextMenu, uint idcmd, bool executeString)
        {
            string command = GetCommandStringW(iContextMenu, idcmd, executeString);

            if (string.IsNullOrEmpty(command))
                command = GetCommandStringA(iContextMenu, idcmd, executeString);

            return command;
        }

        /// <summary>
        /// Retrieves the command string for a specific item from an iContextMenu (Ansi)
        /// </summary>
        /// <param name="iContextMenu">the IContextMenu to receive the string from</param>
        /// <param name="idcmd">the id of the specific item</param>
        /// <param name="executeString">indicating whether it should return an execute string or not</param>
        /// <returns>if executeString is true it will return the executeString for the item, 
        /// otherwise it will return the help info string</returns>
        private string GetCommandStringA(IContextMenu iContextMenu, uint idcmd, bool executeString)
        {
            string info = string.Empty;
            byte[] bytes = new byte[256];
            int index;

            iContextMenu.GetCommandString(
                idcmd,
                (executeString ? GCS.VERBA : GCS.HELPTEXTA),
                0,
                bytes,
                256);

            index = 0;
            while (index < bytes.Length && bytes[index] != 0)
            { index++; }

            if (index < bytes.Length)
                info = Encoding.Default.GetString(bytes, 0, index);

            return info;
        }

        /// <summary>
        /// Retrieves the command string for a specific item from an iContextMenu (Unicode)
        /// </summary>
        /// <param name="iContextMenu">the IContextMenu to receive the string from</param>
        /// <param name="idcmd">the id of the specific item</param>
        /// <param name="executeString">indicating whether it should return an execute string or not</param>
        /// <returns>if executeString is true it will return the executeString for the item, 
        /// otherwise it will return the help info string</returns>
        private string GetCommandStringW(IContextMenu iContextMenu, uint idcmd, bool executeString)
        {
            string info = string.Empty;
            byte[] bytes = new byte[256];
            int index;

            iContextMenu.GetCommandString(
                idcmd,
                (executeString ? GCS.VERBW : GCS.HELPTEXTW),
                0,
                bytes,
                256);

            index = 0;
            while (index < bytes.Length - 1 && (bytes[index] != 0 || bytes[index + 1] != 0))
            { index += 2; }

            if (index < bytes.Length - 1)
                info = Encoding.Unicode.GetString(bytes, 0, index);

            return info;
        }

        /// <summary>
        /// Invokes a specific command from an IContextMenu
        /// </summary>
        /// <param name="iContextMenu">the IContextMenu containing the item</param>
        /// <param name="cmd">the index of the command to invoke</param>
        /// <param name="parentDir">the parent directory from where to invoke</param>
        /// <param name="ptInvoke">the point (in screen coordinates) from which to invoke</param>
        private void InvokeCommand(IContextMenu iContextMenu, uint cmd, Point ptInvoke)
        {
            CMINVOKECOMMANDINFOEX invoke = new CMINVOKECOMMANDINFOEX();
            invoke.cbSize = Interop.cbInvokeCommand;
            invoke.lpVerb = (IntPtr)cmd;
            invoke.lpVerbW = (IntPtr)cmd;
            invoke.fMask = CMIC.ASYNCOK | CMIC.UNICODE | CMIC.PTINVOKE |
                ((Control.ModifierKeys & Keys.Control) != 0 ? CMIC.CONTROL_DOWN : 0) |
                ((Control.ModifierKeys & Keys.Shift) != 0 ? CMIC.SHIFT_DOWN : 0);
            invoke.ptInvoke = new NativeMethods.POINT(ptInvoke.X, ptInvoke.Y);
            invoke.nShow = NativeMethods.WindowShowStyle.ShowNormal;

            iContextMenu.InvokeCommand(ref invoke);
        }

        private bool GetIContextMenu(ShellItem[] items, IntPtr hwndOwner, out IntPtr icontextMenuPtr, out IContextMenu iContextMenu)
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

        /// <summary>
        /// This method receives WindowMessages. It will make the "Open With" and "Send To" work 
        /// by calling HandleMenuMsg and HandleMenuMsg2.
        /// </summary>
        /// <param name="m">the Message of the Browser's WndProc</param>
        /// <returns>true if the message has been handled, false otherwise</returns>
        protected override void WndProc(ref Message m)
        {
            if (iContextMenu != null &&
                m.Msg == (int)NativeMethods.WM.MENUSELECT &&
                ((int)Interop.HiWord(m.WParam) & (int)MFT.SEPARATOR) == 0 &&
                ((int)Interop.HiWord(m.WParam) & (int)MFT.POPUP) == 0)
            {
                string info = GetCommandString(
                    iContextMenu,
                    (uint)Interop.LoWord(m.WParam) - Interop.CMD_FIRST,
                    false);
            }

            if (iContextMenu2 != null &&
                (m.Msg == (int)NativeMethods.WM.INITMENUPOPUP ||
                 m.Msg == (int)NativeMethods.WM.MEASUREITEM ||
                 m.Msg == (int)NativeMethods.WM.DRAWITEM))
            {
                if (iContextMenu2.HandleMenuMsg(
                    (uint)m.Msg, m.WParam, m.LParam) == NativeMethods.S_OK)
                    return;
            }

            if (iContextMenu3 != null &&
                m.Msg == (int)NativeMethods.WM.MENUCHAR)
            {
                if (iContextMenu3.HandleMenuMsg2(
                    (uint)m.Msg, m.WParam, m.LParam, IntPtr.Zero) == NativeMethods.S_OK)
                    return;
            }

            foreach (var subMenu in SubMenus)
            {
                if (subMenu.iContextMenu2 != null &&
                    ((m.Msg == (int)NativeMethods.WM.INITMENUPOPUP && m.WParam == subMenu.subMenuPtr) ||
                        m.Msg == (int)NativeMethods.WM.MEASUREITEM ||
                        m.Msg == (int)NativeMethods.WM.DRAWITEM))
                {
                    if (subMenu.iContextMenu2.HandleMenuMsg(
                        (uint)m.Msg, m.WParam, m.LParam) == NativeMethods.S_OK)
                        return;
                }

                if (subMenu.iContextMenu3 != null &&
                    m.WParam == subMenu.subMenuPtr &&
                    m.Msg == (int)NativeMethods.WM.MENUCHAR)
                {
                    if (subMenu.iContextMenu3.HandleMenuMsg2(
                        (uint)m.Msg, m.WParam, m.LParam, IntPtr.Zero) == NativeMethods.S_OK)
                        return;
                }
            }

            base.WndProc(ref m);
        }
    }
}
