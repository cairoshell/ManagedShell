using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ManagedShell.Common.Logging;
using ManagedShell.ShellFolders.Enums;
using ManagedShell.ShellFolders.Interfaces;

namespace ManagedShell.ShellFolders
{
    public static class ShellLinkHelper
    {
        public static void Save(IShellLink existingLink)
        {
            ((IPersistFile)existingLink).Save(null, true);
        }

        public static void Save(IShellLink link, string destinationPath)
        {
            ((IPersistFile)link).Save(destinationPath, true);
        }

        public static IShellLink Create()
        {
            ShellLink link = new ShellLink();
            IShellLink shellLink = (IShellLink)link;

            return shellLink;
        }

        public static void CreateAndSave(string linkTargetPath, string destinationPath)
        {
            ShellLink link = new ShellLink();
            IShellLink shellLink = (IShellLink)link;

            try
            {
                shellLink.SetPath(linkTargetPath);

                Save(shellLink, destinationPath);
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellLinkHelper: Unable to create link from {linkTargetPath} to {destinationPath}", e);
            }

            Marshal.FinalReleaseComObject(link);
        }

        public static IShellLink Load(IntPtr userInputHwnd, string existingLinkPath)
        {
            ShellLink link = new ShellLink();
            IShellLink shellLink = (IShellLink)link;
            IPersistFile persistFile = (IPersistFile)link;

            try
            {
                // load from disk
                persistFile.Load(existingLinkPath, (int)STGM.READ);

                // attempt to resolve a broken shortcut
                SLR_FLAGS flags = new SLR_FLAGS();

                if (userInputHwnd == IntPtr.Zero)
                {
                    flags = SLR_FLAGS.SLR_NO_UI;
                }

                shellLink.Resolve(userInputHwnd, flags);
            }
            catch (Exception e)
            {
                ShellLogger.Error($"ShellLinkHelper: Unable to load link from path {existingLinkPath}", e);
            }

            return shellLink;
        }
    }
}
