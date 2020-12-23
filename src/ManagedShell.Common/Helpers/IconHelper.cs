using ManagedShell.Interop;
using System;
using System.IO;
using System.Runtime.InteropServices;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.Common.Helpers
{
    public class IconHelper
    {
        public static ComTaskScheduler IconScheduler = new ComTaskScheduler();
        public static object ComLock = new object();
        
        // IImageList references
        private static Guid iidImageList = new Guid("46EB5926-582E-4017-9FDF-E8998DAA0950");
        private static IImageList iml0; // 32pt
        private static IImageList iml1; // 16pt
        private static IImageList iml2; // 48pt

        private static void initIml(int size)
        {
            // Initialize the appropriate IImageList for the desired icon size if it hasn't been already

            if (size == 0 && iml0 == null)
            {
                SHGetImageList(0, ref iidImageList, out iml0);
            }
            else if (size == 1 && iml1 == null)
            {
                SHGetImageList(1, ref iidImageList, out iml1);
            }
            else if (size == 2 && iml2 == null)
            {
                SHGetImageList(2, ref iidImageList, out iml2);
            }
        }

        public static void DisposeIml()
        {
            // Dispose any IImageList objects we instantiated.
            // Called by the main shutdown method.

            lock (ComLock)
            {
                if (iml0 != null)
                {
                    Marshal.ReleaseComObject(iml0);
                    iml0 = null;
                }
                if (iml1 != null)
                {
                    Marshal.ReleaseComObject(iml1);
                    iml1 = null;
                }
                if (iml2 != null)
                {
                    Marshal.ReleaseComObject(iml2);
                    iml2 = null;
                }
            }
        }

        public static IntPtr GetIconByFilename(string fileName, int size)
        {
            return GetIcon(fileName, size);
        }

        private static IntPtr GetIcon(string filename, int size)
        {
            lock (ComLock)
            {
                try
                {
                    filename = translateIconExceptions(filename);

                    SHFILEINFO shinfo = new SHFILEINFO();
                    shinfo.szDisplayName = string.Empty;
                    shinfo.szTypeName = string.Empty;
                    IntPtr hIconInfo;

                    if (!filename.StartsWith("\\") && (File.GetAttributes(filename) & FileAttributes.Directory) == FileAttributes.Directory)
                    {
                        hIconInfo = SHGetFileInfo(filename, FILE_ATTRIBUTE_NORMAL | FILE_ATTRIBUTE_DIRECTORY, ref shinfo, (uint)Marshal.SizeOf(shinfo), (uint)(SHGFI.SysIconIndex));
                    }
                    else
                    {
                        hIconInfo = SHGetFileInfo(filename, FILE_ATTRIBUTE_NORMAL, ref shinfo, (uint)Marshal.SizeOf(shinfo), (uint)(SHGFI.UseFileAttributes | SHGFI.SysIconIndex));
                    }

                    var iconIndex = shinfo.iIcon;

                    // Initialize the IImageList object
                    initIml(size);

                    IntPtr hIcon = IntPtr.Zero;
                    int ILD_TRANSPARENT = 1;

                    switch (size)
                    {
                        case 0:
                            iml0.GetIcon(iconIndex, ILD_TRANSPARENT, ref hIcon);
                            break;
                        case 1:
                            iml1.GetIcon(iconIndex, ILD_TRANSPARENT, ref hIcon);
                            break;
                        case 2:
                            iml2.GetIcon(iconIndex, ILD_TRANSPARENT, ref hIcon);
                            break;
                    }

                    return hIcon;
                }
                catch
                {
                    return IntPtr.Zero;
                }
            }
        }

        private static string translateIconExceptions(string filename)
        {
            if (filename.EndsWith(".settingcontent-ms"))
            {
                return "C:\\Windows\\ImmersiveControlPanel\\SystemSettings.exe";
            }

            return filename;
        }

    }
}
