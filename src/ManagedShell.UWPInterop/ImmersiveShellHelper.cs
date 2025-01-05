using ManagedShell.Common.Helpers;
using ManagedShell.Common.Logging;
using ManagedShell.UWPInterop.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop
{
    public static class ImmersiveShellHelper
    {
        private static Guid CLSID_ShellExperienceManagerFactory = new Guid("2E8FCB18-A0EE-41AD-8EF8-77FB3A370CA5");
        private static Guid IID_TrayClockFlyoutExperienceManager = new Guid("b1604325-6b59-427b-bf1b-80a2db02d3d8");

        private static Interfaces.IServiceProvider _immersiveShell;
        private static IShellExperienceManagerFactory _shellExperienceManagerFactory;
        private static ITrayClockFlyoutExperienceManager _trayClockFlyoutExperienceManager;

        public static Interfaces.IServiceProvider GetImmersiveShell()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: ImmersiveShell unsupported");
                return null;
            }

            try
            {
                _immersiveShell ??= (Interfaces.IServiceProvider)new CImmersiveShell();
                return _immersiveShell;
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to create ImmersiveShell: {ex}");
                return null;
            }
        }

        public static IShellExperienceManagerFactory GetShellExperienceManagerFactory()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: IShellExperienceManagerFactory unsupported");
                return null;
            }

            try
            {
                if (GetImmersiveShell().QueryService(CLSID_ShellExperienceManagerFactory, CLSID_ShellExperienceManagerFactory, out object factoryObj) == 0)
                {
                    return (IShellExperienceManagerFactory)factoryObj;
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query IShellExperienceManagerFactory");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to create IShellExperienceManagerFactory: {ex}");
            }
            return null;
        }

        internal static ITrayClockFlyoutExperienceManager GetTrayClockFlyoutExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter || EnvironmentHelper.IsWindows1124H2OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: ITrayClockFlyoutExperienceManager unsupported");
                return null;
            }

            _shellExperienceManagerFactory ??= GetShellExperienceManagerFactory();
            if (_shellExperienceManagerFactory == null) return null;

            try
            {
                string str = "Windows.Internal.ShellExperience.TrayClockFlyout";
                IntPtr hString = IntPtr.Zero;
                if (NativeMethods.WindowsCreateString(str, str.Length, ref hString) != 0)
                {
                    ShellLogger.Warning("ImmersiveShell: Unable to create experience manager string");
                    return null;
                }

                _shellExperienceManagerFactory.GetExperienceManager(hString, out IntPtr pExperienceManagerInterface);
                NativeMethods.WindowsDeleteString(hString);

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_TrayClockFlyoutExperienceManager, out IntPtr pClockFlyoutManager) == 0)
                {
                    return (ITrayClockFlyoutExperienceManager)Marshal.GetObjectForIUnknown(pClockFlyoutManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query ITrayClockFlyoutExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get ITrayClockFlyoutExperienceManager: {ex}");
            }
            return null;
        }

        public static void ShowClockFlyout(Interop.NativeMethods.Rect anchorRect)
        {
            _trayClockFlyoutExperienceManager ??= GetTrayClockFlyoutExperienceManager();
            if (_trayClockFlyoutExperienceManager == null) return;

            try
            {
                // Allow Explorer to steal focus
                Interop.NativeMethods.GetWindowThreadProcessId(Interop.NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Progman", "Program Manager"), out uint procId);
                Interop.NativeMethods.AllowSetForegroundWindow(procId);

                if (_trayClockFlyoutExperienceManager.ShowFlyout(new Windows.Foundation.Rect(anchorRect.Left, anchorRect.Top, anchorRect.Width, anchorRect.Height)) == 0)
                {
                    return;
                }

                ShellLogger.Warning("ImmersiveShell: Unable to show clock flyout");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show clock flyout: {ex}");
            }
        }
    }
}
