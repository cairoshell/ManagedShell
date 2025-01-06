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
        private static Guid IID_ActionCenterExperienceManager = new Guid("df65db57-d504-456e-8bd7-004ce308d8d9");
        private static Guid IID_ControlCenterExperienceManager = new Guid("d669a58e-6b18-4d1d-9004-a8862adb0a20");
        private static Guid IID_TrayMtcUvcFlyoutExperienceManager = new Guid("7154c95d-c519-49bd-a97e-645bbfabe111");
        private static Guid IID_NetworkFlyoutExperienceManager = new Guid("E44F17E6-AB85-409C-8D01-17D74BEC150E");
        private static Guid IID_NetworkFlyoutExperienceManager_19H2 = new Guid("c9ddc674-b44b-4c67-9d79-2b237d9be05a");
        private static Guid IID_TrayBatteryFlyoutExperienceManager = new Guid("0a73aedc-1c68-410d-8d53-63af80951e8f");

        private static Interfaces.IServiceProvider _immersiveShell;
        private static IShellExperienceManagerFactory _shellExperienceManagerFactory;
        private static ITrayClockFlyoutExperienceManager _trayClockFlyoutExperienceManager;

        public static void AllowExplorerFocus()
        {
            try
            {
                Interop.NativeMethods.GetWindowThreadProcessId(Interop.NativeMethods.FindWindowEx(IntPtr.Zero, IntPtr.Zero, "Progman", "Program Manager"), out uint procId);
                Interop.NativeMethods.AllowSetForegroundWindow(procId);
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to allow Explorer to set the foreground window: {ex}");
            }
        }

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

        private static IntPtr GetExperienceManagerFromFactory(string experienceManager)
        {
            _shellExperienceManagerFactory ??= GetShellExperienceManagerFactory();
            if (_shellExperienceManagerFactory == null) return IntPtr.Zero;

            try
            {
                IntPtr hString = IntPtr.Zero;
                if (NativeMethods.WindowsCreateString(experienceManager, experienceManager.Length, ref hString) != 0)
                {
                    ShellLogger.Warning("ImmersiveShell: Unable to create experience manager string");
                    return IntPtr.Zero;
                }

                _shellExperienceManagerFactory.GetExperienceManager(hString, out IntPtr pExperienceManagerInterface);
                NativeMethods.WindowsDeleteString(hString);
                return pExperienceManagerInterface;
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to create experience manager: {ex}");
                return IntPtr.Zero;
            }
        }

        internal static ITrayClockFlyoutExperienceManager GetTrayClockFlyoutExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: ITrayClockFlyoutExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.TrayClockFlyout");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

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

        public static IActionCenterExperienceManager GetActionCenterExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: IActionCenterExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.ActionCenter");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_ActionCenterExperienceManager, out IntPtr pActionCenterManager) == 0)
                {
                    return (IActionCenterExperienceManager)Marshal.GetObjectForIUnknown(pActionCenterManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query IActionCenterExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get IActionCenterExperienceManager: {ex}");
            }
            return null;
        }

        public static IControlCenterExperienceManager GetControlCenterExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter || EnvironmentHelper.IsWindows1124H2OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: IControlCenterExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.ControlCenter");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_ControlCenterExperienceManager, out IntPtr pControlCenterManager) == 0)
                {
                    return (IControlCenterExperienceManager)Marshal.GetObjectForIUnknown(pControlCenterManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query IControlCenterExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get IControlCenterExperienceManager: {ex}");
            }
            return null;
        }

        internal static ITrayMtcUvcFlyoutExperienceManager GetMtcUtcExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: ITrayMtcUvcFlyoutExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.MtcUvc");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_TrayMtcUvcFlyoutExperienceManager, out IntPtr pMtcUvcManager) == 0)
                {
                    return (ITrayMtcUvcFlyoutExperienceManager)Marshal.GetObjectForIUnknown(pMtcUvcManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query ITrayMtcUvcFlyoutExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get ITrayMtcUvcFlyoutExperienceManager: {ex}");
            }
            return null;
        }

        internal static INetworkFlyoutExperienceManager GetNetworkExperienceManager()
        {
            // TODO: The version is wrong
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: INetworkFlyoutExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.NetworkFlyout");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_NetworkFlyoutExperienceManager, out IntPtr pNetworkManager) == 0)
                {
                    return (INetworkFlyoutExperienceManager)Marshal.GetObjectForIUnknown(pNetworkManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query INetworkFlyoutExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get INetworkFlyoutExperienceManager: {ex}");
            }
            return null;
        }

        internal static INetworkFlyoutExperienceManager_19H2 GetNetworkExperienceManager_19H2()
        {
            // TODO: The version is wrong
            if (!EnvironmentHelper.IsWindows10OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: INetworkFlyoutExperienceManager_19H2 unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.NetworkFlyout");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_NetworkFlyoutExperienceManager_19H2, out IntPtr pNetworkManager) == 0)
                {
                    return (INetworkFlyoutExperienceManager_19H2)Marshal.GetObjectForIUnknown(pNetworkManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query INetworkFlyoutExperienceManager_19H2");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get INetworkFlyoutExperienceManager_19H2: {ex}");
            }
            return null;
        }

        internal static ITrayBatteryFlyoutExperienceManager GetBatteryExperienceManager()
        {
            if (!EnvironmentHelper.IsWindows10OrBetter || EnvironmentHelper.IsWindows1124H2OrBetter)
            {
                ShellLogger.Error("ImmersiveShell: ITrayBatteryFlyoutExperienceManager unsupported");
                return null;
            }

            try
            {
                IntPtr pExperienceManagerInterface = GetExperienceManagerFromFactory("Windows.Internal.ShellExperience.TrayBatteryFlyout");
                if (pExperienceManagerInterface == IntPtr.Zero) return null;

                if (Marshal.QueryInterface(pExperienceManagerInterface, ref IID_TrayBatteryFlyoutExperienceManager, out IntPtr pBatteryManager) == 0)
                {
                    return (ITrayBatteryFlyoutExperienceManager)Marshal.GetObjectForIUnknown(pBatteryManager);
                }

                ShellLogger.Warning("ImmersiveShell: Unable to query ITrayBatteryFlyoutExperienceManager");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to get ITrayBatteryFlyoutExperienceManager: {ex}");
            }
            return null;
        }

        public static void ShowClockFlyout(Interop.NativeMethods.Rect anchorRect)
        {
            _trayClockFlyoutExperienceManager ??= GetTrayClockFlyoutExperienceManager();
            if (_trayClockFlyoutExperienceManager == null) return;

            AllowExplorerFocus();

            try
            {
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

        public static void ShowMediaFlyout(Interop.NativeMethods.Rect anchorRect)
        {
            ITrayMtcUvcFlyoutExperienceManager manager = GetMtcUtcExperienceManager();
            if (manager == null) return;

            AllowExplorerFocus();

            try
            {
                if (manager.ShowFlyout(new Windows.Foundation.Rect(anchorRect.Left, anchorRect.Top, anchorRect.Width, anchorRect.Height)) == 0)
                {
                    return;
                }

                ShellLogger.Warning("ImmersiveShell: Unable to show media flyout");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show media flyout: {ex}");
            }
        }

        public static void ShowNetworkFlyout(Interop.NativeMethods.Rect anchorRect)
        {
            AllowExplorerFocus();

            try
            {
                // TODO: The version is not right
                if (EnvironmentHelper.IsWindows10RS4OrBetter)
                {
                    INetworkFlyoutExperienceManager_19H2 manager = GetNetworkExperienceManager_19H2();
                    if (manager == null) return;

                    if (manager.ShowFlyout(new Windows.Foundation.Rect(anchorRect.Left, anchorRect.Top, anchorRect.Width, anchorRect.Height), 0) == 0)
                    {
                        return;
                    }
                }
                else
                {
                    INetworkFlyoutExperienceManager manager = GetNetworkExperienceManager();
                    if (manager == null) return;

                    if (manager.ShowFlyout(new Windows.Foundation.Rect(anchorRect.Left, anchorRect.Top, anchorRect.Width, anchorRect.Height)) == 0)
                    {
                        return;
                    }
                }

                ShellLogger.Warning("ImmersiveShell: Unable to show network flyout");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show network flyout: {ex}");
            }
        }

        public static void ShowBatteryFlyout(Interop.NativeMethods.Rect anchorRect)
        {
            ITrayBatteryFlyoutExperienceManager manager = GetBatteryExperienceManager();
            if (manager == null) return;

            AllowExplorerFocus();

            try
            {
                if (manager.ShowFlyout(new Windows.Foundation.Rect(anchorRect.Left, anchorRect.Top, anchorRect.Width, anchorRect.Height)) == 0)
                {
                    return;
                }

                ShellLogger.Warning("ImmersiveShell: Unable to show battery flyout");
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show battery flyout: {ex}");
            }
        }

        public static void ShowActionCenter()
        {
            IActionCenterExperienceManager manager = GetActionCenterExperienceManager();
            if (manager == null) return;

            AllowExplorerFocus();

            try
            {
                manager.HotKeyInvoked(0);
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show action center: {ex}");
            }
        }

        public static void ShowControlCenter()
        {
            IControlCenterExperienceManager manager = GetControlCenterExperienceManager();
            if (manager == null) return;

            AllowExplorerFocus();

            try
            {
                manager.HotKeyInvoked(0);
            }
            catch (Exception ex)
            {
                ShellLogger.Warning($"ImmersiveShell: Unable to show control center: {ex}");
            }
        }
    }
}
