using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("E44F17E6-AB85-409C-8D01-17D74BEC150E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface INetworkFlyoutExperienceManager
    {
        // Note: Invoking methods on ComInterfaceType.InterfaceIsIInspectable interfaces
        // is no longer supported in the CLR, but can be simulated with IUnknown.
        void GetIids(out int iidCount, out IntPtr iids);
        void GetRuntimeClassName(out IntPtr className);
        void GetTrustLevel(out int trustLevel);

        [PreserveSig]
        int ShowFlyout(Windows.Foundation.Rect rect);

        [PreserveSig]
        int HideFlyout();
    }
}
