using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("c9ddc674-b44b-4c67-9d79-2b237d9be05a"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface INetworkFlyoutExperienceManager_19H2
    {
        // Note: Invoking methods on ComInterfaceType.InterfaceIsIInspectable interfaces
        // is no longer supported in the CLR, but can be simulated with IUnknown.
        void GetIids(out int iidCount, out IntPtr iids);
        void GetRuntimeClassName(out IntPtr className);
        void GetTrustLevel(out int trustLevel);

        [PreserveSig]
        int ShowFlyout(Windows.Foundation.Rect rect, int unk);

        [PreserveSig]
        int HideFlyout();
    }
}
