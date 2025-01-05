using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("b1604325-6b59-427b-bf1b-80a2db02d3d8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITrayClockFlyoutExperienceManager
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
