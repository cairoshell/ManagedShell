using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("7154c95d-c519-49bd-a97e-645bbfabe111"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITrayMtcUvcFlyoutExperienceManager
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
