﻿using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("7154c95d-c519-49bd-a97e-645bbfabe111"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ITrayMtcUvcFlyoutExperienceManager
    {
        void GetIids(out int iidCount, out IntPtr iids);
        void GetRuntimeClassName(out IntPtr className);
        void GetTrustLevel(out int trustLevel);

        [PreserveSig]
        int ShowFlyout(Windows.Foundation.Rect rect);

        [PreserveSig]
        int HideFlyout();
    }
}
