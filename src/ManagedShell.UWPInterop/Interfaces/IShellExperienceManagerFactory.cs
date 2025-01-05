using System;
using System.Runtime.InteropServices;

namespace ManagedShell.UWPInterop.Interfaces
{
    [ComImport, Guid("2E8FCB18-A0EE-41AD-8EF8-77FB3A370CA5"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellExperienceManagerFactory
    {
        // Note: Invoking methods on ComInterfaceType.InterfaceIsIInspectable interfaces
        // is no longer supported in the CLR, but can be simulated with IUnknown.
        void GetIids(out int iidCount, out IntPtr iids);
        void GetRuntimeClassName(out IntPtr className);
        void GetTrustLevel(out int trustLevel);

        void GetExperienceManager(IntPtr hStrExperience, out IntPtr pp);
    }
}
