using System;
using System.Runtime.InteropServices;

namespace ManagedShell.Interop
{
    public partial class NativeMethods
    {
        const string Comctl32_DllName = "comctl32.dll";

        [DllImport(Comctl32_DllName, SetLastError = true)]
        public static extern IntPtr ImageList_GetIcon(IntPtr himl, uint i, uint flags);

        [DllImport(Comctl32_DllName, SetLastError = true)]
        public static extern IntPtr ImageList_Duplicate(IntPtr himl);
    }
}
