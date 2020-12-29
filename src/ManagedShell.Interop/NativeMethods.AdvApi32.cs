using System;
using System.Runtime.InteropServices;

namespace ManagedShell.Interop
{
    public partial class NativeMethods
    {
        const string AdvApi32_DllName = "advapi32.dll";
        public const uint TOKENADJUSTPRIVILEGES = 0x00000020;
        public const uint TOKENQUERY = 0x00000008;

        /// <summary>
        /// Structure for the token privileges request.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct TOKEN_PRIVILEGES
        {
            /// <summary>
            /// The number of privileges.
            /// </summary>
            public uint PrivilegeCount;

            /// <summary>
            /// The local UID for the request.
            /// </summary>
            public long Luid;

            /// <summary>
            /// Attributes for the request.
            /// </summary>
            public uint Attributes;
        }

        [DllImport(AdvApi32_DllName, SetLastError = true)]
        public static extern bool OpenProcessToken(IntPtr processHandle, uint desiredAccess, out IntPtr tokenHandle);

        [DllImport(AdvApi32_DllName, SetLastError = true)]
        public static extern bool AdjustTokenPrivileges(IntPtr tokenHandle, bool disableAllPrivileges, ref TOKEN_PRIVILEGES newState, uint bufferLength, IntPtr previousState, IntPtr returnLength);

        [DllImport(AdvApi32_DllName, SetLastError = true)]
        public static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);
    }
}
