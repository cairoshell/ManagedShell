using ManagedShell.Interop;

namespace ManagedShell.Common.Helpers
{
    public static class LayoutHelper
    {
        public static int GetKeyboardLayout(bool currentThread = false)
        {
            return NativeMethods.GetKeyboardLayout(currentThread ? 0 : NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out _));
        }

        public static long[] GetKeyboardLayoutList()
        {
            var size = NativeMethods.GetKeyboardLayoutList(0, null);
            var result = new long[size];
            NativeMethods.GetKeyboardLayoutList(size, result);

            return result;
        }

        public static bool SetKeyboardLayout(int layoutId)
        {
            return NativeMethods.PostMessage(0xffff,
                (uint) NativeMethods.WM.INPUTLANGCHANGEREQUEST,
                0,
                NativeMethods.LoadKeyboardLayout(layoutId.ToString("x8"), (uint)(NativeMethods.KLF.SUBSTITUTE_OK | NativeMethods.KLF.ACTIVATE)));
        }
    }
}
