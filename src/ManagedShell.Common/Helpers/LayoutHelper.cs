using ManagedShell.Interop;

namespace ManagedShell.Common.Helpers
{
    public static class LayoutHelper
    {
        public static int GetCurrentLayoutId(bool currentThread = false)
        {
            return NativeMethods.GetKeyboardLayout(currentThread ? 0 : NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out _)).ToInt32();
        }
    }
}
