using ManagedShell.Common.Structs;
using ManagedShell.Interop;
using System.Collections.Generic;
using System.Linq;
using System;

namespace ManagedShell.Common.Helpers
{
    public static class KeyboardLayoutHelper
    {
        public static KeyboardLayout GetKeyboardLayout()
        {
            uint threadId = NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out _);
            var layout = NativeMethods.GetKeyboardLayout(threadId);

            return new KeyboardLayout(layout);
        }

        public static List<KeyboardLayout> GetKeyboardLayoutList()
        {
            var size = NativeMethods.GetKeyboardLayoutList(0, null);
            var result = new long[size];
            NativeMethods.GetKeyboardLayoutList(size, result);

            return result.Select(x => new KeyboardLayout((int)x)).ToList();
        }

        public static bool SetKeyboardLayout(int layoutId)
        {
            return NativeMethods.PostMessage(NativeMethods.GetForegroundWindow(), (int)NativeMethods.WM.INPUTLANGCHANGEREQUEST, IntPtr.Zero, new IntPtr(layoutId));
        }
    }
}
