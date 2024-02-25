using ManagedShell.Common.Structs;
using ManagedShell.Interop;
using System.Collections.Generic;
using System;

namespace ManagedShell.Common.Helpers
{
    public static class KeyboardLayoutHelper
    {
        public static KeyboardLayout GetKeyboardLayout()
        {
            uint threadId = NativeMethods.GetWindowThreadProcessId(NativeMethods.GetForegroundWindow(), out _);
            var layout = NativeMethods.GetKeyboardLayout(threadId);

            return new KeyboardLayout((uint)layout);
        }

        public static List<KeyboardLayout> GetKeyboardLayoutList()
        {
            var keyboardLayouts = new List<KeyboardLayout>();

            var size = NativeMethods.GetKeyboardLayoutList(0, null);
            var layoutIds = new IntPtr[size];
            NativeMethods.GetKeyboardLayoutList(layoutIds.Length, layoutIds);

            foreach (var layoutId in layoutIds)
            {
                var keyboardLayout = new KeyboardLayout((uint)layoutId);
                keyboardLayouts.Add(keyboardLayout);
            }

            return keyboardLayouts;
        }

        public static bool SetKeyboardLayout(uint layoutId)
        {
            return NativeMethods.PostMessage(NativeMethods.GetForegroundWindow(), (int)NativeMethods.WM.INPUTLANGCHANGEREQUEST, IntPtr.Zero, new IntPtr(layoutId));
        }
    }
}