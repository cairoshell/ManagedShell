using ManagedShell.Common.Structs;
using ManagedShell.Interop;
using System.Collections.Generic;
using System.Globalization;
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

            return new KeyboardLayout()
            {
                HKL = layout,
                NativeName = CultureInfo.GetCultureInfo((short)layout).NativeName,
                ThreeLetterName = CultureInfo.GetCultureInfo((short)layout).ThreeLetterISOLanguageName.ToUpper()
            };
        }

        public static List<KeyboardLayout> GetKeyboardLayoutList()
        {
            var size = NativeMethods.GetKeyboardLayoutList(0, null);
            var result = new long[size];
            NativeMethods.GetKeyboardLayoutList(size, result);

            return result.Select(x => new KeyboardLayout()
            {
                HKL = (int)x,
                NativeName = CultureInfo.GetCultureInfo((short)x).NativeName,
                ThreeLetterName = CultureInfo.GetCultureInfo((short)x).ThreeLetterISOLanguageName.ToUpper()
            }).ToList();
        }

        public static bool SetKeyboardLayout(int layoutId)
        {
            var loadedHkl = new IntPtr(NativeMethods.LoadKeyboardLayout(((short)layoutId).ToString("x8"), (uint)NativeMethods.KLF.SUBSTITUTE_OK));
            return NativeMethods.PostMessage(NativeMethods.GetForegroundWindow(), (int)NativeMethods.WM.INPUTLANGCHANGEREQUEST, IntPtr.Zero, loadedHkl);
        }
    }
}
