using ManagedShell.Common.Structs;
using ManagedShell.Interop;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.Common.Helpers
{
    public static class KeyboardLayoutHelper
    {
        private static uint GetFocusedThreadId()
        {
            var gti = new GUITHREADINFO();
            gti.cbSize = Marshal.SizeOf(typeof(GUITHREADINFO));

            // Some apps (e.g. Electron/WebView2 apps such as Teams, and WinUI3 apps such as the
            // modern Notepad) keep their real keyboard focus on a child window that belongs to a
            // different thread than the top-level foreground window. GetGUIThreadInfo(0, ...)
            // resolves the true focused window first.
            // idThread 0 asks for the foreground thread's queue; hwndFocus within it is the
            // window that truly owns keyboard focus, which may belong to a different thread
            // than the top-level foreground window itself.
            if (GetGUIThreadInfo(0, ref gti))
            {
                IntPtr focusedWindow = gti.hwndFocus != IntPtr.Zero ? gti.hwndFocus : gti.hwndActive;

                if (focusedWindow != IntPtr.Zero)
                {
                    uint focusedThreadId = GetWindowThreadProcessId(focusedWindow, out _);

                    if (focusedThreadId != 0)
                    {
                        return focusedThreadId;
                    }
                }
            }

            return GetWindowThreadProcessId(GetForegroundWindow(), out _);
        }

        public static KeyboardLayout GetKeyboardLayout(bool currentThread = false)
        {
            uint threadId = 0;
            if (!currentThread)
                threadId = GetFocusedThreadId();
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
            return PostMessage(0xffff,
                (uint) WM.INPUTLANGCHANGEREQUEST,
                0,
                (long)LoadKeyboardLayout(layoutId.ToString("x8"), (uint)(KLF.SUBSTITUTE_OK | KLF.ACTIVATE)));
        }
    }
}
