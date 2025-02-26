using System;

namespace ManagedShell.AppBar
{
    public class FullScreenApp
    {
        public IntPtr hWnd;
        public ScreenInfo screen;
        public string title;
        public bool fromTasksService;
    }
}