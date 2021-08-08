using System;
using ManagedShell.Common.Enums;
using ManagedShell.Common.Interfaces;
using ManagedShell.Common.Logging;
using ManagedShell.Common.SupportingClasses;

namespace ManagedShell.Common.Helpers
{
    public class AppVisibilityHelper : IDisposable
    {
        public event EventHandler<AppVisibilityEventArgs> AppVisibilityChanged;
        public event EventHandler<LauncherVisibilityEventArgs> LauncherVisibilityChanged;

        private IAppVisibility appVis;
        private int eventCookie;

        public AppVisibilityHelper()
        {
            if (!EnvironmentHelper.IsWindows8OrBetter)
            {
                return;
            }

            // register for app visibility events
            appVis = (IAppVisibility)new AppVisibility();

            if (appVis == null)
            {
                return;
            }

            AppVisibilityEvents events = new AppVisibilityEvents();
            events.AppVisibilityChanged += Events_AppVisibilityChanged;
            events.LauncherVisibilityChanged += Events_LauncherVisibilityChanged;

            if (appVis.Advise(events, out eventCookie) == 0)
            {
                // subscribed to events successfully
                ShellLogger.Debug("AppVisibilityHelper: Subscribed to change events");
            }
        }

        private void Events_AppVisibilityChanged(object sender, AppVisibilityEventArgs e)
        {
            AppVisibilityChanged?.Invoke(sender, e);
        }

        private void Events_LauncherVisibilityChanged(object sender, LauncherVisibilityEventArgs e)
        {
            LauncherVisibilityChanged?.Invoke(sender, e);
        }

        public bool IsLauncherVisible()
        {
            if (!EnvironmentHelper.IsWindows8OrBetter)
            {
                return false;
            }

            if (appVis == null)
            {
                return false;
            }

            appVis.IsLauncherVisible(out bool pfVisible);

            return pfVisible;
        }

        public MONITOR_APP_VISIBILITY GetAppVisibilityOnMonitor(IntPtr hMonitor)
        {
            if (!EnvironmentHelper.IsWindows8OrBetter)
            {
                return MONITOR_APP_VISIBILITY.MAV_UNKNOWN;
            }

            if (appVis == null)
            {
                return MONITOR_APP_VISIBILITY.MAV_UNKNOWN;
            }

            appVis.GetAppVisibilityOnMonitor(hMonitor, out MONITOR_APP_VISIBILITY pMode);

            return pMode;
        }

        public void Dispose()
        {
            if (!EnvironmentHelper.IsWindows8OrBetter)
            {
                return;
            }

            if (appVis == null)
            {
                return;
            }

            if (eventCookie > 0)
            {
                // unregister from events
                if (appVis.Unadvise(eventCookie) == 0)
                {
                    ShellLogger.Debug("AppVisibilityHelper: Unsubscribed from change events");
                }
            }
        }
    }
}
