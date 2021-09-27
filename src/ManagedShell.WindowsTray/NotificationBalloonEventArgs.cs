using System;

namespace ManagedShell.WindowsTray
{
    public class NotificationBalloonEventArgs : EventArgs
    {
        public NotificationBalloonInfo BalloonInfo;
        public NotifyIcon NotifyIcon;
    }
}
