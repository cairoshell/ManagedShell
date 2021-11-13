using System;
using static ManagedShell.Interop.NativeMethods;

namespace ManagedShell.WindowsTray
{
    /// <summary>
    /// Delegate signature for the system tray callback.
    /// </summary>
    /// <param name="msg">The system tray message number.</param>
    /// <param name="nicData">The NotifyIconData structure</param>
    /// <returns>Indication of message outcome.</returns>
    public delegate bool SystrayDelegate(uint msg, SafeNotifyIconData nicData);

    /// <summary>
    /// Delegate signature for the icon data callback.
    /// </summary>
    /// <param name="dwMessage">The message sent</param>
    /// <param name="hWnd">The handle of the icon</param>
    /// <param name="uID">The the ID ofthe icon</param>
    /// <param name="guidItem">The GUID of the icon</param>
    /// <returns>Indication of message outcome.</returns>
    public delegate IntPtr IconDataDelegate(int dwMessage, uint hWnd, uint uID, Guid guidItem);

    /// <summary>
    /// Delegate signature for the tray host size callback.
    /// </summary>
    /// <returns>Indication of message outcome.</returns>
    public delegate TrayHostSizeData TrayHostSizeDelegate();

    /// <summary>
    /// Delegate signature for the AppBar autohide callback.
    /// </summary>
    /// <param name="edge">The AppBar edge to check for an autohide bar. Null to check any edge.</param>
    /// <returns>Window handle to the autohide AppBar.</returns>
    public delegate IntPtr AutoHideBarDelegate(ABEdge? edge);
}
