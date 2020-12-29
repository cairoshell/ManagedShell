namespace ManagedShell.Common.Helpers
{
    public class MouseHelper
    {
        public static uint GetCursorPositionParam()
        {
            return ((uint)System.Windows.Forms.Cursor.Position.Y << 16) | (uint)System.Windows.Forms.Cursor.Position.X;
        }
    }
}
