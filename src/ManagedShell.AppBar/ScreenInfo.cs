using System.Drawing;
using System.Windows.Forms;

namespace ManagedShell.AppBar
{
	public class ScreenInfo
	{
		private ScreenInfo(string deviceName, Rectangle bounds)
		{
			DeviceName = deviceName;
			Bounds = bounds;
		}

		public static ScreenInfo Create(Screen screen)
		{
			return new ScreenInfo(screen.DeviceName, screen.Bounds);
		}

		public static ScreenInfo CreateVirtualScreen()
		{
			return new ScreenInfo(nameof(SystemInformation.VirtualScreen), SystemInformation.VirtualScreen);
		}

		public string DeviceName { get; }

		public Rectangle Bounds { get; }

		public bool IsVirtualScreen => DeviceName == nameof(SystemInformation.VirtualScreen);

		protected bool Equals(ScreenInfo other)
		{
			return DeviceName == other.DeviceName;
		}

		public override bool Equals(object obj)
		{
			if (ReferenceEquals(null, obj)) return false;
			if (ReferenceEquals(this, obj)) return true;
			if (obj.GetType() != this.GetType()) return false;
			return Equals((ScreenInfo) obj);
		}

		public override int GetHashCode()
		{
			return (DeviceName != null ? DeviceName.GetHashCode() : 0);
		}
	}
}