using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ManagedShell.AppBar
{
    public class AppBarScreen
    {
        public Rectangle Bounds { get; set; }
        
        public string DeviceName { get; set; }
        
        public bool Primary { get; set; }
        
        public Rectangle WorkingArea { get; set; }

        public static AppBarScreen FromScreen(Screen screen)
        {
            return new AppBarScreen
            {
                Bounds = screen.Bounds,
                DeviceName = screen.DeviceName,
                Primary = screen.Primary,
                WorkingArea = screen.WorkingArea
            };
        }

        public static AppBarScreen FromPrimaryScreen()
        {
            return FromScreen(Screen.PrimaryScreen);
        }

        public static List<AppBarScreen> FromAllScreens()
        {
            List<AppBarScreen> screens = new List<AppBarScreen>();
            
            foreach (var screen in Screen.AllScreens)
            {
                screens.Add(FromScreen(screen));
            }

            return screens;
        }
    }
}
