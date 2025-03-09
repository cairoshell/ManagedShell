using ManagedShell.Common.Logging;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;

namespace ManagedShell.Common.Helpers
{
    public class SoundHelper
    {
        private const string SYSTEM_SOUND_ROOT_KEY = @"AppEvents\Schemes\Apps";
        private const int SND_FILENAME = 0x00020000;
        private const int SND_ASYNC = 0x0001;
        private const long SND_SYSTEM = 0x00200000L;

        [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool PlaySound(string pszSound, IntPtr hmod, uint fdwSound);

        /// <summary>
        /// Plays the specified system sound using the audio session for system notification sounds.
        /// </summary>
        /// <param name="app">The name of the app that the sound belongs to. For example, ".Default" contains system sounds, "Explorer" contains Explorer sounds.</param>
        /// <param name="name">The name of the system sound to play.</param>
        public static bool PlaySystemSound(string app, string name)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey($"{SYSTEM_SOUND_ROOT_KEY}\\{app}\\{name}\\.Current"))
                {
                    if (key == null)
                    {
                        ShellLogger.Debug($"SoundHelper: Unable to find sound {name} for app {app}");
                        return false;
                    }

                    if (key.GetValue(null) is string soundFileName)
                    {
                        if (string.IsNullOrEmpty(soundFileName))
                        {
                            ShellLogger.Debug($"SoundHelper: Missing file for sound {name} for app {app}");
                            return false;
                        }

                        return PlaySound(soundFileName, IntPtr.Zero, (uint)(SND_ASYNC | SND_FILENAME | SND_SYSTEM));
                    }
                    else
                    {
                        ShellLogger.Debug($"SoundHelper: Missing file for sound {name} for app {app}");
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                ShellLogger.Debug($"SoundHelper: Unable to play sound {name} for app {app}: {e.Message}");
                return false;
            }
        }

        /// <summary>
        /// Plays the system notification sound.
        /// </summary>
        public static void PlayNotificationSound()
        {
            // System default sound for the classic notification balloon.
            if (!PlaySystemSound("Explorer", "SystemNotification"))
            {
                if (EnvironmentHelper.IsWindows8OrBetter)
                {
                    // Toast notification sound.
                    if (!PlaySystemSound(".Default", "Notification.Default"))
                        PlayXPNotificationSound();
                }
                else
                {
                    PlayXPNotificationSound();
                }
            }
        }

        public static bool PlayXPNotificationSound()
        {
            return PlaySystemSound(".Default", "SystemNotification");
        }
    }
}