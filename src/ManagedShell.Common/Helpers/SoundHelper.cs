using ManagedShell.Common.Logging;
using Microsoft.Win32;
using System;
using System.Media;

namespace ManagedShell.Common.Helpers
{
    public class SoundHelper
    {
        private const string SYSTEM_SOUND_ROOT_KEY = @"AppEvents\Schemes\Apps";

        /// <summary>
        /// Plays the specified system sound.
        /// </summary>
        /// <param name="app">The name of the app that the sound belongs to. For example, ".Default" contains system sounds, "Explorer" contains Explorer sounds.</param>
        /// <param name="name">The name of the system sound to play.</param>
        public static void PlaySystemSound(string app, string name)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey($"{SYSTEM_SOUND_ROOT_KEY}\\{app}\\{name}\\.Current"))
                {
                    if (key == null)
                    {
                        ShellLogger.Error($"SoundHelper: Unable to find sound {name} for app {app}");
                        return;
                    }

                    if (key.GetValue(null) is string soundFileName)
                    {
                        if (string.IsNullOrEmpty(soundFileName))
                        {
                            ShellLogger.Error($"SoundHelper: Missing file for sound {name} for app {app}");
                            return;
                        }

                        using (SoundPlayer soundPlayer = new SoundPlayer(soundFileName))
                        {
                            soundPlayer.Play();
                        }
                    }
                    else
                    {
                        ShellLogger.Error($"SoundHelper: Missing file for sound {name} for app {app}");
                    }
                }
            }
            catch (Exception e)
            {
                ShellLogger.Error($"SoundHelper: Unable to play sound {name} for app {app}: {e.Message}");
            }
        }

        /// <summary>
        /// Plays the system notification sound.
        /// </summary>
        public static void PlayNotificationSound()
        {
            if (EnvironmentHelper.IsWindows8OrBetter)
            {
                PlaySystemSound(".Default", "Notification.Default");
            }
            else
            {
                PlaySystemSound(".Default", "SystemNotification");
            }
        }
    }
}
