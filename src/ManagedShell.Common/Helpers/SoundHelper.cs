using ManagedShell.Common.Logging;
using Microsoft.Win32;
using System;
using System.Runtime.InteropServices;

namespace ManagedShell.Common.Helpers
{
    public class SoundHelper
    {
        private const string SYSTEM_SOUND_ROOT_KEY = @"AppEvents\Schemes\Apps";

        [Flags]
        private enum PlaySoundFlags
        {
            SND_ASYNC = 0x00000001,
            SND_NODEFAULT = 0x00000002,
            SND_ALIAS = 0x00010000,
            SND_FILENAME = 0x00020000,
            SND_SYSTEM = 0x00200000,
        }

        private const PlaySoundFlags DEFAULT_SYSTEM_SOUND_FLAGS = PlaySoundFlags.SND_ASYNC | PlaySoundFlags.SND_NODEFAULT | PlaySoundFlags.SND_SYSTEM;

        [DllImport("winmm.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool PlaySound(string pszSound, IntPtr hmod, PlaySoundFlags soundFlags);

        /// <summary>
        /// Plays the specified system sound using the audio session for system notification sounds.
        /// </summary>
        /// <param name="app">The name of the app that the sound belongs to. For example, ".Default" contains system sounds, "Explorer" contains Explorer sounds.</param>
        /// <param name="name">The name of the system sound to play.</param>
        public static bool PlaySystemSound(string app, string name)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey($@"{SYSTEM_SOUND_ROOT_KEY}\{app}\{name}\.Current");
                if (key == null)
                {
                    ShellLogger.Error($"SoundHelper: Unable to find sound {name} for app {app}");
                    return false;
                }

                var soundFileName = key.GetValue(null) as string;
                if (string.IsNullOrEmpty(soundFileName))
                {
                    ShellLogger.Error($"SoundHelper: Missing file for sound {name} for app {app}");
                    return false;
                }

                return PlaySound(soundFileName, IntPtr.Zero, DEFAULT_SYSTEM_SOUND_FLAGS | PlaySoundFlags.SND_FILENAME);
            }
            catch (Exception e)
            {
                ShellLogger.Debug($"SoundHelper: Unable to play sound {name} for app {app}: {e.Message}");
                return false;
            }
        }

        /// <summary>
        /// Plays the specified system sound using the audio session for system notification sounds.
        /// </summary>
        /// <param name="alias">The name of the system sound for ".Default" to play.</param>
        public static void PlaySystemSound(string alias)
        {
            try
            {
                PlaySound(alias, IntPtr.Zero, DEFAULT_SYSTEM_SOUND_FLAGS | PlaySoundFlags.SND_ALIAS);
            }
            catch (Exception e)
            {
                ShellLogger.Error($"SoundHelper: Unable to play sound {alias}: {e.Message}");
            }
        }

        /// <summary>
        /// Plays the system notification sound.
        /// </summary>
        public static void PlayNotificationSound()
        {
            // System default sound for the classic notification balloon.
            if (PlaySystemSound("Explorer", "SystemNotification")) return;
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

        public static void PlayXPNotificationSound()
        {
            PlaySystemSound("SystemNotification");
        }
    }
}