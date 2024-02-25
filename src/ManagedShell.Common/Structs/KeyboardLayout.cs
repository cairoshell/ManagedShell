using System.Globalization;

namespace ManagedShell.Common.Structs
{
    public struct KeyboardLayout
    {
        public uint Id { get; set; }
        public ushort LanguageId { get; set; }
        public ushort KeyboardId { get; set; }

        public string NativeName { get; set; }
        public string ThreeLetterName { get; set; }
        public string DisplayName { get; set; }

        public KeyboardLayout(uint layoutId)
        {
            Id = layoutId;
            LanguageId = (ushort)(layoutId & 0xFFFF);
            KeyboardId = (ushort)(layoutId >> 16);

            var cultureInfo = CultureInfo.GetCultureInfo(LanguageId);
            NativeName = cultureInfo.NativeName;
            ThreeLetterName = cultureInfo.ThreeLetterWindowsLanguageName;
            DisplayName = cultureInfo.DisplayName;
        }
    }
}