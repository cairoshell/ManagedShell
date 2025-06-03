using System.Globalization;

namespace ManagedShell.Common.Structs
{
    public struct KeyboardLayout
    {
        public int HKL { get; set; }
        public string NativeName { get; set; }
        public string ThreeLetterName { get; set; }
        public string DisplayName { get; set; }

        public KeyboardLayout(int hkl)
        {
            HKL = hkl;
            var cultureInfo = CultureInfo.GetCultureInfo((short)hkl);

            NativeName = cultureInfo.NativeName;
            ThreeLetterName = cultureInfo.ThreeLetterISOLanguageName.ToUpper();
            DisplayName = cultureInfo.DisplayName;
        }
    }
}
