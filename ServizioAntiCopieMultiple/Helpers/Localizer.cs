using System.Globalization;
using System.Resources;

namespace ServizioAntiCopieMultiple.Helpers
{
    internal static class Localizer
    {
        private static string _lang = "it";
        public static string Language
        {
            get => _lang;
            set
            {
                _lang = string.IsNullOrEmpty(value) ? "it" : value;
                try
                {
                    var culture = new CultureInfo(_lang);
                    CultureInfo.CurrentUICulture = culture;
                }
                catch { }
            }
        }

        private static readonly ResourceManager _rm = new ResourceManager("ServizioAntiCopieMultiple.Resources.Strings", typeof(Localizer).Assembly);

        public static string T(string key)
        {
            try
            {
                var str = _rm.GetString(key, CultureInfo.CurrentUICulture);
                if (!string.IsNullOrEmpty(str)) return str;
            }
            catch { }
            // fallback to invariant/italian
            try
            {
                var str2 = _rm.GetString(key, new CultureInfo("it"));
                if (!string.IsNullOrEmpty(str2)) return str2;
            }
            catch { }
            return key;
        }
    }
}
