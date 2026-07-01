using System.Globalization;

namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocLanguageMapper {
        internal static ushort? TryReadLanguageId(string? languageTag, string description) {
            if (string.IsNullOrWhiteSpace(languageTag)) {
                return null;
            }

            string tag = languageTag!.Trim();
            CultureInfo culture = GetCulture(tag, description);
            if (culture.LCID < 0 || culture.LCID > ushort.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports {description} only when the culture maps to a Word 97-2003 language identifier.");
            }

            return checked((ushort)culture.LCID);
        }

        internal static string? TryGetLanguageTag(ushort languageId) {
            if (languageId == 0) {
                return null;
            }

            try {
                CultureInfo culture = CultureInfo.GetCultureInfo(languageId);
                return string.IsNullOrWhiteSpace(culture.Name) ? null : culture.Name;
            } catch (CultureNotFoundException) {
                return null;
            }
        }

        private static CultureInfo GetCulture(string languageTag, string description) {
            try {
                return CultureInfo.GetCultureInfo(languageTag.Trim());
            } catch (CultureNotFoundException exception) {
                throw new NotSupportedException($"Native DOC saving supports {description} only when '{languageTag}' is a recognized culture tag.", exception);
            }
        }
    }
}
