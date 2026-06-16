using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class HtmlStyleDeclarationParser {
    internal static RtfTextDirection? ParseDirection(string? value) {
        string normalized = string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : value!.Trim().ToLowerInvariant();
        switch (normalized) {
            case "rtl":
            case "right-to-left":
                return RtfTextDirection.RightToLeft;
            case "ltr":
            case "left-to-right":
                return RtfTextDirection.LeftToRight;
            default:
                return null;
        }
    }

    internal static bool TryParseLanguageId(string? value, out int languageId) {
        languageId = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string normalized = value!.Trim();
        if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0) {
            languageId = parsed;
            return true;
        }

        try {
            CultureInfo culture = CultureInfo.GetCultureInfo(normalized);
            languageId = culture.LCID;
            return true;
        } catch (CultureNotFoundException) {
            return false;
        }
    }
}
