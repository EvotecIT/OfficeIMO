using System;
using System.Globalization;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointCustomShowAction {
        internal const string Prefix = "ppaction://customshow?id=";
        private const string ReturnSuffix = "&return=true";

        internal static bool IsCustomShowAction(string? action) =>
            action?.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase) == true;

        internal static bool TryParseSupported(string? action, out uint customShowId,
            out bool returnsToSlide) {
            customShowId = 0;
            returnsToSlide = false;
            if (!IsCustomShowAction(action)) return false;
            string value = action!.Substring(Prefix.Length);
            if (value.EndsWith(ReturnSuffix, StringComparison.OrdinalIgnoreCase)) {
                returnsToSlide = true;
                value = value.Substring(0, value.Length - ReturnSuffix.Length);
            }
            return value.Length > 0 && value.IndexOf('&') < 0
                && uint.TryParse(value, NumberStyles.None,
                    CultureInfo.InvariantCulture, out customShowId);
        }
    }
}
