using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointTextDefaults {
        internal const double DefaultFontSizePoints = 18D;
        internal const string LegacyFallbackFontFamily = "Calibri";

        internal static string ResolveBodyLatinFont(PowerPointSlide? slide) {
            A.FontScheme? overrideScheme = slide?.SlidePart.ThemeOverridePart?.ThemeOverride?.FontScheme
                ?? slide?.SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme;
            string? overrideTypeface = overrideScheme?.MinorFont?.LatinFont?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(overrideTypeface)) {
                return overrideTypeface!;
            }

            string? masterTypeface = slide?.SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?
                .ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
            return string.IsNullOrWhiteSpace(masterTypeface)
                ? LegacyFallbackFontFamily
                : masterTypeface!;
        }
    }
}
