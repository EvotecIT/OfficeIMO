using System;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointTextDefaults {
        internal const double DefaultFontSizePoints = 18D;
        internal const string LegacyFallbackFontFamily = "Calibri";

        internal static int? ToDrawingFontSize(double? points, string parameterName) {
            if (!points.HasValue) return null;
            if (double.IsNaN(points.Value) || double.IsInfinity(points.Value)
                || points.Value < 1D || points.Value > 4000D) {
                throw new ArgumentOutOfRangeException(parameterName, points,
                    "DrawingML font size must be a finite value from 1 through 4000 points.");
            }
            return checked((int)Math.Round(points.Value * 100D, MidpointRounding.AwayFromZero));
        }

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
