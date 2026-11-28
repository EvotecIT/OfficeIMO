using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointTextDefaults {
        internal const double DefaultFontSizePoints = 18D;
        // DrawingML bodyPr defaults: 0.1in horizontally, 0.05in vertically (ECMA-376).
        internal const double DefaultHorizontalInsetPoints = 7.2D;
        internal const double DefaultVerticalInsetPoints = 3.6D;
        internal const string LegacyFallbackFontFamily = "Calibri";

        internal static string ResolveBodyLatinFont(PowerPointSlide? slide) {
            if (slide == null) return LegacyFallbackFontFamily;
            PowerPointThemeScope scope = PowerPointThemeScope.For(slide.PartRef);
            string? typeface = MinorLatinTypeface(scope.FontScheme);
            if (!string.IsNullOrWhiteSpace(typeface)) return typeface!;

            // An override can omit the Latin face. In that case fall back to the master font.
            const int drawing = OfficeXmlNamespaces.DrawingML;
            typeface = MinorLatinTypeface(scope.Theme.Child(drawing, "themeElements").Child(drawing, "fontScheme"));
            return string.IsNullOrWhiteSpace(typeface) ? LegacyFallbackFontFamily : typeface!;
        }

        private static string? MinorLatinTypeface(OfficeXmlNode scheme) =>
            scheme.Child(OfficeXmlNamespaces.DrawingML, "minorFont").Child(OfficeXmlNamespaces.DrawingML, "latin").Attribute("typeface");

        internal static int? ToDrawingFontSize(double? points, string parameterName) {
            if (!points.HasValue) return null;
            if (double.IsNaN(points.Value) || double.IsInfinity(points.Value)
                || points.Value < 1D || points.Value > 4000D) {
                throw new ArgumentOutOfRangeException(parameterName, points,
                    "DrawingML font size must be a finite value from 1 through 4000 points.");
            }
            return checked((int)Math.Round(points.Value * 100D, MidpointRounding.AwayFromZero));
        }
    }
}
