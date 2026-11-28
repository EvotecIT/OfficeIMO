using OfficeIMO.Drawing;
using static OfficeIMO.Word.OpenXml.WordXml;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryGetNativeDrawingSolidFillColor(OfficeXmlNode owner, out OfficeColor color, IReadOnlyDictionary<string, OfficeColor>? themeColors = null) {
            color = default;
            OfficeXmlNode solidFill = owner.IsNone ? default : owner.Child(A, "solidFill");
            return TryGetNativeDrawingColor(solidFill, out color, themeColors);
        }

        private static bool TryGetNativeDrawingOutlineColor(OfficeXmlNode owner, out OfficeColor color, IReadOnlyDictionary<string, OfficeColor>? themeColors = null) {
            color = default;
            OfficeXmlNode outline = owner.IsNone ? default : owner.Child(A, "ln");
            if (outline.IsNone || outline.Child(A, "noFill").Exists) {
                return false;
            }

            return TryGetNativeDrawingSolidFillColor(outline, out color, themeColors);
        }

        private static bool HasNativeDrawingNoFill(OfficeXmlNode owner) => owner.Exists && owner.Child(A, "noFill").Exists;

        private static bool HasNativeDrawingOutlineNoFill(OfficeXmlNode owner) =>
            owner.Exists && owner.Child(A, "ln") is { IsNone: false } outline && outline.Child(A, "noFill").Exists;

        private static bool TryGetNativeDrawingGradientFill(OfficeXmlNode owner, out OfficeLinearGradient? gradient, IReadOnlyDictionary<string, OfficeColor>? themeColors = null) {
            gradient = null;
            OfficeXmlNode gradientFill = owner.IsNone ? default : owner.Child(A, "gradFill");
            OfficeXmlNode stopList = gradientFill.IsNone ? default : gradientFill.Child(A, "gsLst");
            if (gradientFill.IsNone || stopList.IsNone) {
                return false;
            }

            var stops = new List<(double Offset, OfficeColor Color)>();
            foreach (OfficeXmlNode stop in stopList.Children(A, "gs")) {
                if (!TryGetNativeDrawingColor(stop, out OfficeColor color, themeColors)) {
                    continue;
                }

                double offset = Math.Max(0D, Math.Min(100000D, ParseInt32(stop.Attr("pos")) ?? 0)) / 100000D;
                stops.Add((offset, color));
            }

            if (stops.Count < 2) {
                return false;
            }

            stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            OfficeColor startColor = stops[0].Color;
            OfficeColor endColor = stops[stops.Count - 1].Color;
            gradient = CreateNativeDrawingLinearGradient(gradientFill.Child(A, "lin"), startColor, endColor);
            return true;
        }

        private static OfficeLinearGradient CreateNativeDrawingLinearGradient(OfficeXmlNode linear, OfficeColor startColor, OfficeColor endColor) {
            double angle = NormalizeNativeDrawingAngleDegrees(ParseInt32(linear.Attr("ang")));
            if (IsNativeDrawingAngleNear(angle, 0D) || IsNativeDrawingAngleNear(angle, 180D)) {
                return OfficeLinearGradient.Horizontal(startColor, endColor);
            }

            return IsNativeDrawingAngleNear(angle, 90D) || IsNativeDrawingAngleNear(angle, 270D)
                ? OfficeLinearGradient.Vertical(startColor, endColor)
                : OfficeLinearGradient.DiagonalDown(startColor, endColor);
        }

        private static double NormalizeNativeDrawingAngleDegrees(int? angle) {
            if (!angle.HasValue) {
                return 0D;
            }

            double degrees = angle.Value / 60000D;
            degrees %= 360D;
            return degrees < 0D ? degrees + 360D : degrees;
        }

        private static bool IsNativeDrawingAngleNear(double angle, double target) {
            double distance = Math.Abs(angle - target);
            distance = Math.Min(distance, 360D - distance);
            return distance <= 22.5D;
        }

        private static bool TryGetNativeDrawingColor(OfficeXmlNode owner, out OfficeColor color, IReadOnlyDictionary<string, OfficeColor>? themeColors = null) {
            color = default;
            if (owner.IsNone) {
                return false;
            }

            OfficeXmlNode scheme = owner.Is(A, "schemeClr") ? owner : owner.Child(A, "schemeClr");
            if (scheme.Exists) {
                string? schemeName = GetNativeDrawingEnumAttributeValue(scheme, "val");
                if (scheme.Attr("val") is string schemeToken && themeColors != null &&
                    themeColors.TryGetValue(schemeToken, out color) ||
                    TryGetNativeDefaultThemeColor(schemeName, out color)) {
                    color = ApplyNativeDrawingColorTransforms(color, scheme);
                    return true;
                }
                return false;
            }

            // Non-scheme decoding, including linear scRGB, belongs to the shared resolver.
            OfficeColor? resolved = OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ResolveColor(owner, colorScheme: default);
            if (!resolved.HasValue) {
                return false;
            }
            color = resolved.Value;
            return true;
        }

        /// <summary>
        /// The theme colours of the document that holds <paramref name="sourceElement"/>, keyed by the
        /// <c>a:schemeClr</c> token (<c>"accent1"</c>, <c>"tx1"</c>, ...).
        /// </summary>
        private static IReadOnlyDictionary<string, OfficeColor> GetNativeDrawingThemeColors(OfficeXmlNode sourceElement) {
            var colors = new Dictionary<string, OfficeColor>(StringComparer.Ordinal);
            OfficeXmlNode theme = sourceElement.IsNone ? default : sourceElement.PackageThemeNode();
            OfficeXmlNode themeElements = theme.IsNone ? default : theme.Child(A, "themeElements");
            OfficeXmlNode colorScheme = themeElements.IsNone ? default : themeElements.Child(A, "clrScheme");
            if (colorScheme.IsNone) {
                return colors;
            }

            AddNativeDrawingThemeColor(colors, "dk1", colorScheme.Child(A, "dk1"));
            AddNativeDrawingThemeColor(colors, "lt1", colorScheme.Child(A, "lt1"));
            AddNativeDrawingThemeColor(colors, "dk2", colorScheme.Child(A, "dk2"));
            AddNativeDrawingThemeColor(colors, "lt2", colorScheme.Child(A, "lt2"));
            AddNativeDrawingThemeColor(colors, "accent1", colorScheme.Child(A, "accent1"));
            AddNativeDrawingThemeColor(colors, "accent2", colorScheme.Child(A, "accent2"));
            AddNativeDrawingThemeColor(colors, "accent3", colorScheme.Child(A, "accent3"));
            AddNativeDrawingThemeColor(colors, "accent4", colorScheme.Child(A, "accent4"));
            AddNativeDrawingThemeColor(colors, "accent5", colorScheme.Child(A, "accent5"));
            AddNativeDrawingThemeColor(colors, "accent6", colorScheme.Child(A, "accent6"));
            AddNativeDrawingThemeColor(colors, "hlink", colorScheme.Child(A, "hlink"));
            AddNativeDrawingThemeColor(colors, "folHlink", colorScheme.Child(A, "folHlink"));
            AddNativeDrawingThemeAlias(colors, "bg1", "lt1");
            AddNativeDrawingThemeAlias(colors, "tx1", "dk1");
            AddNativeDrawingThemeAlias(colors, "bg2", "lt2");
            AddNativeDrawingThemeAlias(colors, "tx2", "dk2");
            return colors;
        }

        private static void AddNativeDrawingThemeColor(Dictionary<string, OfficeColor> colors, string key, OfficeXmlNode element) {
            if (TryGetNativeDrawingColor(element, out OfficeColor color, colors)) {
                colors[key] = color;
            }
        }

        private static void AddNativeDrawingThemeAlias(Dictionary<string, OfficeColor> colors, string alias, string target) {
            if (!colors.ContainsKey(alias) && colors.TryGetValue(target, out OfficeColor color)) {
                colors[alias] = color;
            }
        }

        private static bool TryGetNativeDefaultThemeColor(string? schemeName, out OfficeColor color) {
            color = default;
            if (string.IsNullOrWhiteSpace(schemeName)) {
                return false;
            }

            switch (schemeName!.Trim().ToLowerInvariant()) {
                case "dark1":
                case "dk1":
                case "text1":
                    color = OfficeColor.Black;
                    return true;
                case "light1":
                case "lt1":
                case "background1":
                    color = OfficeColor.White;
                    return true;
                case "dark2":
                case "dk2":
                case "text2":
                    color = OfficeColor.ParseHex("#44546A");
                    return true;
                case "light2":
                case "lt2":
                case "background2":
                    color = OfficeColor.ParseHex("#E7E6E6");
                    return true;
                case "accent1":
                    color = OfficeColor.ParseHex("#4472C4");
                    return true;
                case "accent2":
                    color = OfficeColor.ParseHex("#ED7D31");
                    return true;
                case "accent3":
                    color = OfficeColor.ParseHex("#A5A5A5");
                    return true;
                case "accent4":
                    color = OfficeColor.ParseHex("#FFC000");
                    return true;
                case "accent5":
                    color = OfficeColor.ParseHex("#5B9BD5");
                    return true;
                case "accent6":
                    color = OfficeColor.ParseHex("#70AD47");
                    return true;
                case "hyperlink":
                    color = OfficeColor.ParseHex("#0563C1");
                    return true;
                case "followedhyperlink":
                    color = OfficeColor.ParseHex("#954F72");
                    return true;
                default:
                    return false;
            }
        }

        private static OfficeColor ApplyNativeDrawingColorTransforms(OfficeColor color, OfficeXmlNode colorElement) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ApplyTransforms(color, colorElement);

        private static string? GetNativeDrawingEnumAttributeValue(OfficeXmlNode element, string name) {
            foreach (OfficeXmlAttribute attribute in element.Attributes()) {
                if (attribute.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

    }
}
