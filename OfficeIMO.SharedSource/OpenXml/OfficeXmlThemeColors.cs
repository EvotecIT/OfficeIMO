using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// Resolves DrawingML colors to <see cref="OfficeColor"/> through <see cref="OfficeXmlNode"/>, so it reads the same
    /// from OfficeIMO's own trees and from SDK elements: literal colors, scheme colors through a color map and scheme,
    /// the placeholder color of a style reference, and the color transforms in child order.
    /// </summary>
    internal static class OfficeXmlThemeColors {
        private const int D = OfficeXmlNamespaces.DrawingML;

        /// <summary>
        /// The color of <paramref name="container"/> (a color element or an element holding one), or null when it has
        /// none or it can't be resolved.
        /// </summary>
        internal static OfficeColor? ResolveColor(OfficeXmlNode container, OfficeXmlNode colorScheme, OfficeXmlNode placeholderColor = default, OfficeXmlNode colorMap = default) =>
            ResolveColorElement(FindColorElement(container), colorScheme, FindColorElement(placeholderColor), colorMap);

        /// <summary>The scheme slot <paramref name="scheme"/> names once the color map has mapped it.</summary>
        internal static string MapSchemeColor(string scheme, OfficeXmlNode colorMap) {
            if (colorMap.IsNone) return scheme;
            string attribute = scheme.ToLowerInvariant() switch {
                "dark1" or "dk1" or "text1" or "tx1" => "tx1",
                "light1" or "lt1" or "background1" or "bg1" => "bg1",
                "dark2" or "dk2" or "text2" or "tx2" => "tx2",
                "light2" or "lt2" or "background2" or "bg2" => "bg2",
                "hyperlink" or "hlink" => "hlink",
                "followedhyperlink" or "folhlink" => "folHlink",
                _ => scheme
            };
            string? mapped = colorMap.Attribute(attribute);
            return string.IsNullOrWhiteSpace(mapped) ? scheme : mapped!;
        }

        /// <summary>The color of a scheme slot in <paramref name="colorScheme"/>.</summary>
        internal static OfficeColor? ResolveSchemeColor(OfficeXmlNode colorScheme, string? scheme) {
            if (colorScheme.IsNone || string.IsNullOrWhiteSpace(scheme)) return null;
            string? slot = scheme!.Trim().ToLowerInvariant() switch {
                "dark1" or "dk1" or "text1" or "tx1" => "dk1",
                "light1" or "lt1" or "background1" or "bg1" => "lt1",
                "dark2" or "dk2" or "text2" or "tx2" => "dk2",
                "light2" or "lt2" or "background2" or "bg2" => "lt2",
                "accent1" => "accent1",
                "accent2" => "accent2",
                "accent3" => "accent3",
                "accent4" => "accent4",
                "accent5" => "accent5",
                "accent6" => "accent6",
                "hyperlink" or "hlink" => "hlink",
                "followedhyperlink" or "folhlink" => "folHlink",
                _ => null
            };
            if (slot == null) return null;
            OfficeXmlNode entry = colorScheme.Child(D, slot);
            return entry.IsNone ? null : ResolveColorElement(FindColorElement(entry), default, default);
        }

        private static OfficeColor? ResolveColorElement(OfficeXmlNode colorElement, OfficeXmlNode colorScheme, OfficeXmlNode placeholderColor, OfficeXmlNode colorMap = default) {
            if (colorElement.IsNone) return null;
            OfficeColor? color;
            switch (colorElement.LocalName) {
                case "srgbClr":
                    color = OfficeDrawingColorTransforms.ParseRgb(colorElement.Attribute("val"));
                    break;
                case "scrgbClr":
                    color = OfficeDrawingColorTransforms.ParseScRgb(Percentage(colorElement, "r"), Percentage(colorElement, "g"), Percentage(colorElement, "b"));
                    break;
                case "hslClr":
                    color = OfficeDrawingColorTransforms.ParseHsl(colorElement.Int32Attribute("hue"), Percentage(colorElement, "sat"), Percentage(colorElement, "lum"));
                    break;
                case "sysClr":
                    color = OfficeDrawingColorTransforms.ParseRgb(colorElement.Attribute("lastClr"));
                    break;
                case "schemeClr": {
                    string? scheme = colorElement.Attribute("val");
                    color = IsPlaceholderScheme(scheme)
                        ? ResolveColorElement(placeholderColor, colorScheme, default, colorMap)
                        : ResolveSchemeColor(colorScheme, scheme == null ? null : MapSchemeColor(scheme, colorMap));
                    break;
                }
                case "prstClr":
                    color = OfficeDrawingColorTransforms.ParsePreset(colorElement.Attribute("val"));
                    break;
                default:
                    color = null;
                    break;
            }
            return color.HasValue ? ApplyTransforms(color.Value, colorElement) : null;
        }

        /// <summary>Applies DrawingML color transforms in child order.</summary>
        internal static OfficeColor ApplyTransforms(OfficeColor color, OfficeXmlNode colorElement) {
            var transforms = new OfficeDrawingColorTransforms(color);
            for (OfficeXmlNode transform = colorElement.FirstChild; !transform.IsNone; transform = transform.NextSibling) {
                string name = transform.LocalName;
                transforms.Apply(name, OfficeDrawingColorTransforms.TakesValue(name) ? transform.Attribute("val") : null);
            }
            return transforms.Result;
        }

        private static OfficeXmlNode FindColorElement(OfficeXmlNode container) {
            if (container.IsNone) return default;
            if (container.NamespaceId == D && container.LocalName is "srgbClr" or "scrgbClr" or "hslClr" or "sysClr" or "schemeClr" or "prstClr") return container;
            foreach (string name in new[] { "srgbClr", "scrgbClr", "hslClr", "schemeClr", "sysClr", "prstClr" }) {
                OfficeXmlNode color = container.Child(D, name);
                if (!color.IsNone) return color;
            }
            return default;
        }

        private static int? Percentage(OfficeXmlNode element, string localName) => OfficeDrawingColorTransforms.ParsePercentage(element.Attribute(localName));

        private static bool IsPlaceholderScheme(string? scheme) =>
            string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
            || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
            || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);
    }
}
