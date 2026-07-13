using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Shared.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static A.ColorScheme? GetDocumentColorScheme(WordDocument document) =>
            document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;

        private static OfficeColor ResolveParagraphTextColor(WordParagraph? paragraph, A.ColorScheme? colorScheme, OfficeColor? fallback = null) {
            OfficeColor fallbackColor = fallback ?? OfficeColor.Black;
            if (paragraph == null) {
                return fallbackColor;
            }

            string? resolvedThemeColor = ResolveThemeColor(
                GetWordAttribute(paragraph._runProperties?.Color, "themeColor"),
                GetWordAttribute(paragraph._runProperties?.Color, "themeTint"),
                GetWordAttribute(paragraph._runProperties?.Color, "themeShade"),
                colorScheme);
            if (TryParseOfficeColor(resolvedThemeColor, out OfficeColor themeColor)) {
                return themeColor;
            }

            return paragraph.Color ?? fallbackColor;
        }

        private static string? ResolveThemeColor(string? themeColor, string? tint, string? shade, A.ColorScheme? colorScheme) {
            OfficeColor? resolvedColor = OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(colorScheme, themeColor);
            if (!resolvedColor.HasValue) {
                return null;
            }

            return ApplyWordThemeTransforms(resolvedColor.Value, tint, shade).ToRgbHex();
        }

        private static OfficeColor ApplyWordThemeTransforms(OfficeColor color, string? tint, string? shade) {
            OfficeColor transformed = color;
            if (TryParseHexByte(shade, out int shadeValue)) {
                double amount = Math.Max(0D, Math.Min(255D, shadeValue)) / 255D;
                transformed = OfficeColorTransforms.Shade(transformed, amount);
            }

            if (TryParseHexByte(tint, out int tintValue)) {
                double inputRatio = Math.Max(0D, Math.Min(255D, tintValue)) / 255D;
                transformed = OfficeColorTransforms.Tint(transformed, inputRatio);
            }

            return transformed;
        }

        private static string? GetWordAttribute(OpenXmlElement? element, string localName) {
            if (element == null) {
                return null;
            }

            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool TryParseOfficeColor(string? hexColor, out OfficeColor color) =>
            OfficeColor.TryParseHex(hexColor, out color);

        private static bool TryParseHexByte(string? value, out int result) {
            result = 0;
            return !string.IsNullOrWhiteSpace(value) && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
        }
    }
}
