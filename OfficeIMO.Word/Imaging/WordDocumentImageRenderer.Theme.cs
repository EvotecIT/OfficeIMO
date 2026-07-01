using System;
using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
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
            string? resolvedColor = ResolveSchemeColor(colorScheme, themeColor);
            if (string.IsNullOrWhiteSpace(resolvedColor)) {
                return null;
            }

            return ApplyWordThemeTransforms(resolvedColor!, tint, shade);
        }

        private static string? ResolveSchemeColor(A.ColorScheme? colorScheme, string? themeColor) {
            if (colorScheme == null || string.IsNullOrWhiteSpace(themeColor)) {
                return null;
            }

            OpenXmlCompositeElement? colorElement = themeColor switch {
                "dark1" or "dk1" or "text1" or "tx1" => colorScheme.GetFirstChild<A.Dark1Color>(),
                "light1" or "lt1" or "background1" or "bg1" => colorScheme.GetFirstChild<A.Light1Color>(),
                "dark2" or "dk2" or "text2" or "tx2" => colorScheme.GetFirstChild<A.Dark2Color>(),
                "light2" or "lt2" or "background2" or "bg2" => colorScheme.GetFirstChild<A.Light2Color>(),
                "accent1" => colorScheme.GetFirstChild<A.Accent1Color>(),
                "accent2" => colorScheme.GetFirstChild<A.Accent2Color>(),
                "accent3" => colorScheme.GetFirstChild<A.Accent3Color>(),
                "accent4" => colorScheme.GetFirstChild<A.Accent4Color>(),
                "accent5" => colorScheme.GetFirstChild<A.Accent5Color>(),
                "accent6" => colorScheme.GetFirstChild<A.Accent6Color>(),
                "hyperlink" => colorScheme.GetFirstChild<A.Hyperlink>(),
                "followedHyperlink" or "followedhyperlink" => colorScheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };

            return colorElement?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                ?? colorElement?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
        }

        private static string ApplyWordThemeTransforms(string hexColor, string? tint, string? shade) {
            if (hexColor.Length != 6 || !TryParseHexColor(hexColor, out int red, out int green, out int blue)) {
                return hexColor;
            }

            if (TryParseHexByte(shade, out int shadeValue)) {
                double amount = Math.Max(0D, Math.Min(255D, shadeValue)) / 255D;
                red = ClampColor(red * amount);
                green = ClampColor(green * amount);
                blue = ClampColor(blue * amount);
            }

            if (TryParseHexByte(tint, out int tintValue)) {
                double amount = 1D - (Math.Max(0D, Math.Min(255D, tintValue)) / 255D);
                red = ClampColor(red + (255D - red) * amount);
                green = ClampColor(green + (255D - green) * amount);
                blue = ClampColor(blue + (255D - blue) * amount);
            }

            return red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
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

        private static bool TryParseOfficeColor(string? hexColor, out OfficeColor color) {
            color = OfficeColor.Black;
            if (string.IsNullOrWhiteSpace(hexColor)) {
                return false;
            }

            try {
                color = Helpers.ParseColor(hexColor!);
                return true;
            } catch (ArgumentException) {
                return false;
            } catch (FormatException) {
                return false;
            }
        }

        private static bool TryParseHexColor(string hexColor, out int red, out int green, out int blue) {
            red = 0;
            green = 0;
            blue = 0;
            if (hexColor.Length != 6) {
                return false;
            }

            try {
                red = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                green = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                blue = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        private static bool TryParseHexByte(string? value, out int result) {
            result = 0;
            return !string.IsNullOrWhiteSpace(value) && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
        }

        private static int ClampColor(double value) =>
            (int)Math.Max(0D, Math.Min(255D, Math.Round(value)));
    }
}
