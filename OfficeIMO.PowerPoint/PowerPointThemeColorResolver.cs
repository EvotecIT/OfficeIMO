using System;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static class PowerPointThemeColorResolver {
        internal static string? ResolveSolidFillColor(A.SolidFill? solidFill, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            if (solidFill == null) {
                return null;
            }

            string? rgbColor = solidFill.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbColor)) {
                return ApplyColorTransforms(rgbColor, solidFill.RgbColorModelHex);
            }

            A.SchemeColor? schemeColor = solidFill.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                string? placeholderScheme = GetSchemeColorValue(placeholderColor);
                string? placeholderResolvedColor = ResolveSchemeColor(colorScheme, placeholderScheme);
                placeholderResolvedColor = ApplyColorTransforms(placeholderResolvedColor, placeholderColor);
                return ApplyColorTransforms(placeholderResolvedColor, schemeColor);
            }

            return ApplyColorTransforms(ResolveSchemeColor(colorScheme, scheme), schemeColor);
        }

        internal static OfficeColor? ResolveSolidFillOfficeColor(A.SolidFill? solidFill, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            string? colorValue = ResolveSolidFillColor(solidFill, colorScheme, placeholderColor);
            if (!OfficeColor.TryParseHex(colorValue, out OfficeColor color)) {
                return null;
            }

            double alpha = ResolveSolidFillAlpha(solidFill, placeholderColor);
            byte alphaByte = (byte)Math.Round(255D * alpha);
            return OfficeColor.FromRgba(color.R, color.G, color.B, alphaByte);
        }

        internal static string? ResolveGradientStopColor(A.GradientStop stop, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            A.RgbColorModelHex? rgbColor = stop.GetFirstChild<A.RgbColorModelHex>();
            string? rgbValue = rgbColor?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbValue)) {
                return ApplyColorTransforms(rgbValue, rgbColor);
            }

            A.SchemeColor? schemeColor = stop.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                string? placeholderScheme = GetSchemeColorValue(placeholderColor);
                string? placeholderResolvedColor = ResolveSchemeColor(colorScheme, placeholderScheme);
                placeholderResolvedColor = ApplyColorTransforms(placeholderResolvedColor, placeholderColor);
                return ApplyColorTransforms(placeholderResolvedColor, schemeColor);
            }

            return ApplyColorTransforms(ResolveSchemeColor(colorScheme, scheme), schemeColor);
        }

        internal static OfficeColor? ResolveGradientStopOfficeColor(A.GradientStop stop, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            string? colorValue = ResolveGradientStopColor(stop, colorScheme, placeholderColor);
            if (!OfficeColor.TryParseHex(colorValue, out OfficeColor color)) {
                return null;
            }

            double alpha = ResolveGradientStopAlpha(stop, placeholderColor);
            byte alphaByte = (byte)Math.Round(255D * alpha);
            return OfficeColor.FromRgba(color.R, color.G, color.B, alphaByte);
        }

        internal static string? ResolveHighlightColor(A.Highlight? highlight, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            if (highlight == null) {
                return null;
            }

            A.RgbColorModelHex? rgbColor = highlight.GetFirstChild<A.RgbColorModelHex>();
            string? rgbValue = rgbColor?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbValue)) {
                return ApplyColorTransforms(rgbValue, rgbColor);
            }

            A.SchemeColor? schemeColor = highlight.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                string? placeholderScheme = GetSchemeColorValue(placeholderColor);
                string? placeholderResolvedColor = ResolveSchemeColor(colorScheme, placeholderScheme);
                placeholderResolvedColor = ApplyColorTransforms(placeholderResolvedColor, placeholderColor);
                return ApplyColorTransforms(placeholderResolvedColor, schemeColor);
            }

            string? resolvedSchemeColor = ApplyColorTransforms(ResolveSchemeColor(colorScheme, scheme), schemeColor);
            if (!string.IsNullOrWhiteSpace(resolvedSchemeColor)) {
                return resolvedSchemeColor;
            }

            A.SystemColor? systemColor = highlight.GetFirstChild<A.SystemColor>();
            return ApplyColorTransforms(systemColor?.LastColor?.Value, systemColor);
        }

        internal static OfficeColor? ResolveHighlightOfficeColor(A.Highlight? highlight, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor = null) {
            string? colorValue = ResolveHighlightColor(highlight, colorScheme, placeholderColor);
            if (!OfficeColor.TryParseHex(colorValue, out OfficeColor color)) {
                return null;
            }

            double alpha = ResolveHighlightAlpha(highlight, placeholderColor);
            byte alphaByte = (byte)Math.Round(255D * alpha);
            return OfficeColor.FromRgba(color.R, color.G, color.B, alphaByte);
        }

        private static string? GetSchemeColorValue(A.SchemeColor? schemeColor) {
            string? attribute = schemeColor?.GetAttribute("val", string.Empty).Value;
            return !string.IsNullOrWhiteSpace(attribute)
                ? attribute
                : schemeColor?.Val?.Value.ToString();
        }

        private static bool IsPlaceholderSchemeColor(string? scheme) {
            return string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveSchemeColor(A.ColorScheme? colorScheme, string? scheme) {
            if (colorScheme == null || string.IsNullOrWhiteSpace(scheme)) {
                return null;
            }

            OpenXmlCompositeElement? colorElement = scheme switch {
                "Dark1" or "dk1" or "Text1" or "tx1" => colorScheme.GetFirstChild<A.Dark1Color>(),
                "Light1" or "lt1" or "Background1" or "bg1" => colorScheme.GetFirstChild<A.Light1Color>(),
                "Dark2" or "dk2" or "Text2" or "tx2" => colorScheme.GetFirstChild<A.Dark2Color>(),
                "Light2" or "lt2" or "Background2" or "bg2" => colorScheme.GetFirstChild<A.Light2Color>(),
                "Accent1" or "accent1" => colorScheme.GetFirstChild<A.Accent1Color>(),
                "Accent2" or "accent2" => colorScheme.GetFirstChild<A.Accent2Color>(),
                "Accent3" or "accent3" => colorScheme.GetFirstChild<A.Accent3Color>(),
                "Accent4" or "accent4" => colorScheme.GetFirstChild<A.Accent4Color>(),
                "Accent5" or "accent5" => colorScheme.GetFirstChild<A.Accent5Color>(),
                "Accent6" or "accent6" => colorScheme.GetFirstChild<A.Accent6Color>(),
                "Hyperlink" or "hlink" => colorScheme.GetFirstChild<A.Hyperlink>(),
                "FollowedHyperlink" or "folHlink" => colorScheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };

            return colorElement?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                ?? colorElement?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
        }

        private static string? ApplyColorTransforms(string? hexColor, OpenXmlElement? colorElement) {
            if (string.IsNullOrWhiteSpace(hexColor) || colorElement == null) {
                return hexColor;
            }

            string color = hexColor!;
            if (color.Length != 6) {
                return hexColor;
            }

            if (!TryParseHexColor(color, out int red, out int green, out int blue)) {
                return hexColor;
            }

            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                int? rawValue = GetTransformValue(transform);
                if (!rawValue.HasValue) {
                    continue;
                }

                double amount = Math.Max(0D, Math.Min(100000D, rawValue.Value)) / 100000D;
                switch (transform.LocalName) {
                    case "lumMod":
                        red = ClampColor(red * amount);
                        green = ClampColor(green * amount);
                        blue = ClampColor(blue * amount);
                        break;
                    case "lumOff":
                        red = ClampColor(red + 255D * amount);
                        green = ClampColor(green + 255D * amount);
                        blue = ClampColor(blue + 255D * amount);
                        break;
                    case "tint":
                        red = ClampColor(red + (255D - red) * amount);
                        green = ClampColor(green + (255D - green) * amount);
                        blue = ClampColor(blue + (255D - blue) * amount);
                        break;
                    case "shade":
                        red = ClampColor(red * amount);
                        green = ClampColor(green * amount);
                        blue = ClampColor(blue * amount);
                        break;
                }
            }

            return red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }

        private static double ResolveSolidFillAlpha(A.SolidFill? solidFill, A.SchemeColor? placeholderColor) {
            if (solidFill == null) {
                return 1D;
            }

            double alpha = 1D;
            if (solidFill.RgbColorModelHex != null) {
                return ApplyAlphaTransforms(alpha, solidFill.RgbColorModelHex);
            }

            A.SchemeColor? schemeColor = solidFill.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                alpha = ApplyAlphaTransforms(alpha, placeholderColor);
            }

            return ApplyAlphaTransforms(alpha, schemeColor);
        }

        private static double ResolveGradientStopAlpha(A.GradientStop stop, A.SchemeColor? placeholderColor) {
            double alpha = 1D;
            A.RgbColorModelHex? rgbColor = stop.GetFirstChild<A.RgbColorModelHex>();
            if (rgbColor != null) {
                return ApplyAlphaTransforms(alpha, rgbColor);
            }

            A.SchemeColor? schemeColor = stop.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                alpha = ApplyAlphaTransforms(alpha, placeholderColor);
            }

            return ApplyAlphaTransforms(alpha, schemeColor);
        }

        private static double ResolveHighlightAlpha(A.Highlight? highlight, A.SchemeColor? placeholderColor) {
            if (highlight == null) {
                return 1D;
            }

            double alpha = 1D;
            A.RgbColorModelHex? rgbColor = highlight.GetFirstChild<A.RgbColorModelHex>();
            if (rgbColor != null) {
                return ApplyAlphaTransforms(alpha, rgbColor);
            }

            A.SchemeColor? schemeColor = highlight.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                alpha = ApplyAlphaTransforms(alpha, placeholderColor);
            }

            A.SystemColor? systemColor = highlight.GetFirstChild<A.SystemColor>();
            if (systemColor != null && schemeColor == null) {
                return ApplyAlphaTransforms(alpha, systemColor);
            }

            return ApplyAlphaTransforms(alpha, schemeColor);
        }

        private static double ApplyAlphaTransforms(double alpha, OpenXmlElement? colorElement) {
            if (colorElement == null) {
                return alpha;
            }

            double resolved = alpha;
            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                int? rawValue = GetTransformValue(transform);
                if (!rawValue.HasValue) {
                    continue;
                }

                double amount = Math.Max(0D, Math.Min(100000D, rawValue.Value)) / 100000D;
                switch (transform.LocalName) {
                    case "alpha":
                        resolved = amount;
                        break;
                    case "alphaMod":
                        resolved *= amount;
                        break;
                    case "alphaOff":
                        resolved += amount;
                        break;
                }
            }

            return Math.Max(0D, Math.Min(1D, resolved));
        }

        private static int? GetTransformValue(OpenXmlElement transform) {
            string? value = transform.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "val", StringComparison.Ordinal))
                .Value;
            return int.TryParse(value, out int result) ? result : null;
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

        private static int ClampColor(double value) {
            return (int)Math.Max(0D, Math.Min(255D, Math.Round(value)));
        }
    }
}
