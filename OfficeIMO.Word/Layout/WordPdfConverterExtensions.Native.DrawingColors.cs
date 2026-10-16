using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryGetNativeDrawingSolidFillColor(OpenXmlElement? owner, out OfficeColor color, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor>? themeColors = null) {
            color = default;
            A.SolidFill? solidFill = owner?.GetFirstChild<A.SolidFill>();
            return TryGetNativeDrawingColor(solidFill, out color, themeColors);
        }

        private static bool TryGetNativeDrawingOutlineColor(OpenXmlElement? owner, out OfficeColor color, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor>? themeColors = null) {
            color = default;
            A.Outline? outline = owner?.GetFirstChild<A.Outline>();
            if (outline == null || outline.GetFirstChild<A.NoFill>() != null) {
                return false;
            }

            return TryGetNativeDrawingSolidFillColor(outline, out color, themeColors);
        }

        private static bool HasNativeDrawingNoFill(OpenXmlElement? owner) => owner?.GetFirstChild<A.NoFill>() != null;

        private static bool HasNativeDrawingOutlineNoFill(OpenXmlElement? owner) =>
            owner?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() != null;

        private static bool TryGetNativeDrawingGradientFill(OpenXmlElement? owner, out OfficeLinearGradient? gradient, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor>? themeColors = null) {
            gradient = null;
            A.GradientFill? gradientFill = owner?.GetFirstChild<A.GradientFill>();
            A.GradientStopList? stopList = gradientFill?.GetFirstChild<A.GradientStopList>();
            if (gradientFill == null || stopList == null) {
                return false;
            }

            var stops = new List<(double Offset, OfficeColor Color)>();
            foreach (A.GradientStop stop in stopList.Elements<A.GradientStop>()) {
                if (!TryGetNativeDrawingColor(stop, out OfficeColor color, themeColors)) {
                    continue;
                }

                double offset = Math.Max(0D, Math.Min(100000D, stop.Position?.Value ?? 0)) / 100000D;
                stops.Add((offset, color));
            }

            if (stops.Count < 2) {
                return false;
            }

            stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            OfficeColor startColor = stops[0].Color;
            OfficeColor endColor = stops[stops.Count - 1].Color;
            gradient = CreateNativeDrawingLinearGradient(gradientFill.GetFirstChild<A.LinearGradientFill>(), startColor, endColor);
            return true;
        }

        private static OfficeLinearGradient CreateNativeDrawingLinearGradient(A.LinearGradientFill? linear, OfficeColor startColor, OfficeColor endColor) {
            double angle = NormalizeNativeDrawingAngleDegrees(linear?.Angle?.Value);
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

        private static bool TryGetNativeDrawingFillOpacity(OpenXmlElement? owner, out double opacity) {
            A.SolidFill? solidFill = owner?.GetFirstChild<A.SolidFill>();
            if (TryGetNativeDrawingOpacity(solidFill, out opacity)) {
                return true;
            }

            A.GradientFill? gradientFill = owner?.GetFirstChild<A.GradientFill>();
            if (TryGetNativeDrawingOpacity(gradientFill, out opacity)) {
                return true;
            }

            opacity = 1D;
            return false;
        }

        private static bool TryGetNativeDrawingOpacity(OpenXmlElement? owner, out double opacity) {
            opacity = 1D;
            if (owner == null) {
                return false;
            }

            foreach (OpenXmlElement colorElement in owner.Descendants()) {
                switch (colorElement.LocalName) {
                    case "srgbClr":
                    case "schemeClr":
                    case "scrgbClr":
                    case "sysClr":
                        if (TryGetNativeDrawingColorAlpha(colorElement, out opacity)) {
                            return true;
                        }
                        break;
                }
            }

            return false;
        }

        private static bool TryGetNativeDrawingColorAlpha(OpenXmlElement colorElement, out double opacity) {
            opacity = 1D;
            bool found = false;
            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                int? value = GetNativeDrawingColorTransformValue(transform);
                if (!value.HasValue) {
                    continue;
                }

                double amount = Math.Max(0D, Math.Min(100000D, value.Value)) / 100000D;
                switch (transform.LocalName) {
                    case "alpha":
                        opacity = amount;
                        found = true;
                        break;
                    case "alphaMod":
                    case "alphaModFix":
                        opacity *= amount;
                        found = true;
                        break;
                }
            }

            return found;
        }

        private static bool TryGetNativeDrawingColor(OpenXmlElement? owner, out OfficeColor color, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor>? themeColors = null) {
            color = default;
            if (owner == null) {
                return false;
            }

            A.RgbColorModelHex? rgb = owner.GetFirstChild<A.RgbColorModelHex>();
            if (rgb?.Val?.Value is string rgbValue && OfficeColor.TryParseHex(rgbValue, out color)) {
                color = ApplyNativeDrawingColorTransforms(color, rgb);
                return true;
            }

            A.SystemColor? system = owner.GetFirstChild<A.SystemColor>();
            if (system?.LastColor?.Value is string systemValue && OfficeColor.TryParseHex(systemValue, out color)) {
                color = ApplyNativeDrawingColorTransforms(color, system);
                return true;
            }

            A.SchemeColor? scheme = owner.GetFirstChild<A.SchemeColor>();
            string? schemeName = scheme != null ? GetNativeDrawingEnumAttributeValue(scheme, "val") : null;
            if (scheme?.Val?.Value != null &&
                themeColors != null &&
                themeColors.TryGetValue(scheme.Val.Value, out color)) {
                color = ApplyNativeDrawingColorTransforms(color, scheme);
                return true;
            }

            if (scheme != null && TryGetNativeDefaultThemeColor(schemeName, out color)) {
                color = ApplyNativeDrawingColorTransforms(color, scheme);
                return true;
            }

            A.RgbColorModelPercentage? percentage = owner.GetFirstChild<A.RgbColorModelPercentage>();
            if (percentage != null) {
                color = OfficeColor.FromRgb(
                    ConvertNativeDrawingPercentageToByte(percentage.RedPortion?.Value),
                    ConvertNativeDrawingPercentageToByte(percentage.GreenPortion?.Value),
                    ConvertNativeDrawingPercentageToByte(percentage.BluePortion?.Value));
                color = ApplyNativeDrawingColorTransforms(color, percentage);
                return true;
            }

            A.PresetColor? preset = owner.GetFirstChild<A.PresetColor>();
            string? presetName = preset != null ? GetNativeDrawingEnumAttributeValue(preset, "val") : null;
            if (preset != null && TryGetNativeDrawingPresetColor(presetName, out color)) {
                color = ApplyNativeDrawingColorTransforms(color, preset);
                return true;
            }

            return false;
        }

        private static IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> GetNativeDrawingThemeColors(OpenXmlPart? sourcePart) {
            var colors = new Dictionary<A.SchemeColorValues, OfficeColor>();
            if (!(sourcePart?.OpenXmlPackage is WordprocessingDocument wordDocument)) {
                return colors;
            }

            A.ColorScheme? colorScheme = wordDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (colorScheme == null) {
                return colors;
            }

            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Dark1, colorScheme.GetFirstChild<A.Dark1Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Light1, colorScheme.GetFirstChild<A.Light1Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Dark2, colorScheme.GetFirstChild<A.Dark2Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Light2, colorScheme.GetFirstChild<A.Light2Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent1, colorScheme.GetFirstChild<A.Accent1Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent2, colorScheme.GetFirstChild<A.Accent2Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent3, colorScheme.GetFirstChild<A.Accent3Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent4, colorScheme.GetFirstChild<A.Accent4Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent5, colorScheme.GetFirstChild<A.Accent5Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Accent6, colorScheme.GetFirstChild<A.Accent6Color>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.Hyperlink, colorScheme.GetFirstChild<A.Hyperlink>());
            AddNativeDrawingThemeColor(colors, A.SchemeColorValues.FollowedHyperlink, colorScheme.GetFirstChild<A.FollowedHyperlinkColor>());
            AddNativeDrawingThemeAlias(colors, A.SchemeColorValues.Background1, A.SchemeColorValues.Light1);
            AddNativeDrawingThemeAlias(colors, A.SchemeColorValues.Text1, A.SchemeColorValues.Dark1);
            AddNativeDrawingThemeAlias(colors, A.SchemeColorValues.Background2, A.SchemeColorValues.Light2);
            AddNativeDrawingThemeAlias(colors, A.SchemeColorValues.Text2, A.SchemeColorValues.Dark2);
            return colors;
        }

        private static void AddNativeDrawingThemeColor(Dictionary<A.SchemeColorValues, OfficeColor> colors, A.SchemeColorValues key, OpenXmlElement? element) {
            if (TryGetNativeDrawingColor(element, out OfficeColor color, colors)) {
                colors[key] = color;
            }
        }

        private static void AddNativeDrawingThemeAlias(Dictionary<A.SchemeColorValues, OfficeColor> colors, A.SchemeColorValues alias, A.SchemeColorValues target) {
            if (!colors.ContainsKey(alias) && colors.TryGetValue(target, out OfficeColor color)) {
                colors[alias] = color;
            }
        }

        private static bool TryGetNativeDrawingPresetColor(string? presetName, out OfficeColor color) {
            color = default;
            if (string.IsNullOrWhiteSpace(presetName)) {
                return false;
            }

            switch (presetName!.Trim().ToLowerInvariant()) {
                case "black":
                    color = OfficeColor.Black;
                    return true;
                case "white":
                    color = OfficeColor.White;
                    return true;
                case "red":
                    color = OfficeColor.FromRgb(255, 0, 0);
                    return true;
                case "green":
                    color = OfficeColor.FromRgb(0, 128, 0);
                    return true;
                case "blue":
                    color = OfficeColor.FromRgb(0, 0, 255);
                    return true;
                case "yellow":
                    color = OfficeColor.FromRgb(255, 255, 0);
                    return true;
                case "cyan":
                    color = OfficeColor.FromRgb(0, 255, 255);
                    return true;
                case "magenta":
                    color = OfficeColor.FromRgb(255, 0, 255);
                    return true;
                case "gray":
                case "grey":
                    color = OfficeColor.Gray;
                    return true;
                case "dkgray":
                case "darkgray":
                case "darkgrey":
                    color = OfficeColor.FromRgb(128, 128, 128);
                    return true;
                case "ltgray":
                case "lightgray":
                case "lightgrey":
                    color = OfficeColor.LightGray;
                    return true;
                case "orange":
                    color = OfficeColor.FromRgb(255, 165, 0);
                    return true;
                case "purple":
                    color = OfficeColor.FromRgb(128, 0, 128);
                    return true;
                case "brown":
                    color = OfficeColor.FromRgb(165, 42, 42);
                    return true;
                case "pink":
                    color = OfficeColor.FromRgb(255, 192, 203);
                    return true;
                default:
                    return OfficeColor.TryParse(presetName, out color);
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

        private static OfficeColor ApplyNativeDrawingColorTransforms(OfficeColor color, OpenXmlElement colorElement) {
            double red = color.R;
            double green = color.G;
            double blue = color.B;
            double alpha = color.A;

            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                int? value = GetNativeDrawingColorTransformValue(transform);
                if (!value.HasValue) {
                    continue;
                }

                double amount = Math.Max(0D, Math.Min(100000D, value.Value)) / 100000D;
                switch (transform.LocalName) {
                    case "tint":
                        red = red + (255D - red) * amount;
                        green = green + (255D - green) * amount;
                        blue = blue + (255D - blue) * amount;
                        break;
                    case "shade":
                    case "lumMod":
                        ApplyNativeDrawingLuminanceTransform(ref red, ref green, ref blue, amount, 0D);
                        break;
                    case "lumOff":
                        ApplyNativeDrawingLuminanceTransform(ref red, ref green, ref blue, 1D, amount);
                        break;
                    case "alpha":
                        alpha = 255D * amount;
                        break;
                    case "alphaMod":
                    case "alphaModFix":
                        alpha *= amount;
                        break;
                }
            }

            return OfficeColor.FromRgba(
                ClampNativeDrawingColorByte(red),
                ClampNativeDrawingColorByte(green),
                ClampNativeDrawingColorByte(blue),
                ClampNativeDrawingColorByte(alpha));
        }

        private static void ApplyNativeDrawingLuminanceTransform(ref double red, ref double green, ref double blue, double luminanceMultiplier, double luminanceOffset) {
            RgbToHsl(red, green, blue, out double hue, out double saturation, out double luminance);
            luminance = Math.Max(0D, Math.Min(1D, luminance * luminanceMultiplier + luminanceOffset));
            HslToRgb(hue, saturation, luminance, out red, out green, out blue);
        }

        private static void RgbToHsl(double red, double green, double blue, out double hue, out double saturation, out double luminance) {
            double r = Math.Max(0D, Math.Min(255D, red)) / 255D;
            double g = Math.Max(0D, Math.Min(255D, green)) / 255D;
            double b = Math.Max(0D, Math.Min(255D, blue)) / 255D;
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            luminance = (max + min) / 2D;

            if (Math.Abs(max - min) < double.Epsilon) {
                hue = 0D;
                saturation = 0D;
                return;
            }

            double delta = max - min;
            saturation = luminance > 0.5D
                ? delta / (2D - max - min)
                : delta / (max + min);

            if (Math.Abs(max - r) < double.Epsilon) {
                hue = (g - b) / delta + (g < b ? 6D : 0D);
            } else if (Math.Abs(max - g) < double.Epsilon) {
                hue = (b - r) / delta + 2D;
            } else {
                hue = (r - g) / delta + 4D;
            }

            hue /= 6D;
        }

        private static void HslToRgb(double hue, double saturation, double luminance, out double red, out double green, out double blue) {
            if (saturation <= 0D) {
                red = green = blue = luminance * 255D;
                return;
            }

            double q = luminance < 0.5D
                ? luminance * (1D + saturation)
                : luminance + saturation - luminance * saturation;
            double p = 2D * luminance - q;
            red = HueToRgb(p, q, hue + 1D / 3D) * 255D;
            green = HueToRgb(p, q, hue) * 255D;
            blue = HueToRgb(p, q, hue - 1D / 3D) * 255D;
        }

        private static double HueToRgb(double p, double q, double hue) {
            if (hue < 0D) {
                hue += 1D;
            } else if (hue > 1D) {
                hue -= 1D;
            }

            if (hue < 1D / 6D) {
                return p + (q - p) * 6D * hue;
            }

            if (hue < 1D / 2D) {
                return q;
            }

            return hue < 2D / 3D
                ? p + (q - p) * (2D / 3D - hue) * 6D
                : p;
        }

        private static int? GetNativeDrawingColorTransformValue(OpenXmlElement transform) {
            string? raw = GetNativeDrawingEnumAttributeValue(transform, "val") ?? GetNativeDrawingEnumAttributeValue(transform, "amt");
            return int.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int value)
                ? value
                : (int?)null;
        }

        private static string? GetNativeDrawingEnumAttributeValue(OpenXmlElement element, string name) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

        private static byte ConvertNativeDrawingPercentageToByte(int? value) {
            double normalized = Math.Max(0D, Math.Min(100000D, value ?? 0)) / 100000D;
            return ClampNativeDrawingColorByte(255D * normalized);
        }

        private static byte ClampNativeDrawingColorByte(double value) =>
            (byte)Math.Max(0, Math.Min(255, (int)Math.Round(value)));
    }
}
