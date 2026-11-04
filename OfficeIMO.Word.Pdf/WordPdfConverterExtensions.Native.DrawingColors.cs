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

        private static bool TryGetNativeDrawingColor(OpenXmlElement? owner, out OfficeColor color, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor>? themeColors = null) {
            color = default;
            if (owner == null) {
                return false;
            }

            A.SchemeColor? scheme = owner as A.SchemeColor ?? owner.GetFirstChild<A.SchemeColor>();
            if (scheme != null) {
                string? schemeName = GetNativeDrawingEnumAttributeValue(scheme, "val");
                if (scheme.Val?.Value != null && themeColors != null &&
                    themeColors.TryGetValue(scheme.Val.Value, out color) ||
                    TryGetNativeDefaultThemeColor(schemeName, out color)) {
                    color = ApplyNativeDrawingColorTransforms(color, scheme);
                    return true;
                }
                return false;
            }

            // Non-scheme decoding, including linear scRGB, belongs to the shared resolver.
            OfficeColor? resolved = OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ResolveColor(owner, colorScheme: null);
            if (!resolved.HasValue) {
                return false;
            }
            color = resolved.Value;
            return true;
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

        private static OfficeColor ApplyNativeDrawingColorTransforms(OfficeColor color, OpenXmlElement colorElement) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ApplyTransforms(color, colorElement);

        private static string? GetNativeDrawingEnumAttributeValue(OpenXmlElement element, string name) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName.Equals(name, StringComparison.OrdinalIgnoreCase)) {
                    return string.IsNullOrWhiteSpace(attribute.Value) ? null : attribute.Value;
                }
            }

            return null;
        }

    }
}
