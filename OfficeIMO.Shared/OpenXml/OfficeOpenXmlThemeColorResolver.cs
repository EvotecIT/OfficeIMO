using System;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Shared.OpenXml;

/// <summary>
/// Resolves DrawingML theme colors and applies their ordered color transformations.
/// This source is shared by the Word, Excel, and PowerPoint OpenXML adapters.
/// </summary>
internal static class OfficeOpenXmlThemeColorResolver {
    private static readonly string[] DefaultSpreadsheetThemeColors = {
        "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D",
        "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"
    };

    internal static OfficeColor? ResolveColor(
        OpenXmlElement? container,
        A.ColorScheme? colorScheme,
        A.SchemeColor? placeholderColor = null) {
        OpenXmlElement? colorElement = FindColorElement(container);
        return ResolveColorElement(colorElement, colorScheme, placeholderColor);
    }

    internal static OfficeColor? ResolveSchemeColor(A.ColorScheme? colorScheme, string? scheme) {
        if (colorScheme == null || string.IsNullOrWhiteSpace(scheme)) {
            return null;
        }

        string normalized = scheme!.Trim().ToLowerInvariant();
        OpenXmlCompositeElement? colorElement = normalized switch {
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
            "hyperlink" or "hlink" => colorScheme.GetFirstChild<A.Hyperlink>(),
            "followedhyperlink" or "folhlink" => colorScheme.GetFirstChild<A.FollowedHyperlinkColor>(),
            _ => null
        };

        return ResolveThemeEntry(colorElement);
    }

    internal static OfficeColor? ResolveSpreadsheetThemeColor(A.ColorScheme? colorScheme, uint themeIndex) {
        string? scheme = themeIndex switch {
            0U => "light1",
            1U => "dark1",
            2U => "light2",
            3U => "dark2",
            4U => "accent1",
            5U => "accent2",
            6U => "accent3",
            7U => "accent4",
            8U => "accent5",
            9U => "accent6",
            10U => "hyperlink",
            11U => "followedHyperlink",
            _ => null
        };
        OfficeColor? resolved = ResolveSchemeColor(colorScheme, scheme);
        if (resolved.HasValue || themeIndex >= DefaultSpreadsheetThemeColors.Length) {
            return resolved;
        }

        return OfficeColor.TryParseHex(DefaultSpreadsheetThemeColors[themeIndex], out OfficeColor fallback)
            ? fallback
            : (OfficeColor?)null;
    }

    private static OfficeColor? ResolveColorElement(
        OpenXmlElement? colorElement,
        A.ColorScheme? colorScheme,
        A.SchemeColor? placeholderColor) {
        if (colorElement == null) {
            return null;
        }

        OfficeColor? color;
        if (colorElement is A.RgbColorModelHex rgbColor) {
            color = ParseRgb(rgbColor.Val?.Value);
        } else if (colorElement is A.SystemColor systemColor) {
            color = ParseRgb(systemColor.LastColor?.Value);
        } else if (colorElement is A.SchemeColor schemeColor) {
            string? scheme = GetSchemeValue(schemeColor);
            if (IsPlaceholderScheme(scheme)) {
                color = ResolveColorElement(placeholderColor, colorScheme, null);
            } else {
                color = ResolveSchemeColor(colorScheme, scheme);
            }
        } else if (colorElement is A.PresetColor presetColor) {
            color = OfficeColor.TryParse(presetColor.Val?.Value.ToString(), out OfficeColor preset)
                ? preset
                : (OfficeColor?)null;
        } else {
            color = null;
        }

        return color.HasValue ? ApplyTransforms(color.Value, colorElement) : null;
    }

    private static OfficeColor ApplyTransforms(OfficeColor color, OpenXmlElement colorElement) {
        OfficeColor resolved = color;
        foreach (OpenXmlElement transform in colorElement.ChildElements) {
            switch (transform.LocalName) {
                case "comp":
                    resolved = OfficeColorTransforms.Complement(resolved);
                    continue;
                case "inv":
                    resolved = OfficeColor.FromRgba(
                        (byte)(255 - resolved.R),
                        (byte)(255 - resolved.G),
                        (byte)(255 - resolved.B),
                        resolved.A);
                    continue;
                case "gray":
                    byte gray = ToChannel((resolved.R * 0.299D) + (resolved.G * 0.587D) + (resolved.B * 0.114D));
                    resolved = OfficeColor.FromRgba(gray, gray, gray, resolved.A);
                    continue;
            }

            if (!TryReadTransformValue(transform, out double value)) {
                continue;
            }

            switch (transform.LocalName) {
                case "alpha":
                    resolved = OfficeColorTransforms.WithAlpha(resolved, ClampUnit(value));
                    break;
                case "alphaMod":
                    resolved = OfficeColorTransforms.ModulateAlpha(resolved, Math.Max(0D, value));
                    break;
                case "alphaOff":
                    resolved = OfficeColorTransforms.OffsetAlpha(resolved, value);
                    break;
                case "tint":
                    resolved = OfficeColorTransforms.Tint(resolved, ClampUnit(value));
                    break;
                case "shade":
                    resolved = OfficeColorTransforms.Shade(resolved, ClampUnit(value));
                    break;
                case "lumMod":
                    resolved = OfficeColorTransforms.ModulateLuminance(resolved, Math.Max(0D, value));
                    break;
                case "lumOff":
                    resolved = OfficeColorTransforms.OffsetLuminance(resolved, value);
                    break;
                case "red":
                    resolved = OfficeColor.FromRgba(ToChannel(255D * value), resolved.G, resolved.B, resolved.A);
                    break;
                case "redMod":
                    resolved = OfficeColor.FromRgba(ToChannel(resolved.R * value), resolved.G, resolved.B, resolved.A);
                    break;
                case "redOff":
                    resolved = OfficeColor.FromRgba(ToChannel(resolved.R + (255D * value)), resolved.G, resolved.B, resolved.A);
                    break;
                case "green":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(255D * value), resolved.B, resolved.A);
                    break;
                case "greenMod":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G * value), resolved.B, resolved.A);
                    break;
                case "greenOff":
                    resolved = OfficeColor.FromRgba(resolved.R, ToChannel(resolved.G + (255D * value)), resolved.B, resolved.A);
                    break;
                case "blue":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(255D * value), resolved.A);
                    break;
                case "blueMod":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B * value), resolved.A);
                    break;
                case "blueOff":
                    resolved = OfficeColor.FromRgba(resolved.R, resolved.G, ToChannel(resolved.B + (255D * value)), resolved.A);
                    break;
            }
        }

        return resolved;
    }

    private static OpenXmlElement? FindColorElement(OpenXmlElement? container) {
        if (container == null) {
            return null;
        }

        if (container is A.RgbColorModelHex or A.SystemColor or A.SchemeColor or A.PresetColor) {
            return container;
        }

        return container.GetFirstChild<A.RgbColorModelHex>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.SchemeColor>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.SystemColor>()
            ?? container.GetFirstChild<A.PresetColor>();
    }

    private static OfficeColor? ResolveThemeEntry(OpenXmlCompositeElement? colorElement) {
        if (colorElement == null) {
            return null;
        }

        return ParseRgb(colorElement.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value)
            ?? ParseRgb(colorElement.GetFirstChild<A.SystemColor>()?.LastColor?.Value);
    }

    private static OfficeColor? ParseRgb(string? value) =>
        OfficeColor.TryParseHex(value, out OfficeColor color) ? color : (OfficeColor?)null;

    private static string? GetSchemeValue(A.SchemeColor? schemeColor) {
        string? attribute = schemeColor?.GetAttribute("val", string.Empty).Value;
        return !string.IsNullOrWhiteSpace(attribute)
            ? attribute
            : schemeColor?.Val?.Value.ToString();
    }

    private static bool IsPlaceholderScheme(string? scheme) =>
        string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadTransformValue(OpenXmlElement transform, out double value) {
        value = 0D;
        string? raw = transform.GetAttribute("val", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(raw) || !int.TryParse(raw, out int scaled)) {
            return false;
        }

        value = scaled / 100000D;
        return true;
    }

    private static byte ToChannel(double value) =>
        (byte)Math.Max(0D, Math.Min(255D, Math.Round(value, MidpointRounding.ToEven)));

    private static double ClampUnit(double value) => Math.Max(0D, Math.Min(1D, value));
}
