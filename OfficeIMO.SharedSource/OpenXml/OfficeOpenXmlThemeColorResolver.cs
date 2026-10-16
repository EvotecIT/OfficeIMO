using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>
/// Resolves DrawingML theme colors and applies their ordered color transformations.
/// This source is shared by the Word, Excel, and PowerPoint OpenXML adapters.
/// </summary>
internal static class OfficeOpenXmlThemeColorResolver {
    private static readonly string[] DefaultSpreadsheetThemeColors = {
        "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D",
        "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080"
    };

    /// <summary>Reads the automatic per-point palette of a radial chart that varies colours by point.</summary>
    /// <remarks>Office derives the palette from the classic chart style and the theme; see
    /// <see cref="OfficeOpenXmlChartAutomaticColors"/>.</remarks>
    internal static OfficeColor[]? ReadRadialPalette(ChartPart chartPart, OpenXmlCompositeElement series,
        int pointCount, A.ColorScheme? scheme, bool requireQualifiedPalette = false,
        bool useSpreadsheetDefaultTheme = false) {
        if (pointCount < 1 ||
            series.Parent is not C.PieChart and not C.DoughnutChart and not C.Pie3DChart and not C.OfPieChart ||
            series.Parent.GetFirstChild<C.VaryColors>() is not C.VaryColors varyColors ||
            varyColors.Val?.Value == false)
            return null;
        // Point fills take precedence over the theme palette in every static renderer.
        // An unsupported inherited palette is immaterial when every rendered ring has
        // an explicit appearance for every category.
        if (AllRadialPointsHaveExplicitFill((OpenXmlCompositeElement)series.Parent, pointCount)) return null;
        OfficeColor[]? colors = scheme == null && !useSpreadsheetDefaultTheme ? null :
            OfficeOpenXmlChartAutomaticColors.ResolvePoints(chartPart.ChartSpace, pointCount, scheme,
                useSpreadsheetDefaultTheme);
        if (colors == null && requireQualifiedPalette)
            throw new NotSupportedException("The automatic radial palette has no resolvable theme colours.");
        return colors;
    }

    private static bool AllRadialPointsHaveExplicitFill(OpenXmlCompositeElement chart, int pointCount) {
        bool hasSeries = false;
        bool multipleRings = chart is C.DoughnutChart &&
            chart.Elements<C.PieChartSeries>().Skip(1).Any();
        foreach (C.PieChartSeries series in chart.Elements<C.PieChartSeries>()) {
            hasSeries = true;
            C.ChartShapeProperties? seriesAppearance = series.GetFirstChild<C.ChartShapeProperties>();
            if (multipleRings && (seriesAppearance?.GetFirstChild<A.SolidFill>() != null ||
                seriesAppearance?.GetFirstChild<A.NoFill>() != null ||
                seriesAppearance?.GetFirstChild<A.PatternFill>() != null))
                continue;
            var covered = new bool[pointCount];
            foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
                uint? index = point.Index?.Val?.Value;
                if (!index.HasValue || index.Value >= (uint)pointCount) continue;
                C.ChartShapeProperties? appearance = point.GetFirstChild<C.ChartShapeProperties>();
                if (appearance?.GetFirstChild<A.SolidFill>() != null ||
                    appearance?.GetFirstChild<A.NoFill>() != null ||
                    appearance?.GetFirstChild<A.PatternFill>() != null)
                    covered[(int)index.Value] = true;
            }
            if (covered.Any(value => !value)) return false;
        }
        return hasSeries;
    }

    internal static OpenXmlElement? ResolveChartColorMap(ChartPart chartPart) {
        if (chartPart.ChartSpace?.GetFirstChild<C.ColorMapOverride>() is C.ColorMapOverride chartOverride)
            return chartOverride;
        foreach (OpenXmlPart owner in chartPart.GetParentParts()) {
            if (owner is SlidePart slide)
                return slide.Slide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? slide.SlideLayoutPart?.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? (OpenXmlElement?)slide.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap;
            if (owner is SlideLayoutPart layout)
                return layout.SlideLayout?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? (OpenXmlElement?)layout.SlideMasterPart?.SlideMaster?.ColorMap;
            if (owner is SlideMasterPart master)
                return master.SlideMaster?.ColorMap;
            if (owner is NotesSlidePart notes)
                return notes.NotesSlide?.ColorMapOverride?.GetFirstChild<A.OverrideColorMapping>()
                    ?? (OpenXmlElement?)notes.NotesMasterPart?.NotesMaster?.ColorMap;
            if (owner is NotesMasterPart notesMaster)
                return notesMaster.NotesMaster?.ColorMap;
            if (owner is HandoutMasterPart handoutMaster)
                return handoutMaster.HandoutMaster?.ColorMap;
            if (owner is MainDocumentPart document)
                return document.DocumentSettingsPart?.Settings?
                    .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping>();
            foreach (OpenXmlPart parent in owner.GetParentParts())
                if (parent is MainDocumentPart parentDocument)
                    return parentDocument.DocumentSettingsPart?.Settings?
                        .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping>();
        }
        return null;
    }

    internal static string MapSchemeColor(string scheme, OpenXmlElement? colorMap) {
        if (colorMap == null) return scheme;
        bool wordMap = colorMap is DocumentFormat.OpenXml.Wordprocessing.ColorSchemeMapping;
        string attribute = scheme.ToLowerInvariant() switch {
            "dark1" or "dk1" or "text1" or "tx1" => wordMap ? "t1" : "tx1",
            "light1" or "lt1" or "background1" or "bg1" => "bg1",
            "dark2" or "dk2" or "text2" or "tx2" => wordMap ? "t2" : "tx2",
            "light2" or "lt2" or "background2" or "bg2" => "bg2",
            "hyperlink" or "hlink" => "hlink",
            "followedhyperlink" or "folhlink" => "folHlink",
            _ => scheme
        };
        string? mapped = colorMap.GetAttributes()
            .FirstOrDefault(item => item.LocalName == attribute).Value;
        return string.IsNullOrWhiteSpace(mapped) ? scheme : mapped!;
    }

    internal static OfficeColor? ResolveColor(
        OpenXmlElement? container,
        A.ColorScheme? colorScheme,
        OpenXmlElement? placeholderColor = null,
        OpenXmlElement? colorMap = null) {
        OpenXmlElement? colorElement = FindColorElement(container);
        return ResolveColorElement(colorElement, colorScheme,
            FindColorElement(placeholderColor), colorMap);
    }

    internal static bool HasUnsupportedTransforms(OpenXmlElement? container) {
        OpenXmlElement? color = FindColorElement(container);
        if (color == null) return false;
        foreach (OpenXmlElement transform in color.ChildElements) {
            if (transform.LocalName is "comp" or "inv" or "gray") continue;
            if (transform.LocalName is not ("alpha" or "alphaMod" or "alphaOff" or "tint" or "shade" or
                "lumMod" or "lumOff" or "hueOff" or "hueMod" or "satOff" or "satMod" or
                "red" or "redMod" or "redOff" or "green" or "greenMod" or
                "greenOff" or "blue" or "blueMod" or "blueOff") || !TryReadTransformValue(transform, out _)) return true;
        }
        return false;
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

    private static OfficeColor? ResolveDefaultSpreadsheetSchemeColor(string? token) {
        uint? index = token switch {
            "light1" or "lt1" => 0U,
            "dark1" or "dk1" => 1U,
            "light2" or "lt2" => 2U,
            "dark2" or "dk2" => 3U,
            "accent1" => 4U,
            "accent2" => 5U,
            "accent3" => 6U,
            "accent4" => 7U,
            "accent5" => 8U,
            "accent6" => 9U,
            "hyperlink" or "hlink" => 10U,
            "followedHyperlink" or "folHlink" => 11U,
            _ => null
        };
        return index.HasValue ? ResolveSpreadsheetThemeColor(null, index.Value) : null;
    }

    private static OfficeColor? ResolveColorElement(
        OpenXmlElement? colorElement,
        A.ColorScheme? colorScheme,
        OpenXmlElement? placeholderColor,
        OpenXmlElement? colorMap = null) {
        if (colorElement == null) {
            return null;
        }

        OfficeColor? color;
        if (colorElement is A.RgbColorModelHex rgbColor) {
            color = ParseRgb(rgbColor.Val?.Value);
        } else if (colorElement is A.RgbColorModelPercentage scRgbColor) {
            color = ParseScRgb(scRgbColor);
        } else if (colorElement is A.HslColor hslColor) {
            color = ParseHsl(hslColor);
        } else if (colorElement is A.SystemColor systemColor) {
            color = ParseRgb(systemColor.LastColor?.Value);
        } else if (colorElement is A.SchemeColor schemeColor) {
            string? scheme = GetSchemeValue(schemeColor);
            if (IsPlaceholderScheme(scheme)) {
                color = ResolveColorElement(placeholderColor, colorScheme, null, colorMap);
            } else {
                color = ResolveSchemeColor(colorScheme, scheme == null ? null : MapSchemeColor(scheme, colorMap));
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

    /// <summary>Applies DrawingML color transforms in child order, retaining each transform's alpha semantics.</summary>
    /// <remarks>Tint and shade retain the specified input proportion in linear sRGB, independently of encoded-RGB helpers.</remarks>
    internal static OfficeColor ApplyTransforms(OfficeColor color, OpenXmlElement colorElement) {
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
                    resolved = OfficeGradientColors.Interpolate(
                        OfficeColor.FromRgba(255, 255, 255, resolved.A), resolved,
                        ClampUnit(value), OfficeGradientColorInterpolation.LinearRgb);
                    break;
                case "shade":
                    resolved = OfficeGradientColors.Interpolate(
                        OfficeColor.FromRgba(0, 0, 0, resolved.A), resolved,
                        ClampUnit(value), OfficeGradientColorInterpolation.LinearRgb);
                    break;
                case "lumMod":
                    resolved = OfficeColorTransforms.ModulateLuminance(resolved, Math.Max(0D, value));
                    break;
                case "lumOff":
                    resolved = OfficeColorTransforms.OffsetLuminance(resolved, value);
                    break;
                // Hue offsets are 60000ths of a degree; the shared reader scales by 100000.
                case "hueOff":
                    resolved = OfficeColorTransforms.OffsetHue(resolved, value * 100000D / 60000D);
                    break;
                case "hueMod":
                    resolved = OfficeColorTransforms.ModulateHue(resolved, Math.Max(0D, value));
                    break;
                case "satOff":
                    resolved = OfficeColorTransforms.OffsetSaturation(resolved, value);
                    break;
                case "satMod":
                    resolved = OfficeColorTransforms.ModulateSaturation(resolved, Math.Max(0D, value));
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

        if (container is A.RgbColorModelHex or A.RgbColorModelPercentage
            or A.HslColor or A.SystemColor or A.SchemeColor or A.PresetColor) {
            return container;
        }

        return container.GetFirstChild<A.RgbColorModelHex>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.RgbColorModelPercentage>()
            ?? container.GetFirstChild<A.HslColor>()
            ?? container.GetFirstChild<A.SchemeColor>()
            ?? (OpenXmlElement?)container.GetFirstChild<A.SystemColor>()
            ?? container.GetFirstChild<A.PresetColor>();
    }

    private static OfficeColor? ResolveThemeEntry(OpenXmlCompositeElement? colorElement) {
        if (colorElement == null) {
            return null;
        }

        return ResolveColorElement(FindColorElement(colorElement),
            colorScheme: null, placeholderColor: null);
    }

    private static OfficeColor? ParseRgb(string? value) =>
        OfficeColor.TryParseHex(value, out OfficeColor color) ? color : (OfficeColor?)null;

    private static OfficeColor? ParseScRgb(A.RgbColorModelPercentage color) {
        if (color.RedPortion?.Value is not int red
            || color.GreenPortion?.Value is not int green
            || color.BluePortion?.Value is not int blue) {
            return null;
        }
        return OfficeColor.FromRgb(
            ToChannel(255D * LinearScRgbToSrgb(red / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(green / 100000D)),
            ToChannel(255D * LinearScRgbToSrgb(blue / 100000D)));
    }

    private static double LinearScRgbToSrgb(double value) {
        double linear = ClampUnit(value);
        return linear <= 0.0031308D
            ? 12.92D * linear
            : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D;
    }

    private static OfficeColor? ParseHsl(A.HslColor color) {
        if (color.HueValue?.Value is not int hue
            || color.SatValue?.Value is not int saturation
            || color.LumValue?.Value is not int luminance) {
            return null;
        }
        double normalizedHue = ((hue / 60000D) % 360D + 360D) % 360D;
        double normalizedSaturation = ClampUnit(saturation / 100000D);
        double normalizedLuminance = ClampUnit(luminance / 100000D);
        double chroma = (1D - Math.Abs(2D * normalizedLuminance - 1D))
            * normalizedSaturation;
        double sector = normalizedHue / 60D;
        double intermediate = chroma * (1D - Math.Abs(sector % 2D - 1D));
        double red;
        double green;
        double blue;
        if (sector < 1D) {
            red = chroma; green = intermediate; blue = 0D;
        } else if (sector < 2D) {
            red = intermediate; green = chroma; blue = 0D;
        } else if (sector < 3D) {
            red = 0D; green = chroma; blue = intermediate;
        } else if (sector < 4D) {
            red = 0D; green = intermediate; blue = chroma;
        } else if (sector < 5D) {
            red = intermediate; green = 0D; blue = chroma;
        } else {
            red = chroma; green = 0D; blue = intermediate;
        }
        double match = normalizedLuminance - chroma / 2D;
        return OfficeColor.FromRgb(ToChannel(255D * (red + match)),
            ToChannel(255D * (green + match)),
            ToChannel(255D * (blue + match)));
    }

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
