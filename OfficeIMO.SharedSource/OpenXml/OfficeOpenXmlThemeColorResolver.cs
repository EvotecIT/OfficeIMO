using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;

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
    internal static OfficeColor[]? ReadRadialPalette(IOfficeChartPart chartPart, OfficeXmlNode series,
        int pointCount, OfficeXmlNode scheme, bool requireQualifiedPalette = false,
        bool useSpreadsheetDefaultTheme = false) {
        OfficeXmlNode group = series.Parent;
        bool radial = IsRadialChart(group);
        OfficeXmlNode plot = group.Parent;
        if (!plot.Is(NsC, "plotArea")) plot = OfficeXmlNode.None;
        // A single bar, line, radar or scatter series with c:varyColors takes the same per-point colours.
        if (pointCount < 1 ||
            !radial && OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plot) != series)
            return null;
        OfficeXmlNode varyColors = group.Child(NsC, "varyColors");
        if (varyColors.IsNone || varyColors.BoolVal() == false)
            return null;
        // Point fills take precedence over the theme palette in every static renderer.
        // An unsupported inherited palette is immaterial when every rendered ring has
        // an explicit appearance for every category.
        if (radial && AllRadialPointsHaveExplicitFill(group, pointCount)) return null;
        OfficeColor[]? colors = scheme.IsNone && !useSpreadsheetDefaultTheme ? null :
            OfficeOpenXmlChartAutomaticColors.ResolvePoints(chartPart.ChartSpace, pointCount, scheme,
                useSpreadsheetDefaultTheme, chartPart);
        if (colors == null && requireQualifiedPalette)
            throw new NotSupportedException("The automatic radial palette has no resolvable theme colours.");
        return colors;
    }

    private static bool IsRadialChart(OfficeXmlNode group) =>
        group.NamespaceId == NsC && group.LocalName is "pieChart" or "doughnutChart" or "pie3DChart" or "ofPieChart";

    private static bool AllRadialPointsHaveExplicitFill(OfficeXmlNode chart, int pointCount) {
        bool hasSeries = false;
        bool multipleRings = chart.Is(NsC, "doughnutChart") && chart.Children(NsC, "ser").Count > 1;
        foreach (OfficeXmlNode series in chart.Children(NsC, "ser")) {
            hasSeries = true;
            OfficeXmlNode seriesAppearance = series.Child(NsC, "spPr");
            if (multipleRings && HasExplicitFill(seriesAppearance))
                continue;
            var covered = new bool[pointCount];
            foreach (OfficeXmlNode point in series.Children(NsC, "dPt")) {
                uint? index = point.Child(NsC, "idx").UIntVal();
                if (!index.HasValue || index.Value >= (uint)pointCount) continue;
                if (HasExplicitFill(point.Child(NsC, "spPr")))
                    covered[(int)index.Value] = true;
            }
            if (covered.Any(value => !value)) return false;
        }
        return hasSeries;
    }

    private static bool HasExplicitFill(OfficeXmlNode appearance) =>
        !appearance.IsNone && (appearance.Child(NsA, "solidFill").Exists || appearance.Child(NsA, "noFill").Exists ||
            appearance.Child(NsA, "pattFill").Exists);

    /// <summary>The chart space's own color map override, or the color map of the part that holds the chart.</summary>
    internal static OfficeXmlNode ResolveChartColorMap(IOfficeChartPart chartPart) {
        OfficeXmlNode chartOverride = chartPart.ChartSpace.Child(NsC, "clrMapOvr");
        return chartOverride.Exists ? chartOverride : chartPart.HostColorMap;
    }

    internal static string MapSchemeColor(string scheme, OfficeXmlNode colorMap) {
        if (colorMap.IsNone) return scheme;
        bool wordMap = colorMap.Is(OfficeXmlNamespaces.WordprocessingML, "clrSchemeMapping");
        string attribute = scheme.ToLowerInvariant() switch {
            "dark1" or "dk1" or "text1" or "tx1" => wordMap ? "t1" : "tx1",
            "light1" or "lt1" or "background1" or "bg1" => "bg1",
            "dark2" or "dk2" or "text2" or "tx2" => wordMap ? "t2" : "tx2",
            "light2" or "lt2" or "background2" or "bg2" => "bg2",
            "hyperlink" or "hlink" => "hlink",
            "followedhyperlink" or "folhlink" => "folHlink",
            _ => scheme
        };
        // The first attribute with this local name, in any namespace (w:clrSchemeMapping uses w: attributes).
        string? mapped = null;
        foreach (OfficeXmlAttribute item in colorMap.Attributes()) {
            if (item.LocalName == attribute) {
                mapped = item.Value;
                break;
            }
        }
        return string.IsNullOrWhiteSpace(mapped) ? scheme : mapped!;
    }

    internal static OfficeColor? ResolveColor(
        OfficeXmlNode container,
        OfficeXmlNode colorScheme,
        OfficeXmlNode placeholderColor = default,
        OfficeXmlNode colorMap = default) {
        OfficeXmlNode colorElement = FindColorElement(container);
        return ResolveColorElement(colorElement, colorScheme,
            FindColorElement(placeholderColor), colorMap);
    }

    internal static bool HasUnsupportedTransforms(OfficeXmlNode container) {
        OfficeXmlNode color = FindColorElement(container);
        if (color.IsNone) return false;
        foreach (OfficeXmlNode transform in color.Children()) {
            if (transform.LocalName is "comp" or "inv" or "gray") continue;
            if (transform.LocalName is not ("alpha" or "alphaMod" or "alphaOff" or "tint" or "shade" or
                "lumMod" or "lumOff" or "hueOff" or "hueMod" or "satOff" or "satMod" or
                "red" or "redMod" or "redOff" or "green" or "greenMod" or
                "greenOff" or "blue" or "blueMod" or "blueOff") || !TryReadTransformValue(transform, out _)) return true;
        }
        return false;
    }

    internal static OfficeColor? ResolveSchemeColor(OfficeXmlNode colorScheme, string? scheme) {
        if (colorScheme.IsNone || string.IsNullOrWhiteSpace(scheme)) {
            return null;
        }

        string normalized = scheme!.Trim().ToLowerInvariant();
        string? slot = normalized switch {
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

        return slot == null ? null : ResolveThemeEntry(colorScheme.Child(NsA, slot));
    }

    internal static OfficeColor? ResolveSpreadsheetThemeColor(OfficeXmlNode colorScheme, uint themeIndex) {
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

    /// <summary>Maps shared theme slot identities to SpreadsheetML's light-first color indexes.</summary>
    internal static uint? GetSpreadsheetThemeColorIndex(string? token) => token switch {
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

    private static int? Percentage(OfficeXmlNode element, string localName) =>
        OfficeDrawingColorTransforms.ParsePercentage(element.Attribute(localName));

    private static OfficeColor? ResolveColorElement(
        OfficeXmlNode colorElement,
        OfficeXmlNode colorScheme,
        OfficeXmlNode placeholderColor,
        OfficeXmlNode colorMap = default) {
        if (colorElement.IsNone || colorElement.NamespaceId != NsA) {
            return null;
        }

        OfficeColor? color;
        switch (colorElement.LocalName) {
            case "srgbClr":
                color = ParseRgb(colorElement.Attribute("val"));
                break;
            case "scrgbClr":
                // Percentages: thousandths of a percent, or a strict document's "50%" form.
                color = OfficeDrawingColorTransforms.ParseScRgb(Percentage(colorElement, "r"), Percentage(colorElement, "g"), Percentage(colorElement, "b"));
                break;
            case "hslClr":
                color = OfficeDrawingColorTransforms.ParseHsl(colorElement.IntAttr("hue"), Percentage(colorElement, "sat"), Percentage(colorElement, "lum"));
                break;
            case "sysClr":
                color = ParseRgb(colorElement.Attribute("lastClr"));
                break;
            case "schemeClr": {
                string? scheme = colorElement.Attribute("val");
                if (IsPlaceholderScheme(scheme)) {
                    color = ResolveColorElement(placeholderColor, colorScheme, OfficeXmlNode.None, colorMap);
                } else {
                    color = ResolveSchemeColor(colorScheme, scheme == null ? null : MapSchemeColor(scheme, colorMap));
                }
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

    /// <summary>Applies DrawingML color transforms in child order, retaining each transform's alpha semantics.</summary>
    /// <remarks>Tint and shade retain the specified input proportion in linear sRGB, independently of encoded-RGB helpers.</remarks>
    internal static OfficeColor ApplyTransforms(OfficeColor color, OfficeXmlNode colorElement) {
        var transforms = new OfficeDrawingColorTransforms(color);
        foreach (OfficeXmlNode transform in colorElement.Children()) {
            string name = transform.LocalName;
            transforms.Apply(name, OfficeDrawingColorTransforms.TakesValue(name) ? transform.Attribute("val") : null);
        }
        return transforms.Result;
    }

    private static OfficeXmlNode FindColorElement(OfficeXmlNode container) {
        if (container.IsNone) {
            return OfficeXmlNode.None;
        }

        if (container.NamespaceId == NsA &&
            container.LocalName is "srgbClr" or "scrgbClr" or "hslClr" or "sysClr" or "schemeClr" or "prstClr") {
            return container;
        }

        OfficeXmlNode found = container.Child(NsA, "srgbClr");
        if (found.IsNone) found = container.Child(NsA, "scrgbClr");
        if (found.IsNone) found = container.Child(NsA, "hslClr");
        if (found.IsNone) found = container.Child(NsA, "schemeClr");
        if (found.IsNone) found = container.Child(NsA, "sysClr");
        if (found.IsNone) found = container.Child(NsA, "prstClr");
        return found;
    }

    private static OfficeColor? ResolveThemeEntry(OfficeXmlNode colorElement) {
        if (colorElement.IsNone) {
            return null;
        }

        return ResolveColorElement(FindColorElement(colorElement),
            colorScheme: OfficeXmlNode.None, placeholderColor: OfficeXmlNode.None);
    }

    private static OfficeColor? ParseRgb(string? value) => OfficeDrawingColorTransforms.ParseRgb(value);

    private static bool IsPlaceholderScheme(string? scheme) =>
        string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
        || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadTransformValue(OfficeXmlNode transform, out double value) {
        value = 0D;
        string? raw = transform.Attribute("val");
        if (string.IsNullOrWhiteSpace(raw) || !int.TryParse(raw, out int scaled)) {
            return false;
        }

        value = scaled / 100000D;
        return true;
    }

    // SDK callers
    internal static OfficeColor[]? ReadRadialPalette(ChartPart chartPart, OpenXmlCompositeElement series, int pointCount, A.ColorScheme? scheme, bool requireQualifiedPalette = false, bool useSpreadsheetDefaultTheme = false) =>
        ReadRadialPalette(new OfficeSdkChartPart(chartPart), OfficeXmlNode.FromSdk(series), pointCount, OfficeXmlNode.FromSdk(scheme), requireQualifiedPalette, useSpreadsheetDefaultTheme);
    internal static OpenXmlElement? ResolveChartColorMap(ChartPart chartPart) => (OpenXmlElement?)ResolveChartColorMap(new OfficeSdkChartPart(chartPart)).SdkElement;
    internal static string MapSchemeColor(string scheme, OpenXmlElement? colorMap) => MapSchemeColor(scheme, OfficeXmlNode.FromSdk(colorMap));
    internal static OfficeColor? ResolveColor(OpenXmlElement? container, A.ColorScheme? colorScheme, OpenXmlElement? placeholderColor = null, OpenXmlElement? colorMap = null) =>
        ResolveColor(OfficeXmlNode.FromSdk(container), OfficeXmlNode.FromSdk(colorScheme), OfficeXmlNode.FromSdk(placeholderColor), OfficeXmlNode.FromSdk(colorMap));
    internal static bool HasUnsupportedTransforms(OpenXmlElement? container) => HasUnsupportedTransforms(OfficeXmlNode.FromSdk(container));
    internal static OfficeColor? ResolveSchemeColor(A.ColorScheme? colorScheme, string? scheme) => ResolveSchemeColor(OfficeXmlNode.FromSdk(colorScheme), scheme);
    internal static OfficeColor? ResolveSpreadsheetThemeColor(A.ColorScheme? colorScheme, uint themeIndex) => ResolveSpreadsheetThemeColor(OfficeXmlNode.FromSdk(colorScheme), themeIndex);
    internal static OfficeColor ApplyTransforms(OfficeColor color, OpenXmlElement colorElement) => ApplyTransforms(color, OfficeXmlNode.FromSdk(colorElement));
}
