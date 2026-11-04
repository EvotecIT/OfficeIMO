using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>
/// Resolves the colours Office gives chart series and data points that store no fill or line colour.
/// </summary>
/// <remarks>
/// Word, Excel and PowerPoint 16 derive these colours from the classic <c>c:style</c> number
/// (2 when absent) and the theme. The <c>cs:chartStyle</c> and <c>cs:colorStyle</c> parts only
/// drive later restyling in the Office UI; Office does not consult them when it renders a series
/// without <c>c:spPr</c>. Style column one (styles 1, 9, 17, ...) cycles six tints of text 1,
/// column two cycles accents 1 to 6, and columns three to eight shade a single accent. Each
/// pass through the pattern is darkened or lightened in linear light, spread between -70% and
/// +70%: a pattern of length L used by n items has n / L + 1 passes (n passes for one accent).
/// The fixtures in
/// <c>OfficeIMO.TestAssets/Documents/Charts/AutomaticColors</c> pin these rules to Office exports.
/// </remarks>
internal static class OfficeOpenXmlChartAutomaticColors {
    internal const string AutomaticSeriesColorExtensionUri = "urn:officeimo:chart:auto-series-color:1";

    internal static bool HasAutomaticSeriesColor(OpenXmlElement element) =>
        element.ChildElements.Where(child => child.LocalName == "extLst")
            .SelectMany(list => list.ChildElements)
            .Any(extension => extension.GetAttribute("uri", string.Empty).Value == AutomaticSeriesColorExtensionUri);

    internal static void ClearAutomaticSeriesColor(OpenXmlElement element) {
        foreach (var extension in element.ChildElements.Where(child => child.LocalName == "extLst")
            .SelectMany(list => list.ChildElements).Where(child =>
                child.GetAttribute("uri", string.Empty).Value == AutomaticSeriesColorExtensionUri).ToArray())
            extension.Remove();
    }

    private const string Chart2010Namespace = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
    private static readonly double[] GreyscaleTints = { 0.88D, 0.55D, 0.78D, 0.92D, 0.70D, 0.30D };
    private static readonly OfficeChartMarkerShape[] AutomaticMarkers = {
        OfficeChartMarkerShape.Diamond, OfficeChartMarkerShape.Square, OfficeChartMarkerShape.Triangle,
        OfficeChartMarkerShape.X, OfficeChartMarkerShape.Star, OfficeChartMarkerShape.Dot,
        OfficeChartMarkerShape.Dash, OfficeChartMarkerShape.Circle, OfficeChartMarkerShape.Plus
    };

    /// <summary>Office's automatic marker size in points.</summary>
    internal const int AutomaticMarkerSize = 7;

    /// <summary>Reads the classic chart style number, preferring the Office 2010 <c>c14:style</c> choice.</summary>
    internal static int ReadStyle(OpenXmlElement? chartSpace) {
        if (chartSpace == null) return 2;
        foreach (OpenXmlElement child in chartSpace.ChildElements) {
            if (child.LocalName == "AlternateContent") {
                OpenXmlElement? choice = child.ChildElements.FirstOrDefault(element => element.LocalName == "Choice");
                OpenXmlElement? modern = choice?.ChildElements.FirstOrDefault(element =>
                    element.LocalName == "style" && element.NamespaceUri == Chart2010Namespace);
                if (TryReadStyleValue(modern, out int modernStyle) && modernStyle > 100 && modernStyle <= 148)
                    return modernStyle - 100;
                OpenXmlElement? fallback = child.ChildElements.FirstOrDefault(element => element.LocalName == "Fallback")?
                    .ChildElements.FirstOrDefault(element => element.LocalName == "style");
                if (TryReadStyleValue(fallback, out int fallbackStyle) && fallbackStyle >= 1 && fallbackStyle <= 48)
                    return fallbackStyle;
            } else if (child is C.Style style && style.Val?.Value is byte value && value >= 1 && value <= 48) {
                return value;
            }
        }
        return 2;
    }

    /// <summary>Resolves the automatic colour of item <paramref name="index"/> of <paramref name="count"/>.</summary>
    /// <param name="style">The classic chart style number.</param>
    /// <param name="index">The series <c>c:idx</c>, or the data-point index when colours vary by point.</param>
    /// <param name="count">The number of colour slots in use: the highest series index plus one, or the point count.</param>
    /// <param name="scheme">The theme colour scheme.</param>
    /// <param name="colorMap">The owner's colour map, when one applies.</param>
    /// <param name="useSpreadsheetDefaultTheme">Whether to fall back to Excel's built-in theme colours.</param>
    internal static OfficeColor? Resolve(int style, int index, int count, A.ColorScheme? scheme,
        OpenXmlElement? colorMap, bool useSpreadsheetDefaultTheme = false) {
        if (index < 0) return null;
        count = Math.Max(count, index + 1);
        int column = (Math.Max(1, Math.Min(48, style)) - 1) % 8;
        int patternLength = column == 0 ? GreyscaleTints.Length : column == 1 ? 6 : 1;
        int slot = index % patternLength;
        OfficeColor? color;
        if (column == 0) {
            color = ResolveTheme("tx1", scheme, colorMap, useSpreadsheetDefaultTheme);
            if (color.HasValue) color = ChartTint(color.Value, GreyscaleTints[slot]);
        } else {
            string accent = "accent" + (column == 1 ? slot + 1 : column - 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            color = ResolveTheme(accent, scheme, colorMap, useSpreadsheetDefaultTheme);
        }
        if (!color.HasValue) return null;
        // Office spreads the passes of a multi-colour pattern by count / length, so six accent
        // series are already darkened, but spreads a single-accent ramp across all count items.
        int cycle = index / patternLength;
        int lastCycle = patternLength == 1 ? count - 1 : count / patternLength;
        // The shade or tint runs from -70% to +70% in whole percent, truncated as Office does.
        long percent = (cycle + 1L) * 140L / (lastCycle + 2L) - 70L;
        if (percent == 0L) return color;
        double variation = percent / 100D;
        return variation < 0D ? ChartShade(color.Value, 1D + variation) : ChartTint(color.Value, 1D - variation);
    }

    /// <summary>Resolves the automatic colour of one series element from its chart's style and theme.</summary>
    internal static OfficeColor? ResolveSeries(OpenXmlElement series, A.ColorScheme? scheme,
        bool useSpreadsheetDefaultTheme = false) =>
        ResolveSeries(series, scheme, new SeriesResolutionContext(), useSpreadsheetDefaultTheme);

    /// <summary>Reuses plot metadata for the series in one chart projection.</summary>
    internal sealed class SeriesResolutionContext {
        private readonly Dictionary<C.PlotArea, int> _counts = new(PlotReferenceComparer.Instance);

        internal int Count(C.PlotArea? plot, int minimum) {
            if (plot == null) return minimum;
            if (!_counts.TryGetValue(plot, out int count)) {
                count = 0;
                foreach (C.Index item in plot.ChildElements.SelectMany(group => group.ChildElements)
                    .Where(element => element.LocalName == "ser").Select(element => element.GetFirstChild<C.Index>())
                    .OfType<C.Index>()) {
                    if (item.Val?.Value is uint value && value < int.MaxValue)
                        count = Math.Max(count, (int)value + 1);
                }
                _counts.Add(plot, count);
            }
            return Math.Max(count, minimum);
        }
    }

    private sealed class PlotReferenceComparer : IEqualityComparer<C.PlotArea> {
        internal static readonly PlotReferenceComparer Instance = new();
        public bool Equals(C.PlotArea? left, C.PlotArea? right) => ReferenceEquals(left, right);
        public int GetHashCode(C.PlotArea value) => RuntimeHelpers.GetHashCode(value);
    }

    internal static OfficeColor? ResolveSeries(OpenXmlElement series, A.ColorScheme? scheme,
        SeriesResolutionContext context, bool useSpreadsheetDefaultTheme = false) {
        C.ChartSpace? chartSpace = series.Ancestors<C.ChartSpace>().FirstOrDefault();
        uint index = series.GetFirstChild<C.Index>()?.Val?.Value ?? 0U;
        if (index > int.MaxValue - 1) return null;
        C.PlotArea? plot = series.Ancestors<C.PlotArea>().FirstOrDefault();
        int count = context.Count(plot, (int)index + 1);
        return Resolve(ReadStyle(chartSpace), (int)index, count, scheme, ResolveColorMap(chartSpace),
            useSpreadsheetDefaultTheme);
    }

    /// <summary>Resolves the automatic colours of <paramref name="count"/> varied data points.</summary>
    internal static OfficeColor[]? ResolvePoints(OpenXmlElement? chartSpace, int count, A.ColorScheme? scheme,
        bool useSpreadsheetDefaultTheme = false) {
        if (count < 1) return null;
        int style = ReadStyle(chartSpace);
        OpenXmlElement? colorMap = ResolveColorMap(chartSpace as C.ChartSpace);
        var colors = new OfficeColor[count];
        for (int index = 0; index < count; index++) {
            OfficeColor? color = Resolve(style, index, count, scheme, colorMap, useSpreadsheetDefaultTheme);
            if (!color.HasValue) return null;
            colors[index] = color.Value;
        }
        return colors;
    }

    /// <summary>Returns Office's automatic marker symbol for a series index.</summary>
    internal static OfficeChartMarkerShape ResolveMarker(uint seriesIndex) =>
        AutomaticMarkers[(int)(seriesIndex % (uint)AutomaticMarkers.Length)];

    /// <summary>Darkens a colour in linear light, retaining <paramref name="ratio"/> of it.</summary>
    internal static OfficeColor ChartShade(OfficeColor color, double ratio) => MapLinear(color, value => value * ratio);

    /// <summary>Lightens a colour in linear light, retaining <paramref name="ratio"/> of it and mixing in white.</summary>
    internal static OfficeColor ChartTint(OfficeColor color, double ratio) =>
        MapLinear(color, value => value * ratio + (1D - ratio));

    private static OfficeColor? ResolveTheme(string token, A.ColorScheme? scheme, OpenXmlElement? colorMap,
        bool useSpreadsheetDefaultTheme) {
        string mapped = OfficeOpenXmlThemeColorResolver.MapSchemeColor(token, colorMap);
        OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(scheme, mapped);
        if (color.HasValue || !useSpreadsheetDefaultTheme) return color;
        uint? themeIndex = mapped switch {
            "dk1" or "tx1" => 1U,
            "accent1" => 4U, "accent2" => 5U, "accent3" => 6U,
            "accent4" => 7U, "accent5" => 8U, "accent6" => 9U,
            _ => null
        };
        return themeIndex.HasValue ? OfficeOpenXmlThemeColorResolver.ResolveSpreadsheetThemeColor(null, themeIndex.Value) : null;
    }

    private static OpenXmlElement? ResolveColorMap(C.ChartSpace? chartSpace) =>
        chartSpace?.ChartPart is ChartPart part ? OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(part) : null;

    private static OfficeColor MapLinear(OfficeColor color, Func<double, double> map) =>
        OfficeColor.FromRgba(MapChannel(color.R, map), MapChannel(color.G, map), MapChannel(color.B, map), color.A);

    private static byte MapChannel(byte channel, Func<double, double> map) {
        double value = channel / 255D;
        double linear = value <= 0.04045D ? value / 12.92D : Math.Pow((value + 0.055D) / 1.055D, 2.4D);
        linear = Math.Max(0D, Math.Min(1D, map(linear)));
        double encoded = linear <= 0.0031308D ? linear * 12.92D : 1.055D * Math.Pow(linear, 1D / 2.4D) - 0.055D;
        return (byte)Math.Max(0D, Math.Min(255D, Math.Round(encoded * 255D, MidpointRounding.AwayFromZero)));
    }

    private static bool TryReadStyleValue(OpenXmlElement? element, out int value) {
        value = 0;
        string? raw = element?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return raw != null && int.TryParse(raw, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out value);
    }
}
