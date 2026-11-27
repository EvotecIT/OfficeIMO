using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads chart-wide plotting rules that every native chart adapter projects the same way.</summary>
internal static class OfficeOpenXmlChartPlotSemantics {
    /// <summary>Applies the chart's blank-cell rule, manual layouts, colour variation, the physical axis
    /// crossings and scales, and the value-axis display units.</summary>
    internal static OfficeChartLayout WithPlotSemantics(this OfficeChartLayout layout, OfficeXmlNode chart, bool includeDisplayUnits = true) {
        if (chart.IsNone) return layout;
        layout = layout.WithBlankDisplay(ReadBlankDisplay(chart));
        OfficeXmlNode overMaximum = chart.Child(NsC, "showDLblsOverMax");
        layout = layout.WithDataLabelsOverMaximum(overMaximum.IsNone || overMaximum.BoolVal() != false);
        OfficeXmlNode plot = chart.Child(NsC, "plotArea");
        layout = layout.WithManualLayouts(ReadManualLayout(plot), ReadManualLayout(chart.Child(NsC, "title")),
            ReadManualLayout(chart.Child(NsC, "legend")));
        layout = layout.WithVaryColorsByPoint(!GetPointColoredSeries(plot).IsNone);
        if (!TryGetPhysicalAxes(plot, out OfficeXmlNode horizontal, out OfficeXmlNode vertical))
            return layout;
        layout = layout.WithLogarithmicValueAxes(ReadLogBase(horizontal), ReadLogBase(vertical))
            .WithReversedValueAxes(IsReversedValueAxis(horizontal), IsReversedValueAxis(vertical));
        // An axis's c:crosses or c:crossesAt places that axis on the perpendicular one: a category axis with
        // crosses="autoZero" runs along the value axis's zero line, and a value axis with crosses="max" sits after
        // the last category.
        layout = layout.WithAxisCrossings(ReadCrossing(horizontal), ReadCrossesAt(horizontal),
            ReadCrossing(vertical), ReadCrossesAt(vertical));
        if (!includeDisplayUnits) return layout;
        (double? Divisor, string? Label) horizontalUnits = horizontal.IsChart("valAx") ? ReadDisplayUnits(horizontal) : default;
        (double? Divisor, string? Label) verticalUnits = vertical.IsChart("valAx") ? ReadDisplayUnits(vertical) : default;
        return layout.WithAxisDisplayUnits(horizontalUnits.Divisor, horizontalUnits.Label, verticalUnits.Divisor, verticalUnits.Label);
    }

    /// <summary>The placeholder Office shows for an automatic title it cannot take from a series name, in English.</summary>
    internal const string AutomaticTitlePlaceholder = "Chart Title";

    /// <summary>Resolves the text Office draws for a c:title that has no text of its own, or for a chart with no
    /// c:title whose c:autoTitleDeleted is off.</summary>
    /// <remarks>Office titles a chart with one series by that series' name. Otherwise it shows a placeholder in the
    /// user interface language; the placeholder is returned only when the chart's c:lang is English or absent, so a
    /// chart in another language is not given English text. Returns null when the chart has no such title, when its
    /// automatic title was deleted (c:autoTitleDeleted), or when no text applies.</remarks>
    /// <param name="chart">The native chart.</param>
    /// <param name="singleSeriesName">The series name a caller resolved from current source cells, used in place of the cached name.</param>
    internal static string? ReadAutomaticTitle(OfficeXmlNode chart, string? singleSeriesName = null) {
        OfficeXmlNode title = chart.Child(NsC, "title");
        if (!title.Child(NsC, "tx").IsNone) return null;
        OfficeXmlNode deleted = chart.Child(NsC, "autoTitleDeleted");
        if (!deleted.IsNone && deleted.BoolVal() != false) return null;
        // Without a c:title element Excel still titles a one-series chart with the series name when
        // c:autoTitleDeleted is present and off; with the element missing, or several series, it shows no title.
        bool implicitTitle = title.IsNone;
        if (implicitTitle && deleted.IsNone) return null;
        List<OfficeXmlNode> series = FirstSeries(chart.Child(NsC, "plotArea"), 2);
        if (series.Count == 0) return null;
        if (series.Count == 1) {
            string? name = string.IsNullOrWhiteSpace(singleSeriesName) ? ReadSeriesName(series[0]) : singleSeriesName;
            if (!string.IsNullOrWhiteSpace(name)) return name!.Trim();
        }

        if (implicitTitle) return null;
        OfficeXmlNode chartSpace = chart.Parent;
        string? language = chartSpace.IsChart("chartSpace") ? chartSpace.ChildVal(NsC, "lang") : null;
        return string.IsNullOrWhiteSpace(language) || language!.StartsWith("en", StringComparison.OrdinalIgnoreCase)
            ? AutomaticTitlePlaceholder : null;
    }

    // The first series (elements named ser) of the plot's chart layers, at most limit of them.
    private static List<OfficeXmlNode> FirstSeries(OfficeXmlNode plot, int limit) {
        var series = new List<OfficeXmlNode>(limit);
        foreach (OfficeXmlNode layer in plot.Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal)) continue;
            foreach (OfficeXmlNode item in layer.Children()) {
                if (item.LocalName != "ser") continue;
                series.Add(item);
                if (series.Count == limit) return series;
            }
        }
        return series;
    }

    private static string? ReadSeriesName(OfficeXmlNode series) {
        OfficeXmlNode text = series.Child(NsC, "tx");
        if (text.IsNone) return null;
        string? direct = text.ChildText(NsC, "v");
        if (!string.IsNullOrWhiteSpace(direct)) return direct;
        OfficeXmlNode point = text.ContentDescendants(NsC, "pt")
            .FirstOrDefault(item => (item.UIntAttr("idx") ?? 0U) == 0U);
        return point.ChildText(NsC, "v");
    }

    /// <summary>Returns the only series of a chart whose group sets c:varyColors, so Office colours its points
    /// individually; none for radial charts (which always vary), area charts and charts with several series.</summary>
    internal static OfficeXmlNode GetPointColoredSeries(OfficeXmlNode plot) {
        if (plot.IsNone) return default;
        List<OfficeXmlNode> series = FirstSeries(plot, 2);
        if (series.Count != 1) return default;
        OfficeXmlNode layer = series[0].Parent;
        if (!(layer.IsChart("barChart") || layer.IsChart("bar3DChart") || layer.IsChart("lineChart") ||
                layer.IsChart("line3DChart") || layer.IsChart("radarChart") || layer.IsChart("scatterChart"))) return default;
        OfficeXmlNode vary = layer.Child(NsC, "varyColors");
        return !vary.IsNone && vary.BoolVal() != false ? series[0] : default;
    }

    // c:layout/c:manualLayout. A missing mode is factor and a missing target is outer, as the schema defaults.
    private static OfficeChartManualLayout? ReadManualLayout(OfficeXmlNode owner) {
        OfficeXmlNode manual = owner.Child(NsC, "layout").Child(NsC, "manualLayout");
        if (manual.IsNone) return null;
        double? x = manual.Child(NsC, "x").DoubleVal();
        double? y = manual.Child(NsC, "y").DoubleVal();
        double? width = manual.Child(NsC, "w").DoubleVal();
        double? height = manual.Child(NsC, "h").DoubleVal();
        if (!x.HasValue && !y.HasValue && !width.HasValue && !height.HasValue) return null;
        OfficeChartLayoutMode xMode = ReadMode(manual.ChildVal(NsC, "xMode"));
        OfficeChartLayoutMode yMode = ReadMode(manual.ChildVal(NsC, "yMode"));
        try {
            return new OfficeChartManualLayout(x ?? 0D, y ?? 0D, width, height,
                x.HasValue ? xMode : OfficeChartLayoutMode.Factor, y.HasValue ? yMode : OfficeChartLayoutMode.Factor,
                ReadMode(manual.ChildVal(NsC, "wMode")), ReadMode(manual.ChildVal(NsC, "hMode")),
                manual.ChildVal(NsC, "layoutTarget") == "inner"
                    ? OfficeChartLayoutTarget.Inner : OfficeChartLayoutTarget.Outer);
        } catch (ArgumentOutOfRangeException) {
            // Office ignores a layout far outside the chart and places the element automatically.
            return null;
        }
    }

    private static OfficeChartLayoutMode ReadMode(string? value) =>
        value == "edge" ? OfficeChartLayoutMode.Edge : OfficeChartLayoutMode.Factor;

    private static double? ReadLogBase(OfficeXmlNode axis) {
        if (!axis.IsChart("valAx")) return null;
        double? value = axis.Child(NsC, "scaling").Child(NsC, "logBase").DoubleVal();
        return value is double logBase && logBase >= 2D && logBase <= 1000D ? logBase : null;
    }

    private static bool IsReversedValueAxis(OfficeXmlNode axis) =>
        axis.IsChart("valAx") && axis.Child(NsC, "scaling").ChildVal(NsC, "orientation") == "maxMin";

    /// <summary>Reads c:dispBlanksAs. Office treats a missing element or value as gaps.</summary>
    internal static OfficeChartBlankDisplay ReadBlankDisplay(OfficeXmlNode chart) {
        string? value = chart.ChildVal(NsC, "dispBlanksAs");
        if (value == "span") return OfficeChartBlankDisplay.Span;
        if (value == "zero") return OfficeChartBlankDisplay.Zero;
        return OfficeChartBlankDisplay.Gap;
    }

    /// <summary>Resolves the primary axes drawn horizontally (bottom or top) and vertically (left or right).</summary>
    internal static bool TryGetPhysicalAxes(OfficeXmlNode plot, out OfficeXmlNode horizontal, out OfficeXmlNode vertical) {
        horizontal = default;
        vertical = default;
        if (plot.IsNone) return false;
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
        OfficeXmlNode layer = default;
        foreach (OfficeXmlNode element in plot.Children()) {
            if (element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && !element.Child(NsC, "axId").IsNone &&
                groups.Read(element) == OfficeChartAxisGroup.Primary) {
                layer = element;
                break;
            }
        }
        if (layer.IsNone) return false;
        OfficeXmlNode[] axes = layer.Children(NsC, "axId").ToList().Select(reference => groups.Resolve(reference.UIntVal())).ToArray();
        horizontal = axes.FirstOrDefault(axis => IsHorizontal(axis) == true);
        vertical = axes.FirstOrDefault(axis => IsHorizontal(axis) == false);
        return !horizontal.IsNone && !vertical.IsNone;
    }

    private static bool? IsHorizontal(OfficeXmlNode axis) {
        string? position = axis.ChildVal(NsC, "axPos");
        if (position == "b" || position == "t") return true;
        if (position == "l" || position == "r") return false;
        return null;
    }

    private static OfficeChartAxisCrossingPosition ReadCrossing(OfficeXmlNode axis) {
        string? value = axis.ChildVal(NsC, "crosses");
        if (value == "max") return OfficeChartAxisCrossingPosition.Maximum;
        if (value == "min") return OfficeChartAxisCrossingPosition.Minimum;
        return OfficeChartAxisCrossingPosition.AutoZero;
    }

    private static double? ReadCrossesAt(OfficeXmlNode axis) {
        double? value = axis.Child(NsC, "crossesAt").DoubleVal();
        return value.HasValue && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value) ? value : null;
    }

    /// <summary>Reads a value axis's c:dispUnits divisor and, when c:dispUnitsLbl is present, its label.</summary>
    internal static (double? Divisor, string? Label) ReadDisplayUnits(OfficeXmlNode axis) {
        OfficeXmlNode displayUnits = axis.Child(NsC, "dispUnits");
        if (displayUnits.IsNone) return default;
        OfficeXmlNode builtIn = displayUnits.Child(NsC, "builtInUnit");
        // The attribute as written, or the element's text when it has none, as the SDK reads it.
        string? builtInUnit = builtIn.IsNone ? null : builtIn.Attribute("val") ?? builtIn.InnerText;
        double? divisor = builtInUnit == null
            ? displayUnits.Child(NsC, "custUnit").DoubleVal()
            : GetBuiltInDivisor(builtInUnit);
        if (!divisor.HasValue || divisor.Value <= 0D || double.IsNaN(divisor.Value) || double.IsInfinity(divisor.Value))
            return default;
        OfficeXmlNode displayUnitsLabel = displayUnits.Child(NsC, "dispUnitsLbl");
        string? label = ReadChartTextValue(displayUnitsLabel.Child(NsC, "tx"));
        if (!displayUnitsLabel.IsNone && string.IsNullOrWhiteSpace(label)) {
            label = builtInUnit == null
                ? divisor.Value.ToString("0.##", CultureInfo.InvariantCulture)
                : GetBuiltInLabel(builtInUnit);
        }

        return (divisor.Value, string.IsNullOrWhiteSpace(label) ? null : label!.Trim());
    }

    // The displayed text of c:tx: its rich-text runs, or the first cached point of a formula reference.
    // InnerText would also include the formula itself.
    private static string? ReadChartTextValue(OfficeXmlNode text) {
        if (text.IsNone) return null;
        OfficeXmlNode rich = text.Child(NsC, "rich");
        if (!rich.IsNone) return string.Concat(rich.ContentDescendants(NsA, "t").Select(run => run.InnerText));
        OfficeXmlNode point = default;
        foreach (OfficeXmlNode item in text.Child(NsC, "strRef").Child(NsC, "strCache").Children(NsC, "pt")) {
            if ((item.UIntAttr("idx") ?? 0U) == 0U) {
                point = item;
                break;
            }
        }
        return point.ChildText(NsC, "v") ?? text.Child(NsC, "strLit").Children(NsC, "pt").First().ChildText(NsC, "v");
    }

    // SDK callers
    internal static OfficeChartLayout WithPlotSemantics(this OfficeChartLayout layout, C.Chart? chart, bool includeDisplayUnits = true) =>
        layout.WithPlotSemantics(OfficeXmlNode.FromSdk(chart), includeDisplayUnits);

    internal static string? ReadAutomaticTitle(C.Chart chart, string? singleSeriesName = null) =>
        ReadAutomaticTitle(OfficeXmlNode.FromSdk(chart), singleSeriesName);

    internal static OpenXmlCompositeElement? GetPointColoredSeries(C.PlotArea? plot) =>
        (OpenXmlCompositeElement?)GetPointColoredSeries(OfficeXmlNode.FromSdk(plot)).SdkElement;

    internal static OfficeChartBlankDisplay ReadBlankDisplay(C.Chart chart) => ReadBlankDisplay(OfficeXmlNode.FromSdk(chart));

    internal static (double? Divisor, string? Label) ReadDisplayUnits(OpenXmlCompositeElement? axis) =>
        ReadDisplayUnits(OfficeXmlNode.FromSdk(axis));

    private static double? GetBuiltInDivisor(string builtInUnit) => NormalizeBuiltInUnit(builtInUnit) switch {
        "hundreds" => 100D,
        "thousands" => 1_000D,
        "tenthousands" => 10_000D,
        "hundredthousands" => 100_000D,
        "millions" => 1_000_000D,
        "tenmillions" => 10_000_000D,
        "hundredmillions" => 100_000_000D,
        "billions" => 1_000_000_000D,
        "trillions" => 1_000_000_000_000D,
        _ => null
    };

    private static string GetBuiltInLabel(string builtInUnit) => NormalizeBuiltInUnit(builtInUnit) switch {
        "hundreds" => "Hundreds",
        "thousands" => "Thousands",
        "tenthousands" => "Ten Thousands",
        "hundredthousands" => "Hundred Thousands",
        "millions" => "Millions",
        "tenmillions" => "Ten Millions",
        "hundredmillions" => "Hundred Millions",
        "billions" => "Billions",
        "trillions" => "Trillions",
        _ => builtInUnit
    };

    private static string NormalizeBuiltInUnit(string builtInUnit) =>
        builtInUnit.Replace(" ", string.Empty).Replace("-", string.Empty).Replace("_", string.Empty).ToLowerInvariant();
}
