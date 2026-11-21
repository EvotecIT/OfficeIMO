using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads chart-wide plotting rules that every native chart adapter projects the same way.</summary>
internal static class OfficeOpenXmlChartPlotSemantics {
    /// <summary>Applies the chart's blank-cell rule, manual layouts, colour variation, the physical axis
    /// crossings and scales, and the value-axis display units.</summary>
    internal static OfficeChartLayout WithPlotSemantics(this OfficeChartLayout layout, C.Chart? chart, bool includeDisplayUnits = true) {
        if (chart == null) return layout;
        layout = layout.WithBlankDisplay(ReadBlankDisplay(chart));
        C.ShowDataLabelsOverMaximum? overMaximum = chart.GetFirstChild<C.ShowDataLabelsOverMaximum>();
        layout = layout.WithDataLabelsOverMaximum(overMaximum == null || overMaximum.Val?.Value != false);
        layout = layout.WithManualLayouts(ReadManualLayout(chart.PlotArea), ReadManualLayout(chart.GetFirstChild<C.Title>()),
            ReadManualLayout(chart.GetFirstChild<C.Legend>()));
        layout = layout.WithVaryColorsByPoint(GetPointColoredSeries(chart.PlotArea) != null);
        if (!TryGetPhysicalAxes(chart.PlotArea, out OpenXmlCompositeElement? horizontal, out OpenXmlCompositeElement? vertical))
            return layout;
        layout = layout.WithLogarithmicValueAxes(ReadLogBase(horizontal), ReadLogBase(vertical))
            .WithReversedValueAxes(IsReversedValueAxis(horizontal), IsReversedValueAxis(vertical));
        // An axis's c:crosses or c:crossesAt places that axis on the perpendicular one: a category axis with
        // crosses="autoZero" runs along the value axis's zero line, and a value axis with crosses="max" sits after
        // the last category.
        layout = layout.WithAxisCrossings(ReadCrossing(horizontal), ReadCrossesAt(horizontal),
            ReadCrossing(vertical), ReadCrossesAt(vertical));
        if (!includeDisplayUnits) return layout;
        (double? Divisor, string? Label) horizontalUnits = horizontal is C.ValueAxis ? ReadDisplayUnits(horizontal) : default;
        (double? Divisor, string? Label) verticalUnits = vertical is C.ValueAxis ? ReadDisplayUnits(vertical) : default;
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
    internal static string? ReadAutomaticTitle(C.Chart chart, string? singleSeriesName = null) {
        C.Title? title = chart.GetFirstChild<C.Title>();
        if (title?.GetFirstChild<C.ChartText>() != null) return null;
        C.AutoTitleDeleted? deleted = chart.GetFirstChild<C.AutoTitleDeleted>();
        if (deleted != null && deleted.Val?.Value != false) return null;
        // Without a c:title element Excel still titles a one-series chart with the series name when
        // c:autoTitleDeleted is present and off; with the element missing, or several series, it shows no title.
        bool implicitTitle = title == null;
        if (implicitTitle && deleted == null) return null;
        OpenXmlCompositeElement[] series = chart.PlotArea?.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .SelectMany(layer => layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser"))
            .Take(2).ToArray() ?? Array.Empty<OpenXmlCompositeElement>();
        if (series.Length == 0) return null;
        if (series.Length == 1) {
            string? name = string.IsNullOrWhiteSpace(singleSeriesName) ? ReadSeriesName(series[0]) : singleSeriesName;
            if (!string.IsNullOrWhiteSpace(name)) return name!.Trim();
        }

        if (implicitTitle) return null;
        string? language = (chart.Parent as C.ChartSpace)?.GetFirstChild<C.EditingLanguage>()?.Val?.Value;
        return string.IsNullOrWhiteSpace(language) || language!.StartsWith("en", StringComparison.OrdinalIgnoreCase)
            ? AutomaticTitlePlaceholder : null;
    }

    private static string? ReadSeriesName(OpenXmlCompositeElement series) {
        C.SeriesText? text = series.GetFirstChild<C.SeriesText>();
        if (text == null) return null;
        string? direct = text.GetFirstChild<C.NumericValue>()?.Text;
        if (!string.IsNullOrWhiteSpace(direct)) return direct;
        C.StringPoint? point = text.Descendants<C.StringPoint>()
            .FirstOrDefault(item => (item.Index?.Value ?? 0U) == 0U);
        return point?.NumericValue?.Text;
    }

    /// <summary>Returns the only series of a chart whose group sets c:varyColors, so Office colours its points
    /// individually; null for radial charts (which always vary), area charts and charts with several series.</summary>
    internal static OpenXmlCompositeElement? GetPointColoredSeries(C.PlotArea? plot) {
        if (plot == null) return null;
        OpenXmlCompositeElement[] series = plot.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .SelectMany(layer => layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser"))
            .Take(2).ToArray();
        if (series.Length != 1 || series[0].Parent is not (C.BarChart or C.Bar3DChart or C.LineChart or C.Line3DChart or
                C.RadarChart or C.ScatterChart)) return null;
        C.VaryColors? vary = series[0].Parent!.GetFirstChild<C.VaryColors>();
        return vary != null && vary.Val?.Value != false ? series[0] : null;
    }

    // c:layout/c:manualLayout. A missing mode is factor and a missing target is outer, as the schema defaults.
    private static OfficeChartManualLayout? ReadManualLayout(OpenXmlCompositeElement? owner) {
        C.ManualLayout? manual = owner?.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>();
        if (manual == null) return null;
        double? x = manual.GetFirstChild<C.Left>()?.Val?.Value;
        double? y = manual.GetFirstChild<C.Top>()?.Val?.Value;
        double? width = manual.GetFirstChild<C.Width>()?.Val?.Value;
        double? height = manual.GetFirstChild<C.Height>()?.Val?.Value;
        if (!x.HasValue && !y.HasValue && !width.HasValue && !height.HasValue) return null;
        OfficeChartLayoutMode xMode = ReadMode(manual.GetFirstChild<C.LeftMode>()?.Val?.Value);
        OfficeChartLayoutMode yMode = ReadMode(manual.GetFirstChild<C.TopMode>()?.Val?.Value);
        try {
            return new OfficeChartManualLayout(x ?? 0D, y ?? 0D, width, height,
                x.HasValue ? xMode : OfficeChartLayoutMode.Factor, y.HasValue ? yMode : OfficeChartLayoutMode.Factor,
                ReadMode(manual.GetFirstChild<C.WidthMode>()?.Val?.Value), ReadMode(manual.GetFirstChild<C.HeightMode>()?.Val?.Value),
                manual.GetFirstChild<C.LayoutTarget>()?.Val?.Value == C.LayoutTargetValues.Inner
                    ? OfficeChartLayoutTarget.Inner : OfficeChartLayoutTarget.Outer);
        } catch (ArgumentOutOfRangeException) {
            // Office ignores a layout far outside the chart and places the element automatically.
            return null;
        }
    }

    private static OfficeChartLayoutMode ReadMode(C.LayoutModeValues? value) =>
        value == C.LayoutModeValues.Edge ? OfficeChartLayoutMode.Edge : OfficeChartLayoutMode.Factor;

    private static double? ReadLogBase(OpenXmlCompositeElement? axis) {
        if (axis is not C.ValueAxis) return null;
        double? value = axis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.LogBase>()?.Val?.Value;
        return value is double logBase && logBase >= 2D && logBase <= 1000D ? logBase : null;
    }

    private static bool IsReversedValueAxis(OpenXmlCompositeElement? axis) =>
        axis is C.ValueAxis && axis.GetFirstChild<C.Scaling>()?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin;

    /// <summary>Reads c:dispBlanksAs. Office treats a missing element or value as gaps.</summary>
    internal static OfficeChartBlankDisplay ReadBlankDisplay(C.Chart chart) {
        C.DisplayBlanksAsValues? value = chart.GetFirstChild<C.DisplayBlanksAs>()?.Val?.Value;
        if (value == C.DisplayBlanksAsValues.Span) return OfficeChartBlankDisplay.Span;
        if (value == C.DisplayBlanksAsValues.Zero) return OfficeChartBlankDisplay.Zero;
        return OfficeChartBlankDisplay.Gap;
    }

    /// <summary>Resolves the primary axes drawn horizontally (bottom or top) and vertically (left or right).</summary>
    internal static bool TryGetPhysicalAxes(C.PlotArea? plot, out OpenXmlCompositeElement? horizontal,
        out OpenXmlCompositeElement? vertical) {
        horizontal = null;
        vertical = null;
        if (plot == null) return false;
        var groups = OfficeOpenXmlChartAxisGroups.Create(plot);
        OpenXmlCompositeElement? layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().FirstOrDefault(element =>
            element.LocalName.EndsWith("Chart", StringComparison.Ordinal) && element.Elements<C.AxisId>().Any() &&
            groups.Read(element) == OfficeChartAxisGroup.Primary);
        if (layer == null) return false;
        OpenXmlCompositeElement?[] axes = layer.Elements<C.AxisId>().Select(reference => groups.Resolve(reference.Val?.Value)).ToArray();
        horizontal = axes.FirstOrDefault(axis => IsHorizontal(axis) == true);
        vertical = axes.FirstOrDefault(axis => IsHorizontal(axis) == false);
        return horizontal != null && vertical != null;
    }

    private static bool? IsHorizontal(OpenXmlCompositeElement? axis) {
        C.AxisPositionValues? position = axis?.GetFirstChild<C.AxisPosition>()?.Val?.Value;
        if (position == C.AxisPositionValues.Bottom || position == C.AxisPositionValues.Top) return true;
        if (position == C.AxisPositionValues.Left || position == C.AxisPositionValues.Right) return false;
        return null;
    }

    private static OfficeChartAxisCrossingPosition ReadCrossing(OpenXmlCompositeElement? axis) {
        C.CrossesValues? value = axis?.GetFirstChild<C.Crosses>()?.Val?.Value;
        if (value == C.CrossesValues.Maximum) return OfficeChartAxisCrossingPosition.Maximum;
        if (value == C.CrossesValues.Minimum) return OfficeChartAxisCrossingPosition.Minimum;
        return OfficeChartAxisCrossingPosition.AutoZero;
    }

    private static double? ReadCrossesAt(OpenXmlCompositeElement? axis) {
        double? value = axis?.GetFirstChild<C.CrossesAt>()?.Val?.Value;
        return value.HasValue && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value) ? value : null;
    }

    /// <summary>Reads a value axis's c:dispUnits divisor and, when c:dispUnitsLbl is present, its label.</summary>
    internal static (double? Divisor, string? Label) ReadDisplayUnits(OpenXmlCompositeElement? axis) {
        C.DisplayUnits? displayUnits = axis?.GetFirstChild<C.DisplayUnits>();
        if (displayUnits == null) return default;
        C.BuiltInUnit? builtIn = displayUnits.GetFirstChild<C.BuiltInUnit>();
        string? builtInUnit = builtIn?.Val?.InnerText ?? builtIn?.Val?.Value.ToString() ?? builtIn?.InnerText;
        double? divisor = builtInUnit == null
            ? displayUnits.GetFirstChild<C.CustomDisplayUnit>()?.Val?.Value
            : GetBuiltInDivisor(builtInUnit);
        if (!divisor.HasValue || divisor.Value <= 0D || double.IsNaN(divisor.Value) || double.IsInfinity(divisor.Value))
            return default;
        C.DisplayUnitsLabel? displayUnitsLabel = displayUnits.GetFirstChild<C.DisplayUnitsLabel>();
        string? label = ReadChartTextValue(displayUnitsLabel?.GetFirstChild<C.ChartText>());
        if (displayUnitsLabel != null && string.IsNullOrWhiteSpace(label)) {
            label = builtInUnit == null
                ? divisor.Value.ToString("0.##", CultureInfo.InvariantCulture)
                : GetBuiltInLabel(builtInUnit);
        }

        return (divisor.Value, string.IsNullOrWhiteSpace(label) ? null : label!.Trim());
    }

    // The displayed text of c:tx: its rich-text runs, or the first cached point of a formula reference.
    // InnerText would also include the formula itself.
    private static string? ReadChartTextValue(C.ChartText? text) {
        if (text == null) return null;
        C.RichText? rich = text.GetFirstChild<C.RichText>();
        if (rich != null) return string.Concat(rich.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(run => run.Text));
        C.StringPoint? point = text.GetFirstChild<C.StringReference>()?.StringCache?.Elements<C.StringPoint>()
            .FirstOrDefault(item => (item.Index?.Value ?? 0U) == 0U);
        return point?.NumericValue?.Text ?? text.GetFirstChild<C.StringLiteral>()?.Elements<C.StringPoint>().FirstOrDefault()?.NumericValue?.Text;
    }

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
