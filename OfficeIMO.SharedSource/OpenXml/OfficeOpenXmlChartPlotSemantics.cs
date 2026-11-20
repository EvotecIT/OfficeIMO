using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads chart-wide plotting rules that every native chart adapter projects the same way.</summary>
internal static class OfficeOpenXmlChartPlotSemantics {
    /// <summary>Applies the chart's blank-cell rule, the physical axis crossings and the value-axis display units.</summary>
    internal static OfficeChartLayout WithPlotSemantics(this OfficeChartLayout layout, C.Chart? chart, bool includeDisplayUnits = true) {
        if (chart == null) return layout;
        layout = layout.WithBlankDisplay(ReadBlankDisplay(chart));
        C.ShowDataLabelsOverMaximum? overMaximum = chart.GetFirstChild<C.ShowDataLabelsOverMaximum>();
        layout = layout.WithDataLabelsOverMaximum(overMaximum == null || overMaximum.Val?.Value != false);
        if (!TryGetPhysicalAxes(chart.PlotArea, out OpenXmlCompositeElement? horizontal, out OpenXmlCompositeElement? vertical))
            return layout;
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

    /// <summary>Resolves the text Office draws for a c:title that has no text of its own.</summary>
    /// <remarks>Office titles a chart with one series by that series' name. Otherwise it shows a placeholder in the
    /// user interface language; the placeholder is returned only when the chart's c:lang is English or absent, so a
    /// chart in another language is not given English text. Returns null when the chart has no such title, when its
    /// automatic title was deleted (c:autoTitleDeleted), or when no text applies.</remarks>
    internal static string? ReadAutomaticTitle(C.Chart chart) {
        C.Title? title = chart.GetFirstChild<C.Title>();
        if (title == null || title.GetFirstChild<C.ChartText>() != null) return null;
        C.AutoTitleDeleted? deleted = chart.GetFirstChild<C.AutoTitleDeleted>();
        if (deleted != null && deleted.Val?.Value != false) return null;
        OpenXmlCompositeElement[] series = chart.PlotArea?.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(layer => layer.LocalName.EndsWith("Chart", StringComparison.Ordinal))
            .SelectMany(layer => layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser"))
            .Take(2).ToArray() ?? Array.Empty<OpenXmlCompositeElement>();
        if (series.Length == 0) return null;
        if (series.Length == 1) {
            string? name = ReadSeriesName(series[0]);
            if (!string.IsNullOrWhiteSpace(name)) return name!.Trim();
        }

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
        string? label = displayUnitsLabel?.GetFirstChild<C.ChartText>()?.InnerText;
        if (displayUnitsLabel != null && string.IsNullOrWhiteSpace(label)) {
            label = builtInUnit == null
                ? divisor.Value.ToString("0.##", CultureInfo.InvariantCulture)
                : GetBuiltInLabel(builtInUnit);
        }

        return (divisor.Value, string.IsNullOrWhiteSpace(label) ? null : label!.Trim());
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
