using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native series error bars into the shared chart model.</summary>
internal static class OfficeOpenXmlChartErrorBars {
    // Office draws error bars without their own line colour in the chart's dark text colour.
    private static readonly OfficeColor DefaultColor = OfficeColor.FromRgb(0x59, 0x59, 0x59);

    internal static IReadOnlyList<OfficeChartErrorBars>? Read(OpenXmlCompositeElement series, A.ColorScheme? scheme, int maximumPoints) {
        List<OfficeChartErrorBars>? result = null;
        bool scatter = series.Parent is C.ScatterChart or C.BubbleChart;
        // Bars are measured on the source values; a logarithmic or reversed value axis would need them restated.
        if (series.Elements<C.ErrorBars>().Any() && HasScaledValueAxis(series)) {
            OfficeOpenXmlChartFidelityScope.Approximate("Error bars on logarithmic or reversed axes cannot be projected.",
                "Error bars on a logarithmic or reversed axis are left out.");
            return null;
        }
        foreach (C.ErrorBars bars in series.Elements<C.ErrorBars>().Take(2)) {
            string? direction = Value(bars, "errDir");
            string? type = Value(bars, "errBarType");
            string? valueType = Value(bars, "errValType");
            OfficeChartErrorBarValueType kind = valueType switch {
                "fixedVal" => OfficeChartErrorBarValueType.Fixed,
                "percentage" => OfficeChartErrorBarValueType.Percentage,
                "stdDev" => OfficeChartErrorBarValueType.StandardDeviation,
                "stdErr" => OfficeChartErrorBarValueType.StandardError,
                "cust" => OfficeChartErrorBarValueType.Custom,
                _ => OfficeChartErrorBarValueType.Fixed
            };
            double amount = double.TryParse(Value(bars, "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) &&
                !double.IsNaN(parsed) && !double.IsInfinity(parsed) ? parsed : 1D;
            IReadOnlyList<double>? plus = null, minus = null;
            if (kind == OfficeChartErrorBarValueType.Custom) {
                plus = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(Child(bars, "plus"), maximumPoints);
                minus = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(Child(bars, "minus"), maximumPoints);
            }
            A.Outline? outline = bars.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.Outline>();
            if (outline?.GetFirstChild<A.NoFill>() != null) continue;
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(outline?.GetFirstChild<A.SolidFill>(), scheme);
            if (outline?.ChildElements.Any(child => child is A.GradientFill or A.PatternFill) == true)
                OfficeOpenXmlChartFidelityScope.Approximate("The error bar line fill cannot be projected.",
                    "Error bars with a gradient or pattern line are drawn in a solid colour.");
            double? width = outline?.Width?.Value is int emus && emus > 0 ? Math.Min(emus / 12700D, OfficeChartStyleBounds.MaximumLineWidthPoints) : null;
            OfficeStrokeDashStyle? dash = OfficeOpenXmlChartTrendlines.ReadDash(outline);
            if (outline?.GetFirstChild<A.PresetDash>() != null && !dash.HasValue)
                OfficeOpenXmlChartFidelityScope.Approximate("The error bar dash pattern cannot be projected.",
                    "An error bar dash pattern the renderer cannot draw is drawn solid.");
            (result ??= new List<OfficeChartErrorBars>()).Add(new OfficeChartErrorBars(kind, amount,
                type == "plus" ? OfficeChartErrorBarDirection.Plus : type == "minus" ? OfficeChartErrorBarDirection.Minus : OfficeChartErrorBarDirection.Both,
                Value(bars, "noEndCap") is not ("1" or "true"), scatter && direction == "x", plus, minus,
                color ?? DefaultColor, width, dash));
        }
        return result;
    }

    private static bool HasScaledValueAxis(OpenXmlCompositeElement series) {
        if (series.Parent is not OpenXmlCompositeElement layer || layer.Parent is not C.PlotArea plot) return false;
        var ids = new HashSet<uint>(layer.Elements<C.AxisId>().Select(reference => reference.Val?.Value).OfType<uint>());
        return plot.Elements<C.ValueAxis>().Any(axis => axis.AxisId?.Val?.Value is uint id && ids.Contains(id) &&
            (axis.Scaling?.GetFirstChild<C.LogBase>() != null ||
             axis.Scaling?.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin));
    }

    private static OpenXmlElement? Child(OpenXmlElement parent, string localName) =>
        parent.ChildElements.FirstOrDefault(child => child.LocalName == localName);

    private static string? Value(OpenXmlElement parent, string localName) =>
        Child(parent, localName)?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
}
