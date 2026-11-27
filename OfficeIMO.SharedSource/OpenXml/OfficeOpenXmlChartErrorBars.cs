using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Reads native series error bars into the shared chart model.</summary>
internal static class OfficeOpenXmlChartErrorBars {
    // Office draws error bars without their own line colour in the chart's dark text colour.
    private static readonly OfficeColor DefaultColor = OfficeColor.FromRgb(0x59, 0x59, 0x59);

    /// <param name="series">The native series element.</param>
    /// <param name="scheme">The theme colours that resolve scheme colours.</param>
    /// <param name="maximumPoints">The most cached custom amounts read per direction.</param>
    /// <param name="selectVisible">Keeps the cached custom amounts of the points a visible-cells-only chart plots,
    /// given the amounts' source element (c:plus or c:minus); null when every cached point is plotted.</param>
    internal static IReadOnlyList<OfficeChartErrorBars>? Read(OfficeXmlNode series, OfficeXmlNode scheme, int maximumPoints,
        Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisible = null) {
        List<OfficeChartErrorBars>? result = null;
        bool scatter = series.Parent.NamespaceId == NsC && series.Parent.LocalName is "scatterChart" or "bubbleChart";
        int read = 0;
        foreach (OfficeXmlNode bars in series.Children(NsC, "errBars")) {
            if (read++ == 2) break;
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
            double amount = OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(Value(bars, "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) &&
                !double.IsNaN(parsed) && !double.IsInfinity(parsed) ? parsed : 1D;
            IReadOnlyList<double>? plus = null, minus = null;
            if (kind == OfficeChartErrorBarValueType.Custom) {
                plus = ReadCustomAmounts(Child(bars, "plus"), maximumPoints, selectVisible);
                minus = ReadCustomAmounts(Child(bars, "minus"), maximumPoints, selectVisible);
            }
            OfficeXmlNode outline = bars.Child(NsC, "spPr").Child(NsA, "ln");
            if (outline.Child(NsA, "noFill").Exists) continue;
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(outline.Child(NsA, "solidFill"), scheme);
            if (outline.Child(NsA, "gradFill").Exists || outline.Child(NsA, "pattFill").Exists)
                Approximate("The error bar line fill cannot be projected.",
                    "Error bars with a gradient or pattern line are drawn in a solid colour.");
            double? width = outline.IntAttr("w") is int emus && emus > 0 ? Math.Min(emus / 12700D, OfficeChartStyleBounds.MaximumLineWidthPoints) : null;
            OfficeStrokeDashStyle? dash = OfficeOpenXmlChartTrendlines.ReadDash(outline);
            if ((outline.Child(NsA, "prstDash").Exists || outline.Child(NsA, "custDash").Exists) && !dash.HasValue)
                Approximate("The error bar dash pattern cannot be projected.",
                    "An error bar dash pattern the renderer cannot draw is drawn solid.");
            (result ??= new List<OfficeChartErrorBars>()).Add(new OfficeChartErrorBars(kind, amount,
                type == "plus" ? OfficeChartErrorBarDirection.Plus : type == "minus" ? OfficeChartErrorBarDirection.Minus : OfficeChartErrorBarDirection.Both,
                Value(bars, "noEndCap") is not ("1" or "true"), scatter && direction == "x", plus, minus,
                color ?? DefaultColor, width, dash));
        }
        return result;
    }

    // How the bars would be drawn matters only while a chart is projected for rendering; reading chart data alone
    // never fails on it.
    private static void Approximate(string strictReason, string approximation) {
        if (OfficeOpenXmlChartFidelityScope.IsProjecting) OfficeOpenXmlChartFidelityScope.Approximate(strictReason, approximation);
    }

    private static IReadOnlyList<double>? ReadCustomAmounts(OfficeXmlNode source, int maximumPoints,
        Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisible) {
        IReadOnlyList<double>? amounts = OfficeOpenXmlChartCacheReader.ReadCachedNumbers(source, maximumPoints);
        return amounts != null && source.Exists && selectVisible != null ? selectVisible(source, amounts) : amounts;
    }

    // The first child with this local name, in any namespace.
    private static OfficeXmlNode Child(OfficeXmlNode parent, string localName) {
        foreach (OfficeXmlNode child in parent.Children())
            if (child.LocalName == localName) return child;
        return OfficeXmlNode.None;
    }

    // The first attribute named val, in any namespace, of that child.
    private static string? Value(OfficeXmlNode parent, string localName) {
        OfficeXmlNode child = Child(parent, localName);
        if (child.IsNone) return null;
        foreach (OfficeXmlAttribute attribute in child.Attributes())
            if (attribute.LocalName == "val") return attribute.Value;
        return null;
    }

    // SDK callers
    internal static IReadOnlyList<OfficeChartErrorBars>? Read(OpenXmlCompositeElement series, A.ColorScheme? scheme, int maximumPoints,
        Func<OpenXmlElement, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisible = null) =>
        Read(OfficeXmlNode.FromSdk(series), OfficeXmlNode.FromSdk(scheme), maximumPoints,
            selectVisible == null ? null : new Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>(
                (source, amounts) => selectVisible((OpenXmlElement)source.SdkElement!, amounts)));
}
