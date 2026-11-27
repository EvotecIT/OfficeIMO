using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    internal void ApplyPointStyles(PowerPointChartData data) {
        List<OfficeXmlNode> nativeSeries = SeriesInIndexOrder();
        if (nativeSeries.Count != data.Series.Count) throw new InvalidOperationException("Chart series metadata does not match the native chart.");
        for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
            PowerPointChartSeries source = data.Series[seriesIndex];
            if (source.PointStyles == null && source.PointColors == null) continue;
            OfficeOpenXmlChartPointStyles.ApplyPoints(nativeSeries[seriesIndex],
                source.PointStyles, source.PointColors, source.Values.Count);
        }
        MarkChartEdited();
    }

    /// <summary>Replaces a point's fill/outline overrides. Null restores series/theme inheritance.</summary>
    public PowerPointChart SetDataPointStyle(int seriesIndex, uint pointIndex, OfficeChartPointStyle? style) {
        ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
        List<OfficeXmlNode> series = SeriesInIndexOrder();
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(seriesIndex, series.Count);
        OfficeOpenXmlChartPointStyles.ApplyPoint(series[seriesIndex], pointIndex, style);
        MarkChartEdited();
        return this;
    }

    // Every c:ser of the chart space, ordered by c:idx, those without one last.
    private List<OfficeXmlNode> SeriesInIndexOrder() {
        OfficeXmlNode chartSpace = ChartSpaceNode;
        if (chartSpace.IsNone) throw new InvalidOperationException("Chart content is missing.");
        return chartSpace.ContentDescendants(OfficeXmlNamespaces.Chart, "ser")
            .OrderBy(element => element.Child(OfficeXmlNamespaces.Chart, "idx").UIntVal() ?? uint.MaxValue).ToList();
    }
}
