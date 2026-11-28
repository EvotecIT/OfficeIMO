using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Replaces a point's fill/outline overrides. Null restores series/theme inheritance.</summary>
    public WordChart SetDataPointStyle(int seriesIndex, uint pointIndex, OfficeChartPointStyle? style) {
        ArgumentOutOfRangeException.ThrowIfNegative(seriesIndex);
        OfficeXmlNode chart = CurrentChart;
        // Only plotted series count; series in extension or vendor markup, such as filtered series, do not.
        List<OfficeXmlNode>? series = chart.IsNone ? null : PlottedSeries(PlotAreaOf(chart))
            .OrderBy(element => element.Child(NsC, "idx").IsNone ? uint.MaxValue : element.Child(NsC, "idx").UIntVal() ?? uint.MaxValue).ToList();
        if (series == null || seriesIndex >= series.Count) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        OfficeOpenXmlChartPointStyles.ApplyPoint(series[seriesIndex], pointIndex, style);
        SaveChartSpace();
        return this;
    }
}
