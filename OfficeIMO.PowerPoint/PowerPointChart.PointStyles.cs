using System;
using System.Linq;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    internal void ApplyPointStyles(PowerPointChartData data) {
        var part = GetChartPart();
        var chartSpace = part.ChartSpace ?? throw new InvalidOperationException("Chart content is missing.");
        var nativeSeries = chartSpace.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>()
            .Where(element => element.LocalName == "ser" && element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            .OrderBy(element => element.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue).ToList();
        if (nativeSeries.Count != data.Series.Count) throw new InvalidOperationException("Chart series metadata does not match the native chart.");
        for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
            PowerPointChartSeries source = data.Series[seriesIndex];
            if (source.PointStyles == null && source.PointColors == null) continue;
            for (int pointIndex = 0; pointIndex < source.Values.Count; pointIndex++)
                OfficeOpenXmlChartPointStyles.ApplyPoint(nativeSeries[seriesIndex], (uint)pointIndex,
                    source.PointStyles?[pointIndex], source.PointColors?[pointIndex]);
        }
        chartSpace.Save();
    }

    /// <summary>Replaces a point's fill/outline overrides. Null restores series/theme inheritance.</summary>
    public PowerPointChart SetDataPointStyle(int seriesIndex, uint pointIndex, OfficeChartPointStyle? style) {
        if (seriesIndex < 0) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        var part = GetChartPart();
        var chartSpace = part.ChartSpace ?? throw new InvalidOperationException("Chart content is missing.");
        var series = chartSpace.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>()
            .Where(element => element.LocalName == "ser" && element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            .OrderBy(element => element.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue).ToList();
        if (seriesIndex >= series.Count) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        OfficeOpenXmlChartPointStyles.ApplyPoint(series[seriesIndex], pointIndex, style);
        chartSpace.Save();
        return this;
    }
}
