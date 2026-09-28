using System;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    internal void ApplySharedPointStyles(System.Collections.Generic.IReadOnlyList<OfficeChartSeries> series) {
        for (int index = 0; index < series.Count; index++) {
            OfficeChartSeries data = series[index];
            if (data.PointStyles != null)
                ApplySeriesByIndex(index, native => OfficeOpenXmlChartPointStyles.ApplySeries(native, data));
        }
        Save();
    }

    internal void ApplySharedPointExplosions(System.Collections.Generic.IReadOnlyList<OfficeChartSeries> series) {
        for (int index = 0; index < series.Count; index++) {
            OfficeChartSeries data = series[index];
            if (data.PointExplosions != null)
                ApplySeriesByIndex(index, native => OfficeOpenXmlChartExplosions.ApplySeries(native, data));
        }
        Save();
    }

    /// <summary>Replaces a point's fill/outline overrides. Null restores series/theme inheritance.</summary>
    public ExcelChart SetDataPointStyle(int seriesIndex, uint pointIndex, OfficeChartPointStyle? style) {
        if (seriesIndex < 0) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        if (!ApplySeriesByIndex(seriesIndex, series => OfficeOpenXmlChartPointStyles.ApplyPoint(series, pointIndex, style)))
            throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        Save();
        return this;
    }

    /// <summary>Sets one pie or doughnut slice's outward offset as a percentage of its radius. Null restores the series offset; zero explicitly removes the slice offset.</summary>
    public ExcelChart SetDataPointExplosion(int seriesIndex, uint pointIndex, int? percent) {
        if (seriesIndex < 0) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        if (!ApplySeriesByIndex(seriesIndex, series => OfficeOpenXmlChartExplosions.ApplyPoint(series, pointIndex, percent)))
            throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        Save();
        return this;
    }
}
