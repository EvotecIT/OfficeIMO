using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Replaces a point's fill/outline overrides. Null restores series/theme inheritance.</summary>
    public WordChart SetDataPointStyle(int seriesIndex, uint pointIndex, OfficeChartPointStyle? style) {
        if (seriesIndex < 0) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        var chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
        var series = chart?.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>()
            .Where(element => element.LocalName == "ser" && element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/chart")
            .OrderBy(element => element.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue).ToList();
        if (series == null || seriesIndex >= series.Count) throw new ArgumentOutOfRangeException(nameof(seriesIndex));
        OfficeOpenXmlChartPointStyles.ApplyPoint(series[seriesIndex], pointIndex, style);
        _chartPart?.ChartSpace?.Save();
        return this;
    }
}
