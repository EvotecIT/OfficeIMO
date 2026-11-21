namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // Office's "Vary colors by point" applies only to a chart with one series. Each point without
    // its own colour takes the palette colour of its category, and the legend lists the categories.
    private static OfficeSeries? GetPointColoredSeries(OfficeChartSnapshot snapshot) {
        if (!snapshot.Layout.VaryColorsByPoint || snapshot.Data.Series.Count != 1) return null;
        OfficeSeries series = snapshot.Data.Series[0];
        OfficeChartType kind = series.RenderKind ?? snapshot.ChartKind;
        return IsPieChart(kind) || IsDoughnutChart(kind) || IsAreaChart(kind) ? null : series;
    }

    private static OfficeChartSnapshot ApplyVaryColorsByPoint(OfficeChartSnapshot snapshot) {
        OfficeSeries? series = GetPointColoredSeries(snapshot);
        if (series == null) return snapshot;
        var colors = new OfficeColor?[series.Values.Count];
        for (int i = 0; i < colors.Length; i++)
            colors[i] = series.PointColors != null && i < series.PointColors.Count && series.PointColors[i].HasValue
                ? series.PointColors[i]
                : GetSeriesColor(snapshot.Style, i);
        return new OfficeChartSnapshot(snapshot.Name, snapshot.Title, snapshot.ChartKind,
            new OfficeChartData(snapshot.Data.Categories, new[] { series.WithPointColors(colors) }), snapshot.Width, snapshot.Height,
            snapshot.Style, snapshot.Layout, snapshot.BubbleScalePercent, snapshot.BubbleSizeMode, snapshot.RadialLayout);
    }

    // One legend entry per category, painted like that category's point.
    private static IReadOnlyList<OfficeSeries>? GetPointColoredLegendSeries(OfficeChartSnapshot snapshot) {
        OfficeSeries? series = GetPointColoredSeries(snapshot);
        if (series == null) return null;
        var entries = new List<OfficeSeries>(snapshot.Data.Categories.Count);
        IReadOnlyCollection<int>? hidden = snapshot.Layout.HiddenCategoryLegendIndexes;
        for (int i = 0; i < snapshot.Data.Categories.Count && i < series.Values.Count; i++) {
            if (hidden?.Contains(i) == true) continue;
            OfficeColor color = GetPointStyle(series, i)?.FillColor ??
                GetPointColor(series.PointColors, i, GetSeriesColor(snapshot.Style, i));
            entries.Add(new OfficeSeries(snapshot.Data.Categories[i], new[] { 0D }, null, color, null,
                series.ShowMarkers, showInLegend: true, connectLine: series.ConnectLine, markerSize: series.MarkerSize,
                markerShape: series.MarkerShape, renderKind: series.RenderKind ?? snapshot.ChartKind));
        }
        return entries;
    }
}
