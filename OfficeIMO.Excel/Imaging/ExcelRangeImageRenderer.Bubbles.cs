using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

internal static partial class ExcelRangeImageRenderer {
    // One Excel-to-Core mapping for both PDF and image adapters. Zero-size entries retain
    // X/Y alignment and axis bounds without painting negatives that Core cannot represent.
    internal static OfficeChartSeries CreateBubbleChartSeries(ExcelChartSeries series, bool callerPalette = false) =>
        OfficeChartSeries.CreateBubble(series.Name, series.XValues!, series.Values,
            series.BubbleSizes!.Select(size => Math.Max(0D, size)),
            callerPalette && series.HasAutomaticColor ? null : ResolveArgb(series.SeriesColorArgb),
            ResolvePointColors(series.PointColorArgb),
            markerOutlineColor: ResolveArgb(series.MarkerOutlineColorArgb),
            markerOutlineWidth: series.MarkerOutlineWidth ?? series.SeriesLineWidth)
            .WithPointStyles(series.PointStyles);
}
