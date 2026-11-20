namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // A legend with a manual layout keeps the arrangement of its position (a band for top and bottom,
    // a column otherwise) and moves to, or resizes into, the rectangle the layout gives it.
    private static void AddPlacedSeriesLegend(OfficeDrawing drawing, IReadOnlyList<OfficeSeries> legendSeries,
        OfficeChartManualLayout legendLayout, double width, double height, double contentTop, double topLegendHeight,
        double bottomLegendHeight, double legendWidth, double plotTop, double plotHeight, OfficeChartStyle style,
        OfficeChartLayout layout, OfficeChartType chartKind) {
        OfficeChartLegendPosition position = layout.LegendPosition;
        bool band = position is OfficeChartLegendPosition.Top or OfficeChartLegendPosition.Bottom;
        double automaticLeft, automaticTop, automaticWidth, automaticHeight;
        if (band) {
            automaticLeft = 8D;
            automaticWidth = Math.Max(1D, width - 16D);
            automaticHeight = Math.Max(1D, position == OfficeChartLegendPosition.Top ? topLegendHeight : bottomLegendHeight);
            automaticTop = position == OfficeChartLegendPosition.Top ? contentTop + 2D : height - bottomLegendHeight + 2D;
        } else {
            automaticLeft = position == OfficeChartLegendPosition.Left ? 6D : width - legendWidth + 6D;
            automaticWidth = Math.Max(1D, legendWidth - 12D);
            automaticTop = plotTop;
            automaticHeight = plotHeight;
        }

        (double left, double placedWidth) = legendLayout.Resolve(horizontal: true, width, automaticLeft, automaticWidth);
        (double top, double placedHeight) = legendLayout.Resolve(horizontal: false, height, automaticTop, automaticHeight);
        if (band) {
            AddSeriesLegendBand(drawing, legendSeries, left, top, placedWidth, style, layout, chartKind);
        } else {
            AddSeriesLegend(drawing, legendSeries, left, top, placedWidth, placedHeight, style, layout, chartKind);
        }
    }

    private static double MeasureChartTextWidth(string text, string? fontFamily, double fontSize, OfficeFontStyle fontStyle) {
        var fontInfo = new OfficeFontInfo(
            string.IsNullOrWhiteSpace(fontFamily) ? OfficeFontInfo.Default.FamilyName : fontFamily!, fontSize, fontStyle);
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(fontInfo);
        return measurer.MeasureWidth(text, measurer.CreateStyle(fontInfo)) * OfficeTextMeasurer.PointsPerInch / OfficeTextMeasurer.DefaultDpi;
    }
}
