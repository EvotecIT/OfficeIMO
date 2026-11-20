namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    // A legend with a manual layout keeps the arrangement of its position (a band for top and bottom,
    // a column otherwise) and moves to, or resizes into, the rectangle the layout gives it.
    private static void AddPlacedSeriesLegend(OfficeDrawing drawing, IReadOnlyList<OfficeSeries> legendSeries,
        OfficeChartManualLayout legendLayout, double width, double height, double contentTop, double topLegendHeight,
        double bottomLegendHeight, double legendWidth, double plotTop, double plotHeight, OfficeChartStyle style,
        OfficeChartLayout layout, OfficeChartType chartKind) {
        PlacedLegend placed = ResolvePlacedLegend(legendLayout, layout, width, height, contentTop, topLegendHeight,
            bottomLegendHeight, legendWidth, plotTop, plotHeight);
        if (placed.Band) {
            AddSeriesLegendBand(drawing, legendSeries, placed.Left, placed.Top, placed.Width, style, layout, chartKind, placed.Height);
        } else {
            AddSeriesLegend(drawing, legendSeries, placed.Left, placed.Top, placed.Width, placed.Height, style, layout, chartKind);
        }
    }

    // Pie and doughnut legends list categories and are placed the same way.
    private static void AddPlacedCategoryLegend(OfficeDrawing drawing, IReadOnlyList<string> categories,
        OfficeChartManualLayout legendLayout, double width, double height, double contentTop, double topLegendHeight,
        double bottomLegendHeight, double legendWidth, double plotTop, double plotHeight, OfficeChartStyle style,
        OfficeChartLayout layout, IReadOnlyList<OfficeColor?>? pointColors, IReadOnlyList<OfficeChartPointStyle?>? pointStyles) {
        PlacedLegend placed = ResolvePlacedLegend(legendLayout, layout, width, height, contentTop, topLegendHeight,
            bottomLegendHeight, legendWidth, plotTop, plotHeight);
        if (placed.Band) {
            AddCategoryLegendBand(drawing, categories, placed.Left, placed.Top, placed.Width, style, layout, pointColors, pointStyles);
        } else {
            AddCategoryLegend(drawing, categories, placed.Left, placed.Top, placed.Width, placed.Height, style, layout, pointColors, pointStyles);
        }
    }

    private readonly struct PlacedLegend {
        internal PlacedLegend(double left, double top, double width, double height, bool band) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
            Band = band;
        }

        internal double Left { get; }
        internal double Top { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal bool Band { get; }
    }

    private static PlacedLegend ResolvePlacedLegend(OfficeChartManualLayout legendLayout, OfficeChartLayout layout,
        double width, double height, double contentTop, double topLegendHeight, double bottomLegendHeight,
        double legendWidth, double plotTop, double plotHeight) {
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
        return new PlacedLegend(left, top, placedWidth, placedHeight, band);
    }

    // Office keeps a radial plot area square: the pie, doughnut or radar web is centred in the placed rectangle
    // and sized by its shorter side. An outer rectangle also holds the radar's category labels; a pie has no
    // axis labels, so its inner and outer rectangles are the same and outside data labels fall beyond it.
    private static bool PlaceRadialPlotArea(OfficeChartLayout layout, double width, double height,
        double labelWidth, double labelHeight, ref double centerX, ref double centerY, ref double radius) {
        if (layout.PlotAreaLayout is not OfficeChartManualLayout plotArea) return false;
        if (plotArea.Target != OfficeChartLayoutTarget.Outer) labelWidth = labelHeight = 0D;
        (double left, double placedWidth) = plotArea.Resolve(horizontal: true, width,
            centerX - radius - labelWidth / 2D, radius * 2D + labelWidth);
        (double top, double placedHeight) = plotArea.Resolve(horizontal: false, height,
            centerY - radius - labelHeight / 2D, radius * 2D + labelHeight);
        double side = Math.Min(placedWidth - labelWidth, placedHeight - labelHeight);
        if ((labelWidth > 0D || labelHeight > 0D) && side < 8D)
            throw new NotSupportedException("Radar category labels need a larger plot area.");
        centerX = left + placedWidth / 2D;
        centerY = top + placedHeight / 2D;
        radius = Math.Max(0.5D, side / 2D);
        return true;
    }

    private static double MeasureChartTextWidth(string text, string? fontFamily, double fontSize, OfficeFontStyle fontStyle) {
        var fontInfo = new OfficeFontInfo(
            string.IsNullOrWhiteSpace(fontFamily) ? OfficeFontInfo.Default.FamilyName : fontFamily!, fontSize, fontStyle);
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(fontInfo);
        return measurer.MeasureWidth(text, measurer.CreateStyle(fontInfo)) * OfficeTextMeasurer.PointsPerInch / OfficeTextMeasurer.DefaultDpi;
    }
}
