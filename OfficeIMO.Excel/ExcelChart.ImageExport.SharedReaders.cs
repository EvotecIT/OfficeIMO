using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    // The image export's reads through the shared chart readers, with the workbook theme's colour scheme as a node.
    public sealed partial class ExcelChart {
        private static bool ImageExportHasAttributes(OfficeXmlNode element) => element.HasAttributes;

        private static OfficeXmlNode ImageExportPointColoredSeries(OfficeXmlNode plotArea) =>
            OfficeOpenXmlChartPlotSemantics.GetPointColoredSeries(plotArea);

        private static OfficeColor[]? ImageExportReadRadialPalette(ExcelPartRef chartPart, OfficeXmlNode series, int pointCount,
            OfficeXmlNode scheme) =>
            OfficeOpenXmlThemeColorResolver.ReadRadialPalette(chartPart.AsChartPart(), series, pointCount, scheme,
                useSpreadsheetDefaultTheme: true);

        private static OfficeChartLayout ImageExportWithPlotSemantics(OfficeChartLayout layout, OfficeXmlNode chart) =>
            layout.WithPlotSemantics(chart, includeDisplayUnits: false);

        private static OfficeChartLayout ImageExportWithCategoryGeometry(OfficeChartLayout layout, OfficeXmlNode plotArea,
            OfficeXmlNode scheme) =>
            layout.WithCategoryGeometry(plotArea, scheme);

        private static OfficeXmlNode ImageExportResolveSecondaryAxis(OfficeXmlNode plotArea) =>
            OfficeOpenXmlChartSecondaryAxis.Resolve(plotArea);

        private static OfficeChartValueAxisLayout? ImageExportReadSecondaryAxis(OfficeXmlNode plotArea, Func<string> resolveSourceLinkedFormat) =>
            OfficeOpenXmlChartSecondaryAxis.Read(plotArea, resolveSourceLinkedFormat);

        private static OfficeChartPointStyle? ImageExportReadSeriesStyle(OfficeXmlNode series, OfficeXmlNode scheme) =>
            OfficeOpenXmlChartPointStyles.ReadSeriesStyle(series, scheme);

        private static IReadOnlyList<OfficeChartErrorBars>? ImageExportReadErrorBars(OfficeXmlNode series, OfficeXmlNode scheme,
            int maximumPoints, Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisibleAmounts) =>
            OfficeOpenXmlChartErrorBars.Read(series, scheme, maximumPoints, selectVisibleAmounts);

        // Data points where the series' schema has none (surface series) were unknown elements to the SDK, which its
        // typed point reader never saw; the export keeps ignoring them.
        private static IReadOnlyList<OfficeXmlNode> ImageExportBoundedPoints(OfficeXmlNode series) {
            IReadOnlyList<OfficeXmlNode> points = OfficeOpenXmlChartPointStyles.GetBoundedPoints(series);
            return points.Count > 0 && !IsImageExportSchemaPlaced(points[0]) ? Array.Empty<OfficeXmlNode>() : points;
        }

        private static IReadOnlyList<OfficeChartPointStyle?>? ImageExportReadPointStyles(IReadOnlyList<OfficeXmlNode> points, int count,
            OfficeXmlNode scheme) =>
            OfficeOpenXmlChartPointStyles.Read(points, count, scheme);

        private static bool ImageExportTryReadExplosions(OfficeXmlNode series, IReadOnlyList<OfficeXmlNode> points, int count,
            out int[]? explosions) =>
            OfficeOpenXmlChartExplosions.TryRead(series, points, count, out explosions);

        private static (OfficeChartPointStyle Paint, OfficeColor? Color, bool Automatic) ImageExportReadRibbonPaint(OfficeXmlNode series,
            OfficeXmlNode scheme, OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext automaticColors) =>
            OfficeOpenXmlChartPointStyles.ReadRibbonPaint(series, scheme, automaticColors, useSpreadsheetDefaultTheme: true);

        private static OfficeColor? ImageExportResolveAutomaticSeriesColor(OfficeXmlNode series, OfficeXmlNode scheme,
            OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext automaticColors) =>
            OfficeOpenXmlChartAutomaticColors.ResolveSeries(series, scheme, automaticColors, useSpreadsheetDefaultTheme: true);

        private static string? ImageExportReadTrendlineLegendName(OfficeXmlNode series, OfficeXmlNode trendline,
            OfficeChartTrendlineType type, int? order, int? period, string seriesName) =>
            OfficeOpenXmlChartTrendlines.ReadLegendName(series, trendline, type, order, period, seriesName);

        private static OfficeStrokeDashStyle? ImageExportReadTrendlineDash(OfficeXmlNode outline, bool projecting) =>
            projecting ? OfficeOpenXmlChartTrendlines.ReadDashOrApproximate(outline) : OfficeOpenXmlChartTrendlines.ReadDash(outline);

        // Both callers pass a data point's or its marker's c:spPr, which the SDK always read as ChartShapeProperties.
        private static bool ImageExportIsSupportedPointStyle(OfficeXmlNode properties, OfficeXmlNode scheme) =>
            OfficeOpenXmlChartPointStyles.IsSupported(properties, scheme);

        private static (double? Divisor, string? Label) ImageExportReadDisplayUnits(OfficeXmlNode axis) =>
            OfficeOpenXmlChartPlotSemantics.ReadDisplayUnits(axis);

        private static OfficeOpenXmlChartTextReader.NativeText ImageExportReadNativeText(OfficeXmlNode chart, OfficeXmlNode title,
            OfficeXmlNode scheme) =>
            OfficeOpenXmlChartTextReader.ReadNativeText(chart, title, scheme);
    }
}
