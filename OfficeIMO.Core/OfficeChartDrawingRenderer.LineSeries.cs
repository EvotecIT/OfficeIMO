using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void AddLineSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft,
        double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        ValueRange? sharedValueAxisRange = null, OfficeChartAxisGroup? axisGroup = null) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeSeries> series = snapshot.Data.Series;
        List<(OfficeSeries Series, int SourceIndex, OfficeChartType Kind)> lineSeries = GetRenderableLineSeries(snapshot);
        if (axisGroup.HasValue) lineSeries = lineSeries.Where(item =>
            item.Series.AxisGroup == axisGroup.Value).ToList();
        if (categories.Count == 0 || lineSeries.Count == 0) {
            return;
        }

        ValueRange range = sharedValueAxisRange ?? ApplyValueAxisScale(GetLineSeriesRenderRange(snapshot, lineSeries, categories.Count, layout), layout, horizontal: false);
        bool clipPlot = HasExplicitValueBounds(layout, axisGroup, IsBarChart(snapshot.ChartKind));
        var plotBounds = new ChartPlotBounds(plotLeft, plotTop, plotWidth, plotHeight);
        OfficeDrawing geometry = clipPlot ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
        OfficeDrawing labels = clipPlot ? new OfficeDrawing(drawing.Width, drawing.Height) : drawing;
        if (layout.ShowHighLowLines) {
            AddHighLowLines(geometry, snapshot, lineSeries, plotLeft, plotTop, plotWidth, plotHeight, range, style, layout);
        }
        var positiveCumulativeByKind = new Dictionary<OfficeChartType, double[]>();
        var negativeCumulativeByKind = new Dictionary<OfficeChartType, double[]>();
        var stackedSeriesByKind = lineSeries
            .Where(item => IsStackedLineChart(item.Kind) || IsPercentStackedLineChart(item.Kind))
            .GroupBy(item => item.Kind)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Series).ToList());
        var percentStackedTotalsByKind = stackedSeriesByKind
            .Where(group => IsPercentStackedLineChart(group.Key))
            .ToDictionary(group => group.Key, group => BuildPercentStackedTotals(group.Value, categories.Count));
        foreach ((OfficeSeries currentSeries, int sourceSeriesIndex, OfficeChartType kind) in lineSeries) {
            bool currentStacked = IsStackedLineChart(kind) || IsPercentStackedLineChart(kind);
            bool currentPercentStacked = IsPercentStackedLineChart(kind);
            double[] positiveCumulative = currentStacked ? GetCumulative(positiveCumulativeByKind, kind, categories.Count) : Array.Empty<double>();
            double[] negativeCumulative = currentStacked ? GetCumulative(negativeCumulativeByKind, kind, categories.Count) : Array.Empty<double>();
            OfficeColor color = GetSeriesColor(style, series, sourceSeriesIndex);
            double strokeWidth = GetSeriesStrokeWidth(currentSeries, 1.75D);
            OfficeStrokeDashStyle dashStyle = GetSeriesStrokeDashStyle(currentSeries);
            var points = new OfficePoint[categories.Count];
            var plotted = new bool[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(currentSeries, i, out double value)) {
                    continue;
                }

                double rawValue = currentPercentStacked ? NormalizePercentStackedValue(percentStackedTotalsByKind[kind], i, value) : value;
                double baseline = currentStacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double plottedValue = currentStacked ? baseline + rawValue : value;

                points[i] = new OfficePoint(GetCartesianCategoryPointX(snapshot, plotLeft, plotWidth, i, categories.Count, layout),
                    clipPlot ? plottedValue
                        : ToPlotY(plottedValue, range.Min, range.Max, plotTop, plotHeight));
                plotted[i] = true;
            }

            if (currentSeries.ConnectLine && currentSeries.SeriesStyle?.ShowOutline != false) {
                if (currentSeries.SeriesStyle?.OutlineJoin is OfficeStrokeLineJoin join) {
                    AddJoinedSeriesOutline(geometry, points, currentSeries.SeriesStyle.OutlineColor ?? color,
                        strokeWidth, dashStyle, join, plotted, plotBounds: clipPlot ? plotBounds : null,
                        valueRange: clipPlot ? range : null);
                } else {
                    for (int i = 1; i < categories.Count; i++) {
                        if (!plotted[i - 1] || !plotted[i]) {
                            continue;
                        }

                        OfficePoint start = points[i - 1], end = points[i];
                        if (clipPlot && !TryClipValueSegment(start, end, range, plotBounds, out start, out end)) continue;
                        double x1 = start.X;
                        double y1 = start.Y;
                        double x2 = end.X;
                        double y2 = end.Y;
                        if (x1 == x2 && y1 == y2) continue;
                        double minX = Math.Min(x1, x2);
                        double minY = Math.Min(y1, y2);
                        AddShape(geometry, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, currentSeries.SeriesStyle?.OutlineColor ?? color, currentSeries.SeriesStyle?.OutlineWidth?.ToPoints() ?? strokeWidth, dashStyle);
                    }
                }
            }

            for (int i = 0; i < categories.Count; i++) {
                if (!plotted[i]) {
                    continue;
                }

                if (clipPlot && (points[i].Y < range.Min || points[i].Y > range.Max)) continue;
                OfficePoint renderPoint = clipPlot
                    ? new OfficePoint(points[i].X, ToPlotY(points[i].Y, range.Min, range.Max, plotTop, plotHeight))
                    : points[i];

                if (layout.ShowMarkers && currentSeries.ShowMarkers) {
                    OfficeColor pointColor = GetPointColor(currentSeries.PointColors, i, color);
                    AddMarker(geometry, currentSeries, renderPoint, 4D, pointColor, 1D, GetPointStyle(currentSeries.PointStyles, i));
                }

                double value = GetSeriesValue(currentSeries, i);
                AddPointDataLabel(
                    labels,
                    layout,
                    style,
                    categories[i],
                    currentSeries,
                    value,
                    GetDataLabelCategoryTotal(series, i),
                    renderPoint.X,
                    renderPoint.Y,
                    sourceSeriesIndex,
                    i);
            }

            if (currentStacked) {
                for (int i = 0; i < categories.Count; i++) {
                    if (!TryGetSeriesValue(currentSeries, i, out double seriesValue)) {
                        continue;
                    }

                    double value = currentPercentStacked ? NormalizePercentStackedValue(percentStackedTotalsByKind[kind], i, seriesValue) : seriesValue;
                    if (value >= 0D) {
                        positiveCumulative[i] += value;
                    } else {
                        negativeCumulative[i] += value;
                    }
                }
            }
        }
        if (clipPlot) AddClippedPlotGeometry(drawing, geometry, labels, plotBounds);
    }

}
