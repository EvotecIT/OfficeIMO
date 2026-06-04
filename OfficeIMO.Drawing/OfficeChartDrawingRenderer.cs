using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Renders dependency-free chart snapshots into vector drawing primitives shared by OfficeIMO exporters.
/// </summary>
public static partial class OfficeChartDrawingRenderer {
    /// <summary>
    /// Renders a chart snapshot into an <see cref="OfficeDrawing"/> scene.
    /// </summary>
    /// <param name="snapshot">Chart snapshot to render.</param>
    /// <returns>Vector drawing containing the chart plot area and series marks.</returns>
    public static OfficeDrawing Render(OfficeChartSnapshot snapshot) {
        if (snapshot == null) {
            throw new ArgumentNullException(nameof(snapshot));
        }

        double width = Math.Min(420D, Math.Max(240D, snapshot.WidthPoints));
        double height = Math.Min(260D, Math.Max(150D, snapshot.HeightPoints));
        OfficeChartStyle style = snapshot.Style;
        OfficeChartLayout layout = snapshot.Layout;
        var drawing = new OfficeDrawing(width, height);

        AddShape(drawing, OfficeShape.Rectangle(width, height), 0D, 0D, style.BackgroundColor, style.BorderColor, 0.75D);
        double contentTop = 0D;
        if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
            double titleHeight = Math.Min(22D, Math.Max(16D, height * 0.12D));
            drawing.AddText(
                snapshot.Title!,
                8D,
                5D,
                Math.Max(1D, width - 16D),
                Math.Max(1D, titleHeight - 4D),
                new OfficeFontInfo(style.FontFamily, Math.Min(12D, Math.Max(8D, titleHeight - 7D)), OfficeFontStyle.Bold),
                style.TitleColor,
                OfficeTextAlignment.Center);
            contentTop = titleHeight;
        }

        if (IsPieChart(snapshot.ChartKind) || IsDoughnutChart(snapshot.ChartKind)) {
            AddPieSeries(drawing, snapshot, width, height, contentTop, IsDoughnutChart(snapshot.ChartKind), style, layout);
            return drawing;
        }

        if (IsRadarChart(snapshot.ChartKind)) {
            AddRadarSeries(drawing, snapshot, width, height, contentTop, style, layout);
            return drawing;
        }

        double plotLeft = 36D;
        double plotTop = 18D + contentTop;
        double legendWidth = GetSeriesLegendWidth(snapshot.Data.Series, width, layout);
        double plotRight = 12D + legendWidth;
        double plotBottom = 40D;
        double plotWidth = Math.Max(20D, width - plotLeft - plotRight);
        double plotHeight = Math.Max(20D, height - plotTop - plotBottom);
        double plotBottomY = plotTop + plotHeight;
        ValueRange axisRange = GetCartesianValueRange(snapshot);

        AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, plotBottomY, null, style.AxisColor, 0.75D);
        AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), plotLeft, plotTop, null, style.AxisColor, 0.75D);
        for (int i = 1; i <= 3; i++) {
            double y = plotTop + plotHeight * i / 4D;
            AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, y, null, style.GridLineColor, 0.5D);
        }

        if (IsAreaChart(snapshot.ChartKind)) {
            AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style);
        } else if (IsScatterChart(snapshot.ChartKind)) {
            AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style);
        } else if (IsLineChart(snapshot.ChartKind)) {
            AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style);
        } else {
            AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style);
        }

        if (IsBarChart(snapshot.ChartKind)) {
            AddHorizontalCategoryAxisLabels(drawing, snapshot.Data.Categories, plotLeft, plotTop, plotHeight, style, layout);
            AddHorizontalValueAxisLabels(drawing, axisRange, plotLeft, plotBottomY, plotWidth, style, layout);
        } else {
            AddValueAxisLabels(drawing, axisRange, plotLeft, plotTop, plotHeight, style, layout);
            AddCategoryAxisLabels(drawing, snapshot.Data.Categories, plotLeft, plotBottomY, plotWidth, style, layout);
        }

        AddSeriesLegend(drawing, snapshot.Data.Series, width - legendWidth + 6D, plotTop, Math.Max(0D, legendWidth - 12D), plotHeight, style, layout);
        return drawing;
    }

    /// <summary>
    /// Renders a chart snapshot and returns reusable drawing quality diagnostics for the rendered scene.
    /// </summary>
    /// <param name="snapshot">Chart snapshot to render.</param>
    /// <param name="qualityOptions">Optional drawing quality analysis options.</param>
    /// <returns>Rendered chart drawing plus quality report.</returns>
    public static OfficeChartRenderingResult RenderWithQuality(OfficeChartSnapshot snapshot, OfficeDrawingQualityOptions? qualityOptions = null) {
        OfficeDrawing drawing = Render(snapshot);
        OfficeDrawingQualityReport qualityReport = OfficeDrawingQualityAnalyzer.Analyze(drawing, qualityOptions);
        return new OfficeChartRenderingResult(drawing, qualityReport);
    }

    /// <summary>
    /// Gets the default premium chart palette color for the zero-based series or slice index.
    /// </summary>
    public static OfficeColor GetSeriesColor(int index) {
        return OfficeChartStyle.Default.GetSeriesColor(index);
    }

    private static OfficeColor GetSeriesColor(OfficeChartStyle style, int index) => style.GetSeriesColor(index);

    private static void AddBarSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        double slot = plotWidth / categories.Count;
        double groupWidth = slot * 0.68D;
        bool horizontal = IsBarChart(snapshot.ChartKind);
        bool stacked = IsStackedBarOrColumnChart(snapshot.ChartKind) || IsPercentStackedBarOrColumnChart(snapshot.ChartKind);
        bool percentStacked = IsPercentStackedBarOrColumnChart(snapshot.ChartKind);
        double barWidth = Math.Max(2D, stacked ? groupWidth : groupWidth / series.Count);
        ValueRange range = percentStacked
            ? new ValueRange(0D, 1D)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetFiniteSeriesRange(series);
        double min = Math.Min(0D, range.Min);
        double max = Math.Max(0D, range.Max);
        if (max <= min) {
            max = min + 1D;
        }

        for (int category = 0; category < categories.Count; category++) {
            double positiveBase = 0D;
            double negativeBase = 0D;
            double percentTotal = percentStacked ? GetPositiveCategoryTotal(series, category) : 0D;
            for (int s = 0; s < series.Count; s++) {
                double value = GetSeriesValue(series[s], category);
                if (value == 0D) {
                    continue;
                }

                double baseline = 0D;
                double plottedValue = value;
                if (stacked) {
                    if (percentStacked) {
                        plottedValue = percentTotal <= 0D ? 0D : Math.Max(0D, value) / percentTotal;
                    }

                    baseline = plottedValue >= 0D ? positiveBase : negativeBase;
                    if (plottedValue >= 0D) {
                        positiveBase += plottedValue;
                    } else {
                        negativeBase += plottedValue;
                    }
                }

                OfficeColor color = GetSeriesColor(style, s);
                if (horizontal) {
                    double categoryHeight = plotHeight / categories.Count;
                    double rowHeight = Math.Max(2D, categoryHeight * 0.68D / (stacked ? 1D : series.Count));
                    double y = plotTop + categoryHeight * category + categoryHeight * 0.16D + (stacked ? 0D : rowHeight * s);
                    double x1 = ToPlotX(baseline, min, max, plotLeft, plotWidth);
                    double x2 = ToPlotX(stacked ? baseline + plottedValue : plottedValue, min, max, plotLeft, plotWidth);
                    double x = Math.Min(x1, x2);
                    double w = Math.Max(1D, Math.Abs(x2 - x1));
                    AddShape(drawing, OfficeShape.Rectangle(w, rowHeight), x, y, color, null, 0D);
                } else {
                    double x = plotLeft + slot * category + (slot - groupWidth) / 2D + (stacked ? 0D : barWidth * s);
                    double y1 = ToPlotY(baseline, min, max, plotTop, plotHeight);
                    double y2 = ToPlotY(stacked ? baseline + plottedValue : plottedValue, min, max, plotTop, plotHeight);
                    double y = Math.Min(y1, y2);
                    double h = Math.Max(1D, Math.Abs(y2 - y1));
                    AddShape(drawing, OfficeShape.Rectangle(barWidth * 0.88D, h), x, y, color, null, 0D);
                }
            }
        }
    }

    private static void AddAreaSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count < 2 || series.Count == 0) {
            return;
        }

        bool stacked = IsStackedAreaChart(snapshot.ChartKind) || IsPercentStackedAreaChart(snapshot.ChartKind);
        bool percentStacked = IsPercentStackedAreaChart(snapshot.ChartKind);
        ValueRange range = percentStacked
            ? new ValueRange(0D, 1D)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetFiniteSeriesRange(series);
        double step = plotWidth / (categories.Count - 1);
        var positiveCumulative = new double[categories.Count];
        var negativeCumulative = new double[categories.Count];

        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, s);
            var topPoints = new List<OfficePoint>(categories.Count);
            var bottomPoints = new List<OfficePoint>(categories.Count);

            for (int i = 0; i < categories.Count; i++) {
                double value = GetSeriesValue(series[s], i);
                double rawValue = percentStacked ? Math.Max(0D, value) : value;
                double baseline = stacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double topValue = baseline + rawValue;

                if (percentStacked) {
                    double total = GetPositiveCategoryTotal(series, i);
                    baseline = total <= 0D ? 0D : baseline / total;
                    topValue = total <= 0D ? 0D : topValue / total;
                }

                double x = plotLeft + step * i;
                topPoints.Add(new OfficePoint(x, ToPlotY(topValue, range.Min, range.Max, plotTop, plotHeight)));
                bottomPoints.Add(new OfficePoint(x, ToPlotY(baseline, range.Min, range.Max, plotTop, plotHeight)));
            }

            var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
            areaPoints.AddRange(topPoints);
            for (int i = bottomPoints.Count - 1; i >= 0; i--) {
                areaPoints.Add(bottomPoints[i]);
            }

            AddPolygonShape(drawing, areaPoints, color, color, 0.5D, 0.32D);
            AddPointLine(drawing, topPoints, color, 1.4D);

            if (stacked) {
                for (int i = 0; i < categories.Count; i++) {
                    double value = percentStacked ? Math.Max(0D, GetSeriesValue(series[s], i)) : GetSeriesValue(series[s], i);
                    if (value >= 0D) {
                        positiveCumulative[i] += value;
                    } else {
                        negativeCumulative[i] += value;
                    }
                }
            }
        }
    }

    private static void AddLineSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count < 2 || series.Count == 0) {
            return;
        }

        bool stacked = IsStackedLineChart(snapshot.ChartKind) || IsPercentStackedLineChart(snapshot.ChartKind);
        bool percentStacked = IsPercentStackedLineChart(snapshot.ChartKind);
        ValueRange range = percentStacked
            ? new ValueRange(0D, 1D)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetFiniteSeriesRange(series);
        double step = plotWidth / (categories.Count - 1);
        var positiveCumulative = new double[categories.Count];
        var negativeCumulative = new double[categories.Count];
        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, s);
            var points = new OfficePoint[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                double value = GetSeriesValue(series[s], i);
                double rawValue = percentStacked ? Math.Max(0D, value) : value;
                double baseline = stacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double plottedValue = stacked ? baseline + rawValue : value;
                if (percentStacked) {
                    double total = GetPositiveCategoryTotal(series, i);
                    plottedValue = total <= 0D ? 0D : plottedValue / total;
                }

                points[i] = new OfficePoint(plotLeft + step * i, ToPlotY(plottedValue, range.Min, range.Max, plotTop, plotHeight));
            }

            for (int i = 1; i < categories.Count; i++) {
                double x1 = points[i - 1].X;
                double y1 = points[i - 1].Y;
                double x2 = points[i].X;
                double y2 = points[i].Y;
                double minX = Math.Min(x1, x2);
                double minY = Math.Min(y1, y2);
                AddShape(drawing, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, color, 1.75D);
            }

            for (int i = 0; i < categories.Count; i++) {
                double x = points[i].X - 2D;
                double y = points[i].Y - 2D;
                AddShape(drawing, OfficeShape.Ellipse(4D, 4D), x, y, OfficeColor.White, color, 1D);
            }

            if (stacked) {
                for (int i = 0; i < categories.Count; i++) {
                    double value = percentStacked ? Math.Max(0D, GetSeriesValue(series[s], i)) : GetSeriesValue(series[s], i);
                    if (value >= 0D) {
                        positiveCumulative[i] += value;
                    } else {
                        negativeCumulative[i] += value;
                    }
                }
            }
        }
    }

    private static void AddScatterSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        IReadOnlyList<double> sharedXValues = GetScatterXValues(categories);
        ValueRange xRange = GetScatterXRange(series, sharedXValues);
        ValueRange yRange = GetFiniteSeriesRange(series);
        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, s);
            IReadOnlyList<double> xValues = series[s].XValues ?? sharedXValues;
            int pointCount = Math.Min(xValues.Count, series[s].Values.Count);
            var points = new List<OfficePoint>(pointCount);
            for (int i = 0; i < pointCount; i++) {
                double yValue = GetSeriesValue(series[s], i);
                double x = ToPlotX(xValues[i], xRange.Min, xRange.Max, plotLeft, plotWidth);
                double y = ToPlotY(yValue, yRange.Min, yRange.Max, plotTop, plotHeight);
                points.Add(new OfficePoint(x, y));
            }

            AddPointLine(drawing, points, color, 1.25D);
            for (int i = 0; i < points.Count; i++) {
                OfficePoint point = points[i];
                AddShape(drawing, OfficeShape.Ellipse(5D, 5D), point.X - 2.5D, point.Y - 2.5D, OfficeColor.White, color, 1.25D);
            }
        }
    }

    private static void AddRadarSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count < 3 || series.Count == 0) {
            return;
        }

        double legendWidth = GetSeriesLegendWidth(series, width, layout);
        double visualWidth = Math.Max(80D, width - legendWidth);
        double centerX = visualWidth / 2D;
        double contentHeight = Math.Max(40D, height - contentTop);
        double centerY = contentTop + contentHeight / 2D;
        double radius = Math.Max(28D, Math.Min(visualWidth - 52D, contentHeight - 42D) / 2D);
        ValueRange range = GetRadarValueRange(series);

        for (int ring = 1; ring <= 4; ring++) {
            double ringRadius = radius * ring / 4D;
            IReadOnlyList<OfficePoint> ringPoints = CreateRadarPoints(categories.Count, centerX, centerY, ringRadius);
            AddPolygonShape(drawing, ringPoints, null, style.GridLineColor, 0.5D);
        }

        IReadOnlyList<OfficePoint> outerPoints = CreateRadarPoints(categories.Count, centerX, centerY, radius);
        for (int i = 0; i < outerPoints.Count; i++) {
            OfficePoint point = outerPoints[i];
            double minX = Math.Min(centerX, point.X);
            double minY = Math.Min(centerY, point.Y);
            AddShape(
                drawing,
                OfficeShape.Line(centerX - minX, centerY - minY, point.X - minX, point.Y - minY),
                minX,
                minY,
                null,
                style.GridLineColor,
                0.5D);
        }

        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, s);
            var points = new List<OfficePoint>(categories.Count);
            for (int i = 0; i < categories.Count; i++) {
                double value = GetSeriesValue(series[s], i);
                double pointRadius = radius * ToRadarRadiusRatio(value, range.Min, range.Max);
                points.Add(CreateRadarPoint(i, categories.Count, centerX, centerY, pointRadius));
            }

            AddPolygonShape(drawing, points, color, color, 1D, 0.18D);
            for (int i = 0; i < points.Count; i++) {
                OfficePoint point = points[i];
                AddShape(drawing, OfficeShape.Ellipse(4D, 4D), point.X - 2D, point.Y - 2D, OfficeColor.White, color, 1D);
            }
        }

        AddRadarCategoryLabels(drawing, categories, centerX, centerY, radius, style, layout);
        AddSeriesLegend(drawing, series, width - legendWidth + 6D, contentTop + 12D, Math.Max(0D, legendWidth - 12D), Math.Max(20D, contentHeight - 24D), style, layout);
    }

    private static void AddPieSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, bool doughnut, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        OfficeChartSeries values = series[0];
        double total = 0D;
        for (int i = 0; i < categories.Count; i++) {
            double value = GetSeriesValue(values, i);
            if (!double.IsNaN(value) && !double.IsInfinity(value) && value > 0D) {
                total += value;
            }
        }

        if (total <= 0D) {
            return;
        }

        double legendWidth = GetCategoryLegendWidth(categories, width, layout);
        double contentHeight = Math.Max(40D, height - contentTop);
        double visualWidth = Math.Max(80D, width - legendWidth);
        double radius = Math.Max(28D, Math.Min(visualWidth - 48D, contentHeight - 36D) / 2D);
        double centerX = visualWidth / 2D;
        double centerY = contentTop + contentHeight / 2D;
        double start = -Math.PI / 2D;
        for (int i = 0; i < categories.Count; i++) {
            double value = Math.Max(0D, GetSeriesValue(values, i));
            if (value <= 0D) {
                continue;
            }

            double sweep = value / total * Math.PI * 2D;
            double end = start + sweep;
            var points = new List<OfficePoint> {
                new OfficePoint(centerX, centerY)
            };
            int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
            for (int segment = 0; segment <= segments; segment++) {
                double angle = start + sweep * segment / segments;
                points.Add(new OfficePoint(
                    centerX + Math.Cos(angle) * radius,
                    centerY + Math.Sin(angle) * radius));
            }

            AddPolygonShape(drawing, points, GetSeriesColor(style, i), OfficeColor.White, 0.5D);
            start = end;
        }

        if (doughnut) {
            double innerDiameter = radius * 1.02D;
            AddShape(
                drawing,
                OfficeShape.Ellipse(innerDiameter, innerDiameter),
                centerX - innerDiameter / 2D,
                centerY - innerDiameter / 2D,
                style.BackgroundColor,
                null,
                0D);
        }

        AddCategoryLegend(drawing, categories, width - legendWidth + 6D, contentTop + 12D, Math.Max(0D, legendWidth - 12D), Math.Max(20D, contentHeight - 24D), style, layout);
    }

    private static void AddShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, OfficeColor? fill, OfficeColor? stroke, double strokeWidth) {
        shape.FillColor = fill;
        shape.StrokeColor = stroke;
        shape.StrokeWidth = strokeWidth;
        drawing.AddShape(shape, x, y);
    }

    private static void AddPolygonShape(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, double? fillOpacity = null) {
        if (points.Count < 3) {
            return;
        }

        double minX = points[0].X;
        double minY = points[0].Y;
        double maxX = points[0].X;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            OfficePoint point = points[i];
            if (point.X < minX) {
                minX = point.X;
            }

            if (point.Y < minY) {
                minY = point.Y;
            }

            if (point.X > maxX) {
                maxX = point.X;
            }

            if (point.Y > maxY) {
                maxY = point.Y;
            }
        }

        if (maxX <= minX || maxY <= minY) {
            return;
        }

        OfficeShape shape = OfficeShape.Polygon(points);
        shape.FillOpacity = fillOpacity;
        AddShape(drawing, shape, minX, minY, fill, stroke, strokeWidth);
    }

    private static void AddPointLine(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth) {
        for (int i = 1; i < points.Count; i++) {
            OfficePoint previous = points[i - 1];
            OfficePoint current = points[i];
            if (previous.Equals(current)) {
                continue;
            }

            double minX = Math.Min(previous.X, current.X);
            double minY = Math.Min(previous.Y, current.Y);
            AddShape(
                drawing,
                OfficeShape.Line(previous.X - minX, previous.Y - minY, current.X - minX, current.Y - minY),
                minX,
                minY,
                null,
                color,
                strokeWidth);
        }
    }

}
