using System;
using System.Collections.Generic;
using System.Globalization;

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

        double width = snapshot.WidthPoints;
        double height = snapshot.HeightPoints;
        OfficeChartStyle style = snapshot.Style;
        OfficeChartLayout layout = snapshot.Layout;
        var drawing = new OfficeDrawing(width, height);

        AddShape(drawing, OfficeShape.Rectangle(width, height), 0D, 0D, style.ShowBackground ? style.BackgroundColor : null, style.BorderColor, 0.75D);
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
            AddPieSeries(drawing, snapshot, width, height, contentTop, 0D, IsDoughnutChart(snapshot.ChartKind), style, layout);
            return drawing;
        }

        double topLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetSeriesLegendBandHeight(snapshot.Data.Series, width - 16D, layout)
            : 0D;
        double bottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetSeriesLegendBandHeight(snapshot.Data.Series, width - 16D, layout)
            : 0D;
        if (topLegendHeight > 0D) {
            AddSeriesLegendBand(drawing, snapshot.Data.Series, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout);
        }

        if (IsRadarChart(snapshot.ChartKind)) {
            AddRadarSeries(drawing, snapshot, width, height, contentTop + topLegendHeight, bottomLegendHeight, style, layout);
            return drawing;
        }

        bool barChart = IsBarChart(snapshot.ChartKind);
        bool showHorizontalAxis = barChart ? layout.ShowValueAxis : layout.ShowCategoryAxis;
        bool showVerticalAxis = barChart ? layout.ShowCategoryAxis : layout.ShowValueAxis;
        double verticalAxisTitleHeight = HasVerticalAxisTitle(snapshot.ChartKind, layout) ? 12D : 0D;
        double plotTop = 18D + contentTop + topLegendHeight + verticalAxisTitleHeight;
        double legendWidth = GetSeriesLegendWidth(snapshot.Data.Series, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        double plotLeft = 36D + (leftLegend ? legendWidth : 0D);
        double plotRight = 12D + (leftLegend ? 0D : legendWidth);
        double horizontalAxisTitleHeight = HasHorizontalAxisTitle(snapshot.ChartKind, layout) ? 12D : 0D;
        double plotBottom = 40D + horizontalAxisTitleHeight + bottomLegendHeight;
        double plotWidth = Math.Max(20D, width - plotLeft - plotRight);
        double plotHeight = Math.Max(20D, height - plotTop - plotBottom);
        double plotBottomY = plotTop + plotHeight;
        double axisLabelLeft = leftLegend ? legendWidth + 2D : 2D;
        double axisLabelWidth = Math.Max(12D, plotLeft - axisLabelLeft - 6D);
        ValueRange axisRange = GetCartesianValueRange(snapshot);
        bool valueAxisUsesPercentDefaults =
            IsPercentStackedBarOrColumnChart(snapshot.ChartKind) ||
            IsPercentStackedLineChart(snapshot.ChartKind) ||
            IsPercentStackedAreaChart(snapshot.ChartKind);

        if (style.PlotAreaBackgroundColor.HasValue || style.PlotAreaBorderColor.HasValue) {
            AddShape(
                drawing,
                OfficeShape.Rectangle(plotWidth, plotHeight),
                plotLeft,
                plotTop,
                style.PlotAreaBackgroundColor,
                style.PlotAreaBorderColor,
                style.PlotAreaBorderColor.HasValue ? 0.75D : 0D);
        }

        if (showHorizontalAxis) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, plotBottomY, null, style.AxisColor, 0.75D);
        }

        if (showVerticalAxis) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), plotLeft, plotTop, null, style.AxisColor, 0.75D);
        }

        if (style.ShowGridLines && layout.ShowValueAxis) {
            if (barChart) {
                for (int i = 1; i <= 3; i++) {
                    double x = plotLeft + plotWidth * i / 4D;
                    AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), x, plotTop, null, style.GridLineColor, 0.5D);
                }
            } else {
                for (int i = 1; i <= 3; i++) {
                    double y = plotTop + plotHeight * i / 4D;
                    AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, y, null, style.GridLineColor, 0.5D);
                }
            }
        }

        if (IsAreaChart(snapshot.ChartKind)) {
            AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else if (IsScatterChart(snapshot.ChartKind)) {
            AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else if (IsLineChart(snapshot.ChartKind)) {
            AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else {
            AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        }

        if (barChart) {
            if (layout.ShowCategoryAxis) {
                AddHorizontalCategoryAxisLabels(drawing, snapshot.Data.Categories, plotTop, plotHeight, axisLabelLeft, axisLabelWidth, style, layout);
            }

            if (layout.ShowValueAxis) {
                AddHorizontalValueAxisLabels(drawing, axisRange, plotLeft, plotBottomY, plotWidth, style, layout, valueAxisUsesPercentDefaults);
            }

            AddAxisTitles(drawing, layout.ShowCategoryAxis ? layout.CategoryAxisTitle : null, layout.ShowValueAxis ? layout.ValueAxisTitle : null, plotLeft, plotTop, plotBottomY, plotWidth, plotHeight, style, layout);
        } else {
            if (layout.ShowValueAxis) {
                AddValueAxisLabels(drawing, axisRange, plotTop, plotHeight, axisLabelLeft, axisLabelWidth, style, layout, valueAxisUsesPercentDefaults);
            }

            if (layout.ShowCategoryAxis) {
                if (IsScatterChart(snapshot.ChartKind)) {
                    IReadOnlyList<double> sharedXValues = GetScatterXValues(snapshot.Data.Categories);
                    AddHorizontalValueAxisLabels(drawing, GetScatterXRange(snapshot.Data.Series, sharedXValues), plotLeft, plotBottomY, plotWidth, style, layout, percentDefault: false);
                } else {
                    AddCategoryAxisLabels(drawing, snapshot.Data.Categories, plotLeft, plotBottomY, plotWidth, style, layout);
                }
            }

            AddAxisTitles(drawing, layout.ShowValueAxis ? layout.ValueAxisTitle : null, layout.ShowCategoryAxis ? layout.CategoryAxisTitle : null, plotLeft, plotTop, plotBottomY, plotWidth, plotHeight, style, layout);
        }

        AddSeriesLegend(
            drawing,
            snapshot.Data.Series,
            leftLegend ? 6D : width - legendWidth + 6D,
            plotTop,
            Math.Max(0D, legendWidth - 12D),
            plotHeight,
            style,
            layout);
        if (bottomLegendHeight > 0D) {
            AddSeriesLegendBand(drawing, snapshot.Data.Series, 8D, height - bottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout);
        }

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

    private static OfficeColor GetSeriesColor(OfficeChartStyle style, IReadOnlyList<OfficeChartSeries> series, int index) {
        if (index >= 0 && index < series.Count && series[index].Color.HasValue) {
            return series[index].Color!.Value;
        }

        return GetSeriesColor(style, index);
    }

    private static OfficeColor GetPointColor(OfficeChartStyle style, IReadOnlyList<OfficeColor?>? pointColors, int index) {
        if (pointColors != null && index >= 0 && index < pointColors.Count && pointColors[index].HasValue) {
            return pointColors[index]!.Value;
        }

        return GetSeriesColor(style, index);
    }

    private static OfficeColor GetPointColor(IReadOnlyList<OfficeColor?>? pointColors, int index, OfficeColor fallbackColor) {
        if (pointColors != null && index >= 0 && index < pointColors.Count && pointColors[index].HasValue) {
            return pointColors[index]!.Value;
        }

        return fallbackColor;
    }

    private static OfficeColor GetPointColor(OfficeChartStyle style, OfficeChartSeries series, int index) {
        OfficeColor fallbackColor = series.Color ?? GetPointColor(style, (IReadOnlyList<OfficeColor?>?)null, index);
        return GetPointColor(series.PointColors, index, fallbackColor);
    }

    private static IReadOnlyList<OfficeColor?> GetCategoryPointColors(OfficeChartStyle style, OfficeChartSeries series, int categoryCount) {
        var colors = new OfficeColor?[categoryCount];
        for (int i = 0; i < colors.Length; i++) {
            colors[i] = GetPointColor(style, series, i);
        }

        return colors;
    }

    private static IReadOnlyList<OfficeColor?>? GetLegendPointColors(OfficeChartStyle style, IReadOnlyList<OfficeChartSeries> series, int categoryCount) {
        for (int i = 0; i < series.Count; i++) {
            if (series[i].PointColors != null) {
                return series[i].PointColors;
            }
        }

        for (int i = 0; i < series.Count; i++) {
            if (series[i].Color.HasValue) {
                return GetCategoryPointColors(style, series[i], categoryCount);
            }
        }

        return null;
    }

    private static void AddBarSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
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
            ? GetPercentStackedSeriesRange(series, categories.Count)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetCartesianValueRange(snapshot);
        double min = Math.Min(0D, range.Min);
        double max = Math.Max(0D, range.Max);
        if (max <= min) {
            max = min + 1D;
        }

        for (int category = 0; category < categories.Count; category++) {
            double positiveBase = 0D;
            double negativeBase = 0D;
            for (int s = 0; s < series.Count; s++) {
                if (!TryGetSeriesValue(series[s], category, out double value)) {
                    continue;
                }

                if (value == 0D && !layout.ShowDataLabels) {
                    continue;
                }

                double baseline = 0D;
                double plottedValue = value;
                if (stacked) {
                    if (percentStacked) {
                        plottedValue = NormalizePercentStackedValue(series, category, value);
                    }

                    baseline = plottedValue >= 0D ? positiveBase : negativeBase;
                    if (plottedValue >= 0D) {
                        positiveBase += plottedValue;
                    } else {
                        negativeBase += plottedValue;
                    }
                }

                OfficeColor color = GetSeriesColor(style, series, s);
                if (series[s].PointColors != null && category < series[s].PointColors!.Count && series[s].PointColors![category].HasValue) {
                    color = GetPointColor(style, series[s].PointColors, category);
                }

                if (horizontal) {
                    double categoryHeight = plotHeight / categories.Count;
                    double rowHeight = Math.Max(2D, categoryHeight * 0.68D / (stacked ? 1D : series.Count));
                    int categorySlot = categories.Count - 1 - category;
                    int seriesSlot = stacked ? 0 : series.Count - 1 - s;
                    double y = plotTop + categoryHeight * categorySlot + categoryHeight * 0.16D + (stacked ? 0D : rowHeight * seriesSlot);
                    double x1 = ToPlotX(baseline, min, max, plotLeft, plotWidth);
                    double x2 = ToPlotX(stacked ? baseline + plottedValue : plottedValue, min, max, plotLeft, plotWidth);
                    double x = Math.Min(x1, x2);
                    double w = Math.Max(1D, Math.Abs(x2 - x1));
                    if (value != 0D) {
                        AddShape(drawing, OfficeShape.Rectangle(w, rowHeight), x, y, color, null, 0D);
                    }

                    AddHorizontalDataLabel(
                        drawing,
                        layout,
                        style,
                        categories[category],
                        series[s],
                        value,
                        GetDataLabelCategoryTotal(series, category),
                        x,
                        x + w,
                        y,
                        y + rowHeight);
                } else {
                    double x = plotLeft + slot * category + (slot - groupWidth) / 2D + (stacked ? 0D : barWidth * s);
                    double y1 = ToPlotY(baseline, min, max, plotTop, plotHeight);
                    double y2 = ToPlotY(stacked ? baseline + plottedValue : plottedValue, min, max, plotTop, plotHeight);
                    double y = Math.Min(y1, y2);
                    double h = Math.Max(1D, Math.Abs(y2 - y1));
                    if (value != 0D) {
                        AddShape(drawing, OfficeShape.Rectangle(barWidth * 0.88D, h), x, y, color, null, 0D);
                    }

                    AddVerticalDataLabel(
                        drawing,
                        layout,
                        style,
                        categories[category],
                        series[s],
                        value,
                        GetDataLabelCategoryTotal(series, category),
                        x + barWidth * 0.44D,
                        y,
                        y + h);
                }
            }
        }
    }

    private static void AddAreaSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count < 2 || series.Count == 0) {
            return;
        }

        bool stacked = IsStackedAreaChart(snapshot.ChartKind) || IsPercentStackedAreaChart(snapshot.ChartKind);
        bool percentStacked = IsPercentStackedAreaChart(snapshot.ChartKind);
        ValueRange range = percentStacked
            ? GetPercentStackedSeriesRange(series, categories.Count)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetCartesianValueRange(snapshot);
        double step = plotWidth / (categories.Count - 1);
        var positiveCumulative = new double[categories.Count];
        var negativeCumulative = new double[categories.Count];

        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, series, s);
            var topPoints = new List<OfficePoint>(categories.Count);
            var bottomPoints = new List<OfficePoint>(categories.Count);
            var runCategoryIndices = new List<int>(categories.Count);

            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(series[s], i, out double value)) {
                    AddAreaRun(drawing, topPoints, bottomPoints, color);
                    AddAreaRunDataLabels(drawing, layout, style, categories, series, s, runCategoryIndices, topPoints);
                    topPoints.Clear();
                    bottomPoints.Clear();
                    runCategoryIndices.Clear();
                    continue;
                }

                double rawValue = percentStacked ? NormalizePercentStackedValue(series, i, value) : value;
                double baseline = stacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double topValue = baseline + rawValue;

                double x = plotLeft + step * i;
                topPoints.Add(new OfficePoint(x, ToPlotY(topValue, range.Min, range.Max, plotTop, plotHeight)));
                bottomPoints.Add(new OfficePoint(x, ToPlotY(baseline, range.Min, range.Max, plotTop, plotHeight)));
                runCategoryIndices.Add(i);

                if (stacked) {
                    double stackedValue = percentStacked ? NormalizePercentStackedValue(series, i, value) : value;
                    if (stackedValue >= 0D) {
                        positiveCumulative[i] += stackedValue;
                    } else {
                        negativeCumulative[i] += stackedValue;
                    }
                }
            }

            AddAreaRun(drawing, topPoints, bottomPoints, color);
            AddAreaRunDataLabels(drawing, layout, style, categories, series, s, runCategoryIndices, topPoints);
        }
    }

    private static void AddAreaRun(OfficeDrawing drawing, IReadOnlyList<OfficePoint> topPoints, IReadOnlyList<OfficePoint> bottomPoints, OfficeColor color) {
        if (topPoints.Count < 2 || bottomPoints.Count != topPoints.Count) {
            return;
        }

        var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
        areaPoints.AddRange(topPoints);
        for (int i = bottomPoints.Count - 1; i >= 0; i--) {
            areaPoints.Add(bottomPoints[i]);
        }

        AddPolygonShape(drawing, areaPoints, color, color, 0.5D, 0.32D);
        AddPointLine(drawing, topPoints, color, 1.4D);
    }

    private static void AddAreaRunDataLabels(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        IReadOnlyList<string> categories,
        IReadOnlyList<OfficeChartSeries> series,
        int seriesIndex,
        IReadOnlyList<int> categoryIndices,
        IReadOnlyList<OfficePoint> topPoints) {
        if (categoryIndices.Count != topPoints.Count) {
            return;
        }

        OfficeChartSeries currentSeries = series[seriesIndex];
        for (int i = 0; i < categoryIndices.Count; i++) {
            int categoryIndex = categoryIndices[i];
            AddPointDataLabel(
                drawing,
                layout,
                style,
                categories[categoryIndex],
                currentSeries,
                currentSeries.Values[categoryIndex],
                GetDataLabelCategoryTotal(series, categoryIndex),
                topPoints[i].X,
                topPoints[i].Y);
        }
    }

    private static void AddLineSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        bool stacked = IsStackedLineChart(snapshot.ChartKind) || IsPercentStackedLineChart(snapshot.ChartKind);
        bool percentStacked = IsPercentStackedLineChart(snapshot.ChartKind);
        ValueRange range = percentStacked
            ? GetPercentStackedSeriesRange(series, categories.Count)
            : stacked
                ? GetStackedSeriesRange(series, categories.Count)
                : GetCartesianValueRange(snapshot);
        double step = categories.Count > 1 ? plotWidth / (categories.Count - 1) : 0D;
        var positiveCumulative = new double[categories.Count];
        var negativeCumulative = new double[categories.Count];
        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, series, s);
            var points = new OfficePoint[categories.Count];
            var plotted = new bool[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(series[s], i, out double value)) {
                    continue;
                }

                double rawValue = percentStacked ? NormalizePercentStackedValue(series, i, value) : value;
                double baseline = stacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double plottedValue = stacked ? baseline + rawValue : value;

                points[i] = new OfficePoint(plotLeft + step * i, ToPlotY(plottedValue, range.Min, range.Max, plotTop, plotHeight));
                plotted[i] = true;
            }

            for (int i = 1; i < categories.Count; i++) {
                if (!plotted[i - 1] || !plotted[i]) {
                    continue;
                }

                double x1 = points[i - 1].X;
                double y1 = points[i - 1].Y;
                double x2 = points[i].X;
                double y2 = points[i].Y;
                double minX = Math.Min(x1, x2);
                double minY = Math.Min(y1, y2);
                AddShape(drawing, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, color, 1.75D);
            }

            for (int i = 0; i < categories.Count; i++) {
                if (!plotted[i]) {
                    continue;
                }

                if (layout.ShowMarkers && series[s].ShowMarkers) {
                    double x = points[i].X - 2D;
                    double y = points[i].Y - 2D;
                    OfficeColor pointColor = GetPointColor(series[s].PointColors, i, color);
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), x, y, pointColor, pointColor, 1D);
                }

                double value = GetSeriesValue(series[s], i);
                AddPointDataLabel(
                    drawing,
                    layout,
                    style,
                    categories[i],
                    series[s],
                    value,
                    GetDataLabelCategoryTotal(series, i),
                    points[i].X,
                    points[i].Y);
            }

            if (stacked) {
                for (int i = 0; i < categories.Count; i++) {
                    if (!TryGetSeriesValue(series[s], i, out double seriesValue)) {
                        continue;
                    }

                    double value = percentStacked ? NormalizePercentStackedValue(series, i, seriesValue) : seriesValue;
                    if (value >= 0D) {
                        positiveCumulative[i] += value;
                    } else {
                        negativeCumulative[i] += value;
                    }
                }
            }
        }
    }

    private static void AddScatterSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        IReadOnlyList<double> sharedXValues = GetScatterXValues(categories);
        ValueRange xRange = GetScatterXRange(series, sharedXValues);
        ValueRange yRange = GetFiniteSeriesRange(series);
        for (int s = 0; s < series.Count; s++) {
            OfficeColor color = GetSeriesColor(style, series, s);
            IReadOnlyList<double> xValues = series[s].XValues ?? sharedXValues;
            int pointCount = Math.Min(xValues.Count, series[s].Values.Count);
            var points = new List<(OfficePoint Point, int SourceIndex)>(pointCount);
            var lineSegment = new List<OfficePoint>(pointCount);
            for (int i = 0; i < pointCount; i++) {
                if (!TryGetSeriesValue(series[s], i, out double yValue)) {
                    if (layout.ConnectScatterPoints) {
                        AddPointLine(drawing, lineSegment, color, 1.25D);
                    }

                    lineSegment.Clear();
                    continue;
                }

                double xValue = xValues[i];
                if (!IsFiniteChartValue(xValue)) {
                    if (layout.ConnectScatterPoints) {
                        AddPointLine(drawing, lineSegment, color, 1.25D);
                    }

                    lineSegment.Clear();
                    continue;
                }

                double x = ToPlotX(xValue, xRange.Min, xRange.Max, plotLeft, plotWidth);
                double y = ToPlotY(yValue, yRange.Min, yRange.Max, plotTop, plotHeight);
                var point = new OfficePoint(x, y);
                points.Add((point, i));
                if (layout.ConnectScatterPoints) {
                    lineSegment.Add(point);
                }
            }

            if (layout.ConnectScatterPoints) {
                AddPointLine(drawing, lineSegment, color, 1.25D);
            }
            for (int i = 0; i < points.Count; i++) {
                OfficePoint point = points[i].Point;
                if (layout.ShowMarkers && series[s].ShowMarkers) {
                    OfficeColor pointColor = GetPointColor(series[s].PointColors, points[i].SourceIndex, color);
                    AddShape(drawing, OfficeShape.Ellipse(5D, 5D), point.X - 2.5D, point.Y - 2.5D, pointColor, pointColor, 1.25D);
                }

                int pointIndex = points[i].SourceIndex;
                string labelCategory = series[s].XValues != null && pointIndex < xValues.Count
                    ? xValues[pointIndex].ToString("0.####", CultureInfo.InvariantCulture)
                    : pointIndex < categories.Count ? categories[pointIndex] : string.Empty;
                AddPointDataLabel(
                    drawing,
                    layout,
                    style,
                    labelCategory,
                    series[s],
                    series[s].Values[pointIndex],
                    GetDataLabelCategoryTotal(series, pointIndex),
                    point.X,
                    point.Y);
            }
        }
    }

    private static void AddRadarSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, double bottomLegendHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count < 3 || series.Count == 0) {
            return;
        }

        double legendWidth = GetSeriesLegendWidth(series, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        double visualWidth = Math.Max(80D, width - legendWidth);
        double centerX = (leftLegend ? legendWidth : 0D) + visualWidth / 2D;
        double contentHeight = Math.Max(40D, height - contentTop - bottomLegendHeight);
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
            OfficeColor color = GetSeriesColor(style, series, s);
            var points = new OfficePoint[categories.Count];
            var plotted = new bool[categories.Count];
            bool allPointsPlotted = true;
            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(series[s], i, out double value)) {
                    allPointsPlotted = false;
                    continue;
                }

                double pointRadius = radius * ToRadarRadiusRatio(value, range.Min, range.Max);
                points[i] = CreateRadarPoint(i, categories.Count, centerX, centerY, pointRadius);
                plotted[i] = true;
            }

            if (allPointsPlotted) {
                AddPolygonShape(drawing, points, layout.FillRadarSeries ? color : null, color, 1D, layout.FillRadarSeries ? 0.18D : 1D);
            } else {
                for (int i = 1; i < categories.Count; i++) {
                    if (!plotted[i - 1] || !plotted[i]) {
                        continue;
                    }

                    AddPointLine(drawing, new[] { points[i - 1], points[i] }, color, 1D);
                }
            }

            if (layout.ShowMarkers && series[s].ShowMarkers) {
                for (int i = 0; i < points.Length; i++) {
                    if (!plotted[i]) {
                        continue;
                    }

                    OfficePoint point = points[i];
                    OfficeColor pointColor = GetPointColor(series[s].PointColors, i, color);
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), point.X - 2D, point.Y - 2D, pointColor, pointColor, 1D);
                }
            }

            for (int i = 0; i < points.Length; i++) {
                if (!plotted[i]) {
                    continue;
                }

                AddPointDataLabel(
                    drawing,
                    layout,
                    style,
                    categories[i],
                    series[s],
                    series[s].Values[i],
                    GetDataLabelCategoryTotal(series, i),
                    points[i].X,
                    points[i].Y);
            }
        }

        AddRadarCategoryLabels(drawing, categories, centerX, centerY, radius, style, layout);
        AddSeriesLegend(
            drawing,
            series,
            leftLegend ? 6D : width - legendWidth + 6D,
            contentTop + 12D,
            Math.Max(0D, legendWidth - 12D),
            Math.Max(20D, contentHeight - 24D),
            style,
            layout);
        if (bottomLegendHeight > 0D) {
            AddSeriesLegendBand(drawing, series, 8D, height - bottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout);
        }
    }

    private static void AddPieSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double width, double height, double contentTop, double bottomLegendHeight, bool doughnut, OfficeChartStyle style, OfficeChartLayout layout) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        if (doughnut) {
            AddDoughnutSeries(drawing, categories, series, width, height, contentTop, bottomLegendHeight, style, layout);
            return;
        }

        OfficeChartSeries values = series[0];
        double total = 0D;
        for (int i = 0; i < categories.Count; i++) {
            if (TryGetSeriesValue(values, i, out double value) && value > 0D) {
                total += value;
            }
        }

        if (total <= 0D) {
            return;
        }

        double topCategoryLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : 0D;
        IReadOnlyList<OfficeColor?> categoryPointColors = GetCategoryPointColors(style, values, categories.Count);
        if (topCategoryLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout, categoryPointColors);
            contentTop += topCategoryLegendHeight;
        }

        double categoryBottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : bottomLegendHeight;
        double legendWidth = GetCategoryLegendWidth(categories, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        double contentHeight = Math.Max(40D, height - contentTop - categoryBottomLegendHeight);
        double visualWidth = Math.Max(80D, width - legendWidth);
        double radius = Math.Max(28D, Math.Min(visualWidth - 48D, contentHeight - 36D) / 2D);
        double centerX = (leftLegend ? legendWidth : 0D) + visualWidth / 2D;
        double centerY = contentTop + contentHeight / 2D;
        double start = -Math.PI / 2D;
        int zeroLabelIndex = 0;
        for (int i = 0; i < categories.Count; i++) {
            if (!TryGetSeriesValue(values, i, out double seriesValue)) {
                continue;
            }

            double value = Math.Max(0D, seriesValue);
            double sweep = value / total * Math.PI * 2D;
            if (value > 0D) {
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

                OfficeColor sliceColor = GetPointColor(style, values, i);
                AddPolygonShape(drawing, points, sliceColor, OfficeColor.White, 0.5D);
                if (layout.ShowDataLabels) {
                    AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, value, total, centerX, centerY, radius, start + sweep / 2D, zeroLabelIndex: null);
                }

                start = end;
            } else if (layout.ShowDataLabels) {
                OfficeColor sliceColor = GetPointColor(style, values, 0);
                AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, 0D, total, centerX, centerY, radius, -Math.PI / 2D, zeroLabelIndex);
                zeroLabelIndex++;
            }
        }

        AddCategoryLegend(
            drawing,
            categories,
            leftLegend ? 6D : width - legendWidth + 6D,
            contentTop + 12D,
            Math.Max(0D, legendWidth - 12D),
            Math.Max(20D, contentHeight - 24D),
            style,
            layout,
            categoryPointColors);
        if (categoryBottomLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, height - categoryBottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout, categoryPointColors);
        }
    }

    private static void AddDoughnutSeries(OfficeDrawing drawing, IReadOnlyList<string> categories, IReadOnlyList<OfficeChartSeries> series, double width, double height, double contentTop, double bottomLegendHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        var renderableSeries = new List<OfficeChartSeries>();
        for (int s = 0; s < series.Count; s++) {
            if (GetPositiveSeriesTotal(series[s], categories.Count) > 0D) {
                renderableSeries.Add(series[s]);
            }
        }

        if (renderableSeries.Count == 0) {
            return;
        }

        double topCategoryLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : 0D;
        IReadOnlyList<OfficeColor?>? legendPointColors = GetLegendPointColors(style, renderableSeries, categories.Count);
        if (topCategoryLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout, legendPointColors);
            contentTop += topCategoryLegendHeight;
        }

        double categoryBottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : bottomLegendHeight;
        double legendWidth = GetCategoryLegendWidth(categories, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        double contentHeight = Math.Max(40D, height - contentTop - categoryBottomLegendHeight);
        double visualWidth = Math.Max(80D, width - legendWidth);
        double radius = Math.Max(28D, Math.Min(visualWidth - 48D, contentHeight - 36D) / 2D);
        double centerX = (leftLegend ? legendWidth : 0D) + visualWidth / 2D;
        double centerY = contentTop + contentHeight / 2D;

        double ringThickness = radius / (renderableSeries.Count + 0.9D);
        for (int s = 0; s < renderableSeries.Count; s++) {
            OfficeChartSeries values = renderableSeries[s];
            double outerRadius = radius - s * ringThickness;
            double innerRadius = Math.Max(0D, outerRadius - ringThickness * 0.82D);
            double total = GetPositiveSeriesTotal(values, categories.Count);
            double start = -Math.PI / 2D;
            int zeroLabelIndex = 0;
            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(values, i, out double seriesValue)) {
                    continue;
                }

                double value = Math.Max(0D, seriesValue);
                double sweep = value / total * Math.PI * 2D;
                if (value > 0D) {
                    double end = start + sweep;
                    OfficeColor sliceColor = GetPointColor(style, values, i);
                    AddPieSlice(drawing, centerX, centerY, outerRadius, start, sweep, sliceColor);
                    if (layout.ShowDataLabels) {
                        AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, value, total, centerX, centerY, Math.Max(innerRadius + 8D, outerRadius - ringThickness * 0.42D), start + sweep / 2D, zeroLabelIndex: null);
                    }

                    start = end;
                } else if (layout.ShowDataLabels && s == 0) {
                    OfficeColor sliceColor = GetPointColor(style, values, 0);
                    AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, 0D, total, centerX, centerY, outerRadius, -Math.PI / 2D, zeroLabelIndex);
                    zeroLabelIndex++;
                }
            }

            if (innerRadius > 0D) {
                double innerDiameter = innerRadius * 2D;
                AddShape(
                    drawing,
                    OfficeShape.Ellipse(innerDiameter, innerDiameter),
                    centerX - innerRadius,
                    centerY - innerRadius,
                    style.ShowBackground ? style.BackgroundColor : null,
                    null,
                    0D);
            }
        }

        AddCategoryLegend(
            drawing,
            categories,
            leftLegend ? 6D : width - legendWidth + 6D,
            contentTop + 12D,
            Math.Max(0D, legendWidth - 12D),
            Math.Max(20D, contentHeight - 24D),
            style,
            layout,
            legendPointColors);
        if (categoryBottomLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, height - categoryBottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout, legendPointColors);
        }
    }

    private static void AddPieDataLabel(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        OfficeColor labelColor,
        string category,
        OfficeChartSeries series,
        double value,
        double total,
        double centerX,
        double centerY,
        double radius,
        double angle,
        int? zeroLabelIndex) {
        string label = FormatDataLabel(layout, category, series, value, total);
        if (string.IsNullOrWhiteSpace(label)) {
            return;
        }

        double labelWidth = Math.Min(72D, Math.Max(36D, label.Length * layout.DataLabelFontSize * 0.52D + 6D));
        double labelHeight = Math.Max(9D, layout.DataLabelFontSize + 3D);
        double distance = zeroLabelIndex.HasValue ? radius * 0.9D : radius * 0.58D;
        double x = centerX + Math.Cos(angle) * distance - labelWidth / 2D;
        double y = centerY + Math.Sin(angle) * distance - labelHeight / 2D;
        if (zeroLabelIndex.HasValue) {
            y += zeroLabelIndex.Value * (labelHeight + 1D);
        }

        AddChartText(drawing, label, x, y, labelWidth, labelHeight, layout.DataLabelFontSize, labelColor, OfficeTextAlignment.Center, style);
    }

    private static OfficeColor GetReadableDataLabelColor(OfficeColor fillColor) {
        double srgbR = fillColor.R / 255D;
        double srgbG = fillColor.G / 255D;
        double srgbB = fillColor.B / 255D;
        double luminance = 0.2126D * srgbR + 0.7152D * srgbG + 0.0722D * srgbB;
        return luminance < 0.52D ? OfficeColor.White : OfficeColor.Black;
    }

    private static double GetPositiveSeriesTotal(OfficeChartSeries values, int categoryCount) {
        double total = 0D;
        for (int i = 0; i < categoryCount; i++) {
            if (TryGetSeriesValue(values, i, out double value) && value > 0D) {
                total += value;
            }
        }

        return total;
    }

    private static void AddPieSlice(OfficeDrawing drawing, double centerX, double centerY, double radius, double start, double sweep, OfficeColor color) {
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

        AddPolygonShape(drawing, points, color, OfficeColor.White, 0.5D);
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
