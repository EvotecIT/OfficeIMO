using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Renders dependency-free chart snapshots into vector drawing primitives shared by OfficeIMO exporters.
/// </summary>
public static partial class OfficeChartDrawingRenderer {
    private const double MinimumChartCanvasWidth = 240D;
    private const double MinimumChartCanvasHeight = 150D;

    /// <summary>
    /// Renders a chart snapshot into an <see cref="OfficeDrawing"/> scene.
    /// </summary>
    /// <param name="snapshot">Chart snapshot to render.</param>
    /// <param name="useMinimumCanvas">Whether to expand small snapshots to the shared default minimum chart canvas.</param>
    /// <returns>Vector drawing containing the chart plot area and series marks.</returns>
    public static OfficeDrawing Render(OfficeChartSnapshot snapshot, bool useMinimumCanvas = true) {
        if (snapshot == null) {
            throw new ArgumentNullException(nameof(snapshot));
        }

        double width = useMinimumCanvas ? Math.Max(MinimumChartCanvasWidth, snapshot.WidthPoints) : Math.Max(1D, snapshot.WidthPoints);
        double height = useMinimumCanvas ? Math.Max(MinimumChartCanvasHeight, snapshot.HeightPoints) : Math.Max(1D, snapshot.HeightPoints);
        OfficeChartStyle style = snapshot.Style;
        OfficeChartLayout layout = snapshot.Layout;
        var drawing = new OfficeDrawing(width, height);

        AddShape(
            drawing,
            OfficeShape.Rectangle(width, height),
            0D,
            0D,
            style.ShowBackground ? style.BackgroundColor : null,
            style.ShowBorder ? style.BorderColor : null,
            style.ShowBorder ? style.ChartBorderWidth ?? 0.75D : 0D,
            style.ChartBorderDashStyle ?? OfficeStrokeDashStyle.Solid);
        double contentTop = 0D;
        if (!string.IsNullOrWhiteSpace(snapshot.Title)) {
            double titleHeight = Math.Min(22D, Math.Max(16D, height * 0.12D));
            double titleTop = Math.Min(layout.TitleTopPadding, Math.Max(0D, height - titleHeight));
            string titleFontFamily = style.TitleFontFamily ?? style.FontFamily;
            double titleFontSize = style.TitleFontSize ?? Math.Min(12D, Math.Max(8D, titleHeight - 7D));
            OfficeFontStyle titleFontStyle = style.TitleFontStyle ?? OfficeFontStyle.Bold;
            drawing.AddText(
                snapshot.Title!,
                8D,
                titleTop,
                Math.Max(1D, width - 16D),
                Math.Max(1D, titleHeight - 4D),
                new OfficeFontInfo(titleFontFamily, titleFontSize, titleFontStyle),
                style.TitleColor,
                OfficeTextAlignment.Center);
            if (!layout.OverlayTitle) {
                contentTop = titleHeight + Math.Max(0D, titleTop - 5D);
            }
        }

        if (IsPieChart(snapshot.ChartKind) || IsDoughnutChart(snapshot.ChartKind)) {
            AddPieSeries(drawing, snapshot, width, height, contentTop, 0D, IsDoughnutChart(snapshot.ChartKind), style, layout);
            return drawing;
        }

        IReadOnlyList<OfficeChartSeries> legendSeries = GetRenderableLegendSeries(snapshot);
        double topLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetSeriesLegendBandHeight(legendSeries, width - 16D, layout)
            : 0D;
        double bottomLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Bottom
            ? GetSeriesLegendBandHeight(legendSeries, width - 16D, layout)
            : 0D;
        if (topLegendHeight > 0D) {
            AddSeriesLegendBand(drawing, legendSeries, 8D, contentTop + 2D, Math.Max(1D, width - 16D), style, layout);
        }

        if (IsRadarChart(snapshot.ChartKind)) {
            AddRadarSeries(drawing, snapshot, width, height, contentTop + topLegendHeight, bottomLegendHeight, style, layout);
            return drawing;
        }

        bool barChart = IsBarChart(snapshot.ChartKind);
        bool showHorizontalAxis = barChart ? layout.ShowValueAxisLine : layout.ShowCategoryAxisLine;
        bool showVerticalAxis = barChart ? layout.ShowCategoryAxisLine : layout.ShowValueAxisLine;
        bool showHorizontalAxisLabels = barChart ? layout.ShowValueAxisLabels : layout.ShowCategoryAxisLabels;
        bool showVerticalAxisLabels = barChart ? layout.ShowCategoryAxisLabels : layout.ShowValueAxisLabels;
        bool horizontalAxisCrossesAtMaximum = !barChart && layout.HorizontalAxisCrossingPosition == OfficeChartAxisCrossingPosition.Maximum;
        bool verticalAxisCrossesAtMaximum = !barChart && layout.VerticalAxisCrossingPosition == OfficeChartAxisCrossingPosition.Maximum;
        bool horizontalAxisLabelsHigh = layout.HorizontalAxisTickLabelPosition == OfficeChartAxisTickLabelPosition.High ||
            (layout.HorizontalAxisTickLabelPosition == OfficeChartAxisTickLabelPosition.NextTo && horizontalAxisCrossesAtMaximum);
        bool verticalAxisLabelsHigh = layout.VerticalAxisTickLabelPosition == OfficeChartAxisTickLabelPosition.High ||
            (layout.VerticalAxisTickLabelPosition == OfficeChartAxisTickLabelPosition.NextTo && verticalAxisCrossesAtMaximum);
        SecondaryAxisRenderContext secondaryAxis = CreateSecondaryAxisRenderContext(
            snapshot, layout, barChart, barChart ? showHorizontalAxisLabels : showVerticalAxisLabels);
        bool hasSecondaryAxis = secondaryAxis.HasSeries;
        ValueRange axisRange = GetPrimaryValueAxisRange(snapshot, layout, barChart, hasSecondaryAxis);
        ValueRange secondaryAxisRange = secondaryAxis.Range;
        double? valueAxisMajorUnit = GetValueAxisMajorUnit(layout, horizontal: barChart);
        IReadOnlyList<double> valueAxisMajorTicks = GetValueAxisMajorTicks(axisRange, valueAxisMajorUnit);
        double? valueAxisMinorUnit = GetValueAxisMinorUnit(layout, horizontal: barChart);
        IReadOnlyList<double> valueAxisMinorTicks = GetValueAxisMinorTicks(axisRange, valueAxisMinorUnit, valueAxisMajorTicks);
        bool valueAxisUsesPercentDefaults =
            IsPercentStackedBarOrColumnChart(snapshot.ChartKind) ||
            IsPercentStackedLineChart(snapshot.ChartKind) ||
            IsPercentStackedAreaChart(snapshot.ChartKind);
        double verticalAxisLabelBandWidth = showVerticalAxisLabels
            ? GetVerticalAxisLabelBandWidth(snapshot, axisRange, valueAxisMajorTicks, layout, valueAxisUsesPercentDefaults, horizontalValueAxis: barChart)
            : 28D;
        double horizontalValueAxisLabelWidth = GetHorizontalValueAxisLabelWidth(axisRange, valueAxisMajorTicks, layout, valueAxisUsesPercentDefaults);
        double horizontalAxisTopLabelHeight = showHorizontalAxisLabels &&
            (horizontalAxisLabelsHigh || (barChart && hasSecondaryAxis)) ? 15D : 0D;
        double secondaryAxisLabelBandWidth = secondaryAxis.LabelBandWidth;
        double verticalAxisRightLabelWidth = Math.Max(
            showVerticalAxisLabels && verticalAxisLabelsHigh ? verticalAxisLabelBandWidth + 8D : 0D,
            !barChart && secondaryAxisLabelBandWidth > 0D ? secondaryAxisLabelBandWidth + 8D : 0D);
        double verticalAxisTitleHeight = HasVerticalAxisTitle(snapshot.ChartKind, layout) ? GetAxisTitleBandHeight(layout) : 0D;
        double plotTop = 18D + contentTop + topLegendHeight + verticalAxisTitleHeight + horizontalAxisTopLabelHeight;
        double legendWidth = GetSeriesLegendWidth(legendSeries, width, layout);
        bool leftLegend = layout.LegendPosition == OfficeChartLegendPosition.Left;
        double plotLeft = 8D + verticalAxisLabelBandWidth + (leftLegend ? legendWidth : 0D);
        double plotRight = 12D + verticalAxisRightLabelWidth + (leftLegend ? 0D : legendWidth);
        double horizontalAxisTitleHeight = HasHorizontalAxisTitle(snapshot.ChartKind, layout) ? GetAxisTitleBandHeight(layout) : 0D;
        double plotBottom = 40D + horizontalAxisTitleHeight + bottomLegendHeight;
        double plotWidth = Math.Max(20D, width - plotLeft - plotRight);
        double plotHeight = Math.Max(20D, height - plotTop - plotBottom);
        double plotBottomY = plotTop + plotHeight;
        double horizontalAxisY = horizontalAxisCrossesAtMaximum ? plotTop : plotBottomY;
        double axisLabelLeft = leftLegend ? legendWidth + 2D : 2D;
        double axisLabelWidth = Math.Max(12D, verticalAxisLabelBandWidth);
        double axisLabelRight = plotLeft + plotWidth + 4D;
        double axisLabelRightWidth = Math.Max(12D,
            hasSecondaryAxis ? secondaryAxisLabelBandWidth : verticalAxisLabelBandWidth);
        double verticalAxisX = verticalAxisCrossesAtMaximum ? plotLeft + plotWidth : plotLeft;

        if (style.PlotAreaBackgroundColor.HasValue || style.PlotAreaBorderColor.HasValue) {
            AddShape(
                drawing,
                OfficeShape.Rectangle(plotWidth, plotHeight),
                plotLeft,
                plotTop,
                style.PlotAreaBackgroundColor,
                style.PlotAreaBorderColor,
                style.PlotAreaBorderColor.HasValue ? style.PlotAreaBorderWidth ?? 0.75D : 0D,
                style.PlotAreaBorderDashStyle ?? OfficeStrokeDashStyle.Solid);
        }

        OfficeColor horizontalAxisColor = barChart ? GetValueAxisColor(style) : GetCategoryAxisColor(style);
        OfficeColor verticalAxisColor = barChart ? GetCategoryAxisColor(style) : GetValueAxisColor(style);
        double horizontalAxisLineWidth = barChart ? GetValueAxisLineWidth(style) : GetCategoryAxisLineWidth(style);
        double verticalAxisLineWidth = barChart ? GetCategoryAxisLineWidth(style) : GetValueAxisLineWidth(style);
        OfficeStrokeDashStyle horizontalAxisLineDashStyle = barChart ? GetValueAxisLineDashStyle(style) : GetCategoryAxisLineDashStyle(style);
        OfficeStrokeDashStyle verticalAxisLineDashStyle = barChart ? GetCategoryAxisLineDashStyle(style) : GetValueAxisLineDashStyle(style);
        if (showHorizontalAxis) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, horizontalAxisY, null, horizontalAxisColor, horizontalAxisLineWidth, horizontalAxisLineDashStyle);
            if (barChart) {
                AddHorizontalValueAxisMajorTickMarks(
                    drawing,
                    plotLeft,
                    horizontalAxisY,
                    plotWidth,
                    axisRange,
                    valueAxisMajorTicks,
                    layout.HorizontalAxisMajorTickMark,
                    horizontalAxisColor,
                    horizontalAxisLineWidth,
                    positiveOutside: !horizontalAxisCrossesAtMaximum);
            } else {
                AddHorizontalAxisMajorTickMarks(
                    drawing,
                    plotLeft,
                    horizontalAxisY,
                    plotWidth,
                    layout.HorizontalAxisMajorTickMark,
                    horizontalAxisColor,
                    horizontalAxisLineWidth,
                    positiveOutside: !horizontalAxisCrossesAtMaximum);
            }

            if (barChart && valueAxisMinorTicks.Count > 0) {
                AddHorizontalValueAxisMinorTickMarks(
                    drawing,
                    plotLeft,
                    horizontalAxisY,
                    plotWidth,
                    axisRange,
                    valueAxisMinorTicks,
                    layout.HorizontalAxisMinorTickMark,
                    horizontalAxisColor,
                    horizontalAxisLineWidth,
                    positiveOutside: !horizontalAxisCrossesAtMaximum);
            } else {
                AddHorizontalAxisMinorTickMarks(
                    drawing,
                    plotLeft,
                    horizontalAxisY,
                    plotWidth,
                    layout.HorizontalAxisMinorTickMark,
                    horizontalAxisColor,
                    horizontalAxisLineWidth,
                    positiveOutside: !horizontalAxisCrossesAtMaximum);
            }
        }

        if (showVerticalAxis) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), verticalAxisX, plotTop, null, verticalAxisColor, verticalAxisLineWidth, verticalAxisLineDashStyle);
            if (barChart) {
                AddVerticalAxisMajorTickMarks(
                    drawing,
                    verticalAxisX,
                    plotTop,
                    plotHeight,
                    layout.VerticalAxisMajorTickMark,
                    verticalAxisColor,
                    verticalAxisLineWidth);
            } else {
                AddVerticalValueAxisMajorTickMarks(
                    drawing,
                    verticalAxisX,
                    plotTop,
                    plotHeight,
                    axisRange,
                    valueAxisMajorTicks,
                    layout.VerticalAxisMajorTickMark,
                    verticalAxisColor,
                    verticalAxisLineWidth);
            }

            if (!barChart && valueAxisMinorTicks.Count > 0) {
                AddVerticalValueAxisMinorTickMarks(
                    drawing,
                    verticalAxisX,
                    plotTop,
                    plotHeight,
                    axisRange,
                    valueAxisMinorTicks,
                    layout.VerticalAxisMinorTickMark,
                    verticalAxisColor,
                    verticalAxisLineWidth);
            } else {
                AddVerticalAxisMinorTickMarks(
                    drawing,
                    verticalAxisX,
                    plotTop,
                    plotHeight,
                    layout.VerticalAxisMinorTickMark,
                    verticalAxisColor,
                    verticalAxisLineWidth);
            }
        }

        if (hasSecondaryAxis && (barChart ? showHorizontalAxis : showVerticalAxis)) {
            if (barChart) {
                AddHorizontalSecondaryValueAxis(drawing, secondaryAxis, plotLeft, plotTop, plotWidth,
                    style, layout);
            } else {
                AddSecondaryValueAxis(drawing, secondaryAxis, plotLeft + plotWidth, plotTop, plotHeight,
                    style, layout);
            }
        }

        if (GetShowCategoryMinorGridLines(style)) {
            if (barChart) {
                AddHorizontalGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    divisions: 8,
                    startIndex: 1,
                    step: 2,
                    GetCategoryMinorGridLineColor(style),
                    GetCategoryMinorGridLineWidth(style),
                    GetCategoryMinorGridLineDashStyle(style));
            } else {
                AddVerticalGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    divisions: 8,
                    startIndex: 1,
                    step: 2,
                    GetCategoryMinorGridLineColor(style),
                    GetCategoryMinorGridLineWidth(style),
                    GetCategoryMinorGridLineDashStyle(style));
            }
        }

        if (GetShowValueMinorGridLines(style)) {
            if (barChart) {
                if (valueAxisMinorTicks.Count > 0) {
                    AddVerticalValueGridLines(
                        drawing,
                        plotLeft,
                        plotTop,
                        plotWidth,
                        plotHeight,
                        axisRange,
                        valueAxisMinorTicks,
                        GetValueMinorGridLineColor(style),
                        GetValueMinorGridLineWidth(style),
                        GetValueMinorGridLineDashStyle(style));
                } else {
                    AddVerticalGridLines(
                        drawing,
                        plotLeft,
                        plotTop,
                        plotWidth,
                        plotHeight,
                        divisions: 8,
                        startIndex: 1,
                        step: 2,
                        GetValueMinorGridLineColor(style),
                        GetValueMinorGridLineWidth(style),
                        GetValueMinorGridLineDashStyle(style));
                }
            } else {
                if (valueAxisMinorTicks.Count > 0) {
                    AddHorizontalValueGridLines(
                        drawing,
                        plotLeft,
                        plotTop,
                        plotWidth,
                        plotHeight,
                        axisRange,
                        valueAxisMinorTicks,
                        GetValueMinorGridLineColor(style),
                        GetValueMinorGridLineWidth(style),
                        GetValueMinorGridLineDashStyle(style));
                } else {
                    AddHorizontalGridLines(
                        drawing,
                        plotLeft,
                        plotTop,
                        plotWidth,
                        plotHeight,
                        divisions: 8,
                        startIndex: 1,
                        step: 2,
                        GetValueMinorGridLineColor(style),
                        GetValueMinorGridLineWidth(style),
                        GetValueMinorGridLineDashStyle(style));
                }
            }
        }

        if (GetShowCategoryGridLines(style)) {
            if (barChart) {
                AddHorizontalGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    divisions: 4,
                    startIndex: 1,
                    step: 1,
                    GetCategoryGridLineColor(style),
                    GetCategoryGridLineWidth(style),
                    GetCategoryGridLineDashStyle(style));
            } else {
                AddVerticalGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    divisions: 4,
                    startIndex: 1,
                    step: 1,
                    GetCategoryGridLineColor(style),
                    GetCategoryGridLineWidth(style),
                    GetCategoryGridLineDashStyle(style));
            }
        }

        if (GetShowValueGridLines(style)) {
            if (barChart) {
                AddVerticalValueGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    axisRange,
                    valueAxisMajorTicks,
                    GetValueGridLineColor(style),
                    GetValueGridLineWidth(style),
                    GetValueGridLineDashStyle(style));
            } else {
                AddHorizontalValueGridLines(
                    drawing,
                    plotLeft,
                    plotTop,
                    plotWidth,
                    plotHeight,
                    axisRange,
                    valueAxisMajorTicks,
                    GetValueGridLineColor(style),
                    GetValueGridLineWidth(style),
                    GetValueGridLineDashStyle(style));
            }
        }

        if (HasMixedCartesianSeriesKinds(snapshot) || hasSecondaryAxis) {
            AddMixedCartesianSeries(drawing, snapshot, axisRange, secondaryAxisRange, hasSecondaryAxis,
                plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else if (IsAreaChart(snapshot.ChartKind)) {
            AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else if (IsScatterChart(snapshot.ChartKind)) {
            AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else if (IsLineChart(snapshot.ChartKind)) {
            AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else {
            AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        }

        if (barChart) {
            if (layout.ShowCategoryAxis && layout.ShowCategoryAxisLabels) {
                AddHorizontalCategoryAxisLabels(
                    drawing,
                    snapshot.Data.Categories,
                    plotTop,
                    plotHeight,
                    verticalAxisLabelsHigh ? axisLabelRight : axisLabelLeft,
                    verticalAxisLabelsHigh ? axisLabelRightWidth : axisLabelWidth,
                    verticalAxisLabelsHigh ? OfficeTextAlignment.Left : OfficeTextAlignment.Right,
                    style,
                    layout);
            }

            if (layout.ShowValueAxis && layout.ShowValueAxisLabels) {
                AddHorizontalValueAxisLabels(
                    drawing,
                    axisRange,
                    plotLeft,
                    horizontalAxisLabelsHigh ? plotTop - 13D : plotBottomY + 4D,
                    plotWidth,
                    horizontalValueAxisLabelWidth,
                    horizontalAxisLabelsHigh,
                    style,
                    layout,
                    valueAxisUsesPercentDefaults);
            }
            if (hasSecondaryAxis && layout.ShowValueAxis && layout.ShowValueAxisLabels) {
                AddHorizontalSecondaryValueAxisLabels(drawing, secondaryAxis, plotLeft, plotTop,
                    plotWidth, style, layout);
            }
            AddAxisTitles(drawing, layout.ShowCategoryAxis ? layout.CategoryAxisTitle : null, layout.ShowValueAxis ? layout.ValueAxisTitle : null, plotLeft, plotTop, plotBottomY, plotWidth, plotHeight, style, layout);
        } else {
            if (layout.ShowValueAxis && layout.ShowValueAxisLabels) {
                AddValueAxisLabels(
                    drawing,
                    axisRange,
                    plotTop,
                    plotHeight,
                    verticalAxisLabelsHigh ? axisLabelRight : axisLabelLeft,
                    verticalAxisLabelsHigh ? axisLabelRightWidth : axisLabelWidth,
                    verticalAxisLabelsHigh ? OfficeTextAlignment.Left : OfficeTextAlignment.Right,
                    style,
                    layout,
                    valueAxisUsesPercentDefaults);
            }

            if (hasSecondaryAxis && layout.ShowValueAxis && layout.ShowValueAxisLabels) {
                AddSecondaryValueAxisLabels(drawing, secondaryAxis, plotTop, plotHeight,
                    axisLabelRight, axisLabelRightWidth, style, layout);
            }

            if (layout.ShowCategoryAxis && layout.ShowCategoryAxisLabels) {
                if (IsScatterChart(snapshot.ChartKind)) {
                    IReadOnlyList<double> sharedXValues = GetScatterXValues(snapshot.Data.Categories);
                    List<OfficeChartSeries> scatterSeries = GetRenderableScatterSeries(snapshot).Select(item => item.Series).ToList();
                    ValueRange scatterXRange = ApplyValueAxisScale(GetScatterPointRanges(scatterSeries, sharedXValues).XRange, layout, horizontal: true);
                    AddHorizontalValueAxisLabels(
                        drawing,
                        scatterXRange,
                        plotLeft,
                        horizontalAxisLabelsHigh ? plotTop - 13D : plotBottomY + 4D,
                        plotWidth,
                        horizontalValueAxisLabelWidth,
                        horizontalAxisLabelsHigh,
                        style,
                        layout,
                        percentDefault: false);
                } else {
                    AddCategoryAxisLabels(
                        drawing,
                        snapshot.Data.Categories,
                        plotLeft,
                        horizontalAxisLabelsHigh ? plotTop - 13D : plotBottomY + 7D,
                        plotWidth,
                        style,
                        layout);
                }
            }

            AddAxisTitles(drawing, layout.ShowValueAxis ? layout.ValueAxisTitle : null, layout.ShowCategoryAxis ? layout.CategoryAxisTitle : null, plotLeft, plotTop, plotBottomY, plotWidth, plotHeight, style, layout);
        }

        if (layout.OverlayLegend) {
            AddOverlaySeriesLegend(drawing, legendSeries, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        } else {
            AddSeriesLegend(
                drawing,
                legendSeries,
                leftLegend ? 6D : width - legendWidth + 6D,
                plotTop,
                Math.Max(0D, legendWidth - 12D),
                plotHeight,
                style,
                layout);
        }

        if (!layout.OverlayLegend && bottomLegendHeight > 0D) {
            AddSeriesLegendBand(drawing, legendSeries, 8D, height - bottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout);
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

    private static void AddVerticalGridLines(
        OfficeDrawing drawing,
        double plotLeft,
        double plotTop,
        double plotWidth,
        double plotHeight,
        int divisions,
        int startIndex,
        int step,
        OfficeColor color,
        double lineWidth,
        OfficeStrokeDashStyle dashStyle) {
        for (int i = startIndex; i < divisions; i += step) {
            double x = plotLeft + plotWidth * i / divisions;
            AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), x, plotTop, null, color, lineWidth, dashStyle);
        }
    }

    private static void AddHorizontalGridLines(
        OfficeDrawing drawing,
        double plotLeft,
        double plotTop,
        double plotWidth,
        double plotHeight,
        int divisions,
        int startIndex,
        int step,
        OfficeColor color,
        double lineWidth,
        OfficeStrokeDashStyle dashStyle) {
        for (int i = startIndex; i < divisions; i += step) {
            double y = plotTop + plotHeight * i / divisions;
            AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, y, null, color, lineWidth, dashStyle);
        }
    }

    private static void AddVerticalValueGridLines(
        OfficeDrawing drawing,
        double plotLeft,
        double plotTop,
        double plotWidth,
        double plotHeight,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeColor color,
        double lineWidth,
        OfficeStrokeDashStyle dashStyle) {
        foreach (double tick in ticks) {
            if (tick <= range.Min || tick >= range.Max) {
                continue;
            }

            double x = ToPlotX(tick, range.Min, range.Max, plotLeft, plotWidth);
            AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), x, plotTop, null, color, lineWidth, dashStyle);
        }
    }

    private static void AddHorizontalValueGridLines(
        OfficeDrawing drawing,
        double plotLeft,
        double plotTop,
        double plotWidth,
        double plotHeight,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeColor color,
        double lineWidth,
        OfficeStrokeDashStyle dashStyle) {
        for (int i = ticks.Count - 1; i >= 0; i--) {
            double tick = ticks[i];
            if (tick <= range.Min || tick >= range.Max) {
                continue;
            }

            double y = ToPlotY(tick, range.Min, range.Max, plotTop, plotHeight);
            AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, y, null, color, lineWidth, dashStyle);
        }
    }

    private static void AddHorizontalAxisMajorTickMarks(
        OfficeDrawing drawing,
        double plotLeft,
        double axisY,
        double plotWidth,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth,
        bool positiveOutside = true) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside);
        for (int i = 0; i <= 4; i++) {
            double x = plotLeft + plotWidth * i / 4D;
            AddShape(drawing, OfficeShape.Line(0D, start, 0D, end), x, axisY, null, color, lineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddHorizontalValueAxisMajorTickMarks(
        OfficeDrawing drawing,
        double plotLeft,
        double axisY,
        double plotWidth,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth,
        bool positiveOutside = true) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside);
        foreach (double tick in ticks) {
            double x = ToPlotX(tick, range.Min, range.Max, plotLeft, plotWidth);
            AddShape(drawing, OfficeShape.Line(0D, start, 0D, end), x, axisY, null, color, lineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddHorizontalAxisMinorTickMarks(
        OfficeDrawing drawing,
        double plotLeft,
        double axisY,
        double plotWidth,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth,
        bool positiveOutside = true) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside);
        for (int i = 1; i < 8; i += 2) {
            double x = plotLeft + plotWidth * i / 8D;
            AddShape(
                drawing,
                OfficeShape.Line(0D, start, 0D, end),
                x,
                axisY,
                null,
                color,
                Math.Max(0.5D, lineWidth * 0.8D),
                OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddHorizontalValueAxisMinorTickMarks(
        OfficeDrawing drawing,
        double plotLeft,
        double axisY,
        double plotWidth,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth,
        bool positiveOutside = true) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside);
        double minorLineWidth = Math.Max(0.5D, lineWidth * 0.8D);
        foreach (double tick in ticks) {
            double x = ToPlotX(tick, range.Min, range.Max, plotLeft, plotWidth);
            AddShape(drawing, OfficeShape.Line(0D, start, 0D, end), x, axisY, null, color, minorLineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddVerticalAxisMajorTickMarks(
        OfficeDrawing drawing,
        double axisX,
        double plotTop,
        double plotHeight,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside: false);
        for (int i = 0; i <= 4; i++) {
            double y = plotTop + plotHeight * i / 4D;
            AddShape(drawing, OfficeShape.Line(start, 0D, end, 0D), axisX, y, null, color, lineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddVerticalValueAxisMajorTickMarks(
        OfficeDrawing drawing,
        double axisX,
        double plotTop,
        double plotHeight,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside: false);
        for (int i = ticks.Count - 1; i >= 0; i--) {
            double tick = ticks[i];
            double y = ToPlotY(tick, range.Min, range.Max, plotTop, plotHeight);
            AddShape(drawing, OfficeShape.Line(start, 0D, end, 0D), axisX, y, null, color, lineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddVerticalValueAxisMinorTickMarks(
        OfficeDrawing drawing,
        double axisX,
        double plotTop,
        double plotHeight,
        ValueRange range,
        IReadOnlyList<double> ticks,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside: false);
        double minorLineWidth = Math.Max(0.5D, lineWidth * 0.8D);
        for (int i = ticks.Count - 1; i >= 0; i--) {
            double tick = ticks[i];
            double y = ToPlotY(tick, range.Min, range.Max, plotTop, plotHeight);
            AddShape(drawing, OfficeShape.Line(start, 0D, end, 0D), axisX, y, null, color, minorLineWidth, OfficeStrokeDashStyle.Solid);
        }
    }

    private static void AddVerticalAxisMinorTickMarks(
        OfficeDrawing drawing,
        double axisX,
        double plotTop,
        double plotHeight,
        OfficeChartAxisTickMark tickMark,
        OfficeColor color,
        double lineWidth) {
        if (tickMark == OfficeChartAxisTickMark.None) {
            return;
        }

        (double start, double end) = GetAxisTickMarkOffsets(tickMark, positiveOutside: false);
        for (int i = 1; i < 8; i += 2) {
            double y = plotTop + plotHeight * i / 8D;
            AddShape(
                drawing,
                OfficeShape.Line(start, 0D, end, 0D),
                axisX,
                y,
                null,
                color,
                Math.Max(0.5D, lineWidth * 0.8D),
                OfficeStrokeDashStyle.Solid);
        }
    }

    private static (double Start, double End) GetAxisTickMarkOffsets(OfficeChartAxisTickMark tickMark, bool positiveOutside) {
        const double tickLength = 4D;
        double outside = positiveOutside ? tickLength : -tickLength;
        double inside = -outside;
        return tickMark switch {
            OfficeChartAxisTickMark.Inside => (0D, inside),
            OfficeChartAxisTickMark.Outside => (0D, outside),
            OfficeChartAxisTickMark.Cross => (-tickLength / 2D, tickLength / 2D),
            _ => (0D, 0D)
        };
    }

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

    private static double GetSeriesStrokeWidth(OfficeChartSeries series, double fallbackWidth) =>
        series.StrokeWidth.HasValue ? Math.Max(0.1D, series.StrokeWidth.Value) : fallbackWidth;

    private static OfficeStrokeDashStyle GetSeriesStrokeDashStyle(OfficeChartSeries series) =>
        series.StrokeDashStyle ?? OfficeStrokeDashStyle.Solid;

    private static double GetMarkerDiameter(OfficeChartSeries series, double fallbackDiameter) =>
        series.MarkerSize.HasValue ? Math.Max(1D, series.MarkerSize.Value) : fallbackDiameter;

    private static void AddMarker(OfficeDrawing drawing, OfficeChartSeries series, OfficePoint point, double fallbackDiameter, OfficeColor color, double strokeWidth) {
        double diameter = GetMarkerDiameter(series, fallbackDiameter);
        double left = point.X - diameter / 2D;
        double top = point.Y - diameter / 2D;
        OfficeColor markerStroke = series.MarkerOutlineColor ?? color;
        double markerStrokeWidth = series.MarkerOutlineWidth ?? strokeWidth;
        if (series.MarkerShape == OfficeChartMarkerShape.Dash) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, diameter, 0D), left, point.Y, null, markerStroke, markerStrokeWidth);
            return;
        }

        if (series.MarkerShape == OfficeChartMarkerShape.Dot) {
            double dotDiameter = Math.Max(1D, diameter * 0.45D);
            double dotLeft = point.X - dotDiameter / 2D;
            double dotTop = point.Y - dotDiameter / 2D;
            AddShape(drawing, OfficeShape.Ellipse(dotDiameter, dotDiameter), dotLeft, dotTop, color, markerStroke, markerStrokeWidth);
            return;
        }

        if (series.MarkerShape == OfficeChartMarkerShape.Plus) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, diameter, 0D), left, point.Y, null, markerStroke, markerStrokeWidth);
            AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, diameter), point.X, top, null, markerStroke, markerStrokeWidth);
            return;
        }

        if (series.MarkerShape == OfficeChartMarkerShape.X) {
            AddShape(drawing, OfficeShape.Line(0D, 0D, diameter, diameter), left, top, null, markerStroke, markerStrokeWidth);
            AddShape(drawing, OfficeShape.Line(0D, diameter, diameter, 0D), left, top, null, markerStroke, markerStrokeWidth);
            return;
        }

        OfficeShape shape;
        switch (series.MarkerShape.GetValueOrDefault(OfficeChartMarkerShape.Circle)) {
            case OfficeChartMarkerShape.Square:
                shape = OfficeShape.Rectangle(diameter, diameter);
                break;
            case OfficeChartMarkerShape.Diamond:
                shape = OfficeShape.Polygon(
                    new OfficePoint(diameter / 2D, 0D),
                    new OfficePoint(diameter, diameter / 2D),
                    new OfficePoint(diameter / 2D, diameter),
                    new OfficePoint(0D, diameter / 2D));
                break;
            case OfficeChartMarkerShape.Triangle:
                shape = OfficeShape.Polygon(
                    new OfficePoint(diameter / 2D, 0D),
                    new OfficePoint(diameter, diameter),
                    new OfficePoint(0D, diameter));
                break;
            case OfficeChartMarkerShape.Star:
                if (!OfficeShapePresets.TryCreate("star5", diameter, diameter, out OfficeShape? star) || star == null) {
                    shape = OfficeShape.Ellipse(diameter, diameter);
                    break;
                }

                shape = star;
                break;
            default:
                shape = OfficeShape.Ellipse(diameter, diameter);
                break;
        }

        AddShape(
            drawing,
            shape,
            left,
            top,
            color,
            markerStroke,
            markerStrokeWidth);
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

    private static void AddBarSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft,
        double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        ValueRange? sharedValueAxisRange = null, OfficeChartAxisGroup? axisGroup = null) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        var barSeries = new List<(OfficeChartSeries Series, int SourceIndex)>();
        for (int i = 0; i < series.Count; i++) {
            if (ShouldRenderSeriesAsBarOrColumn(snapshot, series[i]) &&
                (!axisGroup.HasValue || series[i].AxisGroup == axisGroup.Value)) {
                barSeries.Add((series[i], i));
            }
        }

        if (barSeries.Count == 0) {
            return;
        }

        IReadOnlyList<OfficeChartSeries> barSeriesValues = barSeries.Select(item => item.Series).ToArray();
        var stackedSlots = new Dictionary<OfficeChartKind, int>();
        var clusteredSlots = new Dictionary<int, int>();
        var stackedSeriesByKind = new Dictionary<OfficeChartKind, List<OfficeChartSeries>>();
        int slotCount = 0;
        for (int i = 0; i < barSeries.Count; i++) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, barSeries[i].Series);
            if (IsStackedBarOrColumnChart(kind) || IsPercentStackedBarOrColumnChart(kind)) {
                if (!stackedSeriesByKind.TryGetValue(kind, out List<OfficeChartSeries>? stackGroup)) {
                    stackGroup = new List<OfficeChartSeries>();
                    stackedSeriesByKind.Add(kind, stackGroup);
                }

                stackGroup.Add(barSeries[i].Series);
                if (!stackedSlots.ContainsKey(kind)) {
                    stackedSlots[kind] = slotCount++;
                }
            } else {
                clusteredSlots[barSeries[i].SourceIndex] = slotCount++;
            }
        }

        var percentStackedTotalsByKind = new Dictionary<OfficeChartKind, PercentStackedTotals>();
        foreach (KeyValuePair<OfficeChartKind, List<OfficeChartSeries>> stackGroup in stackedSeriesByKind) {
            if (IsPercentStackedBarOrColumnChart(stackGroup.Key)) {
                percentStackedTotalsByKind.Add(stackGroup.Key, BuildPercentStackedTotals(stackGroup.Value, categories.Count));
            }
        }

        double slot = plotWidth / categories.Count;
        double groupWidth = slot * 0.68D;
        int barSeriesCount = Math.Max(1, slotCount);
        double barWidth = Math.Max(2D, groupWidth / barSeriesCount);
        ValueRange horizontalRange;
        ValueRange verticalRange;
        if (sharedValueAxisRange.HasValue) {
            horizontalRange = sharedValueAxisRange.Value;
            verticalRange = sharedValueAxisRange.Value;
        } else {
            ValueRange baseRange = GetBarSeriesRenderRange(snapshot, barSeries, categories.Count);
            horizontalRange = ResolveRenderedBarRange(baseRange, layout, horizontal: true);
            verticalRange = ResolveRenderedBarRange(baseRange, layout, horizontal: false);
        }

        for (int category = 0; category < categories.Count; category++) {
            var positiveBases = new Dictionary<OfficeChartKind, double>();
            var negativeBases = new Dictionary<OfficeChartKind, double>();
            for (int s = 0; s < barSeries.Count; s++) {
                OfficeChartSeries currentSeries = barSeries[s].Series;
                int sourceSeriesIndex = barSeries[s].SourceIndex;
                OfficeChartKind seriesKind = GetEffectiveSeriesKind(snapshot, currentSeries);
                bool seriesStacked = IsStackedBarOrColumnChart(seriesKind) || IsPercentStackedBarOrColumnChart(seriesKind);
                bool seriesPercentStacked = IsPercentStackedBarOrColumnChart(seriesKind);

                if (!TryGetSeriesValue(currentSeries, category, out double value)) {
                    continue;
                }

                if (value == 0D && !ShouldShowDataLabel(layout, sourceSeriesIndex, category)) {
                    continue;
                }

                double baseline = 0D;
                double plottedValue = value;
                if (seriesStacked) {
                    if (seriesPercentStacked) {
                        plottedValue = NormalizePercentStackedValue(percentStackedTotalsByKind[seriesKind], category, value);
                    }

                    positiveBases.TryGetValue(seriesKind, out double positiveBase);
                    negativeBases.TryGetValue(seriesKind, out double negativeBase);
                    baseline = plottedValue >= 0D ? positiveBase : negativeBase;
                    if (plottedValue >= 0D) {
                        positiveBases[seriesKind] = positiveBase + plottedValue;
                    } else {
                        negativeBases[seriesKind] = negativeBase + plottedValue;
                    }
                }

                OfficeColor color = GetSeriesColor(style, series, sourceSeriesIndex);
                if (currentSeries.PointColors != null && category < currentSeries.PointColors!.Count && currentSeries.PointColors![category].HasValue) {
                    color = GetPointColor(style, currentSeries.PointColors, category);
                }

                bool horizontal = IsBarChart(snapshot.ChartKind);
                ValueRange range = horizontal ? horizontalRange : verticalRange;
                double min = range.Min;
                double max = range.Max;
                int layoutSlot = seriesStacked
                    ? stackedSlots[seriesKind]
                    : clusteredSlots[sourceSeriesIndex];

                if (horizontal) {
                    double categoryHeight = plotHeight / categories.Count;
                    double rowHeight = Math.Max(2D, categoryHeight * 0.68D / barSeriesCount);
                    int categorySlot = GetHorizontalBarCategorySlotIndex(category, categories.Count, layout);
                    int seriesSlot = barSeriesCount - 1 - layoutSlot;
                    double y = plotTop + categoryHeight * categorySlot + categoryHeight * 0.16D + rowHeight * seriesSlot;
                    double visibleBaseline = ClampValueToRange(baseline, min, max);
                    double visibleValue = ClampValueToRange(seriesStacked ? baseline + plottedValue : plottedValue, min, max);
                    double x1 = ToPlotX(visibleBaseline, min, max, plotLeft, plotWidth);
                    double x2 = ToPlotX(visibleValue, min, max, plotLeft, plotWidth);
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
                        currentSeries,
                        value,
                        GetDataLabelCategoryTotal(barSeriesValues, category),
                        x,
                        x + w,
                        y,
                        y + rowHeight,
                        sourceSeriesIndex,
                        category);
                } else {
                    int categorySlotIndex = GetCategorySlotIndex(category, categories.Count, layout);
                    double x = plotLeft + slot * categorySlotIndex + (slot - groupWidth) / 2D + barWidth * layoutSlot;
                    double visibleBaseline = ClampValueToRange(baseline, min, max);
                    double visibleValue = ClampValueToRange(seriesStacked ? baseline + plottedValue : plottedValue, min, max);
                    double y1 = ToPlotY(visibleBaseline, min, max, plotTop, plotHeight);
                    double y2 = ToPlotY(visibleValue, min, max, plotTop, plotHeight);
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
                        currentSeries,
                        value,
                        GetDataLabelCategoryTotal(barSeriesValues, category),
                        x + barWidth * 0.44D,
                        y,
                        y + h,
                        sourceSeriesIndex,
                        category);
                }
            }
        }
    }

    private static ValueRange GetBarSeriesRenderRange(OfficeChartSnapshot snapshot, IReadOnlyList<(OfficeChartSeries Series, int SourceIndex)> barSeries, int categoryCount) {
        var ranges = new List<ValueRange>();
        var clusteredSeries = new List<OfficeChartSeries>();
        var stackedGroups = new Dictionary<OfficeChartKind, List<OfficeChartSeries>>();
        for (int i = 0; i < barSeries.Count; i++) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, barSeries[i].Series);
            if (IsStackedBarOrColumnChart(kind) || IsPercentStackedBarOrColumnChart(kind)) {
                if (!stackedGroups.TryGetValue(kind, out List<OfficeChartSeries>? group)) {
                    group = new List<OfficeChartSeries>();
                    stackedGroups[kind] = group;
                }

                group.Add(barSeries[i].Series);
            } else {
                clusteredSeries.Add(barSeries[i].Series);
            }
        }

        if (clusteredSeries.Count > 0) {
            ValueRange finiteRange = GetFiniteSeriesRange(clusteredSeries);
            ranges.Add(ExpandFlatRange(Math.Min(0D, finiteRange.Min), Math.Max(0D, finiteRange.Max)));
        }

        foreach (KeyValuePair<OfficeChartKind, List<OfficeChartSeries>> group in stackedGroups) {
            ranges.Add(IsPercentStackedBarOrColumnChart(group.Key)
                ? GetPercentStackedSeriesRange(group.Value, categoryCount)
                : GetStackedSeriesRange(group.Value, categoryCount));
        }

        if (ranges.Count == 0) {
            return GetCartesianValueRange(snapshot);
        }

        double min = ranges[0].Min;
        double max = ranges[0].Max;
        for (int i = 1; i < ranges.Count; i++) {
            min = Math.Min(min, ranges[i].Min);
            max = Math.Max(max, ranges[i].Max);
        }

        return ExpandFlatRange(min, max);
    }

    private static double ClampValueToRange(double value, double min, double max) =>
        Math.Max(min, Math.Min(max, value));

    private static ValueRange ResolveRenderedBarRange(ValueRange range, OfficeChartLayout layout, bool horizontal) {
        range = ApplyValueAxisScale(range, layout, horizontal);
        bool hasValueAxisScale = HasValueAxisScale(layout, horizontal);
        double min = hasValueAxisScale ? range.Min : Math.Min(0D, range.Min);
        double max = hasValueAxisScale ? range.Max : Math.Max(0D, range.Max);
        if (max <= min) {
            max = min + 1D;
        }

        return new ValueRange(min, max);
    }

    private static void AddAreaSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft,
        double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        ValueRange? sharedValueAxisRange = null, OfficeChartAxisGroup? axisGroup = null) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> areaSeries = GetRenderableAreaSeries(snapshot);
        if (axisGroup.HasValue) areaSeries = areaSeries.Where(item =>
            item.Series.AxisGroup == axisGroup.Value).ToList();
        if (categories.Count < 2 || areaSeries.Count == 0) {
            return;
        }

        ValueRange range = sharedValueAxisRange ?? ApplyValueAxisScale(GetAreaSeriesRenderRange(snapshot, areaSeries, categories.Count, layout), layout, horizontal: false);
        double step = plotWidth / (categories.Count - 1);
        var positiveCumulativeByKind = new Dictionary<OfficeChartKind, double[]>();
        var negativeCumulativeByKind = new Dictionary<OfficeChartKind, double[]>();
        var stackedSeriesByKind = areaSeries
            .Where(item => IsStackedAreaChart(item.Kind) || IsPercentStackedAreaChart(item.Kind))
            .GroupBy(item => item.Kind)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Series).ToList());
        var percentStackedTotalsByKind = stackedSeriesByKind
            .Where(group => IsPercentStackedAreaChart(group.Key))
            .ToDictionary(group => group.Key, group => BuildPercentStackedTotals(group.Value, categories.Count));

        foreach ((OfficeChartSeries currentSeries, int sourceSeriesIndex, OfficeChartKind kind) in areaSeries) {
            bool currentStacked = IsStackedAreaChart(kind) || IsPercentStackedAreaChart(kind);
            bool currentPercentStacked = IsPercentStackedAreaChart(kind);
            double[] positiveCumulative = currentStacked ? GetCumulative(positiveCumulativeByKind, kind, categories.Count) : Array.Empty<double>();
            double[] negativeCumulative = currentStacked ? GetCumulative(negativeCumulativeByKind, kind, categories.Count) : Array.Empty<double>();
            OfficeColor color = GetSeriesColor(style, series, sourceSeriesIndex);
            double strokeWidth = GetSeriesStrokeWidth(currentSeries, 1.4D);
            OfficeStrokeDashStyle dashStyle = GetSeriesStrokeDashStyle(currentSeries);
            var topPoints = new List<OfficePoint>(categories.Count);
            var bottomPoints = new List<OfficePoint>(categories.Count);
            var runCategoryIndices = new List<int>(categories.Count);

            for (int i = 0; i < categories.Count; i++) {
                if (!TryGetSeriesValue(currentSeries, i, out double value)) {
                    AddAreaRun(drawing, topPoints, bottomPoints, color, strokeWidth, dashStyle);
                    AddAreaRunDataLabels(drawing, layout, style, categories, series, sourceSeriesIndex, runCategoryIndices, topPoints);
                    topPoints.Clear();
                    bottomPoints.Clear();
                    runCategoryIndices.Clear();
                    continue;
                }

                double rawValue = currentPercentStacked ? NormalizePercentStackedValue(percentStackedTotalsByKind[kind], i, value) : value;
                double baseline = currentStacked
                    ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                    : 0D;
                double topValue = baseline + rawValue;

                double x = GetCategoryPointX(plotLeft, step, i, categories.Count, layout);
                topPoints.Add(new OfficePoint(x, ToPlotY(topValue, range.Min, range.Max, plotTop, plotHeight)));
                bottomPoints.Add(new OfficePoint(x, ToPlotY(baseline, range.Min, range.Max, plotTop, plotHeight)));
                runCategoryIndices.Add(i);

                if (currentStacked) {
                    if (rawValue >= 0D) {
                        positiveCumulative[i] += rawValue;
                    } else {
                        negativeCumulative[i] += rawValue;
                    }
                }
            }

            AddAreaRun(drawing, topPoints, bottomPoints, color, strokeWidth, dashStyle);
            AddAreaRunDataLabels(drawing, layout, style, categories, series, sourceSeriesIndex, runCategoryIndices, topPoints);
        }
    }

    private static void AddAreaRun(OfficeDrawing drawing, IReadOnlyList<OfficePoint> topPoints, IReadOnlyList<OfficePoint> bottomPoints, OfficeColor color, double strokeWidth, OfficeStrokeDashStyle dashStyle) {
        if (topPoints.Count < 2 || bottomPoints.Count != topPoints.Count) {
            return;
        }

        var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
        areaPoints.AddRange(topPoints);
        for (int i = bottomPoints.Count - 1; i >= 0; i--) {
            areaPoints.Add(bottomPoints[i]);
        }

        AddPolygonShape(drawing, areaPoints, color, color, 0.5D, 0.32D);
        AddPointLine(drawing, topPoints, color, strokeWidth, dashStyle);
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
                topPoints[i].Y,
                seriesIndex,
                categoryIndex);
        }
    }

    private static void AddLineSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft,
        double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        ValueRange? sharedValueAxisRange = null, OfficeChartAxisGroup? axisGroup = null) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> lineSeries = GetRenderableLineSeries(snapshot);
        if (axisGroup.HasValue) lineSeries = lineSeries.Where(item =>
            item.Series.AxisGroup == axisGroup.Value).ToList();
        if (categories.Count == 0 || lineSeries.Count == 0) {
            return;
        }

        ValueRange range = sharedValueAxisRange ?? ApplyValueAxisScale(GetLineSeriesRenderRange(snapshot, lineSeries, categories.Count, layout), layout, horizontal: false);
        double step = categories.Count > 1 ? plotWidth / (categories.Count - 1) : 0D;
        var positiveCumulativeByKind = new Dictionary<OfficeChartKind, double[]>();
        var negativeCumulativeByKind = new Dictionary<OfficeChartKind, double[]>();
        var stackedSeriesByKind = lineSeries
            .Where(item => IsStackedLineChart(item.Kind) || IsPercentStackedLineChart(item.Kind))
            .GroupBy(item => item.Kind)
            .ToDictionary(group => group.Key, group => group.Select(item => item.Series).ToList());
        var percentStackedTotalsByKind = stackedSeriesByKind
            .Where(group => IsPercentStackedLineChart(group.Key))
            .ToDictionary(group => group.Key, group => BuildPercentStackedTotals(group.Value, categories.Count));
        foreach ((OfficeChartSeries currentSeries, int sourceSeriesIndex, OfficeChartKind kind) in lineSeries) {
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

                points[i] = new OfficePoint(GetCategoryPointX(plotLeft, step, i, categories.Count, layout), ToPlotY(plottedValue, range.Min, range.Max, plotTop, plotHeight));
                plotted[i] = true;
            }

            if (currentSeries.ConnectLine) {
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
                    AddShape(drawing, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, color, strokeWidth, dashStyle);
                }
            }

            for (int i = 0; i < categories.Count; i++) {
                if (!plotted[i]) {
                    continue;
                }

                if (layout.ShowMarkers && currentSeries.ShowMarkers) {
                    OfficeColor pointColor = GetPointColor(currentSeries.PointColors, i, color);
                    AddMarker(drawing, currentSeries, points[i], 4D, pointColor, 1D);
                }

                double value = GetSeriesValue(currentSeries, i);
                AddPointDataLabel(
                    drawing,
                    layout,
                    style,
                    categories[i],
                    currentSeries,
                    value,
                    GetDataLabelCategoryTotal(series, i),
                    points[i].X,
                    points[i].Y,
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
    }

    private static ValueRange GetLineSeriesRenderRange(OfficeChartSnapshot snapshot, IReadOnlyList<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> lineSeries, int categoryCount, OfficeChartLayout layout) {
        var ranges = new List<ValueRange>();
        var standardSeries = new List<OfficeChartSeries>();
        var stackedGroups = new Dictionary<OfficeChartKind, List<OfficeChartSeries>>();
        for (int i = 0; i < lineSeries.Count; i++) {
            OfficeChartKind kind = lineSeries[i].Kind;
            if (IsStackedLineChart(kind) || IsPercentStackedLineChart(kind)) {
                if (!stackedGroups.TryGetValue(kind, out List<OfficeChartSeries>? group)) {
                    group = new List<OfficeChartSeries>();
                    stackedGroups[kind] = group;
                }

                group.Add(lineSeries[i].Series);
            } else {
                standardSeries.Add(lineSeries[i].Series);
            }
        }

        if (standardSeries.Count > 0) {
            ranges.Add(GetCartesianValueRange(snapshot, layout, horizontalValueAxis: false));
        }

        foreach (KeyValuePair<OfficeChartKind, List<OfficeChartSeries>> group in stackedGroups) {
            ranges.Add(IsPercentStackedLineChart(group.Key)
                ? GetPercentStackedSeriesRange(group.Value, categoryCount)
                : GetStackedSeriesRange(group.Value, categoryCount));
        }

        if (ranges.Count == 0) {
            return GetCartesianValueRange(snapshot, layout, horizontalValueAxis: false);
        }

        double min = ranges[0].Min;
        double max = ranges[0].Max;
        for (int i = 1; i < ranges.Count; i++) {
            min = Math.Min(min, ranges[i].Min);
            max = Math.Max(max, ranges[i].Max);
        }

        return ExpandFlatRange(min, max);
    }

    private static ValueRange GetAreaSeriesRenderRange(OfficeChartSnapshot snapshot, IReadOnlyList<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> areaSeries, int categoryCount, OfficeChartLayout layout) {
        var ranges = new List<ValueRange>();
        var standardSeries = new List<OfficeChartSeries>();
        var stackedGroups = new Dictionary<OfficeChartKind, List<OfficeChartSeries>>();
        for (int i = 0; i < areaSeries.Count; i++) {
            OfficeChartKind kind = areaSeries[i].Kind;
            if (IsStackedAreaChart(kind) || IsPercentStackedAreaChart(kind)) {
                if (!stackedGroups.TryGetValue(kind, out List<OfficeChartSeries>? group)) {
                    group = new List<OfficeChartSeries>();
                    stackedGroups[kind] = group;
                }

                group.Add(areaSeries[i].Series);
            } else {
                standardSeries.Add(areaSeries[i].Series);
            }
        }

        if (standardSeries.Count > 0) {
            ranges.Add(GetCartesianValueRange(snapshot, layout, horizontalValueAxis: false));
        }

        foreach (KeyValuePair<OfficeChartKind, List<OfficeChartSeries>> group in stackedGroups) {
            ranges.Add(IsPercentStackedAreaChart(group.Key)
                ? GetPercentStackedSeriesRange(group.Value, categoryCount)
                : GetStackedSeriesRange(group.Value, categoryCount));
        }

        if (ranges.Count == 0) {
            return GetCartesianValueRange(snapshot, layout, horizontalValueAxis: false);
        }

        double min = ranges[0].Min;
        double max = ranges[0].Max;
        for (int i = 1; i < ranges.Count; i++) {
            min = Math.Min(min, ranges[i].Min);
            max = Math.Max(max, ranges[i].Max);
        }

        return ExpandFlatRange(min, max);
    }

    private static double[] GetCumulative(Dictionary<OfficeChartKind, double[]> cumulativeByKind, OfficeChartKind kind, int categoryCount) {
        if (!cumulativeByKind.TryGetValue(kind, out double[]? values)) {
            values = new double[categoryCount];
            cumulativeByKind[kind] = values;
        }

        return values;
    }

    private static List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> GetRenderableAreaSeries(OfficeChartSnapshot snapshot) {
        var items = new List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)>();
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        for (int i = 0; i < series.Count; i++) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, series[i]);
            if (!HasMixedCartesianSeriesKinds(snapshot) || IsAreaChart(kind)) {
                items.Add((series[i], i, kind));
            }
        }

        return items;
    }

    private static List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> GetRenderableLineSeries(OfficeChartSnapshot snapshot) {
        var items = new List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)>();
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        for (int i = 0; i < series.Count; i++) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, series[i]);
            if (!HasMixedCartesianSeriesKinds(snapshot) || IsLineChart(kind)) {
                items.Add((series[i], i, kind));
            }
        }

        return items;
    }

    private static List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> GetRenderableScatterSeries(OfficeChartSnapshot snapshot) {
        if (HasMixedScatterSeriesOnCategoryAxes(snapshot)) {
            return new List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)>();
        }

        var items = new List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)>();
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        for (int i = 0; i < series.Count; i++) {
            OfficeChartKind kind = GetEffectiveSeriesKind(snapshot, series[i]);
            if (!HasMixedCartesianSeriesKinds(snapshot) || IsScatterChart(kind)) {
                items.Add((series[i], i, kind));
            }
        }

        return items;
    }

    private static bool HasMixedScatterSeriesOnCategoryAxes(OfficeChartSnapshot snapshot) =>
        HasMixedCartesianSeriesKinds(snapshot) &&
        !IsScatterChart(snapshot.ChartKind) &&
        snapshot.Data.Series.Any(series => IsScatterChart(GetEffectiveSeriesKind(snapshot, series)));

    private static IReadOnlyList<OfficeChartSeries> GetRenderableLegendSeries(OfficeChartSnapshot snapshot) =>
        HasMixedScatterSeriesOnCategoryAxes(snapshot)
            ? snapshot.Data.Series.Where(series => !IsScatterChart(GetEffectiveSeriesKind(snapshot, series))).ToList()
            : snapshot.Data.Series;

    private static void AddScatterSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot, double plotLeft,
        double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        ValueRange? valueAxisRange = null, OfficeChartAxisGroup? axisGroup = null) {
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        IReadOnlyList<OfficeChartSeries> series = snapshot.Data.Series;
        if (categories.Count == 0 || series.Count == 0) {
            return;
        }

        IReadOnlyList<double> sharedXValues = GetScatterXValues(categories);
        List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> allScatterSeries =
            GetRenderableScatterSeries(snapshot);
        List<(OfficeChartSeries Series, int SourceIndex, OfficeChartKind Kind)> scatterSeries = axisGroup.HasValue
            ? allScatterSeries.Where(item => item.Series.AxisGroup == axisGroup.Value).ToList()
            : allScatterSeries;
        if (scatterSeries.Count == 0) {
            return;
        }

        List<OfficeChartSeries> allRangeSeries = allScatterSeries.Select(item => item.Series).ToList();
        List<OfficeChartSeries> rangeSeries = scatterSeries.Select(item => item.Series).ToList();
        ValueRange pairedXRange = GetScatterPointRanges(allRangeSeries, sharedXValues).XRange;
        ValueRange pairedYRange = GetScatterPointRanges(rangeSeries, sharedXValues).YRange;
        ValueRange xRange = ApplyValueAxisScale(pairedXRange, layout, horizontal: true);
        ValueRange yRange = valueAxisRange ?? ApplyValueAxisScale(pairedYRange, layout, horizontal: false);
        for (int s = 0; s < scatterSeries.Count; s++) {
            OfficeChartSeries currentSeries = scatterSeries[s].Series;
            int sourceSeriesIndex = scatterSeries[s].SourceIndex;
            OfficeColor color = GetSeriesColor(style, series, sourceSeriesIndex);
            double strokeWidth = GetSeriesStrokeWidth(currentSeries, 1.25D);
            OfficeStrokeDashStyle dashStyle = GetSeriesStrokeDashStyle(currentSeries);
            IReadOnlyList<double> xValues = currentSeries.XValues ?? sharedXValues;
            int pointCount = Math.Min(xValues.Count, currentSeries.Values.Count);
            var points = new List<(OfficePoint Point, int SourceIndex)>(pointCount);
            var lineSegment = new List<OfficePoint>(pointCount);
            for (int i = 0; i < pointCount; i++) {
                if (!TryGetSeriesValue(currentSeries, i, out double yValue)) {
                    if (layout.ConnectScatterPoints && currentSeries.ConnectLine) {
                        AddPointLine(drawing, lineSegment, color, strokeWidth, dashStyle);
                    }

                    lineSegment.Clear();
                    continue;
                }

                double xValue = xValues[i];
                if (!IsFiniteChartValue(xValue)) {
                    if (layout.ConnectScatterPoints && currentSeries.ConnectLine) {
                        AddPointLine(drawing, lineSegment, color, strokeWidth, dashStyle);
                    }

                    lineSegment.Clear();
                    continue;
                }

                double x = ToPlotX(xValue, xRange.Min, xRange.Max, plotLeft, plotWidth);
                double y = ToPlotY(yValue, yRange.Min, yRange.Max, plotTop, plotHeight);
                var point = new OfficePoint(x, y);
                points.Add((point, i));
                if (layout.ConnectScatterPoints && currentSeries.ConnectLine) {
                    lineSegment.Add(point);
                }
            }

            if (layout.ConnectScatterPoints && currentSeries.ConnectLine) {
                AddPointLine(drawing, lineSegment, color, strokeWidth, dashStyle);
            }
            for (int i = 0; i < points.Count; i++) {
                OfficePoint point = points[i].Point;
                if (layout.ShowMarkers && currentSeries.ShowMarkers) {
                    OfficeColor pointColor = GetPointColor(currentSeries.PointColors, points[i].SourceIndex, color);
                    AddMarker(drawing, currentSeries, point, 5D, pointColor, 1.25D);
                }

                int pointIndex = points[i].SourceIndex;
                string labelCategory = currentSeries.XValues != null && pointIndex < xValues.Count
                    ? xValues[pointIndex].ToString("0.####", CultureInfo.InvariantCulture)
                    : pointIndex < categories.Count ? categories[pointIndex] : string.Empty;
                AddPointDataLabel(
                    drawing,
                    layout,
                    style,
                    labelCategory,
                    currentSeries,
                    currentSeries.Values[pointIndex],
                    GetDataLabelCategoryTotal(series, pointIndex),
                    point.X,
                    point.Y,
                    sourceSeriesIndex,
                    pointIndex);
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
            double strokeWidth = GetSeriesStrokeWidth(series[s], 1D);
            OfficeStrokeDashStyle dashStyle = GetSeriesStrokeDashStyle(series[s]);
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

            if (allPointsPlotted && series[s].ConnectLine) {
                AddPolygonShape(
                    drawing,
                    points,
                    layout.FillRadarSeries ? color : null,
                    color,
                    strokeWidth,
                    layout.FillRadarSeries ? 0.18D : null,
                    dashStyle);
            } else if (series[s].ConnectLine) {
                for (int i = 1; i < categories.Count; i++) {
                    if (!plotted[i - 1] || !plotted[i]) {
                        continue;
                    }

                    AddPointLine(drawing, new[] { points[i - 1], points[i] }, color, strokeWidth, dashStyle);
                }
            }

            if (layout.ShowMarkers && series[s].ShowMarkers) {
                for (int i = 0; i < points.Length; i++) {
                    if (!plotted[i]) {
                        continue;
                    }

                    OfficePoint point = points[i];
                    OfficeColor pointColor = GetPointColor(series[s].PointColors, i, color);
                    AddMarker(drawing, series[s], point, 4D, pointColor, 1D);
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
                    points[i].Y,
                    s,
                    i);
            }
        }

        AddRadarCategoryLabels(drawing, categories, centerX, centerY, radius, style, layout);
        if (layout.OverlayLegend) {
            AddOverlaySeriesLegend(drawing, series, leftLegend ? legendWidth : 0D, contentTop + 4D, visualWidth, Math.Max(20D, contentHeight - 8D), style, layout);
        } else {
            AddSeriesLegend(
                drawing,
                series,
                leftLegend ? 6D : width - legendWidth + 6D,
                contentTop + 12D,
                Math.Max(0D, legendWidth - 12D),
                Math.Max(20D, contentHeight - 24D),
                style,
                layout);
        }

        if (!layout.OverlayLegend && bottomLegendHeight > 0D) {
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
                if (ShouldShowDataLabel(layout, 0, i)) {
                    AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, value, total, centerX, centerY, radius, start + sweep / 2D, zeroLabelIndex: null);
                }

                start = end;
            } else if (ShouldShowDataLabel(layout, 0, i)) {
                OfficeColor sliceColor = GetPointColor(style, values, 0);
                AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, 0D, total, centerX, centerY, radius, -Math.PI / 2D, zeroLabelIndex);
                zeroLabelIndex++;
            }
        }

        if (layout.OverlayLegend) {
            AddOverlayCategoryLegend(drawing, categories, leftLegend ? legendWidth : 0D, contentTop + 4D, visualWidth, Math.Max(20D, contentHeight - 8D), style, layout, categoryPointColors);
        } else {
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
        }

        if (!layout.OverlayLegend && categoryBottomLegendHeight > 0D) {
            AddCategoryLegendBand(drawing, categories, 8D, height - categoryBottomLegendHeight + 2D, Math.Max(1D, width - 16D), style, layout, categoryPointColors);
        }
    }

    private static void AddDoughnutSeries(OfficeDrawing drawing, IReadOnlyList<string> categories, IReadOnlyList<OfficeChartSeries> series, double width, double height, double contentTop, double bottomLegendHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        var renderableSeries = new List<(OfficeChartSeries Series, int SourceIndex)>();
        for (int s = 0; s < series.Count; s++) {
            if (GetPositiveSeriesTotal(series[s], categories.Count) > 0D) {
                renderableSeries.Add((series[s], s));
            }
        }

        if (renderableSeries.Count == 0) {
            return;
        }

        double topCategoryLegendHeight = layout.LegendPosition == OfficeChartLegendPosition.Top
            ? GetCategoryLegendBandHeight(categories, width - 16D, layout)
            : 0D;
        IReadOnlyList<OfficeColor?>? legendPointColors = GetLegendPointColors(style, renderableSeries.ConvertAll(item => item.Series), categories.Count);
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
            OfficeChartSeries values = renderableSeries[s].Series;
            int sourceSeriesIndex = renderableSeries[s].SourceIndex;
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
                    AddDoughnutSlice(drawing, centerX, centerY, outerRadius, innerRadius, start, sweep, sliceColor);
                    if (ShouldShowDataLabel(layout, sourceSeriesIndex, i)) {
                        AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, value, total, centerX, centerY, Math.Max(innerRadius + 8D, outerRadius - ringThickness * 0.42D), start + sweep / 2D, zeroLabelIndex: null);
                    }

                    start = end;
                } else if (s == 0 && ShouldShowDataLabel(layout, sourceSeriesIndex, i)) {
                    OfficeColor sliceColor = GetPointColor(style, values, 0);
                    AddPieDataLabel(drawing, layout, style, GetReadableDataLabelColor(sliceColor), categories[i], values, 0D, total, centerX, centerY, outerRadius, -Math.PI / 2D, zeroLabelIndex);
                    zeroLabelIndex++;
                }
            }
        }

        if (layout.OverlayLegend) {
            AddOverlayCategoryLegend(drawing, categories, leftLegend ? legendWidth : 0D, contentTop + 4D, visualWidth, Math.Max(20D, contentHeight - 8D), style, layout, legendPointColors);
        } else {
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
        }

        if (!layout.OverlayLegend && categoryBottomLegendHeight > 0D) {
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

        double labelWidth = Math.Min(78D, Math.Max(40D, label.Length * layout.DataLabelFontSize * 0.52D + 12D));
        double labelHeight = Math.Max(12D, layout.DataLabelFontSize + 6D);
        double distance = zeroLabelIndex.HasValue ? radius * 0.9D : radius * 0.58D;
        double x = centerX + Math.Cos(angle) * distance - labelWidth / 2D;
        double y = centerY + Math.Sin(angle) * distance - labelHeight / 2D;
        if (zeroLabelIndex.HasValue) {
            y += zeroLabelIndex.Value * (labelHeight + 1D);
        }

        AddDataLabel(drawing, layout, style, label, x, y, labelWidth, labelHeight, OfficeTextAlignment.Center, labelColor);
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

    private static void AddDoughnutSlice(OfficeDrawing drawing, double centerX, double centerY, double outerRadius, double innerRadius, double start, double sweep, OfficeColor color) {
        if (innerRadius <= 0D) {
            AddPieSlice(drawing, centerX, centerY, outerRadius, start, sweep, color);
            return;
        }

        int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
        var points = new List<OfficePoint>((segments + 1) * 2);
        for (int segment = 0; segment <= segments; segment++) {
            double angle = start + sweep * segment / segments;
            points.Add(new OfficePoint(
                centerX + Math.Cos(angle) * outerRadius,
                centerY + Math.Sin(angle) * outerRadius));
        }

        for (int segment = segments; segment >= 0; segment--) {
            double angle = start + sweep * segment / segments;
            points.Add(new OfficePoint(
                centerX + Math.Cos(angle) * innerRadius,
                centerY + Math.Sin(angle) * innerRadius));
        }

        AddPolygonShape(drawing, points, color, OfficeColor.White, 0.5D);
    }

    private static void AddShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
        shape.FillColor = fill;
        shape.StrokeColor = stroke;
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = dashStyle;
        drawing.AddShape(shape, x, y);
    }

    private static void AddPolygonShape(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, double? fillOpacity = null, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
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
        AddShape(drawing, shape, minX, minY, fill, stroke, strokeWidth, dashStyle);
    }

    private static void AddPointLine(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
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
                strokeWidth,
                dashStyle);
        }
    }

    private static OfficeColor GetCategoryAxisColor(OfficeChartStyle style) =>
        style.CategoryAxisColor ?? style.AxisColor;

    private static OfficeColor GetValueAxisColor(OfficeChartStyle style) =>
        style.ValueAxisColor ?? style.AxisColor;

    private static double GetCategoryAxisLineWidth(OfficeChartStyle style) =>
        style.CategoryAxisLineWidth ?? style.AxisLineWidth ?? 0.75D;

    private static double GetValueAxisLineWidth(OfficeChartStyle style) =>
        style.ValueAxisLineWidth ?? style.AxisLineWidth ?? 0.75D;

    private static OfficeStrokeDashStyle GetCategoryAxisLineDashStyle(OfficeChartStyle style) =>
        style.CategoryAxisLineDashStyle ?? style.AxisLineDashStyle ?? OfficeStrokeDashStyle.Solid;

    private static OfficeStrokeDashStyle GetValueAxisLineDashStyle(OfficeChartStyle style) =>
        style.ValueAxisLineDashStyle ?? style.AxisLineDashStyle ?? OfficeStrokeDashStyle.Solid;

    private static bool GetShowCategoryGridLines(OfficeChartStyle style) =>
        style.ShowCategoryGridLines.GetValueOrDefault(false);

    private static bool GetShowValueGridLines(OfficeChartStyle style) =>
        style.ShowValueGridLines ?? style.ShowGridLines;

    private static bool GetShowCategoryMinorGridLines(OfficeChartStyle style) =>
        style.ShowCategoryMinorGridLines.GetValueOrDefault(false);

    private static bool GetShowValueMinorGridLines(OfficeChartStyle style) =>
        style.ShowValueMinorGridLines.GetValueOrDefault(false);

    private static OfficeColor GetCategoryGridLineColor(OfficeChartStyle style) =>
        style.CategoryGridLineColor ?? style.GridLineColor;

    private static OfficeColor GetValueGridLineColor(OfficeChartStyle style) =>
        style.ValueGridLineColor ?? style.GridLineColor;

    private static OfficeColor GetCategoryMinorGridLineColor(OfficeChartStyle style) =>
        style.CategoryMinorGridLineColor ?? GetCategoryGridLineColor(style);

    private static OfficeColor GetValueMinorGridLineColor(OfficeChartStyle style) =>
        style.ValueMinorGridLineColor ?? GetValueGridLineColor(style);

    private static double GetCategoryGridLineWidth(OfficeChartStyle style) =>
        style.CategoryGridLineWidth ?? style.GridLineWidth ?? 0.5D;

    private static double GetValueGridLineWidth(OfficeChartStyle style) =>
        style.ValueGridLineWidth ?? style.GridLineWidth ?? 0.5D;

    private static double GetCategoryMinorGridLineWidth(OfficeChartStyle style) =>
        style.CategoryMinorGridLineWidth ?? GetCategoryGridLineWidth(style);

    private static double GetValueMinorGridLineWidth(OfficeChartStyle style) =>
        style.ValueMinorGridLineWidth ?? GetValueGridLineWidth(style);

    private static OfficeStrokeDashStyle GetCategoryGridLineDashStyle(OfficeChartStyle style) =>
        style.CategoryGridLineDashStyle ?? style.GridLineDashStyle ?? OfficeStrokeDashStyle.Solid;

    private static OfficeStrokeDashStyle GetValueGridLineDashStyle(OfficeChartStyle style) =>
        style.ValueGridLineDashStyle ?? style.GridLineDashStyle ?? OfficeStrokeDashStyle.Solid;

    private static OfficeStrokeDashStyle GetCategoryMinorGridLineDashStyle(OfficeChartStyle style) =>
        style.CategoryMinorGridLineDashStyle ?? GetCategoryGridLineDashStyle(style);

    private static OfficeStrokeDashStyle GetValueMinorGridLineDashStyle(OfficeChartStyle style) =>
        style.ValueMinorGridLineDashStyle ?? GetValueGridLineDashStyle(style);

}
