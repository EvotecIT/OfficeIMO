using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private readonly struct SecondaryAxisRenderContext {
        internal SecondaryAxisRenderContext(bool hasSeries, ValueRange range,
            IReadOnlyList<double> majorTicks, bool usesPercentDefaults, double labelBandWidth) {
            HasSeries = hasSeries;
            Range = range;
            MajorTicks = majorTicks;
            UsesPercentDefaults = usesPercentDefaults;
            LabelBandWidth = labelBandWidth;
        }

        internal bool HasSeries { get; }
        internal ValueRange Range { get; }
        internal IReadOnlyList<double> MajorTicks { get; }
        internal bool UsesPercentDefaults { get; }
        internal double LabelBandWidth { get; }
    }

    private static SecondaryAxisRenderContext CreateSecondaryAxisRenderContext(
        OfficeChartSnapshot snapshot, OfficeChartLayout layout, bool barChart, bool showLabels) {
        bool hasSeries = !barChart && snapshot.Data.Series.Any(series =>
            series.AxisGroup == OfficeChartAxisGroup.Secondary);
        if (!hasSeries) {
            return new SecondaryAxisRenderContext(false, GetCartesianValueRange(snapshot),
                Array.Empty<double>(), false, 0D);
        }

        ValueRange range = GetMixedCartesianValueRange(snapshot, OfficeChartAxisGroup.Secondary);
        bool usesPercentDefaults = snapshot.Data.Series.Any(series =>
            series.AxisGroup == OfficeChartAxisGroup.Secondary &&
            IsPercentKind(GetEffectiveSeriesKind(snapshot, series)));
        IReadOnlyList<double> majorTicks = GetValueAxisMajorTicks(range, majorUnit: null);
        double labelBandWidth = showLabels
            ? MeasureValueAxisLabelBandWidth(range, majorTicks, layout, usesPercentDefaults,
                horizontalValueAxis: false)
            : 0D;
        return new SecondaryAxisRenderContext(true, range, majorTicks, usesPercentDefaults, labelBandWidth);
    }

    private static ValueRange GetPrimaryValueAxisRange(OfficeChartSnapshot snapshot,
        OfficeChartLayout layout, bool barChart, bool hasSecondaryAxis) => hasSecondaryAxis
        ? ApplyValueAxisScale(GetMixedCartesianValueRange(snapshot, OfficeChartAxisGroup.Primary), layout,
            horizontal: false)
        : GetCartesianValueRange(snapshot, layout, horizontalValueAxis: barChart);

    private static bool IsPercentKind(OfficeChartKind kind) =>
        IsPercentStackedBarOrColumnChart(kind) || IsPercentStackedLineChart(kind) ||
        IsPercentStackedAreaChart(kind);

    private static void AddSecondaryValueAxis(OfficeDrawing drawing, SecondaryAxisRenderContext axis,
        double axisX, double plotTop, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), axisX, plotTop,
            null, GetValueAxisColor(style), GetValueAxisLineWidth(style), GetValueAxisLineDashStyle(style));
        AddVerticalValueAxisMajorTickMarks(drawing, axisX, plotTop, plotHeight, axis.Range,
            axis.MajorTicks, layout.VerticalAxisMajorTickMark, GetValueAxisColor(style),
            GetValueAxisLineWidth(style));
    }

    private static void AddSecondaryValueAxisLabels(OfficeDrawing drawing, SecondaryAxisRenderContext axis,
        double plotTop, double plotHeight, double labelLeft, double labelWidth, OfficeChartStyle style,
        OfficeChartLayout layout) {
        AddValueAxisLabels(drawing, axis.Range, plotTop, plotHeight, labelLeft, labelWidth,
            OfficeTextAlignment.Left, style, layout, axis.UsesPercentDefaults);
    }

    private static void AddMixedCartesianSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot,
        ValueRange primaryValueAxisRange, ValueRange secondaryValueAxisRange, bool hasSecondaryAxis,
        double plotLeft, double plotTop, double plotWidth, double plotHeight, OfficeChartStyle style,
        OfficeChartLayout layout) {
        AddAxisGroupSeries(drawing, snapshot, primaryValueAxisRange, OfficeChartAxisGroup.Primary,
            plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        if (hasSecondaryAxis) {
            AddAxisGroupSeries(drawing, snapshot, secondaryValueAxisRange, OfficeChartAxisGroup.Secondary,
                plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        }
        if (!HasMixedScatterSeriesOnCategoryAxes(snapshot)) {
            AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        }
    }

    private static void AddAxisGroupSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot,
        ValueRange range, OfficeChartAxisGroup axisGroup, double plotLeft, double plotTop,
        double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
        AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
        AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
    }
}
