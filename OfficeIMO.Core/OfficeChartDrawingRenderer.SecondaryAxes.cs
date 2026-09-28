using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static void AddSecondaryAxisTitle(OfficeDrawing drawing, string title,
        double plotLeft, double plotTop, double plotWidth, bool barChart,
        OfficeChartStyle style, OfficeChartLayout layout) {
        double fontSize = GetAxisTitleFontSize(layout);
        double height = Math.Max(10D, fontSize + 2D);
        double y = Math.Max(0D, plotTop - height - (barChart ? 17D : 4D));
        AddChartText(drawing, title, plotLeft + plotWidth / 2D, y, plotWidth / 2D,
            height, fontSize, style.AxisTitleColor ?? style.MutedTextColor,
            OfficeTextAlignment.Right, style, layout.AxisTitleFontFamily ?? layout.AxisTextFontFamily,
            layout.AxisTitleFontStyle ?? layout.AxisTextFontStyle);
    }

    private readonly struct SecondaryAxisRenderContext {
        internal SecondaryAxisRenderContext(bool hasSeries, ValueRange range,
            IReadOnlyList<double> majorTicks, bool usesPercentDefaults, double labelBandWidth, OfficeChartLayout axisLayout) {
            HasSeries = hasSeries;
            Range = range;
            MajorTicks = majorTicks;
            UsesPercentDefaults = usesPercentDefaults;
            LabelBandWidth = labelBandWidth;
            Layout = axisLayout;
        }

        internal bool HasSeries { get; }
        internal ValueRange Range { get; }
        internal IReadOnlyList<double> MajorTicks { get; }
        internal IReadOnlyList<double> MinorTicks => GetValueAxisMinorTicks(Range,
            Layout.VerticalAxisMinorUnit, MajorTicks,
            Layout.VerticalAxisMinorTickMark != OfficeChartAxisTickMark.None ||
            Layout.HorizontalAxisMinorTickMark != OfficeChartAxisTickMark.None);
        internal bool UsesPercentDefaults { get; }
        internal double LabelBandWidth { get; }
        internal OfficeChartLayout Layout { get; }
    }

    private static SecondaryAxisRenderContext CreateSecondaryAxisRenderContext(
        OfficeChartSnapshot snapshot, OfficeChartLayout layout, bool barChart, bool showLabels) {
        bool hasSeries = snapshot.Data.Series.Any(series =>
            series.AxisGroup == OfficeChartAxisGroup.Secondary);
        if (!hasSeries) {
            return new SecondaryAxisRenderContext(false, GetCartesianValueRange(snapshot),
                Array.Empty<double>(), false, 0D, layout);
        }

        OfficeChartValueAxisLayout? axis = layout.SecondaryValueAxis;
        var axisLayout = new OfficeChartLayout(axisNumberFormat: axis?.NumberFormat ?? "General",
            horizontalAxisMinimum: axis?.Minimum, horizontalAxisMaximum: axis?.Maximum,
            verticalAxisMinimum: axis?.Minimum, verticalAxisMaximum: axis?.Maximum,
            horizontalAxisMajorUnit: axis?.MajorUnit, verticalAxisMajorUnit: axis?.MajorUnit,
            horizontalAxisMinorUnit: axis?.MinorUnit, verticalAxisMinorUnit: axis?.MinorUnit,
            axisLabelFontSize: layout.AxisLabelFontSize, axisTextFontFamily: layout.AxisTextFontFamily,
            axisTextFontStyle: layout.AxisTextFontStyle,
            horizontalAxisMajorTickMark: axis?.MajorTickMark ?? layout.HorizontalAxisMajorTickMark,
            verticalAxisMajorTickMark: axis?.MajorTickMark ?? layout.VerticalAxisMajorTickMark,
            horizontalAxisMinorTickMark: axis?.MinorTickMark ?? layout.HorizontalAxisMinorTickMark,
            verticalAxisMinorTickMark: axis?.MinorTickMark ?? layout.VerticalAxisMinorTickMark);
        ValueRange range = ApplyValueAxisScale(
            GetMixedCartesianValueRange(snapshot, OfficeChartAxisGroup.Secondary), axisLayout,
            horizontal: barChart);
        bool usesPercentDefaults = snapshot.Data.Series.Any(series =>
            series.AxisGroup == OfficeChartAxisGroup.Secondary &&
            IsPercentKind(GetEffectiveSeriesKind(snapshot, series)));
        IReadOnlyList<double> majorTicks = GetValueAxisMajorTicks(range, axis?.MajorUnit);
        double labelBandWidth = showLabels
            ? MeasureValueAxisLabelBandWidth(range, majorTicks, axisLayout, usesPercentDefaults,
                horizontalValueAxis: barChart)
            : 0D;
        return new SecondaryAxisRenderContext(true, range, majorTicks, usesPercentDefaults, labelBandWidth, axisLayout);
    }

    private static ValueRange GetPrimaryValueAxisRange(OfficeChartSnapshot snapshot,
        OfficeChartLayout layout, bool barChart, bool hasSecondaryAxis) => hasSecondaryAxis
        ? ApplyValueAxisScale(GetMixedCartesianValueRange(snapshot, OfficeChartAxisGroup.Primary), layout,
            horizontal: barChart)
        : GetCartesianValueRange(snapshot, layout, horizontalValueAxis: barChart);

    private static bool IsPercentKind(OfficeChartKind kind) =>
        IsPercentStackedBarOrColumnChart(kind) || IsPercentStackedLineChart(kind) ||
        IsPercentStackedAreaChart(kind);

    private static void AddSecondaryValueAxis(OfficeDrawing drawing, SecondaryAxisRenderContext axis,
        double axisX, double plotTop, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight), axisX, plotTop,
            null, GetValueAxisColor(style), GetValueAxisLineWidth(style), GetValueAxisLineDashStyle(style));
        AddVerticalValueAxisMajorTickMarks(drawing, axisX, plotTop, plotHeight, axis.Range,
            axis.MajorTicks, axis.Layout.VerticalAxisMajorTickMark, GetValueAxisColor(style),
            GetValueAxisLineWidth(style), positiveOutside: true);
        AddVerticalValueAxisMinorTickMarks(drawing, axisX, plotTop, plotHeight, axis.Range,
            axis.MinorTicks, axis.Layout.VerticalAxisMinorTickMark, GetValueAxisColor(style),
            GetValueAxisLineWidth(style), positiveOutside: true);
    }

    private static void AddSecondaryValueAxisLabels(OfficeDrawing drawing, SecondaryAxisRenderContext axis,
        double plotTop, double plotHeight, double labelLeft, double labelWidth, OfficeChartStyle style,
        OfficeChartLayout layout) {
        AddValueAxisLabels(drawing, axis.Range, plotTop, plotHeight, labelLeft, labelWidth,
            OfficeTextAlignment.Left, style, axis.Layout, axis.UsesPercentDefaults);
    }

    private static void AddHorizontalSecondaryValueAxis(OfficeDrawing drawing,
        SecondaryAxisRenderContext axis, double plotLeft, double axisY, double plotWidth,
        OfficeChartStyle style, OfficeChartLayout layout) {
        AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D), plotLeft, axisY,
            null, GetValueAxisColor(style), GetValueAxisLineWidth(style), GetValueAxisLineDashStyle(style));
        AddHorizontalValueAxisMajorTickMarks(drawing, plotLeft, axisY, plotWidth, axis.Range,
            axis.MajorTicks, axis.Layout.HorizontalAxisMajorTickMark, GetValueAxisColor(style),
            GetValueAxisLineWidth(style), positiveOutside: false);
        AddHorizontalValueAxisMinorTickMarks(drawing, plotLeft, axisY, plotWidth, axis.Range,
            axis.MinorTicks, axis.Layout.HorizontalAxisMinorTickMark, GetValueAxisColor(style),
            GetValueAxisLineWidth(style), positiveOutside: false);
    }

    private static void AddHorizontalSecondaryValueAxisLabels(OfficeDrawing drawing,
        SecondaryAxisRenderContext axis, double plotLeft, double plotTop, double plotWidth,
        OfficeChartStyle style, OfficeChartLayout layout) {
        AddHorizontalValueAxisLabels(drawing, axis.Range, plotLeft, plotTop - 13D, plotWidth,
            Math.Max(12D, axis.LabelBandWidth), labelsAbovePlot: true, style, axis.Layout,
            axis.UsesPercentDefaults);
    }

    private static void AddMixedCartesianSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot,
        ValueRange primaryValueAxisRange, ValueRange secondaryValueAxisRange, bool hasSecondaryAxis,
        double plotLeft, double plotTop, double plotWidth, double plotHeight,
        double numericPlotLeft, double numericPlotTop, double numericPlotWidth,
        double numericPlotHeight, OfficeChartStyle style, OfficeChartLayout layout,
        double maximumBubbleDiameter) {
        AddAxisGroupSeries(drawing, snapshot, primaryValueAxisRange, OfficeChartAxisGroup.Primary,
            plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        if (hasSecondaryAxis) {
            AddAxisGroupSeries(drawing, snapshot, secondaryValueAxisRange, OfficeChartAxisGroup.Secondary,
                plotLeft, plotTop, plotWidth, plotHeight, style, layout);
        }
        if (!HasMixedScatterSeriesOnCategoryAxes(snapshot)) {
            AddScatterSeries(drawing, snapshot, numericPlotLeft, numericPlotTop,
                numericPlotWidth, numericPlotHeight, style, layout,
                primaryValueAxisRange, OfficeChartAxisGroup.Primary,
                maximumBubbleDiameter);
            if (hasSecondaryAxis) {
                AddScatterSeries(drawing, snapshot, numericPlotLeft, numericPlotTop,
                    numericPlotWidth, numericPlotHeight, style, layout,
                    secondaryValueAxisRange, OfficeChartAxisGroup.Secondary,
                    maximumBubbleDiameter);
            }
        }
    }

    private static void AddAxisGroupSeries(OfficeDrawing drawing, OfficeChartSnapshot snapshot,
        ValueRange range, OfficeChartAxisGroup axisGroup, double plotLeft, double plotTop,
        double plotWidth, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        if (!HasMixedCartesianSeriesKinds(snapshot)) {
            if (IsAreaChart(snapshot.ChartKind)) {
                AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
                    range, axisGroup);
            } else if (IsBarOrColumnChart(snapshot.ChartKind)) {
                AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
                    range, axisGroup);
            } else if (IsLineChart(snapshot.ChartKind)) {
                AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
                    range, axisGroup);
            }
            return;
        }

        AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
        AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
        AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight, style, layout,
            range, axisGroup);
    }
}
