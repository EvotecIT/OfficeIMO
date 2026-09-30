using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const double MinimumValueAxisLabelWidth = 34D;
    private const double MaximumValueAxisLabelWidth = 92D;
    private const double AxisLabelHorizontalPadding = 6D;
    private const int MaximumMeasuredCategoryLabelCharacters = 512;

    private static void AddCategoryGridLines(OfficeDrawing drawing, OfficeChartSnapshot snapshot,
        OfficeChartLayout layout, bool horizontal, double plotLeft, double plotTop,
        double plotWidth, double plotHeight, bool minor, OfficeColor color,
        double lineWidth, OfficeStrokeDashStyle dashStyle) {
        if (IsScatterChart(snapshot.ChartKind)) {
            IReadOnlyList<double> xValues = GetScatterXValues(snapshot.Data.Categories);
            List<OfficeChartSeries> series = GetRenderableScatterSeries(snapshot).Select(item => item.Series).ToList();
            ValueRange range = ApplyValueAxisScale(GetScatterPointRanges(series, xValues).XRange, layout, horizontal: true);
            IReadOnlyList<double> majorTicks = GetValueAxisMajorTicks(range, GetValueAxisMajorUnit(layout, horizontal: true));
            IReadOnlyList<double> ticks = minor
                ? GetValueAxisMinorTicks(range, GetValueAxisMinorUnit(layout, horizontal: true), majorTicks)
                : majorTicks;
            AddVerticalValueGridLines(drawing, plotLeft, plotTop, plotWidth, plotHeight,
                range, ticks, color, lineWidth, dashStyle);
            return;
        }

        int count = snapshot.Data.Categories.Count;
        if (count == 0) return;
        bool pointsAtEdges = IsLineChart(snapshot.ChartKind) || IsAreaChart(snapshot.ChartKind);
        int intervals = pointsAtEdges ? count - 1 : count;
        if (intervals <= 0) return;
        int first = minor ? (pointsAtEdges ? 0 : 1) : (pointsAtEdges ? 1 : 0);
        int exclusiveEnd = minor ? intervals : (pointsAtEdges ? intervals : count);
        for (int index = first; index < exclusiveEnd; index++) {
            double fraction = pointsAtEdges
                ? (index + (minor ? 0.5D : 0D)) / intervals
                : (index + (minor ? 0D : 0.5D)) / intervals;
            if (horizontal) {
                double y = plotTop + plotHeight * fraction;
                AddShape(drawing, OfficeShape.Line(0D, 0D, plotWidth, 0D),
                    plotLeft, y, null, color, lineWidth, dashStyle);
            } else {
                double x = plotLeft + plotWidth * fraction;
                AddShape(drawing, OfficeShape.Line(0D, 0D, 0D, plotHeight),
                    x, plotTop, null, color, lineWidth, dashStyle);
            }
        }
    }

    private static double GetVerticalAxisLabelBandWidth(
        OfficeChartSnapshot snapshot,
        ValueRange valueRange,
        IReadOnlyList<double> valueTicks,
        OfficeChartLayout layout,
        bool percentDefault,
        bool horizontalValueAxis) {
        if (horizontalValueAxis) {
            return MeasureCategoryAxisLabelBandWidth(snapshot.Data.Categories, layout);
        }

        return MeasureValueAxisLabelBandWidth(valueRange, valueTicks, layout, percentDefault, horizontalValueAxis: false);
    }

    private static double GetHorizontalValueAxisLabelWidth(
        ValueRange valueRange,
        IReadOnlyList<double> valueTicks,
        OfficeChartLayout layout,
        bool percentDefault) =>
        MeasureValueAxisLabelBandWidth(valueRange, valueTicks, layout, percentDefault, horizontalValueAxis: true);

    private static double MeasureValueAxisLabelBandWidth(
        ValueRange valueRange,
        IReadOnlyList<double> valueTicks,
        OfficeChartLayout layout,
        bool percentDefault,
        bool horizontalValueAxis) {
        string? numberFormat = horizontalValueAxis ? layout.HorizontalAxisNumberFormat : layout.VerticalAxisNumberFormat;
        double? displayUnitDivisor = horizontalValueAxis ? layout.HorizontalAxisDisplayUnitDivisor : layout.VerticalAxisDisplayUnitDivisor;
        double widest = 0D;
        for (int i = 0; i < valueTicks.Count; i++) {
            string label = FormatAxisValue(valueTicks[i], layout, percentDefault, numberFormat, displayUnitDivisor);
            widest = Math.Max(widest, MeasureAxisLabelWidth(label, layout));
        }

        if (valueTicks.Count == 0) {
            widest = Math.Max(
                MeasureAxisLabelWidth(FormatAxisValue(valueRange.Min, layout, percentDefault, numberFormat, displayUnitDivisor), layout),
                MeasureAxisLabelWidth(FormatAxisValue(valueRange.Max, layout, percentDefault, numberFormat, displayUnitDivisor), layout));
        }

        return ClampAxisLabelWidth(widest + AxisLabelHorizontalPadding);
    }

    private static double MeasureCategoryAxisLabelBandWidth(IReadOnlyList<string> categories, OfficeChartLayout layout) {
        double widest = 0D;
        int stride = Math.Max(1, (int)Math.Ceiling(categories.Count / (double)layout.MaximumHorizontalCategoryAxisLabels));
        for (int i = 0; i < categories.Count; i += stride) {
            string label = categories[i] ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(label)) {
                if (label.Length > MaximumMeasuredCategoryLabelCharacters) {
                    label = label.Substring(0, MaximumMeasuredCategoryLabelCharacters);
                }

                widest = Math.Max(widest, MeasureAxisLabelWidth(label, layout));
            }
        }

        return ClampAxisLabelWidth(widest + AxisLabelHorizontalPadding);
    }

    private static double MeasureAxisLabelWidth(string? label, OfficeChartLayout layout) {
        if (string.IsNullOrEmpty(label)) {
            return 0D;
        }

        var fontInfo = new OfficeFontInfo(
            string.IsNullOrWhiteSpace(layout.AxisTextFontFamily) ? OfficeFontInfo.Default.FamilyName : layout.AxisTextFontFamily!,
            layout.AxisLabelFontSize,
            layout.AxisTextFontStyle ?? OfficeFontStyle.Regular);
        OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(fontInfo);
        double measuredPixels = measurer.MeasureWidth(label, measurer.CreateStyle(fontInfo));
        return measuredPixels * OfficeTextMeasurer.PointsPerInch / OfficeTextMeasurer.DefaultDpi;
    }

    private static double ClampAxisLabelWidth(double width) =>
        Math.Min(MaximumValueAxisLabelWidth, Math.Max(MinimumValueAxisLabelWidth, width));
}
