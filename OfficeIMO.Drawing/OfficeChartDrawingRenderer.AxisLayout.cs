using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const double MinimumValueAxisLabelWidth = 34D;
    private const double MaximumValueAxisLabelWidth = 92D;
    private const double AxisLabelHorizontalPadding = 6D;
    private const int MaximumMeasuredCategoryLabelCharacters = 512;

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
