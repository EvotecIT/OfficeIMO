using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static double GetSeriesLegendWidth(IReadOnlyList<OfficeChartSeries> series, double chartWidth, OfficeChartLayout layout) {
        if (series.Count == 0 || chartWidth < 180D) {
            return 0D;
        }

        double widest = 0D;
        for (int i = 0; i < series.Count; i++) {
            string name = string.IsNullOrWhiteSpace(series[i].Name) ? "Series " + (i + 1).ToString(CultureInfo.InvariantCulture) : series[i].Name;
            widest = Math.Max(widest, Math.Min(72D, name.Length * 4.8D));
        }

        return Math.Min(Math.Max(58D, widest + 26D), Math.Max(0D, chartWidth * layout.SeriesLegendWidthRatio));
    }

    private static double GetCategoryLegendWidth(IReadOnlyList<string> categories, double chartWidth, OfficeChartLayout layout) {
        if (categories.Count == 0 || chartWidth < 180D) {
            return 0D;
        }

        double widest = 0D;
        for (int i = 0; i < categories.Count; i++) {
            string name = string.IsNullOrWhiteSpace(categories[i]) ? "Category " + (i + 1).ToString(CultureInfo.InvariantCulture) : categories[i];
            widest = Math.Max(widest, Math.Min(78D, name.Length * 4.8D));
        }

        return Math.Min(Math.Max(62D, widest + 26D), Math.Max(0D, chartWidth * layout.CategoryLegendWidthRatio));
    }

    private static void AddSeriesLegend(OfficeDrawing drawing, IReadOnlyList<OfficeChartSeries> series, double x, double y, double width, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        if (series.Count == 0 || width < 28D) {
            return;
        }

        double rowHeight = layout.LegendRowHeight;
        double visibleRows = Math.Min(series.Count, Math.Max(1D, Math.Floor(plotHeight / rowHeight)));
        double startY = y + Math.Max(0D, (plotHeight - visibleRows * rowHeight) / 2D);
        for (int i = 0; i < series.Count && i < visibleRows; i++) {
            double rowY = startY + i * rowHeight;
            double swatchOffset = Math.Max(0D, (rowHeight - layout.LegendSwatchSize) / 2D);
            AddShape(drawing, OfficeShape.Rectangle(layout.LegendSwatchSize, layout.LegendSwatchSize), x, rowY + swatchOffset, GetSeriesColor(style, series, i), null, 0D);
            string name = string.IsNullOrWhiteSpace(series[i].Name) ? "Series " + (i + 1).ToString(CultureInfo.InvariantCulture) : series[i].Name;
            double textOffset = layout.LegendSwatchSize + layout.LegendTextGap;
            AddChartText(drawing, name, x + textOffset, rowY, width - textOffset, rowHeight, layout.LegendFontSize, style.TextColor, OfficeTextAlignment.Left, style);
        }
    }

    private static void AddCategoryLegend(OfficeDrawing drawing, IReadOnlyList<string> categories, double x, double y, double width, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout, IReadOnlyList<OfficeColor?>? pointColors = null) {
        if (categories.Count == 0 || width < 28D) {
            return;
        }

        double rowHeight = layout.LegendRowHeight;
        double visibleRows = Math.Min(categories.Count, Math.Max(1D, Math.Floor(plotHeight / rowHeight)));
        double startY = y + Math.Max(0D, (plotHeight - visibleRows * rowHeight) / 2D);
        for (int i = 0; i < categories.Count && i < visibleRows; i++) {
            double rowY = startY + i * rowHeight;
            double swatchOffset = Math.Max(0D, (rowHeight - layout.LegendSwatchSize) / 2D);
            AddShape(drawing, OfficeShape.Rectangle(layout.LegendSwatchSize, layout.LegendSwatchSize), x, rowY + swatchOffset, GetPointColor(style, pointColors, i), null, 0D);
            string name = string.IsNullOrWhiteSpace(categories[i]) ? "Category " + (i + 1).ToString(CultureInfo.InvariantCulture) : categories[i];
            double textOffset = layout.LegendSwatchSize + layout.LegendTextGap;
            AddChartText(drawing, name, x + textOffset, rowY, width - textOffset, rowHeight, layout.LegendFontSize, style.TextColor, OfficeTextAlignment.Left, style);
        }
    }

    private static void AddCategoryAxisLabels(OfficeDrawing drawing, IReadOnlyList<string> categories, double plotLeft, double plotBottomY, double plotWidth, OfficeChartStyle style, OfficeChartLayout layout) {
        if (categories.Count == 0) {
            return;
        }

        int stride = Math.Max(1, (int)Math.Ceiling(categories.Count / (double)layout.MaximumCategoryAxisLabels));
        double slot = plotWidth / categories.Count;
        if (layout.PreventLabelOverlap) {
            double minimumStep = Math.Min(layout.CategoryAxisLabelWidth, Math.Max(18D, slot * stride)) + 2D;
            stride = EnsureLabelStride(stride, slot, minimumStep);
        }

        for (int i = 0; i < categories.Count; i += stride) {
            string label = categories[i];
            if (string.IsNullOrWhiteSpace(label)) {
                continue;
            }

            double labelWidth = Math.Min(layout.CategoryAxisLabelWidth, Math.Max(18D, slot * stride));
            double centerX = plotLeft + slot * i + slot / 2D;
            AddChartText(drawing, label, centerX - labelWidth / 2D, plotBottomY + 7D, labelWidth, 11D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Center, style);
        }
    }

    private static void AddHorizontalCategoryAxisLabels(OfficeDrawing drawing, IReadOnlyList<string> categories, double plotLeft, double plotTop, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        if (categories.Count == 0) {
            return;
        }

        int stride = Math.Max(1, (int)Math.Ceiling(categories.Count / (double)layout.MaximumHorizontalCategoryAxisLabels));
        double slot = plotHeight / categories.Count;
        if (layout.PreventLabelOverlap) {
            stride = EnsureLabelStride(stride, slot, 10D + 2D);
        }

        double labelWidth = Math.Max(12D, plotLeft - 6D);
        for (int i = 0; i < categories.Count; i += stride) {
            string label = categories[i];
            if (string.IsNullOrWhiteSpace(label)) {
                continue;
            }

            int categorySlot = categories.Count - 1 - i;
            double centerY = plotTop + slot * categorySlot + slot / 2D;
            AddChartText(drawing, label, 2D, centerY - 5D, labelWidth, 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
        }
    }

    private static void AddValueAxisLabels(OfficeDrawing drawing, ValueRange range, double plotLeft, double plotTop, double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        AddChartText(drawing, FormatAxisValue(range.Max), 2D, plotTop - 5D, Math.Max(12D, plotLeft - 6D), 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
        AddChartText(drawing, FormatAxisValue(range.Min), 2D, plotTop + plotHeight - 5D, Math.Max(12D, plotLeft - 6D), 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
    }

    private static void AddHorizontalValueAxisLabels(OfficeDrawing drawing, ValueRange range, double plotLeft, double plotBottomY, double plotWidth, OfficeChartStyle style, OfficeChartLayout layout) {
        AddChartText(drawing, FormatAxisValue(range.Min), plotLeft - 12D, plotBottomY + 4D, 28D, 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Left, style);
        AddChartText(drawing, FormatAxisValue(range.Max), plotLeft + plotWidth - 28D, plotBottomY + 4D, 34D, 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
    }

    private static void AddRadarCategoryLabels(OfficeDrawing drawing, IReadOnlyList<string> categories, double centerX, double centerY, double radius, OfficeChartStyle style, OfficeChartLayout layout) {
        if (categories.Count == 0) {
            return;
        }

        int stride = Math.Max(1, (int)Math.Ceiling(categories.Count / (double)layout.MaximumRadarCategoryLabels));
        if (layout.PreventLabelOverlap) {
            double circumferenceStep = Math.PI * 2D * Math.Max(1D, radius + 13D) / categories.Count;
            stride = EnsureLabelStride(stride, circumferenceStep, layout.RadarCategoryLabelWidth * 0.7D);
        }

        for (int i = 0; i < categories.Count; i += stride) {
            string label = categories[i];
            if (string.IsNullOrWhiteSpace(label)) {
                continue;
            }

            double angle = -Math.PI / 2D + Math.PI * 2D * i / categories.Count;
            double labelWidth = layout.RadarCategoryLabelWidth;
            double labelHeight = 10D;
            double x = centerX + Math.Cos(angle) * (radius + 13D) - labelWidth / 2D;
            double y = centerY + Math.Sin(angle) * (radius + 13D) - labelHeight / 2D;
            OfficeTextAlignment alignment = Math.Cos(angle) < -0.25D
                ? OfficeTextAlignment.Right
                : Math.Cos(angle) > 0.25D
                    ? OfficeTextAlignment.Left
                    : OfficeTextAlignment.Center;
            AddChartText(drawing, label, x, y, labelWidth, labelHeight, layout.AxisLabelFontSize, style.MutedTextColor, alignment, style);
        }
    }

    private static ValueRange GetCartesianValueRange(OfficeChartSnapshot snapshot) {
        if (IsScatterChart(snapshot.ChartKind)) {
            return GetFiniteSeriesRange(snapshot.Data.Series);
        }

        if (IsPercentStackedLineChart(snapshot.ChartKind) || IsPercentStackedAreaChart(snapshot.ChartKind) || IsPercentStackedBarOrColumnChart(snapshot.ChartKind)) {
            return GetPercentStackedSeriesRange(snapshot.Data.Series, snapshot.Data.Categories.Count);
        }

        if (IsStackedLineChart(snapshot.ChartKind) || IsStackedAreaChart(snapshot.ChartKind) || IsStackedBarOrColumnChart(snapshot.ChartKind)) {
            return GetStackedSeriesRange(snapshot.Data.Series, snapshot.Data.Categories.Count);
        }

        ValueRange range = GetFiniteSeriesRange(snapshot.Data.Series);
        return ExpandFlatRange(Math.Min(0D, range.Min), Math.Max(0D, range.Max));
    }

    private static string FormatAxisValue(double value) {
        double abs = Math.Abs(value);
        if (abs >= 1000D) {
            return (value / 1000D).ToString("0.#", CultureInfo.InvariantCulture) + "k";
        }

        if (abs > 0D && abs < 1D) {
            return value.ToString("0.#%", CultureInfo.InvariantCulture);
        }

        return value.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static int EnsureLabelStride(int stride, double unitStep, double minimumStep) {
        int safeStride = Math.Max(1, stride);
        while (unitStep * safeStride < minimumStep) {
            safeStride++;
        }

        return safeStride;
    }

    private static void AddChartText(OfficeDrawing drawing, string text, double x, double y, double width, double height, double fontSize, OfficeColor color, OfficeTextAlignment alignment, OfficeChartStyle style) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        double safeX = Math.Max(0D, x);
        double safeY = Math.Max(0D, y);
        double safeWidth = Math.Min(width - (safeX - x), drawing.Width - safeX);
        double safeHeight = Math.Min(height - (safeY - y), drawing.Height - safeY);
        if (safeWidth <= 1D || safeHeight <= 1D) {
            return;
        }

        drawing.AddText(
            text,
            safeX,
            safeY,
            safeWidth,
            safeHeight,
            new OfficeFontInfo(style.FontFamily, fontSize),
            color,
            alignment,
            Math.Max(fontSize + 1D, safeHeight));
    }
}
