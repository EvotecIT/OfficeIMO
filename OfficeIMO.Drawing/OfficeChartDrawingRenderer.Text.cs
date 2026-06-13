using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private static double GetSeriesLegendWidth(IReadOnlyList<OfficeChartSeries> series, double chartWidth, OfficeChartLayout layout) {
        if (!ShouldRenderLegendSide(layout) || series.Count == 0 || chartWidth < 180D) {
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
        if (!ShouldRenderLegendSide(layout) || categories.Count == 0 || chartWidth < 180D) {
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
        if (!layout.ShowLegend || series.Count == 0 || width < 28D) {
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
        if (!layout.ShowLegend || categories.Count == 0 || width < 28D) {
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

    private static double GetSeriesLegendBandHeight(IReadOnlyList<OfficeChartSeries> series, double chartWidth, OfficeChartLayout layout) {
        if (!ShouldRenderLegendBand(layout) || series.Count == 0 || chartWidth < 160D) {
            return 0D;
        }

        return GetLegendBandHeight(series.Select(item => item.Name), chartWidth, layout);
    }

    private static double GetCategoryLegendBandHeight(IReadOnlyList<string> categories, double chartWidth, OfficeChartLayout layout) {
        if (!ShouldRenderLegendBand(layout) || categories.Count == 0 || chartWidth < 160D) {
            return 0D;
        }

        return GetLegendBandHeight(categories, chartWidth, layout);
    }

    private static void AddSeriesLegendBand(OfficeDrawing drawing, IReadOnlyList<OfficeChartSeries> series, double x, double y, double width, OfficeChartStyle style, OfficeChartLayout layout) {
        if (!ShouldRenderLegendBand(layout) || series.Count == 0 || width < 48D) {
            return;
        }

        AddLegendBand(
            drawing,
            series.Select(item => string.IsNullOrWhiteSpace(item.Name) ? null : item.Name),
            x,
            y,
            width,
            style,
            layout);
    }

    private static void AddCategoryLegendBand(OfficeDrawing drawing, IReadOnlyList<string> categories, double x, double y, double width, OfficeChartStyle style, OfficeChartLayout layout, IReadOnlyList<OfficeColor?>? pointColors = null) {
        if (!ShouldRenderLegendBand(layout) || categories.Count == 0 || width < 48D) {
            return;
        }

        AddLegendBand(drawing, categories, x, y, width, style, layout, pointColors);
    }

    private static bool ShouldRenderLegendBand(OfficeChartLayout layout) =>
        layout.ShowLegend &&
        (layout.LegendPosition == OfficeChartLegendPosition.Top || layout.LegendPosition == OfficeChartLegendPosition.Bottom);

    private static bool ShouldRenderLegendSide(OfficeChartLayout layout) =>
        layout.ShowLegend &&
        (layout.LegendPosition == OfficeChartLegendPosition.Left || layout.LegendPosition == OfficeChartLegendPosition.Right);

    private static double GetLegendBandHeight(IEnumerable<string?> labels, double chartWidth, OfficeChartLayout layout) {
        int count = labels.Count();
        if (count == 0) {
            return 0D;
        }

        int columns = GetLegendBandColumns(labels, chartWidth, layout);
        int rows = Math.Min(2, (int)Math.Ceiling(count / (double)Math.Max(1, columns)));
        return rows * layout.LegendRowHeight + 4D;
    }

    private static void AddLegendBand(OfficeDrawing drawing, IEnumerable<string?> labels, double x, double y, double width, OfficeChartStyle style, OfficeChartLayout layout, IReadOnlyList<OfficeColor?>? pointColors = null) {
        List<string?> labelList = labels.ToList();
        if (labelList.Count == 0) {
            return;
        }

        int columns = GetLegendBandColumns(labelList, width, layout);
        int visibleCount = Math.Min(labelList.Count, columns * 2);
        double itemWidth = Math.Max(44D, width / Math.Max(1, columns));
        double rowHeight = layout.LegendRowHeight;
        double swatchOffset = Math.Max(0D, (rowHeight - layout.LegendSwatchSize) / 2D);
        for (int i = 0; i < visibleCount; i++) {
            int row = i / columns;
            int column = i % columns;
            string name = string.IsNullOrWhiteSpace(labelList[i])
                ? "Series " + (i + 1).ToString(CultureInfo.InvariantCulture)
                : labelList[i]!;
            double itemX = x + column * itemWidth;
            double rowY = y + row * rowHeight;
            AddShape(drawing, OfficeShape.Rectangle(layout.LegendSwatchSize, layout.LegendSwatchSize), itemX, rowY + swatchOffset, GetPointColor(style, pointColors, i), null, 0D);
            double textOffset = layout.LegendSwatchSize + layout.LegendTextGap;
            AddChartText(drawing, name, itemX + textOffset, rowY, Math.Max(1D, itemWidth - textOffset - 2D), rowHeight, layout.LegendFontSize, style.TextColor, OfficeTextAlignment.Left, style);
        }
    }

    private static int GetLegendBandColumns(IEnumerable<string?> labels, double width, OfficeChartLayout layout) {
        double widest = 0D;
        int count = 0;
        foreach (string? label in labels) {
            count++;
            string name = string.IsNullOrWhiteSpace(label) ? "Series " + count.ToString(CultureInfo.InvariantCulture) : label!;
            widest = Math.Max(widest, Math.Min(92D, name.Length * 4.8D));
        }

        double itemWidth = Math.Max(48D, widest + layout.LegendSwatchSize + layout.LegendTextGap + 8D);
        return Math.Max(1, Math.Min(count, (int)Math.Floor(width / itemWidth)));
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
        AddChartText(drawing, FormatAxisValue(range.Max, layout), 2D, plotTop - 5D, Math.Max(12D, plotLeft - 6D), 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
        AddChartText(drawing, FormatAxisValue(range.Min, layout), 2D, plotTop + plotHeight - 5D, Math.Max(12D, plotLeft - 6D), 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
    }

    private static void AddHorizontalValueAxisLabels(OfficeDrawing drawing, ValueRange range, double plotLeft, double plotBottomY, double plotWidth, OfficeChartStyle style, OfficeChartLayout layout) {
        AddChartText(drawing, FormatAxisValue(range.Min, layout), plotLeft - 12D, plotBottomY + 4D, 28D, 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Left, style);
        AddChartText(drawing, FormatAxisValue(range.Max, layout), plotLeft + plotWidth - 28D, plotBottomY + 4D, 34D, 10D, layout.AxisLabelFontSize, style.MutedTextColor, OfficeTextAlignment.Right, style);
    }

    private static bool HasHorizontalAxisTitle(OfficeChartKind chartKind, OfficeChartLayout layout) =>
        !string.IsNullOrWhiteSpace(IsBarChart(chartKind) ? layout.ValueAxisTitle : layout.CategoryAxisTitle);

    private static bool HasVerticalAxisTitle(OfficeChartKind chartKind, OfficeChartLayout layout) =>
        !string.IsNullOrWhiteSpace(IsBarChart(chartKind) ? layout.CategoryAxisTitle : layout.ValueAxisTitle);

    private static void AddAxisTitles(
        OfficeDrawing drawing,
        string? verticalTitle,
        string? horizontalTitle,
        double plotLeft,
        double plotTop,
        double plotBottomY,
        double plotWidth,
        double plotHeight,
        OfficeChartStyle style,
        OfficeChartLayout layout) {
        double titleFontSize = Math.Min(8.5D, Math.Max(layout.AxisLabelFontSize + 0.7D, layout.AxisLabelFontSize));
        if (!string.IsNullOrWhiteSpace(verticalTitle)) {
            double titleY = Math.Max(0D, plotTop - 14D);
            AddChartText(drawing, verticalTitle!, plotLeft, titleY, plotWidth, 10D, titleFontSize, style.MutedTextColor, OfficeTextAlignment.Left, style);
        }

        if (!string.IsNullOrWhiteSpace(horizontalTitle)) {
            double titleY = plotBottomY + 20D;
            if (titleY + 10D > drawing.Height) {
                titleY = Math.Max(0D, drawing.Height - 10D);
            }

            AddChartText(drawing, horizontalTitle!, 8D, titleY, Math.Max(1D, drawing.Width - 16D), 10D, titleFontSize, style.MutedTextColor, OfficeTextAlignment.Center, style);
        }
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

    private static string FormatAxisValue(double value, OfficeChartLayout layout) {
        if (TryFormatDataLabelValue(value, layout.AxisNumberFormat, out string? formatted)) {
            return formatted!;
        }

        double abs = Math.Abs(value);
        if (abs >= 1000D) {
            return (value / 1000D).ToString("0.#", CultureInfo.InvariantCulture) + "k";
        }

        if (abs > 0D && abs < 1D) {
            return value.ToString("0.#%", CultureInfo.InvariantCulture);
        }

        return value.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static string FormatDataLabel(OfficeChartLayout layout, string category, OfficeChartSeries series, double value, double total) {
        var parts = new List<string>(4);
        if (layout.ShowDataLabelSeriesNames && !string.IsNullOrWhiteSpace(series.Name)) {
            parts.Add(series.Name);
        }

        if (layout.ShowDataLabelCategoryNames && !string.IsNullOrWhiteSpace(category)) {
            parts.Add(category);
        }

        if (layout.ShowDataLabelValues) {
            parts.Add(FormatDataLabelValue(value, layout.DataLabelNumberFormat));
        }

        if (layout.ShowDataLabelPercentages) {
            double ratio = total > 0D && !double.IsNaN(value) && !double.IsInfinity(value)
                ? Math.Max(0D, value) / total
                : 0D;
            parts.Add(FormatDataLabelPercent(ratio));
        }

        return string.Join(layout.DataLabelSeparator, parts);
    }

    private static string FormatDataLabelValue(double value, string? numberFormat) {
        if (TryFormatDataLabelValue(value, numberFormat, out string? formatted)) {
            return formatted!;
        }

        double rounded = Math.Round(value);
        if (Math.Abs(value - rounded) < 0.0000001D) {
            return rounded.ToString("0", CultureInfo.InvariantCulture);
        }

        return value.ToString("0.##", CultureInfo.InvariantCulture);
    }

    private static bool TryFormatDataLabelValue(double value, string? numberFormat, out string? formatted) {
        formatted = null;
        if (string.IsNullOrWhiteSpace(numberFormat) ||
            string.Equals(numberFormat, "General", StringComparison.OrdinalIgnoreCase) ||
            double.IsNaN(value) ||
            double.IsInfinity(value)) {
            return false;
        }

        string format = numberFormat!.Trim();
        int sectionIndex = format.IndexOf(';');
        if (sectionIndex >= 0) {
            format = format.Substring(0, sectionIndex).Trim();
        }

        if (format.Length == 0) {
            return false;
        }

        bool percent = format.IndexOf('%') >= 0;
        bool grouped = HasDataLabelGrouping(format);
        int decimals = GetDataLabelDecimalPlaces(format);
        double displayValue = percent ? value * 100D : value;
        string numericFormat = (grouped ? "N" : "F") + decimals.ToString(CultureInfo.InvariantCulture);
        formatted = displayValue.ToString(numericFormat, CultureInfo.InvariantCulture);
        if (percent) {
            formatted += "%";
        }

        return true;
    }

    private static bool HasDataLabelGrouping(string format) {
        int decimalIndex = format.IndexOf('.');
        int searchLength = decimalIndex >= 0 ? decimalIndex : format.Length;
        return format.Substring(0, searchLength).IndexOf(',') >= 0;
    }

    private static int GetDataLabelDecimalPlaces(string format) {
        int decimalIndex = format.IndexOf('.');
        if (decimalIndex < 0) {
            return 0;
        }

        int count = 0;
        for (int i = decimalIndex + 1; i < format.Length; i++) {
            char c = format[i];
            if (c == '0' || c == '#') {
                count++;
                continue;
            }

            break;
        }

        return Math.Min(6, count);
    }

    private static string FormatDataLabelPercent(double ratio) =>
        ratio.ToString("0.#%", CultureInfo.InvariantCulture);

    private static void AddVerticalDataLabel(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        string category,
        OfficeChartSeries series,
        double value,
        double total,
        double centerX,
        double barTop,
        double barBottom) {
        if (!layout.ShowDataLabels) {
            return;
        }

        string label = FormatDataLabel(layout, category, series, value, total);
        if (string.IsNullOrWhiteSpace(label)) {
            return;
        }

        (double labelWidth, double labelHeight) = GetDataLabelSize(label, layout);
        double barCenterY = (barTop + barBottom) / 2D;
        double x = centerX - labelWidth / 2D;
        double y = layout.DataLabelPosition switch {
            OfficeChartDataLabelPosition.Center => barCenterY - labelHeight / 2D,
            OfficeChartDataLabelPosition.InsideEnd => value >= 0D ? barTop + 1D : barBottom - labelHeight - 1D,
            OfficeChartDataLabelPosition.InsideBase => value >= 0D ? barBottom - labelHeight - 1D : barTop + 1D,
            OfficeChartDataLabelPosition.Bottom => barBottom + 1D,
            OfficeChartDataLabelPosition.Left or OfficeChartDataLabelPosition.Right => barCenterY - labelHeight / 2D,
            _ => value >= 0D ? barTop - labelHeight - 1D : barBottom + 1D
        };
        if (layout.DataLabelPosition == OfficeChartDataLabelPosition.Left) {
            x = centerX - labelWidth - 4D;
        } else if (layout.DataLabelPosition == OfficeChartDataLabelPosition.Right) {
            x = centerX + 4D;
        }

        AddChartText(drawing, label, FitDataLabelX(drawing, x, labelWidth), FitDataLabelY(drawing, y, labelHeight), labelWidth, labelHeight, layout.DataLabelFontSize, style.TextColor, OfficeTextAlignment.Center, style);
    }

    private static void AddHorizontalDataLabel(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        string category,
        OfficeChartSeries series,
        double value,
        double total,
        double barLeft,
        double barRight,
        double barTop,
        double barBottom) {
        if (!layout.ShowDataLabels) {
            return;
        }

        string label = FormatDataLabel(layout, category, series, value, total);
        if (string.IsNullOrWhiteSpace(label)) {
            return;
        }

        (double labelWidth, double labelHeight) = GetDataLabelSize(label, layout);
        double centerX = (barLeft + barRight) / 2D;
        double centerY = (barTop + barBottom) / 2D;
        double x = layout.DataLabelPosition switch {
            OfficeChartDataLabelPosition.Center => centerX - labelWidth / 2D,
            OfficeChartDataLabelPosition.InsideEnd => value >= 0D ? barRight - labelWidth - 2D : barLeft + 2D,
            OfficeChartDataLabelPosition.InsideBase => value >= 0D ? barLeft + 2D : barRight - labelWidth - 2D,
            OfficeChartDataLabelPosition.Left => barLeft - labelWidth - 2D,
            _ => value >= 0D ? barRight + 2D : barLeft - labelWidth - 2D
        };
        double y = layout.DataLabelPosition switch {
            OfficeChartDataLabelPosition.Top => barTop - labelHeight - 1D,
            OfficeChartDataLabelPosition.Bottom => barBottom + 1D,
            _ => centerY - labelHeight / 2D
        };
        OfficeTextAlignment alignment = layout.DataLabelPosition == OfficeChartDataLabelPosition.Center
            ? OfficeTextAlignment.Center
            : value >= 0D ? OfficeTextAlignment.Left : OfficeTextAlignment.Right;
        AddChartText(drawing, label, FitDataLabelX(drawing, x, labelWidth), FitDataLabelY(drawing, y, labelHeight), labelWidth, labelHeight, layout.DataLabelFontSize, style.TextColor, alignment, style);
    }

    private static void AddPointDataLabel(
        OfficeDrawing drawing,
        OfficeChartLayout layout,
        OfficeChartStyle style,
        string category,
        OfficeChartSeries series,
        double value,
        double total,
        double x,
        double y) {
        if (!layout.ShowDataLabels) {
            return;
        }

        string label = FormatDataLabel(layout, category, series, value, total);
        if (string.IsNullOrWhiteSpace(label)) {
            return;
        }

        (double labelWidth, double labelHeight) = GetDataLabelSize(label, layout);
        double labelX = layout.DataLabelPosition switch {
            OfficeChartDataLabelPosition.Left => x - labelWidth - 4D,
            OfficeChartDataLabelPosition.Right or OfficeChartDataLabelPosition.OutsideEnd => x + 4D,
            _ => x - labelWidth / 2D
        };
        double labelY = layout.DataLabelPosition switch {
            OfficeChartDataLabelPosition.Center => y - labelHeight / 2D,
            OfficeChartDataLabelPosition.Bottom => y + 4D,
            OfficeChartDataLabelPosition.Left or OfficeChartDataLabelPosition.Right => y - labelHeight / 2D,
            _ => value >= 0D ? y - labelHeight - 4D : y + 4D
        };
        AddChartText(drawing, label, FitDataLabelX(drawing, labelX, labelWidth), FitDataLabelY(drawing, labelY, labelHeight), labelWidth, labelHeight, layout.DataLabelFontSize, style.TextColor, OfficeTextAlignment.Center, style);
    }

    private static (double Width, double Height) GetDataLabelSize(string label, OfficeChartLayout layout) {
        double labelWidth = Math.Min(78D, Math.Max(18D, label.Length * layout.DataLabelFontSize * 0.52D + 6D));
        double labelHeight = Math.Max(9D, layout.DataLabelFontSize + 3D);
        return (labelWidth, labelHeight);
    }

    private static double FitDataLabelX(OfficeDrawing drawing, double x, double width) {
        double maxX = Math.Max(0D, drawing.Width - width);
        if (x < 0D) {
            return 0D;
        }

        if (x > maxX) {
            return maxX;
        }

        return x;
    }

    private static double FitDataLabelY(OfficeDrawing drawing, double y, double height) {
        double maxY = Math.Max(0D, drawing.Height - height);
        if (y < 0D) {
            return 0D;
        }

        if (y > maxY) {
            return maxY;
        }

        return y;
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
