using System.Diagnostics;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

public sealed partial class OfficeMarkupPowerPointExporter {
    private static bool ShouldAddChartPanel(OfficeMarkupChartBlock chart) =>
        !chart.Attributes.TryGetValue("panel", out var value) || !TryParseBool(value, out var parsed) || parsed;

    private static void AddChartPanel(PowerPointSlide slide, LayoutCursor box, SlideCanvasMetrics metrics) {
        const double padding = 0.12;
        var left = Math.Max(metrics.Horizontal(0.25), box.Left - metrics.Horizontal(padding));
        var top = Math.Max(metrics.Vertical(0.25), box.Top - metrics.Vertical(padding));
        var right = Math.Min(metrics.Width - metrics.Horizontal(0.25), box.Left + box.Width + metrics.Horizontal(padding));
        var bottom = Math.Min(metrics.Height - metrics.Vertical(0.25), box.Top + box.Height + metrics.Vertical(padding));

        var panel = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            left,
            top,
            Math.Max(0.5, right - left),
            Math.Max(0.5, bottom - top),
            "OfficeIMO Markup Chart Panel");
        panel.FillColor = "F8FAFC";
        panel.OutlineColor = "D9E2EF";
        panel.OutlineWidthPoints = 0.75;
    }

    private static void ApplyChartStyle(PowerPointChart chart, OfficeMarkupChartBlock source, OfficeChartData data) {
        var font = GetAttribute(source.Attributes, "font") ?? "Aptos";
        var textColor = ToPowerPointColor(GetAttribute(source.Attributes, "color")) ?? "172033";
        var gridColor = ToPowerPointColor(GetAttribute(source.Attributes, "grid-color")) ?? "E5E7EB";
        var borderColor = ToPowerPointColor(GetAttribute(source.Attributes, "border")) ?? "D9E2EF";
        var seriesColors = ResolveChartPalette(source);
        var normalizedType = Normalize(source.ChartType);

        chart.SetTitleTextStyle(fontSizePoints: 14, bold: true, color: textColor, fontName: font);
        chart.SetLegend(C.LegendPositionValues.Bottom, overlay: false);
        chart.SetLegendTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetChartAreaStyle(fillColor: "FFFFFF", lineColor: borderColor, lineWidthPoints: 0.5);
        chart.SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "EEF2F7", lineWidthPoints: 0.5);

        for (var index = 0; index < data.Series.Count; index++) {
            var color = seriesColors[index % seriesColors.Count];
            if (normalizedType == "line") {
                chart.SetSeriesLineColor(index, color, widthPoints: 2.25);
            } else {
                chart.SetSeriesFillColor(index, color);
                chart.SetSeriesLineColor(index, color, widthPoints: 0.5);
            }
        }

        if (normalizedType == "pie" || normalizedType == "donut" || normalizedType == "doughnut") {
            chart.SetDataLabels(showValue: true, showCategoryName: false);
            chart.SetDataLabelTextStyle(fontSizePoints: 9, color: textColor, fontName: font);
            ApplyChartSemanticOptions(chart, source, normalizedType, font, textColor, gridColor);
            return;
        }

        chart.SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "4B5563", fontName: font);
        chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
        chart.ClearCategoryAxisGridlines();
        ApplyChartSemanticOptions(chart, source, normalizedType, font, textColor, gridColor);
    }

    private static void ApplyChartSemanticOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string font, string textColor, string gridColor) {
        if (GetAttribute(source.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            chart.SetCategoryAxisTitle(categoryTitle);
            chart.SetCategoryAxisTitleTextStyle(fontSizePoints: 10, bold: true, color: textColor, fontName: font);
        }

        if (GetAttribute(source.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            chart.SetValueAxisTitle(valueTitle);
            chart.SetValueAxisTitleTextStyle(fontSizePoints: 10, bold: true, color: textColor, fontName: font);
        }

        if (GetAttribute(source.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            chart.SetCategoryAxisNumberFormat(categoryFormat);
        }

        if (GetAttribute(source.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            chart.SetValueAxisNumberFormat(valueFormat);
        }

        ApplyLegendOptions(chart, source);
        ApplyDataLabelOptions(chart, source, normalizedType, font, textColor);
        ApplyGridlineOptions(chart, source, normalizedType, gridColor);
    }

    private static void ApplyLegendOptions(PowerPointChart chart, OfficeMarkupChartBlock source) {
        var legend = GetAttribute(source.Attributes, "legend", "legend-position", "legendPosition");
        if (string.IsNullOrWhiteSpace(legend)) {
            return;
        }

        var normalized = Normalize(legend!);
        if (normalized is "false" or "none" or "hidden" or "off") {
            chart.HideLegend();
            return;
        }

        if (TryParseLegendPosition(legend!, out var position)) {
            chart.SetLegend(position);
        }
    }

    private static void ApplyDataLabelOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string font, string textColor) {
        var labels = GetAttribute(source.Attributes, "labels", "data-labels", "dataLabels");
        if (string.IsNullOrWhiteSpace(labels)) {
            return;
        }

        if (!IsTruthy(labels!)) {
            chart.ClearDataLabels();
            return;
        }

        var showPercent = normalizedType is "pie" or "donut" or "doughnut"
            && IsTruthy(GetAttribute(source.Attributes, "percent", "show-percent", "showPercent") ?? "false");
        chart.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: showPercent);

        var labelPosition = GetAttribute(source.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
        if (TryParseDataLabelPosition(labelPosition, out var position)) {
            chart.SetDataLabelPosition(position);
        }

        var labelFormat = GetAttribute(source.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
        if (!string.IsNullOrWhiteSpace(labelFormat)) {
            chart.SetDataLabelNumberFormat(labelFormat!);
        }

        chart.SetDataLabelTextStyle(fontSizePoints: 9, color: textColor, fontName: font);
    }

    private static void ApplyGridlineOptions(PowerPointChart chart, OfficeMarkupChartBlock source, string normalizedType, string gridColor) {
        if (normalizedType is "pie" or "donut" or "doughnut") {
            return;
        }

        var gridlines = GetAttribute(source.Attributes, "gridlines");
        var valueGridlines = GetAttribute(source.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(source.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");

        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            if (IsTruthy(valueGridlines!)) {
                chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
            } else {
                chart.ClearValueAxisGridlines();
            }
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            if (IsTruthy(categoryGridlines!)) {
                chart.SetCategoryAxisGridlines(showMajor: true, showMinor: false, lineColor: gridColor, lineWidthPoints: 0.5);
            } else {
                chart.ClearCategoryAxisGridlines();
            }
        }
    }

    private static IReadOnlyList<string> ResolveChartPalette(OfficeMarkupChartBlock chart) {
        if (chart.Attributes.TryGetValue("palette", out var palette) && !string.IsNullOrWhiteSpace(palette)) {
            var colors = palette.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(ToPowerPointColor)
                .Where(color => !string.IsNullOrWhiteSpace(color))
                .Cast<string>()
                .ToList();
            if (colors.Count > 0) {
                return colors;
            }
        }

        return new[] { "2563EB", "F97316", "10B981", "A855F7", "EF4444", "14B8A6" };
    }

    private static bool TryParseBool(string value, out bool result) {
        if (bool.TryParse(value, out result)) {
            return true;
        }

        switch (Normalize(value)) {
            case "yes":
            case "y":
            case "1":
            case "on":
                result = true;
                return true;
            case "no":
            case "n":
            case "0":
            case "off":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    private static bool IsTruthy(string value) =>
        Normalize(value) is not ("false" or "no" or "off" or "none" or "hidden" or "0");

    private static bool TryParseLegendPosition(string value, out C.LegendPositionValues position) {
        switch (Normalize(value)) {
            case "left":
                position = C.LegendPositionValues.Left;
                return true;
            case "right":
                position = C.LegendPositionValues.Right;
                return true;
            case "top":
                position = C.LegendPositionValues.Top;
                return true;
            case "bottom":
                position = C.LegendPositionValues.Bottom;
                return true;
            case "corner":
            case "topright":
                position = C.LegendPositionValues.TopRight;
                return true;
            default:
                position = C.LegendPositionValues.Bottom;
                return false;
        }
    }

    private static bool TryParseDataLabelPosition(string? value, out C.DataLabelPositionValues position) {
        switch (Normalize(value ?? string.Empty)) {
            case "center":
                position = C.DataLabelPositionValues.Center;
                return true;
            case "insideend":
                position = C.DataLabelPositionValues.InsideEnd;
                return true;
            case "insidebase":
                position = C.DataLabelPositionValues.InsideBase;
                return true;
            case "outsideend":
                position = C.DataLabelPositionValues.OutsideEnd;
                return true;
            case "bestfit":
                position = C.DataLabelPositionValues.BestFit;
                return true;
            case "left":
                position = C.DataLabelPositionValues.Left;
                return true;
            case "right":
                position = C.DataLabelPositionValues.Right;
                return true;
            case "top":
                position = C.DataLabelPositionValues.Top;
                return true;
            case "bottom":
                position = C.DataLabelPositionValues.Bottom;
                return true;
            default:
                position = C.DataLabelPositionValues.OutsideEnd;
                return false;
        }
    }

    private static string? ToPowerPointColor(string? color) {
        if (string.IsNullOrWhiteSpace(color)) {
            return null;
        }

        color = color!.Trim();
        if (color.StartsWith("#", StringComparison.Ordinal)) {
            color = color.Substring(1);
        }

        return color.Length == 6 && color.All(IsHexDigit) ? color.ToUpperInvariant() : null;
    }

    private static bool IsHexDigit(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'a' && value <= 'f')
        || (value >= 'A' && value <= 'F');

    private static void AddChart(
        PowerPointSlide slide,
        OfficeMarkupChartBlock chart,
        LayoutCursor cursor,
        OfficeMarkupPowerPointExportOptions options,
        SlideCanvasMetrics metrics) {
        if (!TryCreateChartData(chart, out var data)) {
            if (options.IncludeUnsupportedBlocksAsText) {
                AddText(slide, $"Chart: {chart.Title ?? chart.ChartType}", cursor, height: 0.55);
            }

            return;
        }

        var box = ResolveBox(chart.Placement, chart.Attributes, cursor, Math.Min(2.4, cursor.RemainingHeight), metrics);
        if (ShouldAddChartPanel(chart)) {
            AddChartPanel(slide, box, metrics);
        }

        var nativeChart = AddNativeChart(slide, chart.ChartType, data, box);
        if (!string.IsNullOrWhiteSpace(chart.Title)) {
            nativeChart.SetTitle(chart.Title!);
        }

        ApplyChartStyle(nativeChart, chart, data);
        if (!HasExplicitPlacement(chart.Placement, chart.Attributes)) {
            cursor.Advance(box.Height);
        }
    }

    private static PowerPointChart AddNativeChart(
        PowerPointSlide slide,
        string chartType,
        OfficeChartData data,
        LayoutCursor box) {
        string normalized = Normalize(chartType);
        OfficeChartKind kind = normalized switch {
            "line" => OfficeChartKind.Line,
            "pie" => OfficeChartKind.Pie,
            "donut" or "doughnut" => OfficeChartKind.Doughnut,
            "bar" or "clusteredbar" => OfficeChartKind.BarClustered,
            _ => OfficeChartKind.ColumnClustered
        };
        OfficeChartData resolved = kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut
            ? FirstSeriesOnly(data)
            : data;
        return slide.AddChartInches(kind, resolved, box.Left, box.Top, box.Width, box.Height);
    }

    private static bool TryCreateChartData(OfficeMarkupChartBlock chart, out OfficeChartData data) {
        data = null!;
        if (chart.Data.Count < 2) {
            return false;
        }

        var headers = chart.Data[0].Select(cell => cell ?? string.Empty).ToList();
        if (headers.Count < 2) {
            return false;
        }

        var categories = new List<string>();
        var seriesValues = new List<List<double>>();
        for (int columnIndex = 1; columnIndex < headers.Count; columnIndex++) {
            seriesValues.Add(new List<double>());
        }

        for (int rowIndex = 1; rowIndex < chart.Data.Count; rowIndex++) {
            var row = chart.Data[rowIndex];
            if (row.Count == 0 || string.IsNullOrWhiteSpace(row[0])) {
                continue;
            }

            categories.Add(row[0]);
            for (int columnIndex = 1; columnIndex < headers.Count; columnIndex++) {
                var value = columnIndex < row.Count ? row[columnIndex] : string.Empty;
                seriesValues[columnIndex - 1].Add(ParseDouble(value));
            }
        }

        if (categories.Count == 0) {
            return false;
        }

        var series = new List<OfficeChartSeries>();
        for (int index = 0; index < seriesValues.Count; index++) {
            var name = string.IsNullOrWhiteSpace(headers[index + 1]) ? $"Series {index + 1}" : headers[index + 1];
            series.Add(new OfficeChartSeries(name, seriesValues[index]));
        }

        data = new OfficeChartData(categories, series);
        return true;
    }

    private static OfficeChartData FirstSeriesOnly(OfficeChartData data) =>
        new OfficeChartData(data.Categories, data.Series.Take(1));

    private static double ParseDouble(string value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0d;
}
