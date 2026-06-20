using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private const double MinimumChartWidth = 240D;
    private const double MinimumChartHeight = 150D;

    private static bool TryRenderChartFencedBlock(PdfCore.PdfDocument pdf, SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (!IsChartSemanticFence(semantic)) {
            return false;
        }

        if (!TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot, out string? warningMessage, visualTheme)) {
            if (!string.IsNullOrWhiteSpace(warningMessage)) {
                AddWarning(options, "UnsupportedChartFence", semantic.Language, warningMessage!);
            }

            return false;
        }

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot!);
        MarkdownPdfFigureStyle figureStyle = visualTheme.FigureStyleSnapshot;
        PdfCore.PdfDrawingStyle drawingStyle = figureStyle.DrawingStyleSnapshot;
        drawingStyle.Decorative = false;
        drawingStyle.AlternativeText = string.IsNullOrWhiteSpace(snapshot!.Title)
            ? "Markdown chart"
            : "Markdown chart: " + snapshot.Title;

        pdf.Drawing(drawing, style: drawingStyle);
        RenderFigureCaption(pdf, semantic.Caption, figureStyle);
        return true;
    }

    private static bool IsChartSemanticFence(SemanticFencedBlock semantic) =>
        string.Equals(semantic.SemanticKind, MarkdownSemanticKinds.Chart, StringComparison.OrdinalIgnoreCase);

    internal static bool TryCreateChartSnapshot(SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, out OfficeChartSnapshot? snapshot, out string? warningMessage, MarkdownPdfVisualTheme? visualTheme = null) {
        snapshot = null;
        warningMessage = null;
        if (string.IsNullOrWhiteSpace(semantic.Content)) {
            warningMessage = "The Markdown chart fence is empty and is rendered as a semantic code panel.";
            return false;
        }

        try {
            MarkdownPdfJsonValue root = MarkdownPdfJsonValue.Parse(semantic.Content);
            if (root.Kind != MarkdownPdfJsonValueKind.Object) {
                warningMessage = "The Markdown chart fence must contain a JSON object and is rendered as a semantic code panel.";
                return false;
            }

            string type = ReadString(root, "type") ?? "bar";
            if (!TryMapChartKind(type, UsesHorizontalIndexAxis(root), UsesStackedScale(root), out OfficeChartKind chartKind)) {
                warningMessage = "The Markdown chart fence uses an unsupported chart type and is rendered as a semantic code panel.";
                return false;
            }

            MarkdownPdfJsonValue dataElement = TryGetProperty(root, "data", out MarkdownPdfJsonValue data)
                ? data
                : root;
            List<string> labels = ReadLabels(dataElement);
            bool defaultConnectLine = chartKind != OfficeChartKind.Scatter || ReadChartShowLine(root) == true;
            List<OfficeChartSeries> series = ReadSeries(dataElement, labels, chartKind, defaultConnectLine);
            if (series.Count == 0) {
                warningMessage = "The Markdown chart fence does not contain renderable chart series and is rendered as a semantic code panel.";
                return false;
            }

            if (labels.Count == 0) {
                int maxValues = series.Max(item => item.Values.Count);
                for (int i = 0; i < maxValues; i++) {
                    labels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }

            if (chartKind != OfficeChartKind.Scatter) {
                series = NormalizeSeriesToLabelCount(series, labels.Count);
            }

            if (chartKind == OfficeChartKind.Scatter && !HasDrawableScatterPoint(series)) {
                warningMessage = "The Markdown scatter chart fence does not contain a finite X/Y point and is rendered as a semantic code panel.";
                return false;
            }

            if (!HasFiniteValue(series)) {
                warningMessage = "The Markdown chart fence does not contain finite renderable chart values and is rendered as a semantic code panel.";
                return false;
            }

            if ((chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut) && !HasPositiveFiniteSlice(series)) {
                warningMessage = "The Markdown pie or doughnut chart fence does not contain a positive finite slice and is rendered as a semantic code panel.";
                return false;
            }

            if (labels.Count == 0) {
                warningMessage = "The Markdown chart fence does not contain renderable chart labels and is rendered as a semantic code panel.";
                return false;
            }

            if (TryGetAvailablePdfContentWidth(options, out double availableWidth) && availableWidth < MinimumChartWidth) {
                warningMessage = "The Markdown chart fence needs at least 240 PDF points of content width for native rendering and is rendered as a semantic code panel.";
                return false;
            }

            MarkdownPdfVisualTheme resolvedVisualTheme = ResolveChartVisualTheme(options, visualTheme);
            if (TryGetAvailableChartContentHeight(options, resolvedVisualTheme, out double availableHeight) && availableHeight < MinimumChartHeight) {
                warningMessage = "The Markdown chart fence needs at least 150 PDF points of content height after figure spacing for native rendering and is rendered as a semantic code panel.";
                return false;
            }

            if (chartKind == OfficeChartKind.Radar && labels.Count < 3) {
                warningMessage = "The Markdown radar chart fence needs at least three categories and is rendered as a semantic code panel.";
                return false;
            }

            if (chartKind == OfficeChartKind.Radar) {
                series = FilterDrawableRadarSeries(series, labels.Count);
                if (series.Count == 0) {
                    warningMessage = "The Markdown radar chart fence does not contain drawable adjacent or complete finite data and is rendered as a semantic code panel.";
                    return false;
                }
            }

            if (IsAreaChart(chartKind) && labels.Count < 2) {
                warningMessage = "The Markdown area chart fence needs at least two categories and is rendered as a semantic code panel.";
                return false;
            }

            if (IsAreaChart(chartKind)) {
                series = FilterDrawableAreaSeries(series, labels.Count);
                if (series.Count == 0) {
                    warningMessage = "The Markdown area chart fence does not contain two adjacent finite values and is rendered as a semantic code panel.";
                    return false;
                }
            }

            if (IsAreaChart(chartKind) && !HasAdjacentFiniteRun(series, labels.Count)) {
                warningMessage = "The Markdown area chart fence does not contain two adjacent finite values and is rendered as a semantic code panel.";
                return false;
            }

            string? title = ReadChartTitle(root) ?? semantic.FenceInfo.Title;
            double width = ReadPositiveDouble(root, "width") ?? options.DefaultImageWidth;
            double height = ReadPositiveDouble(root, "height") ?? options.DefaultImageHeight;
            width = Math.Max(MinimumChartWidth, Math.Min(520D, width));
            height = Math.Max(150D, Math.Min(320D, height));
            FitChartToPageFrame(options, resolvedVisualTheme, ref width, ref height);

            snapshot = new OfficeChartSnapshot(
                "Markdown chart",
                title,
                chartKind,
                new OfficeChartData(labels, series),
                width,
                height,
                OfficeChartStyle.Default,
                CreateMarkdownChartLayout(root, chartKind, series));
            return true;
        } catch (FormatException) {
            warningMessage = "The Markdown chart fence is not valid JSON and is rendered as a semantic code panel.";
            return false;
        } catch (ArgumentException ex) {
            warningMessage = "The Markdown chart fence could not be rendered as a chart: " + ex.Message;
            return false;
        }
    }

    private static bool HasPositiveFiniteSlice(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            IReadOnlyList<double> values = series[seriesIndex].Values;
            for (int valueIndex = 0; valueIndex < values.Count; valueIndex++) {
                double value = values[valueIndex];
                if (value > 0D && !double.IsNaN(value) && !double.IsInfinity(value)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool HasFiniteValue(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            IReadOnlyList<double> values = series[seriesIndex].Values;
            for (int valueIndex = 0; valueIndex < values.Count; valueIndex++) {
                double value = values[valueIndex];
                if (IsFiniteChartValue(value)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool HasDrawableScatterPoint(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            OfficeChartSeries item = series[seriesIndex];
            IReadOnlyList<double>? xValues = item.XValues;
            IReadOnlyList<double> yValues = item.Values;
            if (xValues == null) {
                for (int valueIndex = 0; valueIndex < yValues.Count; valueIndex++) {
                    if (IsFiniteChartValue(yValues[valueIndex])) {
                        return true;
                    }
                }

                continue;
            }

            int count = Math.Min(xValues.Count, yValues.Count);
            for (int valueIndex = 0; valueIndex < count; valueIndex++) {
                if (IsFiniteChartValue(xValues[valueIndex]) && IsFiniteChartValue(yValues[valueIndex])) {
                    return true;
                }
            }
        }

        return false;
    }

    private static List<OfficeChartSeries> FilterDrawableAreaSeries(IReadOnlyList<OfficeChartSeries> series, int labelCount) {
        var filtered = new List<OfficeChartSeries>(series.Count);
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            OfficeChartSeries item = series[seriesIndex];
            if (HasAdjacentFiniteRun(item.Values, labelCount)) {
                filtered.Add(item);
            }
        }

        return filtered;
    }

    private static List<OfficeChartSeries> FilterDrawableRadarSeries(IReadOnlyList<OfficeChartSeries> series, int labelCount) {
        var filtered = new List<OfficeChartSeries>(series.Count);
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            OfficeChartSeries item = series[seriesIndex];
            if (HasCompleteFiniteRun(item.Values, labelCount) || HasAdjacentFiniteRun(item.Values, labelCount)) {
                filtered.Add(item);
            }
        }

        return filtered;
    }

    private static bool HasAdjacentFiniteRun(IReadOnlyList<OfficeChartSeries> series, int labelCount) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            if (HasAdjacentFiniteRun(series[seriesIndex].Values, labelCount)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasAdjacentFiniteRun(IReadOnlyList<double> values, int labelCount) {
        int count = Math.Min(labelCount, values.Count);
        for (int valueIndex = 1; valueIndex < count; valueIndex++) {
            if (IsFiniteChartValue(values[valueIndex - 1]) && IsFiniteChartValue(values[valueIndex])) {
                return true;
            }
        }

        return false;
    }

    private static bool HasCompleteFiniteRun(IReadOnlyList<double> values, int labelCount) {
        if (labelCount <= 0 || values.Count < labelCount) {
            return false;
        }

        for (int valueIndex = 0; valueIndex < labelCount; valueIndex++) {
            if (!IsFiniteChartValue(values[valueIndex])) {
                return false;
            }
        }

        return true;
    }

    private static bool IsFiniteChartValue(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static MarkdownPdfVisualTheme ResolveChartVisualTheme(MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme? visualTheme) {
        if (visualTheme != null) {
            return visualTheme.Clone();
        }

        MarkdownPdfVisualTheme? explicitTheme = options.VisualTheme;
        if (explicitTheme != null) {
            return explicitTheme;
        }

        return options.ApplyWordLikeTheme ? MarkdownPdfVisualTheme.WordLike() : MarkdownPdfVisualTheme.Plain();
    }

    private static void FitChartToPageFrame(MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme, ref double width, ref double height) {
        if (!TryGetAvailablePdfContentWidth(options, out double availableWidth)) {
            availableWidth = width;
        }

        double maxWidth = Math.Max(MinimumChartWidth, availableWidth);
        double maxHeight = height;
        if (TryGetAvailableChartContentHeight(options, visualTheme, out double availableHeight)) {
            maxHeight = Math.Max(MinimumChartHeight, availableHeight);
        }

        if (width <= maxWidth && height <= maxHeight) {
            return;
        }

        double scale = Math.Min(maxWidth / width, maxHeight / height);
        width = Math.Max(MinimumChartWidth, width * scale);
        height = Math.Max(MinimumChartHeight, height * scale);
    }

    private static bool TryGetAvailablePdfContentWidth(MarkdownPdfSaveOptions options, out double availableWidth) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions ?? new PdfCore.PdfOptions();
        availableWidth = pdfOptions.PageWidth - pdfOptions.MarginLeft - pdfOptions.MarginRight;
        return availableWidth > 0D && !double.IsNaN(availableWidth) && !double.IsInfinity(availableWidth);
    }

    private static bool TryGetAvailablePdfContentHeight(MarkdownPdfSaveOptions options, out double availableHeight) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions ?? new PdfCore.PdfOptions();
        availableHeight = pdfOptions.PageHeight - pdfOptions.MarginTop - pdfOptions.MarginBottom;
        return availableHeight > 0D && !double.IsNaN(availableHeight) && !double.IsInfinity(availableHeight);
    }

    private static bool TryGetAvailableChartContentHeight(MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme, out double availableHeight) {
        if (!TryGetAvailablePdfContentHeight(options, out availableHeight)) {
            return false;
        }

        PdfCore.PdfDrawingStyle drawingStyle = visualTheme.FigureStyleSnapshot.DrawingStyleSnapshot;
        availableHeight -= drawingStyle.SpacingBefore + drawingStyle.SpacingAfter;
        return availableHeight > 0D && !double.IsNaN(availableHeight) && !double.IsInfinity(availableHeight);
    }

    private static OfficeChartLayout CreateMarkdownChartLayout(MarkdownPdfJsonValue root, OfficeChartKind chartKind, IReadOnlyList<OfficeChartSeries> series) {
        bool pie = chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut;
        bool showLegend = ReadChartLegendDisplay(root) != false;
        bool connectScatterPoints = chartKind == OfficeChartKind.Scatter && HasScatterSeriesLine(series);
        return new OfficeChartLayout(
            showLegend: showLegend,
            legendPosition: pie ? OfficeChartLegendPosition.Right : OfficeChartLegendPosition.Bottom,
            showDataLabels: pie,
            showDataLabelValues: false,
            showDataLabelPercentages: pie,
            showDataLabelCategoryNames: pie,
            dataLabelFontSize: 7D,
            maximumCategoryAxisLabels: 8,
            maximumHorizontalCategoryAxisLabels: 8,
            showMarkers: IsLineChart(chartKind) || chartKind == OfficeChartKind.Scatter,
            connectScatterPoints: connectScatterPoints);
    }

    private static bool HasScatterSeriesLine(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            if (series[seriesIndex].ConnectLine) {
                return true;
            }
        }

        return false;
    }

    private static List<string> ReadLabels(MarkdownPdfJsonValue dataElement) {
        var labels = new List<string>();
        if (!TryGetProperty(dataElement, "labels", out MarkdownPdfJsonValue labelElement) || labelElement.Kind != MarkdownPdfJsonValueKind.Array) {
            return labels;
        }

        foreach (MarkdownPdfJsonValue item in labelElement.ArrayValues) {
            string? label = item.ReadScalarAsText();
            labels.Add(string.IsNullOrWhiteSpace(label) ? string.Empty : label!);
        }

        return labels;
    }

    private static string? ReadChartTitle(MarkdownPdfJsonValue root) {
        string? title = ReadString(root, "title");
        if (!string.IsNullOrWhiteSpace(title)) {
            return title;
        }

        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
            TryGetProperty(options, "plugins", out MarkdownPdfJsonValue plugins) &&
            TryGetProperty(plugins, "title", out MarkdownPdfJsonValue titleElement)) {
            if (titleElement.Kind == MarkdownPdfJsonValueKind.Object && TryGetProperty(titleElement, "text", out MarkdownPdfJsonValue textElement)) {
                if (ReadBool(titleElement, "display") != true) {
                    return null;
                }

                return textElement.ReadScalarAsText();
            }

            return titleElement.ReadScalarAsText();
        }

        return null;
    }

    private static bool UsesHorizontalIndexAxis(MarkdownPdfJsonValue root) =>
        TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
        string.Equals(ReadString(options, "indexAxis"), "y", StringComparison.OrdinalIgnoreCase);

    private static bool UsesStackedScale(MarkdownPdfJsonValue root) =>
        TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
        TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) &&
        (ReadScaleStacked(scales, "x") == true || ReadScaleStacked(scales, "y") == true);

    private static bool? ReadScaleStacked(MarkdownPdfJsonValue scales, string axisName) {
        if (!TryGetProperty(scales, axisName, out MarkdownPdfJsonValue axis) || axis.Kind != MarkdownPdfJsonValueKind.Object) {
            return null;
        }

        return ReadBool(axis, "stacked");
    }

    private static bool TryMapChartKind(string? type, bool horizontalIndexAxis, bool stacked, out OfficeChartKind kind) {
        string normalized = NormalizeChartType(type);
        switch (normalized) {
            case "bar":
                if (horizontalIndexAxis) {
                    kind = stacked ? OfficeChartKind.BarStacked : OfficeChartKind.BarClustered;
                } else {
                    kind = stacked ? OfficeChartKind.ColumnStacked : OfficeChartKind.ColumnClustered;
                }
                return true;
            case "column":
                kind = stacked ? OfficeChartKind.ColumnStacked : OfficeChartKind.ColumnClustered;
                return true;
            case "horizontalbar":
            case "barhorizontal":
                kind = stacked ? OfficeChartKind.BarStacked : OfficeChartKind.BarClustered;
                return true;
            case "line":
                kind = stacked ? OfficeChartKind.LineStacked : OfficeChartKind.Line;
                return true;
            case "area":
                kind = stacked ? OfficeChartKind.AreaStacked : OfficeChartKind.Area;
                return true;
            case "pie":
                kind = OfficeChartKind.Pie;
                return true;
            case "doughnut":
            case "donut":
                kind = OfficeChartKind.Doughnut;
                return true;
            case "scatter":
                kind = OfficeChartKind.Scatter;
                return true;
            case "radar":
                kind = OfficeChartKind.Radar;
                return true;
            default:
                kind = OfficeChartKind.ColumnClustered;
                return false;
        }
    }

    private static bool IsLineChart(OfficeChartKind chartKind) =>
        chartKind == OfficeChartKind.Line ||
        chartKind == OfficeChartKind.LineStacked ||
        chartKind == OfficeChartKind.LineStacked100;

    private static bool IsAreaChart(OfficeChartKind chartKind) =>
        chartKind == OfficeChartKind.Area ||
        chartKind == OfficeChartKind.AreaStacked ||
        chartKind == OfficeChartKind.AreaStacked100;

    private static string NormalizeChartType(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return string.Empty;
        }

        var builder = new StringBuilder(value!.Length);
        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (char.IsLetterOrDigit(ch)) {
                builder.Append(char.ToLowerInvariant(ch));
            }
        }

        return builder.ToString();
    }

    private static string? ReadString(MarkdownPdfJsonValue element, string propertyName) =>
        TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value) ? value.ReadScalarAsText() : null;

    private static bool? ReadBool(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value)) {
            return null;
        }

        switch (value.Kind) {
            case MarkdownPdfJsonValueKind.True:
                return true;
            case MarkdownPdfJsonValueKind.False:
                return false;
            case MarkdownPdfJsonValueKind.String:
                if (bool.TryParse(value.StringValue, out bool parsed)) {
                    return parsed;
                }

                return null;
            default:
                return null;
        }
    }

    private static bool? ReadChartShowLine(MarkdownPdfJsonValue root) {
        bool? showLine = ReadBool(root, "showLine");
        if (showLine.HasValue) {
            return showLine;
        }

        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options)) {
            showLine = ReadBool(options, "showLine");
            if (showLine.HasValue) {
                return showLine;
            }
        }

        return null;
    }

    private static bool? ReadChartLegendDisplay(MarkdownPdfJsonValue root) {
        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
            TryGetProperty(options, "plugins", out MarkdownPdfJsonValue plugins) &&
            TryGetProperty(plugins, "legend", out MarkdownPdfJsonValue legend) &&
            legend.Kind == MarkdownPdfJsonValueKind.Object) {
            return ReadBool(legend, "display");
        }

        return null;
    }

    private static double? ReadPositiveDouble(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value) || !TryReadNumber(value, out double number)) {
            return null;
        }

        return number > 0D && !double.IsNaN(number) && !double.IsInfinity(number) ? number : null;
    }

    private static bool TryReadNumber(MarkdownPdfJsonValue element, out double value) => element.TryGetDouble(out value);

    private static bool TryGetProperty(MarkdownPdfJsonValue element, string propertyName, out MarkdownPdfJsonValue value) =>
        element.TryGetProperty(propertyName, out value);

}
