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

        if (!TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot, out string? warningMessage)) {
            if (!string.IsNullOrWhiteSpace(warningMessage)) {
                AddWarning(options, "UnsupportedChartFence", semantic.Language, warningMessage!);
            }

            return false;
        }

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot!);
        MarkdownPdfFigureStyle figureStyle = visualTheme.FigureStyleSnapshot;
        PdfCore.PdfDrawingStyle drawingStyle = figureStyle.DrawingStyleSnapshot;
        drawingStyle.AlternativeText = string.IsNullOrWhiteSpace(snapshot!.Title)
            ? "Markdown chart"
            : "Markdown chart: " + snapshot.Title;

        pdf.Drawing(drawing, style: drawingStyle);
        RenderFigureCaption(pdf, semantic.Caption, figureStyle);
        return true;
    }

    private static bool IsChartSemanticFence(SemanticFencedBlock semantic) =>
        string.Equals(semantic.SemanticKind, MarkdownSemanticKinds.Chart, StringComparison.OrdinalIgnoreCase);

    internal static bool TryCreateChartSnapshot(SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, out OfficeChartSnapshot? snapshot, out string? warningMessage) {
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
            if (!TryMapChartKind(type, UsesHorizontalIndexAxis(root), out OfficeChartKind chartKind)) {
                warningMessage = "The Markdown chart fence uses an unsupported chart type and is rendered as a semantic code panel.";
                return false;
            }

            MarkdownPdfJsonValue dataElement = TryGetProperty(root, "data", out MarkdownPdfJsonValue data)
                ? data
                : root;
            List<string> labels = ReadLabels(dataElement);
            List<OfficeChartSeries> series = ReadSeries(dataElement, labels, chartKind);
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

            if (labels.Count == 0) {
                warningMessage = "The Markdown chart fence does not contain renderable chart labels and is rendered as a semantic code panel.";
                return false;
            }

            if (TryGetAvailablePdfContentWidth(options, out double availableWidth) && availableWidth < MinimumChartWidth) {
                warningMessage = "The Markdown chart fence needs at least 240 PDF points of content width for native rendering and is rendered as a semantic code panel.";
                return false;
            }

            if (TryGetAvailablePdfContentHeight(options, out double availableHeight) && availableHeight < MinimumChartHeight) {
                warningMessage = "The Markdown chart fence needs at least 150 PDF points of content height for native rendering and is rendered as a semantic code panel.";
                return false;
            }

            if (chartKind == OfficeChartKind.Radar && labels.Count < 3) {
                warningMessage = "The Markdown radar chart fence needs at least three categories and is rendered as a semantic code panel.";
                return false;
            }

            string? title = ReadChartTitle(root) ?? semantic.FenceInfo.Title;
            double width = ReadPositiveDouble(root, "width") ?? options.DefaultImageWidth;
            double height = ReadPositiveDouble(root, "height") ?? options.DefaultImageHeight;
            width = Math.Max(MinimumChartWidth, Math.Min(520D, width));
            height = Math.Max(150D, Math.Min(320D, height));
            FitChartToPageFrame(options, ref width, ref height);

            snapshot = new OfficeChartSnapshot(
                "Markdown chart",
                title,
                chartKind,
                new OfficeChartData(labels, series),
                width,
                height,
                OfficeChartStyle.Default,
                CreateMarkdownChartLayout(chartKind));
            return true;
        } catch (FormatException) {
            warningMessage = "The Markdown chart fence is not valid JSON and is rendered as a semantic code panel.";
            return false;
        } catch (ArgumentException ex) {
            warningMessage = "The Markdown chart fence could not be rendered as a chart: " + ex.Message;
            return false;
        }
    }

    private static void FitChartToPageFrame(MarkdownPdfSaveOptions options, ref double width, ref double height) {
        if (!TryGetAvailablePdfContentWidth(options, out double availableWidth)) {
            availableWidth = width;
        }

        double maxWidth = Math.Max(MinimumChartWidth, availableWidth);
        double maxHeight = height;
        if (TryGetAvailablePdfContentHeight(options, out double availableHeight)) {
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

    private static OfficeChartLayout CreateMarkdownChartLayout(OfficeChartKind chartKind) {
        bool pie = chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut;
        return new OfficeChartLayout(
            legendPosition: pie ? OfficeChartLegendPosition.Right : OfficeChartLegendPosition.Bottom,
            showDataLabels: pie,
            showDataLabelValues: false,
            showDataLabelPercentages: pie,
            showDataLabelCategoryNames: pie,
            dataLabelFontSize: 7D,
            maximumCategoryAxisLabels: 8,
            maximumHorizontalCategoryAxisLabels: 8,
            showMarkers: chartKind == OfficeChartKind.Line || chartKind == OfficeChartKind.Scatter);
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

    private static List<OfficeChartSeries> ReadSeries(MarkdownPdfJsonValue dataElement, List<string> labels, OfficeChartKind chartKind) {
        var series = new List<OfficeChartSeries>();
        bool captureXValues = chartKind == OfficeChartKind.Scatter;
        bool canUsePointCategories = labels.Count == 0 && !captureXValues;
        if (TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) && datasets.Kind == MarkdownPdfJsonValueKind.Array) {
            int index = 0;
            foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
                if (dataset.Kind != MarkdownPdfJsonValueKind.Object) {
                    continue;
                }

                MarkdownChartSeriesValues seriesValues = ReadDataValues(dataset, captureXValues, canUsePointCategories && series.Count == 0);
                if (seriesValues.Values.Count == 0) {
                    continue;
                }

                if (canUsePointCategories && seriesValues.CategoryLabels != null && seriesValues.CategoryLabels.Count > 0) {
                    labels.AddRange(seriesValues.CategoryLabels);
                    canUsePointCategories = false;
                }

                if (captureXValues) {
                    EnsureLabelsCoverValues(labels, seriesValues.Values.Count);
                } else {
                    NormalizeSeriesLength(seriesValues.Values, labels.Count);
                    if (seriesValues.XValues != null) {
                        NormalizeSeriesLength(seriesValues.XValues, labels.Count);
                    }
                }

                OfficeColor? color = ReadColor(dataset, "borderColor") ?? ReadColor(dataset, "backgroundColor");
                IReadOnlyList<OfficeColor?>? pointColors = ReadPointColors(dataset, "backgroundColor", seriesValues.Values.Count);
                string name = ReadString(dataset, "label") ?? "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
                series.Add(new OfficeChartSeries(name, seriesValues.Values, seriesValues.XValues, color, pointColors));
                index++;
            }
        }

        if (series.Count == 0 && TryGetProperty(dataElement, "values", out MarkdownPdfJsonValue valuesElement)) {
            MarkdownChartSeriesValues values = ReadNumberArray(valuesElement, captureXValues: false, captureCategoryLabels: labels.Count == 0);
            if (values.Values.Count > 0) {
                if (labels.Count == 0 && values.CategoryLabels != null && values.CategoryLabels.Count > 0) {
                    labels.AddRange(values.CategoryLabels);
                }

                NormalizeSeriesLength(values.Values, labels.Count);
                series.Add(new OfficeChartSeries("Values", values.Values));
            }
        }

        if (chartKind == OfficeChartKind.Pie && series.Count > 1) {
            return new List<OfficeChartSeries> { series[0] };
        }

        return series;
    }

    private static MarkdownChartSeriesValues ReadDataValues(MarkdownPdfJsonValue dataset, bool captureXValues, bool captureCategoryLabels) {
        if (!TryGetProperty(dataset, "data", out MarkdownPdfJsonValue data)) {
            return new MarkdownChartSeriesValues(new List<double>(), null, null);
        }

        return ReadNumberArray(data, captureXValues, captureCategoryLabels);
    }

    private static MarkdownChartSeriesValues ReadNumberArray(MarkdownPdfJsonValue element, bool captureXValues, bool captureCategoryLabels) {
        var values = new List<double>();
        List<double>? xValues = captureXValues ? new List<double>() : null;
        List<string>? categoryLabels = captureCategoryLabels ? new List<string>() : null;
        bool hasExplicitXValue = false;
        if (element.Kind != MarkdownPdfJsonValueKind.Array) {
            return new MarkdownChartSeriesValues(values, null, null);
        }

        foreach (MarkdownPdfJsonValue item in element.ArrayValues) {
            bool hasPoint = false;
            double yValue = double.NaN;
            double xValue = double.NaN;
            string? categoryLabel = null;

            if (TryReadNumber(item, out double scalarValue)) {
                yValue = scalarValue;
                hasPoint = true;
            } else if (item.Kind == MarkdownPdfJsonValueKind.Null) {
                hasPoint = true;
            } else if (item.Kind == MarkdownPdfJsonValueKind.Array) {
                IReadOnlyList<MarkdownPdfJsonValue> pointValues = item.ArrayValues;
                if (pointValues.Count >= 2) {
                    hasPoint = true;
                    hasExplicitXValue = true;
                    categoryLabel = pointValues[0].ReadScalarAsText();
                    if (TryReadNumber(pointValues[0], out double parsedX)) {
                        xValue = parsedX;
                    }

                    if (TryReadNumber(pointValues[1], out double parsedY)) {
                        yValue = parsedY;
                    }
                }
            } else if (item.Kind == MarkdownPdfJsonValueKind.Object) {
                bool hasY = TryGetProperty(item, "y", out MarkdownPdfJsonValue y);
                bool hasX = TryGetProperty(item, "x", out MarkdownPdfJsonValue x);
                if (hasY || hasX) {
                    hasPoint = true;
                    if (hasY && TryReadNumber(y, out double parsedY)) {
                        yValue = parsedY;
                    }

                    if (hasX) {
                        hasExplicitXValue = true;
                        categoryLabel = x.ReadScalarAsText();
                        if (TryReadNumber(x, out double parsedX)) {
                            xValue = parsedX;
                        }
                    }
                }
            }

            if (!hasPoint) {
                continue;
            }

            values.Add(yValue);
            if (xValues != null) {
                xValues.Add(xValue);
            }

            if (categoryLabels != null) {
                categoryLabels.Add(string.IsNullOrWhiteSpace(categoryLabel) ? values.Count.ToString(CultureInfo.InvariantCulture) : categoryLabel!);
            }
        }

        return new MarkdownChartSeriesValues(values, hasExplicitXValue ? xValues : null, categoryLabels != null && categoryLabels.Count > 0 ? categoryLabels : null);
    }

    private static void EnsureLabelsCoverValues(List<string> labels, int valueCount) {
        for (int i = labels.Count; i < valueCount; i++) {
            labels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
        }
    }

    private static void NormalizeSeriesLength(List<double> values, int labelCount) {
        if (labelCount <= 0) {
            return;
        }

        while (values.Count < labelCount) {
            values.Add(double.NaN);
        }

        if (values.Count > labelCount) {
            values.RemoveRange(labelCount, values.Count - labelCount);
        }
    }

    private static OfficeColor? ReadColor(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue colorElement)) {
            return null;
        }

        if (colorElement.Kind == MarkdownPdfJsonValueKind.String && OfficeColor.TryParse(colorElement.StringValue, out OfficeColor color)) {
            return color;
        }

        if (colorElement.Kind == MarkdownPdfJsonValueKind.Array) {
            foreach (MarkdownPdfJsonValue item in colorElement.ArrayValues) {
                if (item.Kind == MarkdownPdfJsonValueKind.String && OfficeColor.TryParse(item.StringValue, out color)) {
                    return color;
                }
            }
        }

        return null;
    }

    private static IReadOnlyList<OfficeColor?>? ReadPointColors(MarkdownPdfJsonValue element, string propertyName, int expectedCount) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue colorElement) || colorElement.Kind != MarkdownPdfJsonValueKind.Array) {
            return null;
        }

        var colors = new List<OfficeColor?>();
        foreach (MarkdownPdfJsonValue item in colorElement.ArrayValues) {
            colors.Add(item.Kind == MarkdownPdfJsonValueKind.String && OfficeColor.TryParse(item.StringValue, out OfficeColor color) ? color : null);
        }

        if (colors.Count == 0) {
            return null;
        }

        while (colors.Count < expectedCount) {
            colors.Add(null);
        }

        if (colors.Count > expectedCount) {
            colors.RemoveRange(expectedCount, colors.Count - expectedCount);
        }

        return colors;
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
                return textElement.ReadScalarAsText();
            }

            return titleElement.ReadScalarAsText();
        }

        return null;
    }

    private static bool UsesHorizontalIndexAxis(MarkdownPdfJsonValue root) =>
        TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
        string.Equals(ReadString(options, "indexAxis"), "y", StringComparison.OrdinalIgnoreCase);

    private static bool TryMapChartKind(string? type, bool horizontalIndexAxis, out OfficeChartKind kind) {
        string normalized = NormalizeChartType(type);
        switch (normalized) {
            case "bar":
                kind = horizontalIndexAxis ? OfficeChartKind.BarClustered : OfficeChartKind.ColumnClustered;
                return true;
            case "column":
                kind = OfficeChartKind.ColumnClustered;
                return true;
            case "horizontalbar":
            case "barhorizontal":
                kind = OfficeChartKind.BarClustered;
                return true;
            case "line":
                kind = OfficeChartKind.Line;
                return true;
            case "area":
                kind = OfficeChartKind.Area;
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

    private static double? ReadPositiveDouble(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value) || !TryReadNumber(value, out double number)) {
            return null;
        }

        return number > 0D && !double.IsNaN(number) && !double.IsInfinity(number) ? number : null;
    }

    private static bool TryReadNumber(MarkdownPdfJsonValue element, out double value) => element.TryGetDouble(out value);

    private static bool TryGetProperty(MarkdownPdfJsonValue element, string propertyName, out MarkdownPdfJsonValue value) =>
        element.TryGetProperty(propertyName, out value);

    private sealed class MarkdownChartSeriesValues {
        public MarkdownChartSeriesValues(List<double> values, List<double>? xValues, List<string>? categoryLabels) {
            Values = values;
            XValues = xValues;
            CategoryLabels = categoryLabels;
        }

        public List<double> Values { get; }

        public List<double>? XValues { get; }

        public List<string>? CategoryLabels { get; }
    }
}
