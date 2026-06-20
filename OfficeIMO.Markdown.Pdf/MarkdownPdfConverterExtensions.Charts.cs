using System.Text.Json;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static bool TryRenderChartFencedBlock(PdfCore.PdfDocument pdf, SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, MarkdownPdfVisualTheme visualTheme) {
        if (!IsChartSemanticFence(semantic)) {
            return false;
        }

        if (!TryCreateChartSnapshot(semantic, options, out OfficeChartSnapshot? snapshot)) {
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

    private static bool TryCreateChartSnapshot(SemanticFencedBlock semantic, MarkdownPdfSaveOptions options, out OfficeChartSnapshot? snapshot) {
        snapshot = null;
        if (string.IsNullOrWhiteSpace(semantic.Content)) {
            return false;
        }

        try {
            using JsonDocument document = JsonDocument.Parse(semantic.Content);
            JsonElement root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) {
                return false;
            }

            string type = ReadString(root, "type") ?? "bar";
            if (!TryMapChartKind(type, out OfficeChartKind chartKind)) {
                return false;
            }

            JsonElement dataElement = TryGetProperty(root, "data", out JsonElement data)
                ? data
                : root;
            List<string> labels = ReadLabels(dataElement);
            List<OfficeChartSeries> series = ReadSeries(dataElement, labels.Count, chartKind);
            if (series.Count == 0) {
                return false;
            }

            if (labels.Count == 0) {
                int maxValues = series.Max(item => item.Values.Count);
                for (int i = 0; i < maxValues; i++) {
                    labels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }

            if (labels.Count == 0) {
                return false;
            }

            string? title = ReadChartTitle(root);
            double width = ReadPositiveDouble(root, "width") ?? options.DefaultImageWidth;
            double height = ReadPositiveDouble(root, "height") ?? options.DefaultImageHeight;
            width = Math.Max(240D, Math.Min(520D, width));
            height = Math.Max(150D, Math.Min(320D, height));
            FitChartToPageWidth(options, ref width, ref height);

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
        } catch (JsonException) {
            AddWarning(options, "UnsupportedChartFence", semantic.Language, "The Markdown chart fence is not valid JSON and is rendered as a semantic code panel.");
            return false;
        } catch (ArgumentException ex) {
            AddWarning(options, "UnsupportedChartFence", semantic.Language, "The Markdown chart fence could not be rendered as a chart: " + ex.Message);
            return false;
        }
    }

    private static void FitChartToPageWidth(MarkdownPdfSaveOptions options, ref double width, ref double height) {
        PdfCore.PdfOptions pdfOptions = options.PdfOptions ?? new PdfCore.PdfOptions();
        double availableWidth = pdfOptions.PageWidth - pdfOptions.MarginLeft - pdfOptions.MarginRight;
        if (availableWidth <= 0 || double.IsNaN(availableWidth) || double.IsInfinity(availableWidth)) {
            return;
        }

        double maxWidth = Math.Max(120D, availableWidth);
        if (width <= maxWidth) {
            return;
        }

        double scale = maxWidth / width;
        width = maxWidth;
        height = Math.Max(120D, height * scale);
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

    private static List<string> ReadLabels(JsonElement dataElement) {
        var labels = new List<string>();
        if (!TryGetProperty(dataElement, "labels", out JsonElement labelElement) || labelElement.ValueKind != JsonValueKind.Array) {
            return labels;
        }

        foreach (JsonElement item in labelElement.EnumerateArray()) {
            string? label = ReadJsonScalarAsText(item);
            labels.Add(string.IsNullOrWhiteSpace(label) ? string.Empty : label!);
        }

        return labels;
    }

    private static List<OfficeChartSeries> ReadSeries(JsonElement dataElement, int labelCount, OfficeChartKind chartKind) {
        var series = new List<OfficeChartSeries>();
        if (TryGetProperty(dataElement, "datasets", out JsonElement datasets) && datasets.ValueKind == JsonValueKind.Array) {
            int index = 0;
            foreach (JsonElement dataset in datasets.EnumerateArray()) {
                if (dataset.ValueKind != JsonValueKind.Object) {
                    continue;
                }

                List<double> values = ReadDataValues(dataset);
                if (values.Count == 0) {
                    continue;
                }

                NormalizeSeriesLength(values, labelCount);
                OfficeColor? color = ReadColor(dataset, "borderColor") ?? ReadColor(dataset, "backgroundColor");
                IReadOnlyList<OfficeColor?>? pointColors = ReadPointColors(dataset, "backgroundColor", values.Count);
                string name = ReadString(dataset, "label") ?? "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
                series.Add(new OfficeChartSeries(name, values, xValues: null, color, pointColors));
                index++;
            }
        }

        if (series.Count == 0 && TryGetProperty(dataElement, "values", out JsonElement valuesElement)) {
            List<double> values = ReadNumberArray(valuesElement);
            if (values.Count > 0) {
                NormalizeSeriesLength(values, labelCount);
                series.Add(new OfficeChartSeries("Values", values));
            }
        }

        if ((chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut) && series.Count > 1) {
            return new List<OfficeChartSeries> { series[0] };
        }

        return series;
    }

    private static List<double> ReadDataValues(JsonElement dataset) {
        if (!TryGetProperty(dataset, "data", out JsonElement data)) {
            return new List<double>();
        }

        return ReadNumberArray(data);
    }

    private static List<double> ReadNumberArray(JsonElement element) {
        var values = new List<double>();
        if (element.ValueKind != JsonValueKind.Array) {
            return values;
        }

        foreach (JsonElement item in element.EnumerateArray()) {
            if (TryReadNumber(item, out double value)) {
                values.Add(value);
            } else if (item.ValueKind == JsonValueKind.Object && TryGetProperty(item, "y", out JsonElement y) && TryReadNumber(y, out double yValue)) {
                values.Add(yValue);
            }
        }

        return values;
    }

    private static void NormalizeSeriesLength(List<double> values, int labelCount) {
        if (labelCount <= 0) {
            return;
        }

        while (values.Count < labelCount) {
            values.Add(0D);
        }

        if (values.Count > labelCount) {
            values.RemoveRange(labelCount, values.Count - labelCount);
        }
    }

    private static OfficeColor? ReadColor(JsonElement element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out JsonElement colorElement)) {
            return null;
        }

        if (colorElement.ValueKind == JsonValueKind.String && OfficeColor.TryParse(colorElement.GetString(), out OfficeColor color)) {
            return color;
        }

        if (colorElement.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement item in colorElement.EnumerateArray()) {
                if (item.ValueKind == JsonValueKind.String && OfficeColor.TryParse(item.GetString(), out color)) {
                    return color;
                }
            }
        }

        return null;
    }

    private static IReadOnlyList<OfficeColor?>? ReadPointColors(JsonElement element, string propertyName, int expectedCount) {
        if (!TryGetProperty(element, propertyName, out JsonElement colorElement) || colorElement.ValueKind != JsonValueKind.Array) {
            return null;
        }

        var colors = new List<OfficeColor?>();
        foreach (JsonElement item in colorElement.EnumerateArray()) {
            colors.Add(item.ValueKind == JsonValueKind.String && OfficeColor.TryParse(item.GetString(), out OfficeColor color) ? color : null);
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

    private static string? ReadChartTitle(JsonElement root) {
        string? title = ReadString(root, "title");
        if (!string.IsNullOrWhiteSpace(title)) {
            return title;
        }

        if (TryGetProperty(root, "options", out JsonElement options) &&
            TryGetProperty(options, "plugins", out JsonElement plugins) &&
            TryGetProperty(plugins, "title", out JsonElement titleElement)) {
            if (titleElement.ValueKind == JsonValueKind.Object && TryGetProperty(titleElement, "text", out JsonElement textElement)) {
                return ReadJsonScalarAsText(textElement);
            }

            return ReadJsonScalarAsText(titleElement);
        }

        return null;
    }

    private static bool TryMapChartKind(string? type, out OfficeChartKind kind) {
        string normalized = NormalizeChartType(type);
        switch (normalized) {
            case "bar":
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

    private static string? ReadString(JsonElement element, string propertyName) =>
        TryGetProperty(element, propertyName, out JsonElement value) ? ReadJsonScalarAsText(value) : null;

    private static double? ReadPositiveDouble(JsonElement element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out JsonElement value) || !TryReadNumber(value, out double number)) {
            return null;
        }

        return number > 0D && !double.IsNaN(number) && !double.IsInfinity(number) ? number : null;
    }

    private static bool TryReadNumber(JsonElement element, out double value) {
        value = 0D;
        if (element.ValueKind == JsonValueKind.Number) {
            return element.TryGetDouble(out value);
        }

        if (element.ValueKind == JsonValueKind.String) {
            return double.TryParse(element.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        return false;
    }

    private static string? ReadJsonScalarAsText(JsonElement element) {
        switch (element.ValueKind) {
            case JsonValueKind.String:
                return element.GetString();
            case JsonValueKind.Number:
                return element.ToString();
            case JsonValueKind.True:
                return "true";
            case JsonValueKind.False:
                return "false";
            case JsonValueKind.Array:
                var parts = new List<string>();
                foreach (JsonElement item in element.EnumerateArray()) {
                    string? text = ReadJsonScalarAsText(item);
                    if (!string.IsNullOrWhiteSpace(text)) {
                        parts.Add(text!);
                    }
                }

                return parts.Count == 0 ? null : string.Join(" ", parts);
            default:
                return null;
        }
    }

    private static bool TryGetProperty(JsonElement element, string propertyName, out JsonElement value) {
        if (element.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase)) {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }
}
