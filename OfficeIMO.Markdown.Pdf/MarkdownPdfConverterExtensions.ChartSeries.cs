using OfficeIMO.Drawing;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// First-party Markdown to PDF conversion helpers.
/// </summary>
public static partial class MarkdownPdfConverterExtensions {
    private static List<OfficeChartSeries> ReadSeries(MarkdownPdfJsonValue dataElement, List<string> labels, OfficeChartKind chartKind, bool defaultConnectLine, bool horizontalBarChart = false) {
        var drafts = new List<MarkdownChartSeriesDraft>();
        var series = new List<OfficeChartSeries>();
        bool captureXValues = chartKind == OfficeChartKind.Scatter;
        bool captureCategoryLabels = !captureXValues;
        if (TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) && datasets.Kind == MarkdownPdfJsonValueKind.Array) {
            int index = 0;
            foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
                if (dataset.Kind != MarkdownPdfJsonValueKind.Object) {
                    continue;
                }

                if (ReadBool(dataset, "hidden") == true) {
                    continue;
                }

                MarkdownChartSeriesValues seriesValues = ReadDataValues(dataset, captureXValues, captureCategoryLabels, horizontalBarChart);
                if (seriesValues.Values.Count == 0) {
                    continue;
                }

                OfficeColor? color = ReadSeriesColor(dataset, chartKind);
                IReadOnlyList<OfficeColor?>? pointColors = ReadPointColors(dataset, "backgroundColor", seriesValues.Values.Count);
                string name = ReadString(dataset, "label") ?? "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
                bool connectLine = ReadBool(dataset, "showLine") ?? defaultConnectLine;
                drafts.Add(new MarkdownChartSeriesDraft(name, seriesValues.Values, seriesValues.XValues, seriesValues.CategoryLabels, color, pointColors, connectLine));
                index++;
            }
        }

        if (drafts.Count == 0 && TryGetProperty(dataElement, "values", out MarkdownPdfJsonValue valuesElement)) {
            MarkdownChartSeriesValues values = ReadNumberArray(valuesElement, captureXValues, captureCategoryLabels, horizontalBarChart);
            if (values.Values.Count > 0) {
                drafts.Add(new MarkdownChartSeriesDraft("Values", values.Values, values.XValues, values.CategoryLabels, null, null, defaultConnectLine));
            }
        }

        if (!captureXValues && labels.Count == 0) {
            MergeCategoryLabels(labels, drafts);
            if (labels.Count > 0) {
                EnsureLabelsCoverValues(labels, GetMaximumDraftValueCount(drafts));
            }
        }

        for (int i = 0; i < drafts.Count; i++) {
            MarkdownChartSeriesDraft draft = drafts[i];
            if (captureXValues) {
                EnsureLabelsCoverValues(labels, draft.Values.Count);
            }

            series.Add(CreateOfficeChartSeries(draft, labels, captureXValues));
        }

        return series;
    }

    private static int GetMaximumDraftValueCount(IReadOnlyList<MarkdownChartSeriesDraft> drafts) {
        int maxValues = 0;
        for (int i = 0; i < drafts.Count; i++) {
            maxValues = Math.Max(maxValues, drafts[i].Values.Count);
        }

        return maxValues;
    }

    private static void MergeCategoryLabels(List<string> labels, IReadOnlyList<MarkdownChartSeriesDraft> drafts) {
        for (int draftIndex = 0; draftIndex < drafts.Count; draftIndex++) {
            IReadOnlyList<string?>? categoryLabels = drafts[draftIndex].CategoryLabels;
            if (!HasCategoryLabels(categoryLabels)) {
                continue;
            }

            for (int labelIndex = 0; labelIndex < categoryLabels!.Count; labelIndex++) {
                string? label = categoryLabels[labelIndex];
                if (string.IsNullOrWhiteSpace(label) || ContainsLabel(labels, label!)) {
                    continue;
                }

                labels.Add(label!);
            }
        }
    }

    private static OfficeChartSeries CreateOfficeChartSeries(MarkdownChartSeriesDraft draft, List<string> labels, bool scatter) {
        var values = new List<double>(draft.Values);
        List<double>? xValues = draft.XValues == null ? null : new List<double>(draft.XValues);
        List<OfficeColor?>? pointColors = draft.PointColors == null ? null : new List<OfficeColor?>(draft.PointColors);

        if (!scatter && labels.Count > 0) {
            if (HasCategoryLabels(draft.CategoryLabels)) {
                values = AlignValuesToCategoryLabels(draft.Values, draft.CategoryLabels!, labels);
                if (pointColors != null) {
                    pointColors = AlignPointColorsToCategoryLabels(pointColors, draft.CategoryLabels!, labels);
                }
            } else {
                NormalizeSeriesLength(values, labels.Count);
                if (xValues != null) {
                    NormalizeSeriesLength(xValues, labels.Count);
                }

                if (pointColors != null) {
                    NormalizePointColorLength(pointColors, labels.Count);
                }
            }
        }

        return new OfficeChartSeries(draft.Name, values, xValues, draft.Color, pointColors, showMarkers: true, showInLegend: true, connectLine: draft.ConnectLine);
    }

    private static bool HasCategoryLabels(IReadOnlyList<string?>? categoryLabels) {
        if (categoryLabels == null) {
            return false;
        }

        for (int i = 0; i < categoryLabels.Count; i++) {
            if (!string.IsNullOrWhiteSpace(categoryLabels[i])) {
                return true;
            }
        }

        return false;
    }

    private static List<double> AlignValuesToCategoryLabels(IReadOnlyList<double> values, IReadOnlyList<string?> categoryLabels, IReadOnlyList<string> labels) {
        var aligned = CreateMissingDoubleList(labels.Count);
        int count = Math.Min(values.Count, categoryLabels.Count);
        for (int i = 0; i < count; i++) {
            string? category = categoryLabels[i];
            if (string.IsNullOrWhiteSpace(category)) {
                if (i < labels.Count) {
                    aligned[i] = values[i];
                }

                continue;
            }

            int labelIndex = IndexOfLabel(labels, category!);
            if (labelIndex >= 0) {
                aligned[labelIndex] = values[i];
            }
        }

        return aligned;
    }

    private static List<OfficeColor?> AlignPointColorsToCategoryLabels(IReadOnlyList<OfficeColor?> values, IReadOnlyList<string?> categoryLabels, IReadOnlyList<string> labels) {
        var aligned = new List<OfficeColor?>(labels.Count);
        for (int i = 0; i < labels.Count; i++) {
            aligned.Add(null);
        }

        int count = Math.Min(values.Count, categoryLabels.Count);
        for (int i = 0; i < count; i++) {
            string? category = categoryLabels[i];
            if (string.IsNullOrWhiteSpace(category)) {
                if (i < labels.Count) {
                    aligned[i] = values[i];
                }

                continue;
            }

            int labelIndex = IndexOfLabel(labels, category!);
            if (labelIndex >= 0) {
                aligned[labelIndex] = values[i];
            }
        }

        return aligned;
    }

    private static List<double> CreateMissingDoubleList(int count) {
        var values = new List<double>(count);
        for (int i = 0; i < count; i++) {
            values.Add(double.NaN);
        }

        return values;
    }

    private static bool ContainsLabel(IReadOnlyList<string> labels, string value) => IndexOfLabel(labels, value) >= 0;

    private static int IndexOfLabel(IReadOnlyList<string> labels, string value) {
        for (int i = 0; i < labels.Count; i++) {
            if (string.Equals(labels[i], value, StringComparison.Ordinal)) {
                return i;
            }
        }

        return -1;
    }

    private static MarkdownChartSeriesValues ReadDataValues(MarkdownPdfJsonValue dataset, bool captureXValues, bool captureCategoryLabels, bool horizontalBarChart) {
        if (!TryGetProperty(dataset, "data", out MarkdownPdfJsonValue data)) {
            return new MarkdownChartSeriesValues(new List<double>(), null, null);
        }

        return ReadNumberArray(data, captureXValues, captureCategoryLabels, horizontalBarChart);
    }

    private static MarkdownChartSeriesValues ReadNumberArray(MarkdownPdfJsonValue element, bool captureXValues, bool captureCategoryLabels, bool horizontalBarChart = false) {
        var values = new List<double>();
        List<double>? xValues = captureXValues ? new List<double>() : null;
        List<string?>? categoryLabels = captureCategoryLabels ? new List<string?>() : null;
        bool hasExplicitXValue = false;
        bool hasExplicitCategoryLabel = false;
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
                    hasExplicitCategoryLabel |= !string.IsNullOrWhiteSpace(categoryLabel);
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
                    MarkdownPdfJsonValue valueElement = horizontalBarChart ? x : y;
                    MarkdownPdfJsonValue categoryElement = horizontalBarChart ? y : x;
                    bool hasValue = horizontalBarChart ? hasX : hasY;
                    bool hasCategory = horizontalBarChart ? hasY : hasX;
                    if (hasValue && TryReadNumber(valueElement, out double parsedY)) {
                        yValue = parsedY;
                    }

                    if (hasCategory) {
                        hasExplicitXValue = true;
                        categoryLabel = categoryElement.ReadScalarAsText();
                        hasExplicitCategoryLabel |= !string.IsNullOrWhiteSpace(categoryLabel);
                        if (TryReadNumber(categoryElement, out double parsedX)) {
                            xValue = parsedX;
                        }
                    }
                }
            }

            if (!hasPoint) {
                continue;
            }

            if (xValues != null && hasExplicitXValue && !IsFiniteChartValue(xValue)) {
                yValue = double.NaN;
            }

            if (xValues != null && hasExplicitXValue && !IsFiniteChartValue(yValue)) {
                xValue = double.NaN;
            }

            values.Add(yValue);
            if (xValues != null) {
                xValues.Add(xValue);
            }

            if (categoryLabels != null) {
                categoryLabels.Add(string.IsNullOrWhiteSpace(categoryLabel) ? null : categoryLabel);
            }
        }

        if (xValues != null && hasExplicitXValue) {
            for (int i = 0; i < values.Count && i < xValues.Count; i++) {
                if (!IsFiniteChartValue(xValues[i])) {
                    values[i] = double.NaN;
                }
            }
        }

        return new MarkdownChartSeriesValues(values, hasExplicitXValue ? xValues : null, hasExplicitCategoryLabel && categoryLabels != null && categoryLabels.Count > 0 ? categoryLabels : null);
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

    private static List<OfficeChartSeries> NormalizeSeriesToLabelCount(IReadOnlyList<OfficeChartSeries> series, int labelCount) {
        if (labelCount <= 0) {
            return new List<OfficeChartSeries>(series);
        }

        var normalized = new List<OfficeChartSeries>(series.Count);
        for (int i = 0; i < series.Count; i++) {
            OfficeChartSeries item = series[i];
            var values = new List<double>(item.Values);
            NormalizeSeriesLength(values, labelCount);

            List<double>? xValues = null;
            if (item.XValues != null) {
                xValues = new List<double>(item.XValues);
                NormalizeSeriesLength(xValues, labelCount);
            }

            List<OfficeColor?>? pointColors = null;
            if (item.PointColors != null) {
                pointColors = new List<OfficeColor?>(item.PointColors);
                NormalizePointColorLength(pointColors, labelCount);
            }

            normalized.Add(new OfficeChartSeries(item.Name, values, xValues, item.Color, pointColors, item.ShowMarkers, item.ShowInLegend, item.ConnectLine));
        }

        return normalized;
    }

    private static void NormalizePointColorLength(List<OfficeColor?> values, int labelCount) {
        while (values.Count < labelCount) {
            values.Add(null);
        }

        if (values.Count > labelCount) {
            values.RemoveRange(labelCount, values.Count - labelCount);
        }
    }

    private static OfficeColor? ReadSeriesColor(MarkdownPdfJsonValue dataset, OfficeChartKind chartKind) {
        if (UsesFilledSeriesColor(chartKind)) {
            return ReadColor(dataset, "backgroundColor") ?? ReadColor(dataset, "borderColor");
        }

        return ReadColor(dataset, "borderColor") ?? ReadColor(dataset, "backgroundColor");
    }

    private static bool UsesFilledSeriesColor(OfficeChartKind chartKind) =>
        chartKind == OfficeChartKind.BarClustered ||
        chartKind == OfficeChartKind.BarStacked ||
        chartKind == OfficeChartKind.BarStacked100 ||
        chartKind == OfficeChartKind.ColumnClustered ||
        chartKind == OfficeChartKind.ColumnStacked ||
        chartKind == OfficeChartKind.ColumnStacked100 ||
        chartKind == OfficeChartKind.Pie ||
        chartKind == OfficeChartKind.Doughnut ||
        chartKind == OfficeChartKind.Area ||
        chartKind == OfficeChartKind.AreaStacked ||
        chartKind == OfficeChartKind.AreaStacked100;

    private static OfficeColor? ReadColor(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue colorElement)) {
            return null;
        }

        if (colorElement.Kind == MarkdownPdfJsonValueKind.String && TryParseChartColor(colorElement.StringValue, out OfficeColor color)) {
            return color;
        }

        if (colorElement.Kind == MarkdownPdfJsonValueKind.Array) {
            foreach (MarkdownPdfJsonValue item in colorElement.ArrayValues) {
                if (item.Kind == MarkdownPdfJsonValueKind.String && TryParseChartColor(item.StringValue, out color)) {
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
            colors.Add(item.Kind == MarkdownPdfJsonValueKind.String && TryParseChartColor(item.StringValue, out OfficeColor color) ? color : null);
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

    private static bool TryParseChartColor(string? value, out OfficeColor color) {
        if (OfficeColor.TryParse(value, out color)) {
            return true;
        }

        return TryParseCssRgbColor(value, out color);
    }

    private static bool TryParseCssRgbColor(string? value, out OfficeColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string trimmed = value!.Trim();
        bool rgba = trimmed.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase);
        bool rgb = trimmed.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase);
        if (!rgba && !rgb) {
            return false;
        }

        int open = trimmed.IndexOf('(');
        int close = trimmed.LastIndexOf(')');
        if (open < 0 || close <= open) {
            return false;
        }

        string[] parts = trimmed.Substring(open + 1, close - open - 1).Split(',');
        if ((!rgba && parts.Length != 3) || (rgba && parts.Length != 4)) {
            return false;
        }

        if (!TryParseCssColorByte(parts[0], out byte red) ||
            !TryParseCssColorByte(parts[1], out byte green) ||
            !TryParseCssColorByte(parts[2], out byte blue)) {
            return false;
        }

        byte alpha = 255;
        if (rgba && !TryParseCssAlphaByte(parts[3], out alpha)) {
            return false;
        }

        color = OfficeColor.FromRgba(red, green, blue, alpha);
        return true;
    }

    private static bool TryParseCssColorByte(string value, out byte component) {
        component = 0;
        string trimmed = value.Trim();
        bool percent = trimmed.EndsWith("%", StringComparison.Ordinal);
        if (percent) {
            trimmed = trimmed.Substring(0, trimmed.Length - 1);
        }

        if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.IsNaN(parsed) ||
            double.IsInfinity(parsed)) {
            return false;
        }

        if (percent) {
            parsed = parsed * 255D / 100D;
        }

        component = (byte)Math.Max(0D, Math.Min(255D, Math.Round(parsed, MidpointRounding.AwayFromZero)));
        return true;
    }

    private static bool TryParseCssAlphaByte(string value, out byte alpha) {
        alpha = 255;
        string trimmed = value.Trim();
        bool percent = trimmed.EndsWith("%", StringComparison.Ordinal);
        if (percent) {
            trimmed = trimmed.Substring(0, trimmed.Length - 1);
        }

        if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
            double.IsNaN(parsed) ||
            double.IsInfinity(parsed)) {
            return false;
        }

        if (percent) {
            parsed /= 100D;
        }

        alpha = (byte)Math.Max(0D, Math.Min(255D, Math.Round(parsed * 255D, MidpointRounding.AwayFromZero)));
        return true;
    }

    private sealed class MarkdownChartSeriesValues {
        public MarkdownChartSeriesValues(List<double> values, List<double>? xValues, List<string?>? categoryLabels) {
            Values = values;
            XValues = xValues;
            CategoryLabels = categoryLabels;
        }

        public List<double> Values { get; }

        public List<double>? XValues { get; }

        public List<string?>? CategoryLabels { get; }
    }

    private sealed class MarkdownChartSeriesDraft {
        public MarkdownChartSeriesDraft(string name, List<double> values, List<double>? xValues, List<string?>? categoryLabels, OfficeColor? color, IReadOnlyList<OfficeColor?>? pointColors, bool connectLine) {
            Name = name;
            Values = values;
            XValues = xValues;
            CategoryLabels = categoryLabels;
            Color = color;
            PointColors = pointColors;
            ConnectLine = connectLine;
        }

        public string Name { get; }

        public List<double> Values { get; }

        public List<double>? XValues { get; }

        public List<string?>? CategoryLabels { get; }

        public OfficeColor? Color { get; }

        public IReadOnlyList<OfficeColor?>? PointColors { get; }

        public bool ConnectLine { get; }
    }
}
