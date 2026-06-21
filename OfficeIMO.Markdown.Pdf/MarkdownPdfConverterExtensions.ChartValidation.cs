using OfficeIMO.Drawing;

namespace OfficeIMO.Markdown.Pdf;

public static partial class MarkdownPdfConverterExtensions {
    private static bool HasReversedScales(MarkdownPdfJsonValue root) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            scales.Kind != MarkdownPdfJsonValueKind.Object) {
            return false;
        }

        foreach (KeyValuePair<string, MarkdownPdfJsonValue> scale in scales.ObjectValues) {
            if (scale.Value.Kind == MarkdownPdfJsonValueKind.Object && ReadBool(scale.Value, "reverse") == true) {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedCategoryScaleType(MarkdownPdfJsonValue root, OfficeChartKind chartKind) {
        if (chartKind == OfficeChartKind.Scatter ||
            chartKind == OfficeChartKind.Pie ||
            chartKind == OfficeChartKind.Doughnut ||
            chartKind == OfficeChartKind.Radar) {
            return false;
        }

        string categoryAxis = IsBarChart(chartKind) ? "y" : "x";
        string? scaleType = ReadScaleType(root, categoryAxis);
        return !string.IsNullOrWhiteSpace(scaleType) &&
               !string.Equals(NormalizeChartType(scaleType), "category", StringComparison.Ordinal);
    }

    private static string? ReadScaleType(MarkdownPdfJsonValue root, string axisName) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            !TryGetProperty(scales, axisName, out MarkdownPdfJsonValue axis) ||
            axis.Kind != MarkdownPdfJsonValueKind.Object ||
            !TryGetProperty(axis, "type", out MarkdownPdfJsonValue typeElement) ||
            typeElement.Kind == MarkdownPdfJsonValueKind.Null) {
            return null;
        }

        return typeElement.ReadScalarAsText();
    }

    private static bool HasUnsupportedSpanGaps(MarkdownPdfJsonValue root, MarkdownPdfJsonValue dataElement, string rootType) {
        string normalized = NormalizeChartType(rootType);
        if (!IsLineFamilyType(rootType) && !string.Equals(normalized, "scatter", StringComparison.Ordinal)) {
            return false;
        }

        return HasChartJsSpanGaps(root, dataElement) && HasMissingChartDataPoint(dataElement);
    }

    private static bool HasChartJsSpanGaps(MarkdownPdfJsonValue root, MarkdownPdfJsonValue dataElement) {
        if (ReadBool(root, "spanGaps") == true) {
            return true;
        }

        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options) && ReadBool(options, "spanGaps") == true) {
            return true;
        }

        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
            if (dataset.Kind == MarkdownPdfJsonValueKind.Object &&
                ReadBool(dataset, "hidden") != true &&
                ReadBool(dataset, "spanGaps") == true) {
                return true;
            }
        }

        return false;
    }

    private static bool TryReadChartJsFillDirective(MarkdownPdfJsonValue dataset, out bool fill) {
        if (!TryGetProperty(dataset, "fill", out MarkdownPdfJsonValue fillElement) || fillElement.Kind == MarkdownPdfJsonValueKind.Null) {
            fill = false;
            return false;
        }

        fill = ReadChartJsFill(dataset);
        return true;
    }

    private static bool TryResolveRadarFill(MarkdownPdfJsonValue dataElement, out bool fillRadarSeries) {
        fillRadarSeries = true;
        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return true;
        }

        bool hasFilled = false;
        bool hasUnfilled = false;
        foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            if (!TryReadChartJsFillDirective(dataset, out bool fill)) {
                continue;
            }

            if (fill) {
                hasFilled = true;
            } else {
                hasUnfilled = true;
            }
        }

        if (hasFilled && hasUnfilled) {
            return false;
        }

        fillRadarSeries = !hasUnfilled;
        return true;
    }

    private static bool HasUnsupportedFloatingBarTuples(MarkdownPdfJsonValue dataElement, OfficeChartKind chartKind) {
        if (!IsBarChart(chartKind) && !IsColumnChart(chartKind)) {
            return false;
        }

        return HasNumericTupleData(dataElement);
    }

    private static bool HasNumericTupleData(MarkdownPdfJsonValue dataElement) {
        if (TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) && datasets.Kind == MarkdownPdfJsonValueKind.Array) {
            foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
                if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                    continue;
                }

                if (TryGetProperty(dataset, "data", out MarkdownPdfJsonValue data) && HasNumericTupleArray(data)) {
                    return true;
                }
            }
        }

        return TryGetProperty(dataElement, "values", out MarkdownPdfJsonValue values) && HasNumericTupleArray(values);
    }

    private static bool HasNumericTupleArray(MarkdownPdfJsonValue element) {
        if (element.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        foreach (MarkdownPdfJsonValue item in element.ArrayValues) {
            if (item.Kind == MarkdownPdfJsonValueKind.Array &&
                item.ArrayValues.Count >= 2 &&
                TryReadNumber(item.ArrayValues[0], out _) &&
                TryReadNumber(item.ArrayValues[1], out _)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasMissingChartDataPoint(MarkdownPdfJsonValue dataElement) {
        if (TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) && datasets.Kind == MarkdownPdfJsonValueKind.Array) {
            foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
                if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                    continue;
                }

                if (TryGetProperty(dataset, "data", out MarkdownPdfJsonValue data) && HasMissingChartDataPointArray(data)) {
                    return true;
                }
            }
        }

        return TryGetProperty(dataElement, "values", out MarkdownPdfJsonValue values) && HasMissingChartDataPointArray(values);
    }

    private static bool HasMissingChartDataPointArray(MarkdownPdfJsonValue element) {
        if (element.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        foreach (MarkdownPdfJsonValue item in element.ArrayValues) {
            if (item.Kind == MarkdownPdfJsonValueKind.Null) {
                return true;
            }

            if (item.Kind == MarkdownPdfJsonValueKind.Array && ContainsMissingChartDataValue(item.ArrayValues)) {
                return true;
            }

            if (item.Kind == MarkdownPdfJsonValueKind.Object &&
                ((TryGetProperty(item, "x", out MarkdownPdfJsonValue x) && x.Kind == MarkdownPdfJsonValueKind.Null) ||
                 (TryGetProperty(item, "y", out MarkdownPdfJsonValue y) && y.Kind == MarkdownPdfJsonValueKind.Null))) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsMissingChartDataValue(IReadOnlyList<MarkdownPdfJsonValue> values) {
        for (int i = 0; i < values.Count; i++) {
            if (values[i].Kind == MarkdownPdfJsonValueKind.Null) {
                return true;
            }
        }

        return false;
    }
}
