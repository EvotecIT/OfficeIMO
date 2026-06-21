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

            MarkdownPdfJsonValue dataElement = TryGetProperty(root, "data", out MarkdownPdfJsonValue data)
                ? data
                : root;
            string type = ReadString(root, "type") ?? "bar";
            if (HasMixedVisibleDatasetTypes(dataElement, type)) {
                warningMessage = "The Markdown Chart.js fence uses mixed per-dataset chart types that cannot be rendered as one native Office chart and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedSecondaryAxes(dataElement)) {
                warningMessage = "The Markdown Chart.js fence uses secondary dataset axes that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedScaleTypes(root)) {
                warningMessage = "The Markdown Chart.js fence uses non-linear or otherwise unsupported scale types that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (HasReversedScales(root)) {
                warningMessage = "The Markdown Chart.js fence uses reversed scales that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedDoughnutCutout(root, type)) {
                warningMessage = "The Markdown Chart.js doughnut fence uses a custom cutout that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            bool horizontalIndexAxis = UsesHorizontalIndexAxis(root) || UsesHorizontalBarType(type);
            bool stackedScale = UsesStackedScale(root);
            if (stackedScale && HasUnsupportedChartJsStackGroups(dataElement)) {
                warningMessage = "The Markdown Chart.js fence uses separate stack groups that cannot be rendered as one native Office stacked chart and is rendered as a semantic code panel.";
                return false;
            }

            if (!TryResolveFilledLineChart(dataElement, type, out bool filledLineChart)) {
                warningMessage = "The Markdown Chart.js fence mixes filled and unfilled visible line datasets that cannot be rendered as one native Office chart and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedSpanGaps(root, dataElement, type)) {
                warningMessage = "The Markdown Chart.js fence uses spanGaps across missing line points that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedLineInterpolation(root, dataElement, type)) {
                warningMessage = "The Markdown Chart.js fence uses stepped or curved line interpolation that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (!TryMapChartKind(type, horizontalIndexAxis, stackedScale, filledLineChart, out OfficeChartKind chartKind)) {
                warningMessage = "The Markdown chart fence uses an unsupported chart type and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedCategoryScaleType(root, chartKind)) {
                warningMessage = "The Markdown Chart.js fence uses a linear or otherwise non-category scale for the native chart category axis and is rendered as a semantic code panel.";
                return false;
            }

            if (chartKind == OfficeChartKind.Radar && HasUnsupportedRadarScaleVisibility(root)) {
                warningMessage = "The Markdown Chart.js radar fence uses radial scale visibility settings that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (HasUnsupportedFloatingBarTuples(dataElement, chartKind)) {
                warningMessage = "The Markdown Chart.js bar fence uses floating bar tuples that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            bool fillRadarSeries = true;
            if (chartKind == OfficeChartKind.Radar && !TryResolveRadarFill(dataElement, out fillRadarSeries)) {
                warningMessage = "The Markdown Chart.js radar fence mixes filled and unfilled visible datasets that cannot be rendered as one native PDF radar chart and is rendered as a semantic code panel.";
                return false;
            }

            if (HasExplicitChartScaleBounds(root)) {
                warningMessage = "The Markdown Chart.js fence uses explicit scale min/max bounds that cannot be preserved by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
            }

            if (!TryReadChartLegendPosition(root, chartKind, out OfficeChartLegendPosition legendPosition, out warningMessage)) {
                return false;
            }

            List<string> labels = ReadLabels(dataElement);
            bool defaultConnectLine = chartKind != OfficeChartKind.Scatter || ReadChartShowLine(root) == true;
            bool defaultShowMarkers = ReadDefaultChartPointMarkers(root);
            List<OfficeChartSeries> series = ReadSeries(dataElement, labels, chartKind, defaultConnectLine, defaultShowMarkers, horizontalIndexAxis);
            if (series.Count == 0) {
                warningMessage = "The Markdown chart fence does not contain renderable chart series and is rendered as a semantic code panel.";
                return false;
            }

            if (HasTranslucentChartColors(series)) {
                warningMessage = "The Markdown Chart.js fence uses translucent colors that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
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

            if (chartKind == OfficeChartKind.Pie && series.Count > 1) {
                warningMessage = "The Markdown pie chart fence contains multiple visible datasets that cannot be rendered without dropping data and is rendered as a semantic code panel.";
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
                CreateMarkdownChartStyle(root, chartKind),
                CreateMarkdownChartLayout(root, chartKind, series, legendPosition, fillRadarSeries));
            return true;
        } catch (FormatException) {
            warningMessage = "The Markdown chart fence is not valid JSON and is rendered as a semantic code panel.";
            return false;
        } catch (ArgumentException ex) {
            warningMessage = "The Markdown chart fence could not be rendered as a chart: " + ex.Message;
            return false;
        }
    }

    private static bool HasMixedVisibleDatasetTypes(MarkdownPdfJsonValue dataElement, string rootType) {
        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        string normalizedRootType = NormalizeChartType(rootType);
        for (int i = 0; i < datasets.ArrayValues.Count; i++) {
            MarkdownPdfJsonValue dataset = datasets.ArrayValues[i];
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            string? datasetType = ReadString(dataset, "type");
            if (string.IsNullOrWhiteSpace(datasetType)) {
                continue;
            }

            if (!AreEquivalentChartTypes(normalizedRootType, NormalizeChartType(datasetType))) {
                return true;
            }
        }

        return false;
    }

    private static bool AreEquivalentChartTypes(string first, string second) {
        if (string.Equals(first, second, StringComparison.Ordinal)) {
            return true;
        }

        return (string.Equals(first, "doughnut", StringComparison.Ordinal) && string.Equals(second, "donut", StringComparison.Ordinal)) ||
               (string.Equals(first, "donut", StringComparison.Ordinal) && string.Equals(second, "doughnut", StringComparison.Ordinal));
    }

    private static bool HasUnsupportedSecondaryAxes(MarkdownPdfJsonValue dataElement) {
        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        for (int i = 0; i < datasets.ArrayValues.Count; i++) {
            MarkdownPdfJsonValue dataset = datasets.ArrayValues[i];
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            if (HasUnsupportedAxisId(dataset, "xAxisID", "x") ||
                HasUnsupportedAxisId(dataset, "yAxisID", "y")) {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedAxisId(MarkdownPdfJsonValue dataset, string propertyName, string defaultAxisId) {
        string? axisId = ReadString(dataset, propertyName);
        return !string.IsNullOrWhiteSpace(axisId) &&
               !string.Equals(axisId!.Trim(), defaultAxisId, StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasUnsupportedScaleTypes(MarkdownPdfJsonValue root) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            scales.Kind != MarkdownPdfJsonValueKind.Object) {
            return false;
        }

        foreach (KeyValuePair<string, MarkdownPdfJsonValue> scale in scales.ObjectValues) {
            if (scale.Value.Kind != MarkdownPdfJsonValueKind.Object ||
                !TryGetProperty(scale.Value, "type", out MarkdownPdfJsonValue typeElement) ||
                typeElement.Kind == MarkdownPdfJsonValueKind.Null) {
                continue;
            }

            string? type = typeElement.ReadScalarAsText();
            if (string.IsNullOrWhiteSpace(type)) {
                continue;
            }

            string normalized = NormalizeChartType(type);
            if (string.Equals(normalized, "linear", StringComparison.Ordinal) ||
                string.Equals(normalized, "category", StringComparison.Ordinal) ||
                string.Equals(normalized, "radiallinear", StringComparison.Ordinal)) {
                continue;
            }

            return true;
        }

        return false;
    }

    private static bool HasUnsupportedDoughnutCutout(MarkdownPdfJsonValue root, string type) {
        string normalized = NormalizeChartType(type);
        if (!string.Equals(normalized, "doughnut", StringComparison.Ordinal) &&
            !string.Equals(normalized, "donut", StringComparison.Ordinal)) {
            return false;
        }

        return TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
               ((TryGetProperty(options, "cutout", out MarkdownPdfJsonValue cutout) && cutout.Kind != MarkdownPdfJsonValueKind.Null) ||
                (TryGetProperty(options, "cutoutPercentage", out MarkdownPdfJsonValue cutoutPercentage) && cutoutPercentage.Kind != MarkdownPdfJsonValueKind.Null));
    }

    private static bool HasUnsupportedChartJsStackGroups(MarkdownPdfJsonValue dataElement) {
        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        var stackGroups = new HashSet<string>(StringComparer.Ordinal);
        bool hasExplicitStack = false;
        int visibleDatasetCount = 0;
        for (int i = 0; i < datasets.ArrayValues.Count; i++) {
            MarkdownPdfJsonValue dataset = datasets.ArrayValues[i];
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            visibleDatasetCount++;
            string? stack = ReadString(dataset, "stack");
            hasExplicitStack |= !string.IsNullOrWhiteSpace(stack);
            stackGroups.Add(string.IsNullOrWhiteSpace(stack) ? string.Empty : stack!.Trim());
        }

        return hasExplicitStack && visibleDatasetCount > 1 && stackGroups.Count > 1;
    }

    private static bool TryResolveFilledLineChart(MarkdownPdfJsonValue dataElement, string rootType, out bool filledLineChart) {
        filledLineChart = false;
        if (!string.Equals(NormalizeChartType(rootType), "line", StringComparison.Ordinal)) {
            return true;
        }

        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return true;
        }

        bool hasFilled = false;
        bool hasUnfilled = false;
        foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            if (ReadChartJsFill(dataset)) {
                hasFilled = true;
            } else {
                hasUnfilled = true;
            }
        }

        if (hasFilled && hasUnfilled) {
            return false;
        }

        filledLineChart = hasFilled;
        return true;
    }

    private static bool HasUnsupportedLineInterpolation(MarkdownPdfJsonValue root, MarkdownPdfJsonValue dataElement, string rootType) {
        if (!IsLineFamilyType(rootType)) {
            return false;
        }

        if (HasUnsupportedLineInterpolation(root)) {
            return true;
        }

        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options)) {
            if (HasUnsupportedLineInterpolation(options)) {
                return true;
            }

            if (TryGetProperty(options, "elements", out MarkdownPdfJsonValue elements) &&
                TryGetProperty(elements, "line", out MarkdownPdfJsonValue line) &&
                line.Kind == MarkdownPdfJsonValueKind.Object &&
                HasUnsupportedLineInterpolation(line)) {
                return true;
            }
        }

        if (!TryGetProperty(dataElement, "datasets", out MarkdownPdfJsonValue datasets) || datasets.Kind != MarkdownPdfJsonValueKind.Array) {
            return false;
        }

        foreach (MarkdownPdfJsonValue dataset in datasets.ArrayValues) {
            if (dataset.Kind != MarkdownPdfJsonValueKind.Object || ReadBool(dataset, "hidden") == true) {
                continue;
            }

            if (HasUnsupportedLineInterpolation(dataset)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasUnsupportedLineInterpolation(MarkdownPdfJsonValue element) =>
        IsSteppedLine(element) || HasNonZeroTension(element);

    private static bool IsSteppedLine(MarkdownPdfJsonValue element) {
        if (!TryGetProperty(element, "stepped", out MarkdownPdfJsonValue stepped) || stepped.Kind == MarkdownPdfJsonValueKind.Null) {
            return false;
        }

        if (stepped.Kind == MarkdownPdfJsonValueKind.False) {
            return false;
        }

        string? value = stepped.ReadScalarAsText();
        return !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasNonZeroTension(MarkdownPdfJsonValue element) =>
        TryReadFiniteNumber(element, "tension", out double tension) && Math.Abs(tension) > double.Epsilon;

    private static bool IsLineFamilyType(string? type) {
        string normalized = NormalizeChartType(type);
        return string.Equals(normalized, "line", StringComparison.Ordinal) ||
               string.Equals(normalized, "area", StringComparison.Ordinal);
    }

    private static bool ReadChartJsFill(MarkdownPdfJsonValue dataset) {
        if (!TryGetProperty(dataset, "fill", out MarkdownPdfJsonValue fill)) {
            return false;
        }

        switch (fill.Kind) {
            case MarkdownPdfJsonValueKind.False:
            case MarkdownPdfJsonValueKind.Null:
                return false;
            case MarkdownPdfJsonValueKind.True:
                return true;
            case MarkdownPdfJsonValueKind.String:
                return !string.IsNullOrWhiteSpace(fill.StringValue) &&
                       !string.Equals(fill.StringValue, "false", StringComparison.OrdinalIgnoreCase);
            default:
                return true;
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

        MarkdownVisualTheme? sharedTheme = options.ThemeSnapshot;
        if (sharedTheme != null) {
            return MarkdownPdfVisualTheme.FromMarkdownTheme(sharedTheme);
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

    private static OfficeChartStyle CreateMarkdownChartStyle(MarkdownPdfJsonValue root, OfficeChartKind chartKind) {
        bool showGridLines = ReadRenderedScaleGridDisplay(root, chartKind);
        return showGridLines ? OfficeChartStyle.Default : new OfficeChartStyle(showGridLines: false);
    }

    private static OfficeChartLayout CreateMarkdownChartLayout(MarkdownPdfJsonValue root, OfficeChartKind chartKind, IReadOnlyList<OfficeChartSeries> series, OfficeChartLegendPosition legendPosition, bool fillRadarSeries) {
        bool pie = chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut;
        bool showLegend = ReadChartLegendDisplay(root) != false;
        bool showRequestedDataLabels = ReadChartDataLabelsDisplay(root) == true;
        bool showPieDataLabels = pie && showRequestedDataLabels;
        bool showCartesianDataLabels = !pie && showRequestedDataLabels && IsCartesianChart(chartKind);
        bool connectScatterPoints = chartKind == OfficeChartKind.Scatter && HasScatterSeriesLine(series);
        ChartScaleVisibility xScale = ReadChartScaleVisibility(root, "x");
        ChartScaleVisibility yScale = ReadChartScaleVisibility(root, "y");
        bool barChart = IsBarChart(chartKind);
        ChartScaleVisibility categoryScale = barChart ? yScale : xScale;
        ChartScaleVisibility valueScale = barChart ? xScale : yScale;
        string? categoryAxisTitle = barChart ? ReadChartScaleTitle(root, "y") : ReadChartScaleTitle(root, "x");
        string? valueAxisTitle = barChart ? ReadChartScaleTitle(root, "x") : ReadChartScaleTitle(root, "y");
        return new OfficeChartLayout(
            showLegend: showLegend,
            legendPosition: legendPosition,
            showDataLabels: showPieDataLabels || showCartesianDataLabels,
            showDataLabelValues: showCartesianDataLabels,
            showDataLabelPercentages: showPieDataLabels,
            showDataLabelCategoryNames: showPieDataLabels,
            dataLabelFontSize: 7D,
            maximumCategoryAxisLabels: 8,
            maximumHorizontalCategoryAxisLabels: 8,
            showMarkers: (IsLineChart(chartKind) || chartKind == OfficeChartKind.Scatter || chartKind == OfficeChartKind.Radar) && HasVisibleMarkers(series),
            connectScatterPoints: connectScatterPoints,
            fillRadarSeries: fillRadarSeries,
            categoryAxisTitle: categoryAxisTitle,
            valueAxisTitle: valueAxisTitle,
            showCategoryAxis: categoryScale.ShowAxis,
            showValueAxis: valueScale.ShowAxis,
            showCategoryAxisLine: categoryScale.ShowLine,
            showValueAxisLine: valueScale.ShowLine,
            showCategoryAxisLabels: categoryScale.ShowLabels,
            showValueAxisLabels: valueScale.ShowLabels);
    }

    private static bool TryReadChartLegendPosition(MarkdownPdfJsonValue root, OfficeChartKind chartKind, out OfficeChartLegendPosition position, out string? warningMessage) {
        position = chartKind == OfficeChartKind.Pie || chartKind == OfficeChartKind.Doughnut
            ? OfficeChartLegendPosition.Right
            : OfficeChartLegendPosition.Bottom;
        warningMessage = null;
        if (ReadChartLegendDisplay(root) == false) {
            return true;
        }

        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "plugins", out MarkdownPdfJsonValue plugins) ||
            !TryGetProperty(plugins, "legend", out MarkdownPdfJsonValue legend) ||
            legend.Kind != MarkdownPdfJsonValueKind.Object ||
            !TryGetProperty(legend, "position", out MarkdownPdfJsonValue positionElement)) {
            return true;
        }

        string? value = positionElement.ReadScalarAsText();
        if (string.IsNullOrWhiteSpace(value)) {
            return true;
        }

        switch (value!.Trim().ToLowerInvariant()) {
            case "left":
                position = OfficeChartLegendPosition.Left;
                return true;
            case "right":
                position = OfficeChartLegendPosition.Right;
                return true;
            case "top":
                position = OfficeChartLegendPosition.Top;
                return true;
            case "bottom":
                position = OfficeChartLegendPosition.Bottom;
                return true;
            default:
                warningMessage = "The Markdown Chart.js fence uses a legend position that cannot be represented by the native PDF chart renderer and is rendered as a semantic code panel.";
                return false;
        }
    }

    private static bool? ReadChartDataLabelsDisplay(MarkdownPdfJsonValue root) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "plugins", out MarkdownPdfJsonValue plugins) ||
            !TryGetProperty(plugins, "datalabels", out MarkdownPdfJsonValue dataLabels)) {
            return null;
        }

        switch (dataLabels.Kind) {
            case MarkdownPdfJsonValueKind.True:
                return true;
            case MarkdownPdfJsonValueKind.False:
            case MarkdownPdfJsonValueKind.Null:
                return false;
            case MarkdownPdfJsonValueKind.Object:
                bool? display = ReadBool(dataLabels, "display");
                return display ?? true;
            default:
                return null;
        }
    }

    private static string? ReadChartScaleTitle(MarkdownPdfJsonValue root, string axisName) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            !TryGetProperty(scales, axisName, out MarkdownPdfJsonValue axis) ||
            axis.Kind != MarkdownPdfJsonValueKind.Object ||
            !TryGetProperty(axis, "title", out MarkdownPdfJsonValue title) ||
            title.Kind != MarkdownPdfJsonValueKind.Object ||
            !TryGetProperty(title, "text", out MarkdownPdfJsonValue text)) {
            return null;
        }

        if (ReadBool(title, "display") != true) {
            return null;
        }

        string? value = text.ReadScalarAsText();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static bool HasExplicitChartScaleBounds(MarkdownPdfJsonValue root) =>
        TryGetProperty(root, "options", out MarkdownPdfJsonValue options) &&
        TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) &&
        (HasExplicitScaleBounds(scales, "x") || HasExplicitScaleBounds(scales, "y") || HasExplicitScaleBounds(scales, "r"));

    private static bool HasExplicitScaleBounds(MarkdownPdfJsonValue scales, string axisName) =>
        TryGetProperty(scales, axisName, out MarkdownPdfJsonValue axis) &&
        axis.Kind == MarkdownPdfJsonValueKind.Object &&
        ((TryGetProperty(axis, "min", out MarkdownPdfJsonValue min) && min.Kind != MarkdownPdfJsonValueKind.Null) ||
         (TryGetProperty(axis, "max", out MarkdownPdfJsonValue max) && max.Kind != MarkdownPdfJsonValueKind.Null));

    private static bool HasUnsupportedRadarScaleVisibility(MarkdownPdfJsonValue root) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            !TryGetProperty(scales, "r", out MarkdownPdfJsonValue radialScale) ||
            radialScale.Kind != MarkdownPdfJsonValueKind.Object) {
            return false;
        }

        ChartScaleVisibility visibility = ReadChartScaleVisibility(root, "r");
        return !visibility.ShowAxis ||
               !visibility.ShowLine ||
               !visibility.ShowLabels ||
               !visibility.ShowGrid ||
               ReadNestedDisplay(radialScale, "pointLabels") == false ||
               ReadNestedDisplay(radialScale, "angleLines") == false;
    }

    private static bool ReadRenderedScaleGridDisplay(MarkdownPdfJsonValue root, OfficeChartKind chartKind) {
        ChartScaleVisibility xScale = ReadChartScaleVisibility(root, "x");
        ChartScaleVisibility yScale = ReadChartScaleVisibility(root, "y");
        return IsBarChart(chartKind) ? xScale.ShowGrid : yScale.ShowGrid;
    }

    private static ChartScaleVisibility ReadChartScaleVisibility(MarkdownPdfJsonValue root, string axisName) {
        if (!TryGetProperty(root, "options", out MarkdownPdfJsonValue options) ||
            !TryGetProperty(options, "scales", out MarkdownPdfJsonValue scales) ||
            !TryGetProperty(scales, axisName, out MarkdownPdfJsonValue axis) ||
            axis.Kind != MarkdownPdfJsonValueKind.Object) {
            return ChartScaleVisibility.Default;
        }

        bool showAxis = ReadBool(axis, "display") != false;
        bool showLine = showAxis && ReadNestedDisplay(axis, "border") != false;
        bool showLabels = showAxis && ReadNestedDisplay(axis, "ticks") != false;
        bool showGrid = showAxis && ReadNestedDisplay(axis, "grid") != false;
        return new ChartScaleVisibility(showAxis, showLine, showLabels, showGrid);
    }

    private static bool? ReadNestedDisplay(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue nested) || nested.Kind != MarkdownPdfJsonValueKind.Object) {
            return null;
        }

        return ReadBool(nested, "display");
    }

    private static bool HasScatterSeriesLine(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            if (series[seriesIndex].ConnectLine) {
                return true;
            }
        }

        return false;
    }

    private static bool HasVisibleMarkers(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            if (series[seriesIndex].ShowMarkers) {
                return true;
            }
        }

        return false;
    }

    private static bool HasTranslucentChartColors(IReadOnlyList<OfficeChartSeries> series) {
        for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++) {
            OfficeChartSeries item = series[seriesIndex];
            if (IsTranslucentColor(item.Color) || HasTranslucentPointColor(item.PointColors)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasTranslucentPointColor(IReadOnlyList<OfficeColor?>? colors) {
        if (colors == null) {
            return false;
        }

        for (int i = 0; i < colors.Count; i++) {
            if (IsTranslucentColor(colors[i])) {
                return true;
            }
        }

        return false;
    }

    private static bool IsTranslucentColor(OfficeColor? color) =>
        color.HasValue && color.Value.A < byte.MaxValue;

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
            if (titleElement.Kind == MarkdownPdfJsonValueKind.False || titleElement.Kind == MarkdownPdfJsonValueKind.Null) {
                return null;
            }

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

    private static bool UsesHorizontalBarType(string? type) {
        string normalized = NormalizeChartType(type);
        return string.Equals(normalized, "horizontalbar", StringComparison.Ordinal) ||
               string.Equals(normalized, "barhorizontal", StringComparison.Ordinal);
    }

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

    private static bool TryMapChartKind(string? type, bool horizontalIndexAxis, bool stacked, bool filledLineChart, out OfficeChartKind kind) {
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
                kind = filledLineChart
                    ? (stacked ? OfficeChartKind.AreaStacked : OfficeChartKind.Area)
                    : (stacked ? OfficeChartKind.LineStacked : OfficeChartKind.Line);
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

    private static bool IsCartesianChart(OfficeChartKind chartKind) =>
        IsBarChart(chartKind) ||
        IsColumnChart(chartKind) ||
        IsLineChart(chartKind) ||
        IsAreaChart(chartKind) ||
        chartKind == OfficeChartKind.Scatter;

    private static bool IsBarChart(OfficeChartKind chartKind) =>
        chartKind == OfficeChartKind.BarClustered ||
        chartKind == OfficeChartKind.BarStacked ||
        chartKind == OfficeChartKind.BarStacked100;

    private static bool IsColumnChart(OfficeChartKind chartKind) =>
        chartKind == OfficeChartKind.ColumnClustered ||
        chartKind == OfficeChartKind.ColumnStacked ||
        chartKind == OfficeChartKind.ColumnStacked100;

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
            TryGetProperty(options, "plugins", out MarkdownPdfJsonValue plugins)) {
            if (plugins.Kind == MarkdownPdfJsonValueKind.False) {
                return false;
            }

            if (TryGetProperty(plugins, "legend", out MarkdownPdfJsonValue legend)) {
                if (legend.Kind == MarkdownPdfJsonValueKind.False || legend.Kind == MarkdownPdfJsonValueKind.Null) {
                    return false;
                }

                if (legend.Kind == MarkdownPdfJsonValueKind.Object) {
                    return ReadBool(legend, "display");
                }
            }
        }

        return null;
    }

    private static double? ReadPositiveDouble(MarkdownPdfJsonValue element, string propertyName) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value) || !TryReadNumber(value, out double number)) {
            return null;
        }

        return number > 0D && !double.IsNaN(number) && !double.IsInfinity(number) ? number : null;
    }

    private static bool ReadDefaultChartPointMarkers(MarkdownPdfJsonValue root) {
        if (TryGetProperty(root, "options", out MarkdownPdfJsonValue options)) {
            if (TryReadPointRadius(options, out double optionsRadius)) {
                return optionsRadius > 0D;
            }

            if (TryGetProperty(options, "elements", out MarkdownPdfJsonValue elements) &&
                TryGetProperty(elements, "point", out MarkdownPdfJsonValue point) &&
                point.Kind == MarkdownPdfJsonValueKind.Object &&
                TryReadPointRadius(point, out double elementRadius)) {
                return elementRadius > 0D;
            }
        }

        return true;
    }

    private static bool ReadDatasetShowMarkers(MarkdownPdfJsonValue dataset, bool defaultShowMarkers) =>
        TryReadPointRadius(dataset, out double radius) ? radius > 0D : defaultShowMarkers;

    private static bool TryReadPointRadius(MarkdownPdfJsonValue element, out double radius) {
        if (TryReadFiniteNumber(element, "pointRadius", out radius) ||
            TryReadFiniteNumber(element, "radius", out radius)) {
            return radius >= 0D;
        }

        radius = 0D;
        return false;
    }

    private static bool TryReadFiniteNumber(MarkdownPdfJsonValue element, string propertyName, out double number) {
        if (!TryGetProperty(element, propertyName, out MarkdownPdfJsonValue value) || !TryReadNumber(value, out number)) {
            number = 0D;
            return false;
        }

        return !double.IsNaN(number) && !double.IsInfinity(number);
    }

    private static bool TryReadNumber(MarkdownPdfJsonValue element, out double value) => element.TryGetDouble(out value);

    private static bool TryGetProperty(MarkdownPdfJsonValue element, string propertyName, out MarkdownPdfJsonValue value) =>
        element.TryGetProperty(propertyName, out value);

    private readonly struct ChartScaleVisibility {
        public ChartScaleVisibility(bool showAxis, bool showLine, bool showLabels, bool showGrid) {
            ShowAxis = showAxis;
            ShowLine = showLine;
            ShowLabels = showLabels;
            ShowGrid = showGrid;
        }

        public static ChartScaleVisibility Default { get; } = new ChartScaleVisibility(true, true, true, true);

        public bool ShowAxis { get; }

        public bool ShowLine { get; }

        public bool ShowLabels { get; }

        public bool ShowGrid { get; }
    }

}
