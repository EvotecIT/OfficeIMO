using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Extension methods for importing semantic OfficeIMO PowerPoint HTML.
/// </summary>
public static partial class HtmlPowerPointConverterExtensions {
    /// <summary>
    /// Imports a prepared shared HTML conversion document without reparsing its adapter DOM.
    /// </summary>
    public static PptCore.PowerPointPresentation ToPowerPointPresentation(this HtmlConversionDocument document, HtmlToPowerPointOptions? options = null) {
        return GetPresentationOrThrow(ToPowerPointPresentationResult(document, options));
    }

    /// <summary>
    /// Imports a prepared shared HTML conversion document and returns the presentation plus structured evidence.
    /// </summary>
    public static HtmlToPowerPointResult ToPowerPointPresentationResult(this HtmlConversionDocument document, HtmlToPowerPointOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IHtmlDocument adapterDocument = document.CreateDocumentForConversion(HtmlCssMediaContext.Screen);
        HtmlToPowerPointOptions resolved = options?.Clone() ?? new HtmlToPowerPointOptions();
        return ImportDocument(adapterDocument, resolved, document.Diagnostics);
    }

    private static HtmlToPowerPointResult ImportDocument(
        IHtmlDocument document,
        HtmlToPowerPointOptions options,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics = null) {
        options.Limits.Validate();
        if (!Enum.IsDefined(typeof(HtmlImportMode), options.Mode)) throw new ArgumentOutOfRangeException(nameof(options.Mode));
        PptCore.PowerPointPresentation presentation = PptCore.PowerPointPresentation.Create();
        var result = new HtmlToPowerPointResult(presentation);
        if (initialDiagnostics != null) {
            foreach (HtmlDiagnostic diagnostic in initialDiagnostics) result.AddImportDiagnostic(diagnostic);
        }
        var budget = new HtmlImportBudget(options.Limits);
        OfficeHtmlSemanticEnvelopeInfo envelope = OfficeHtmlSemanticEnvelope.Inspect(document, "powerpoint");
        IReadOnlyList<IElement> slideSections = OfficeHtmlSemanticEnvelope
            .SelectOwnedContainers(document, envelope, "section.officeimo-slide");
        bool useSemantic = options.Mode != HtmlImportMode.Generic
            && (options.Mode == HtmlImportMode.Semantic || envelope.IsPresent || slideSections.Count > 0);
        if (useSemantic && !envelope.IsSupported) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported,
                "The semantic HTML envelope does not use a supported PowerPoint source and schema version.",
                HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Failure,
                detail: "source=" + envelope.ActualSource + "; version=" + envelope.SchemaVersion);
            return result;
        }

        if (!useSemantic) {
            ImportGenericDocument(document, presentation, options, result, budget);
            return result;
        }

        if (envelope.IsPresent && envelope.IsLegacy) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticSchemaLegacy,
                "Legacy PowerPoint semantic HTML without an explicit schema version was imported using version 1 compatibility rules.",
                HtmlDiagnosticSeverity.Info);
        }

        if (slideSections.Count == 0) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticContentMissing,
                "No semantic PowerPoint slide sections were found.", HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Failure);
            return result;
        }

        foreach (IElement slideSection in slideSections) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional semantic slides were omitted because the shared import limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: containerLimit);
                break;
            }

            PptCore.PowerPointSlide slide = presentation.AddSlide();
            ImportSlide(slideSection, slide, options, result, budget);
            if (IsTrueAttribute(slideSection.GetAttribute("data-officeimo-hidden"))) {
                slide.Hidden = true;
            }
        }

        return result;
    }

    private static PptCore.PowerPointPresentation GetPresentationOrThrow(HtmlToPowerPointResult result) {
        if (result.Succeeded) return result.Value;

        result.Value.Dispose();
        throw new HtmlConversionException(result.Report.Diagnostics);
    }

    private static void ImportSlide(IElement section, PptCore.PowerPointSlide slide, HtmlToPowerPointOptions options, HtmlToPowerPointResult result, HtmlImportBudget budget) {
        result.Slides++;
        ImportSemanticShapes(section, slide, options, result, budget);

        if (options.ImportNotes) {
            ImportNotes(section, slide, result, budget);
        }
    }

    private static bool IsTrueAttribute(string? value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "1", StringComparison.Ordinal);

    private static void ImportPicture(IElement item, PptCore.PowerPointSlide slide, HtmlToPowerPointResult result, HtmlImportBudget budget, ref double fallbackTop) {
        IElement? image = IsElement(item, "img") && item.HasAttribute("src") ? item : item.QuerySelector("img[src]");
        if (image == null || !HtmlImageDataUri.TryParse(image.GetAttribute("src"), out HtmlImageDataUri dataUri)) {
            return;
        }

        if (!budget.TryReserveImage(dataUri, out string imageLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An embedded slide picture was omitted because the shared image limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: imageLimit);
            return;
        }

        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "Picture inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' could not be decoded.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        if (!TryGetImagePartType(dataUri.MediaType, out PptCore.ImagePartType imagePartType)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                "Picture inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' used unsupported media type '" + dataUri.MediaType + "' and was not imported.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        ReadPictureSize(item, budget, result, out double width, out double height);
        ReadPicturePosition(item, 720D, fallbackTop, budget, result, out double left, out double pictureTop);
        using var stream = new MemoryStream(bytes);
        PptCore.PowerPointPicture picture = slide.AddPicturePoints(stream, imagePartType, left, pictureTop, width, height);
        string label = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        string alt = NormalizeText(image.GetAttribute("alt"));
        if (label.Length > 0) {
            picture.Name = label;
        }

        if (alt.Length > 0) {
            picture.AltText = alt;
        }

        ApplyPictureTransforms(item, picture, budget, result);
        result.Pictures++;
        fallbackTop = Math.Max(fallbackTop, pictureTop + height + 18D);
    }

    private static void ImportChart(IElement item, PptCore.PowerPointSlide slide, HtmlToPowerPointResult result, HtmlImportBudget budget, ref double fallbackTop) {
        string title = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        ReadChartDataDimensions(item, out int seriesCount, out int categoryCount);
        if (!budget.TryReserveChart(seriesCount, categoryCount, out string chartLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Chart inventory item '" + title + "' was omitted because the shared chart limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: chartLimit);
            return;
        }

        bool restoredFromSemanticData = TryReadChartData(item, out PptCore.PowerPointChartData? semanticData);
        PptCore.PowerPointChartData data = semanticData ?? CreatePlaceholderChartDataFromInventory(item);
        string chartKind = ReadChartKind(item);
        ReadChartGeometry(item, 500D, fallbackTop, 320D, 180D, budget, result, out double left, out double chartTop, out double width, out double height);
        if (!TryAddChartByKind(slide, chartKind, data, left, chartTop, width, height, out PptCore.PowerPointChart? chart, out string? fallbackMessage) || chart == null) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                "Chart inventory item '" + (title.Length == 0 ? "Imported chart" : title) + "' used unsupported chart kind '" + chartKind + "' and was not imported.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        chart.SetTitle(title.Length == 0 ? "Imported chart" : title);
        ApplyShapeTransforms(item, chart, budget, result);
        result.Charts++;
        if (!string.IsNullOrWhiteSpace(fallbackMessage)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Chart inventory item '" + (title.Length == 0 ? "Imported chart" : title) + "' " + fallbackMessage, lossKind: HtmlConversionLossKind.Approximation);
        }

        if (!restoredFromSemanticData) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Chart inventory item '" + (title.Length == 0 ? "Imported chart" : title) + "' was restored as a native chart with reconstructed placeholder values; exact source chart data was not present in semantic HTML.", lossKind: HtmlConversionLossKind.Approximation);
        }

        fallbackTop = Math.Max(fallbackTop + 198D, chartTop + height + 18D);
    }

    private static void ImportNotes(IElement section, PptCore.PowerPointSlide slide, HtmlToPowerPointResult result, HtmlImportBudget budget) {
        string notes = ExtractPresenterNotes(section.QuerySelector("pre.officeimo-source-markdown")?.TextContent);
        if (notes.Length == 0) {
            return;
        }

        string annotationLimit = string.Empty;
        if (!budget.IsMetadataWithinLimit(notes, out string metadataLimit)
            || !budget.TryReserveAnnotation(out annotationLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "Presenter notes were omitted because the shared semantic metadata limit was reached.",
                lossKind: HtmlConversionLossKind.Omission,
                detail: metadataLimit.Length > 0 ? metadataLimit : annotationLimit);
            return;
        }

        slide.Notes.Text = notes;
        result.Notes++;
    }

    private static PptCore.PowerPointChartData CreatePlaceholderChartDataFromInventory(IElement item) {
        ReadChartShape(item, out int seriesCount, out int categoryCount);
        return CreatePlaceholderChartData(seriesCount, categoryCount);
    }

    private static void ReadChartDataDimensions(IElement item, out int series, out int categories) {
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            ReadChartShape(item, out series, out categories);
            return;
        }

        series = Math.Max(1, table.QuerySelectorAll("tbody tr").Length);
        categories = Math.Max(1, table.QuerySelectorAll("thead tr th").Length - 1);
    }

    private static PptCore.PowerPointChartData CreatePlaceholderChartData(int seriesCount, int categoryCount) {
        seriesCount = Math.Max(1, seriesCount);
        categoryCount = Math.Max(1, categoryCount);
        string[] categories = Enumerable.Range(1, categoryCount)
            .Select(index => "C" + index.ToString(CultureInfo.InvariantCulture))
            .ToArray();
        PptCore.PowerPointChartSeries[] series = Enumerable.Range(1, seriesCount)
            .Select(seriesIndex => new PptCore.PowerPointChartSeries(
                "Series " + seriesIndex.ToString(CultureInfo.InvariantCulture),
                Enumerable.Range(1, categoryCount).Select(categoryIndex => (double)(seriesIndex * categoryIndex)).ToArray()))
            .ToArray();
        return new PptCore.PowerPointChartData(categories, series);
    }

    private static bool TryReadChartData(IElement item, out PptCore.PowerPointChartData? data) {
        data = null;
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            return false;
        }

        List<string> categories = table.QuerySelectorAll("thead tr th")
            .Skip(1)
            .Select(header => PreserveText(header.TextContent))
            .ToList();
        if (categories.Count == 0) {
            return false;
        }

        var series = new List<PptCore.PowerPointChartSeries>();
        foreach (IElement row in table.QuerySelectorAll("tbody tr")) {
            string name = PreserveText(row.QuerySelector("th")?.TextContent);
            if (name.Length == 0) {
                name = "Series " + (series.Count + 1).ToString(CultureInfo.InvariantCulture);
            }

            List<IElement> valueCells = row.QuerySelectorAll("td").ToList();
            var values = new double[valueCells.Count];
            var xValues = new double[valueCells.Count];
            bool hasXValues = valueCells.Any(cell => cell.GetAttribute("data-officeimo-x") != null);
            for (int i = 0; i < valueCells.Count; i++) {
                IElement cell = valueCells[i];
                string text = PreserveText(cell.TextContent);
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out values[i])
                    || double.IsNaN(values[i]) || double.IsInfinity(values[i])) {
                    return false;
                }

                if (hasXValues) {
                    string? rawXValue = cell.GetAttribute("data-officeimo-x");
                    if (rawXValue == null
                        || !double.TryParse(rawXValue, NumberStyles.Float, CultureInfo.InvariantCulture, out xValues[i])
                        || double.IsNaN(xValues[i]) || double.IsInfinity(xValues[i])) {
                        return false;
                    }
                }
            }

            if (!hasXValues && values.Length != categories.Count) {
                return false;
            }

            series.Add(hasXValues
                ? new PptCore.PowerPointChartSeries(name, values, xValues)
                : new PptCore.PowerPointChartSeries(name, values));
        }

        if (series.Count == 0) {
            return false;
        }

        data = new PptCore.PowerPointChartData(categories, series);
        return true;
    }

    private static void ReadChartShape(IElement item, out int seriesCount, out int categoryCount) {
        seriesCount = 1;
        categoryCount = 3;
        string meta = string.Join(" ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        TryReadMetaInt(meta, "Series:", ref seriesCount);
        TryReadMetaInt(meta, "Categories:", ref categoryCount);
    }

    private static void TryReadMetaInt(string meta, string marker, ref int value) {
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0) {
            return;
        }

        string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed > 0) {
            value = parsed;
        }
    }

    private static void ReadPictureSize(IElement item, HtmlImportBudget budget, HtmlToPowerPointResult result, out double width, out double height) {
        width = 72D;
        height = 72D;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Size:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0) {
            string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
            string[] parts = text.Replace("pt", string.Empty).Split('x');
            if (parts.Length == 2) {
                _ = double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out width);
                _ = double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out height);
            }
        }

        if (TryReadDoubleAttribute(item, "data-officeimo-width", out double attributeWidth)) width = attributeWidth;
        if (TryReadDoubleAttribute(item, "data-officeimo-height", out double attributeHeight)) height = attributeHeight;
        width = NormalizeGeometry(width, 72D, 1D, budget, result, "picture width");
        height = NormalizeGeometry(height, 72D, 1D, budget, result, "picture height");
    }

    private static void ReadPicturePosition(IElement item, double fallbackLeft, double fallbackTop, HtmlImportBudget budget, HtmlToPowerPointResult result, out double left, out double top) {
        left = fallbackLeft;
        top = fallbackTop;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Position:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0) {
            string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
            string[] parts = text.Replace("pt", string.Empty).Split(',');
            if (parts.Length == 2) {
                _ = double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out left);
                _ = double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out top);
            }
        }

        if (TryReadDoubleAttribute(item, "data-officeimo-left", out double attributeLeft)) left = attributeLeft;
        if (TryReadDoubleAttribute(item, "data-officeimo-top", out double attributeTop)) top = attributeTop;
        left = NormalizeGeometry(left, fallbackLeft, -budget.Limits.MaxAbsoluteGeometry, budget, result, "picture left");
        top = NormalizeGeometry(top, fallbackTop, -budget.Limits.MaxAbsoluteGeometry, budget, result, "picture top");
    }

    private static void ApplyPictureTransforms(IElement item, PptCore.PowerPointPicture picture, HtmlImportBudget budget, HtmlToPowerPointResult result) {
        ApplyShapeTransforms(item, picture, budget, result);

        double left = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-left") ?? 0D;
        double top = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-top") ?? 0D;
        double right = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-right") ?? 0D;
        double bottom = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-bottom") ?? 0D;
        left = NormalizeRange(left, 0D, 0D, 1D, budget, result, "picture crop left");
        top = NormalizeRange(top, 0D, 0D, 1D, budget, result, "picture crop top");
        right = NormalizeRange(right, 0D, 0D, 1D, budget, result, "picture crop right");
        bottom = NormalizeRange(bottom, 0D, 0D, 1D, budget, result, "picture crop bottom");
        if (left > 0D || top > 0D || right > 0D || bottom > 0D) {
            picture.Crop(left * 100D, top * 100D, right * 100D, bottom * 100D);
        }
    }

    private static void ApplyShapeTransforms(IElement item, PptCore.PowerPointShape shape, HtmlImportBudget budget, HtmlToPowerPointResult result) {
        if (TryReadDoubleAttribute(item, "data-officeimo-rotation", out double rotation)) {
            shape.Rotation = NormalizeGeometry(rotation, 0D, -budget.Limits.MaxAbsoluteGeometry, budget, result, "shape rotation");
        }

        if (TryReadBoolAttribute(item, "data-officeimo-flip-horizontal", out bool horizontalFlip)) {
            shape.HorizontalFlip = horizontalFlip;
        }

        if (TryReadBoolAttribute(item, "data-officeimo-flip-vertical", out bool verticalFlip)) {
            shape.VerticalFlip = verticalFlip;
        }

        if (IsTrueAttribute(item.GetAttribute("data-officeimo-hidden"))) {
            shape.Hidden = true;
        }
    }

    private static int? ReadOptionalIntAttribute(IElement item, string name) =>
        TryReadIntAttribute(item, name, out int value) ? value : null;

    private static bool TryReadIntAttribute(IElement item, string name, out int value) {
        value = 0;
        string? raw = item.GetAttribute(name);
        return !string.IsNullOrWhiteSpace(raw)
            && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private static double? ReadOptionalDoubleAttribute(IElement item, string name) =>
        TryReadDoubleAttribute(item, name, out double value) ? value : null;

    private static bool TryReadDoubleAttribute(IElement item, string name, out double value) {
        value = 0D;
        string? raw = item.GetAttribute(name);
        return !string.IsNullOrWhiteSpace(raw)
            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryReadBoolAttribute(IElement item, string name, out bool value) {
        value = false;
        string? raw = item.GetAttribute(name);
        if (string.IsNullOrWhiteSpace(raw)) {
            return false;
        }

        if (raw!.Equals("1", StringComparison.Ordinal) || raw.Equals("true", StringComparison.OrdinalIgnoreCase)) {
            value = true;
            return true;
        }

        if (raw.Equals("0", StringComparison.Ordinal) || raw.Equals("false", StringComparison.OrdinalIgnoreCase)) {
            value = false;
            return true;
        }

        return false;
    }

    private static void ReadChartGeometry(IElement item, double fallbackLeft, double fallbackTop, double fallbackWidth, double fallbackHeight, HtmlImportBudget budget, HtmlToPowerPointResult result, out double left, out double top, out double width, out double height) {
        ReadPicturePosition(item, fallbackLeft, fallbackTop, budget, result, out left, out top);
        width = fallbackWidth;
        height = fallbackHeight;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Size:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0) {
            string text = meta.Substring(index + marker.Length).Split(';')[0].Trim();
            string[] parts = text.Replace("pt", string.Empty).Split('x');
            if (parts.Length == 2) {
                _ = double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out width);
                _ = double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out height);
            }
        }

        if (TryReadDoubleAttribute(item, "data-officeimo-width", out double attributeWidth)) width = attributeWidth;
        if (TryReadDoubleAttribute(item, "data-officeimo-height", out double attributeHeight)) height = attributeHeight;
        width = NormalizeGeometry(width, fallbackWidth, 1D, budget, result, "chart width");
        height = NormalizeGeometry(height, fallbackHeight, 1D, budget, result, "chart height");
    }

    private static double NormalizeGeometry(double value, double fallback, double minimum, HtmlImportBudget budget, HtmlToPowerPointResult result, string source) {
        if (budget.TryNormalizeGeometry(value, fallback, minimum, out double normalized)) return normalized;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "Invalid or out-of-range " + source + " metadata used its safe fallback.",
            lossKind: HtmlConversionLossKind.Approximation, source: source);
        return normalized;
    }

    private static double NormalizeRange(double value, double fallback, double minimum, double maximum, HtmlImportBudget budget, HtmlToPowerPointResult result, string source) {
        if (budget.TryNormalizeRange(value, fallback, minimum, maximum, out double normalized)) return normalized;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "Invalid or out-of-range " + source + " metadata used its safe fallback.",
            lossKind: HtmlConversionLossKind.Approximation, source: source);
        return normalized;
    }

    private static string ExtractPresenterNotes(string? markdown) {
        if (string.IsNullOrWhiteSpace(markdown)) {
            return string.Empty;
        }

        string normalized = markdown!.Replace("\r\n", "\n").Replace('\r', '\n');
        int marker = FindPresenterNotesMarker(normalized);
        if (marker < 0) {
            return string.Empty;
        }

        string tail = normalized.Substring(marker + "### Notes".Length);
        return tail.Trim('\n').Replace("\n", Environment.NewLine);
    }

    private static int FindPresenterNotesMarker(string normalizedMarkdown) {
        int searchStart = normalizedMarkdown.Length - 1;
        while (searchStart >= 0) {
            int marker = normalizedMarkdown.LastIndexOf("### Notes", searchStart, StringComparison.OrdinalIgnoreCase);
            if (marker < 0) {
                return -1;
            }

            int lineStart = marker;
            while (lineStart > 0 && normalizedMarkdown[lineStart - 1] != '\n') {
                lineStart--;
            }

            int lineEnd = normalizedMarkdown.IndexOf('\n', marker);
            if (lineEnd < 0) {
                lineEnd = normalizedMarkdown.Length;
            }

            string line = normalizedMarkdown.Substring(lineStart, lineEnd - lineStart).Trim();
            if (line.Equals("### Notes", StringComparison.OrdinalIgnoreCase)) {
                return marker;
            }

            searchStart = marker - 1;
        }

        return -1;
    }

    private static bool TryAddChartByKind(PptCore.PowerPointSlide slide, string chartKind, PptCore.PowerPointChartData data, double left, double top, double width, double height, out PptCore.PowerPointChart? chart, out string? fallbackMessage) {
        chart = null;
        fallbackMessage = null;
        if (chartKind.Equals("ClusteredColumn", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("ColumnClustered", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Line", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("StackedLine", StringComparison.OrdinalIgnoreCase) ||
            chartKind.Equals("StackedLine100", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddLineChartPoints(data, left, top, width, height);
            if (chartKind.Equals("StackedLine", StringComparison.OrdinalIgnoreCase)) {
                chart.SetLineChartGrouping(C.GroupingValues.Stacked);
            } else if (chartKind.Equals("StackedLine100", StringComparison.OrdinalIgnoreCase)) {
                chart.SetLineChartGrouping(C.GroupingValues.PercentStacked);
            }

            return true;
        }

        if (chartKind.Equals("Pie", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddPieChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Doughnut", StringComparison.OrdinalIgnoreCase)) {
            chart = slide.AddDoughnutChartPoints(data, left, top, width, height);
            return true;
        }

        if (chartKind.Equals("Scatter", StringComparison.OrdinalIgnoreCase)) {
            if (!TryCreateScatterChartData(data, out PptCore.PowerPointScatterChartData? scatterData) || scatterData == null) {
                return false;
            }

            chart = slide.AddScatterChartPoints(scatterData, left, top, width, height);
            return true;
        }

        if (CanImportAsClusteredColumnFallback(chartKind)) {
            chart = slide.AddChartPoints(data, left, top, width, height);
            fallbackMessage = "used chart kind '" + chartKind + "' and was imported as a clustered column fallback.";
            return true;
        }

        return false;
    }

    private static bool CanImportAsClusteredColumnFallback(string chartKind) =>
        chartKind.Equals("ClusteredBar", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedColumn", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedColumn100", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedBar", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedBar100", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("Area", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedArea", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("StackedArea100", StringComparison.OrdinalIgnoreCase) ||
        chartKind.Equals("Radar", StringComparison.OrdinalIgnoreCase);

    private static bool TryCreateScatterChartData(PptCore.PowerPointChartData data, out PptCore.PowerPointScatterChartData? scatterData) {
        scatterData = null;
        var series = new List<PptCore.PowerPointScatterChartSeries>();
        foreach (PptCore.PowerPointChartSeries item in data.Series) {
            if (item.XValues == null || item.XValues.Count != item.Values.Count) {
                return false;
            }

            series.Add(new PptCore.PowerPointScatterChartSeries(item.Name, item.XValues, item.Values));
        }

        if (series.Count == 0) {
            return false;
        }

        scatterData = new PptCore.PowerPointScatterChartData(series);
        return true;
    }

    private static string ReadChartKind(IElement item) {
        string meta = string.Join(" ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Type:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index < 0) {
            return "ClusteredColumn";
        }

        string value = meta.Substring(index + marker.Length).Split(';')[0].Trim();
        return value.Length == 0 ? "ClusteredColumn" : value;
    }

    private static bool TryGetImagePartType(string mediaType, out PptCore.ImagePartType imagePartType) {
        if (mediaType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/jpg", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Jpeg;
            return true;
        }

        if (mediaType.Equals("image/gif", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Gif;
            return true;
        }

        if (mediaType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Bmp;
            return true;
        }

        if (mediaType.Equals("image/tiff", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/tif", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Tiff;
            return true;
        }

        if (mediaType.Equals("image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Svg;
            return true;
        }

        if (mediaType.Equals("image/x-emf", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/emf", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Emf;
            return true;
        }

        if (mediaType.Equals("image/x-wmf", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/wmf", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Wmf;
            return true;
        }

        if (mediaType.Equals("image/x-icon", StringComparison.OrdinalIgnoreCase) ||
            mediaType.Equals("image/vnd.microsoft.icon", StringComparison.OrdinalIgnoreCase) ||
            mediaType.Equals("image/ico", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Icon;
            return true;
        }

        if (mediaType.Equals("image/x-pcx", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/pcx", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Pcx;
            return true;
        }

        if (mediaType.Equals("image/png", StringComparison.OrdinalIgnoreCase) || mediaType.Equals("image/x-png", StringComparison.OrdinalIgnoreCase)) {
            imagePartType = PptCore.ImagePartType.Png;
            return true;
        }

        imagePartType = PptCore.ImagePartType.Png;
        return false;
    }

    private static bool IsElement(IElement element, string name) =>
        string.Equals(element.LocalName, name, StringComparison.OrdinalIgnoreCase);

    private static string NormalizeText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : string.Join(" ", text!.Split((char[]?)null!, StringSplitOptions.RemoveEmptyEntries));

    private static string PreserveText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : text!.Trim();

}
