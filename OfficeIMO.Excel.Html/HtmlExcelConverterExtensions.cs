using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Extension methods for importing semantic OfficeIMO Excel HTML.
/// </summary>
public static partial class HtmlExcelConverterExtensions {
    private const int ExcelCellTextCharacterLimit = 32_767;
    private const int ExcelFormulaCharacterLimit = 8_192;

    /// <summary>
    /// Imports a prepared shared HTML conversion document into a native workbook without reparsing its adapter DOM.
    /// </summary>
    public static ExcelDocument ToExcelDocument(this HtmlConversionDocument document, HtmlToExcelOptions? options = null) {
        return GetWorkbookOrThrow(ToExcelDocumentResult(document, options));
    }

    /// <summary>
    /// Imports a prepared shared HTML conversion document and returns the workbook plus structured evidence.
    /// </summary>
    public static HtmlToExcelResult ToExcelDocumentResult(this HtmlConversionDocument document, HtmlToExcelOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        IHtmlDocument adapterDocument = document.CreateDocumentForConversion(HtmlCssMediaContext.Screen);
        HtmlToExcelOptions resolved = options?.Clone() ?? new HtmlToExcelOptions();
        return ImportDocument(adapterDocument, document.SemanticDocument, document.Trust, resolved, document.Diagnostics);
    }

    private static HtmlToExcelResult ImportDocument(
        IHtmlDocument document,
        HtmlSemanticDocument semanticDocument,
        HtmlInputTrust trust,
        HtmlToExcelOptions options,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics = null) {
        options.Limits.Validate();
        if (!Enum.IsDefined(typeof(HtmlImportMode), options.Mode)) throw new ArgumentOutOfRangeException(nameof(options.Mode));
        ExcelDocument workbook = ExcelDocument.Create();
        var result = new HtmlToExcelResult(workbook);
        if (initialDiagnostics != null) {
            foreach (HtmlDiagnostic diagnostic in initialDiagnostics) result.AddImportDiagnostic(diagnostic);
        }
        var budget = new HtmlImportBudget(options.Limits);
        OfficeHtmlSemanticEnvelopeInfo envelope = OfficeHtmlSemanticEnvelope.Inspect(document, "excel");
        IReadOnlyList<IElement> sheetSections = OfficeHtmlSemanticEnvelope
            .SelectOwnedContainers(document, envelope, "section.officeimo-sheet");
        bool useSemantic = options.Mode != HtmlImportMode.Generic
            && (options.Mode == HtmlImportMode.Semantic || envelope.IsPresent || sheetSections.Count > 0);
        if (useSemantic && !envelope.IsSupported) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticSchemaUnsupported,
                "The semantic HTML envelope does not use a supported Excel source and schema version.",
                HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Failure,
                detail: "source=" + envelope.ActualSource + "; version=" + envelope.SchemaVersion);
            workbook.AddWorksheet("Imported");
            return result;
        }
        if (useSemantic && !envelope.CanRestoreTargetSpecific(trust)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired,
                "Target-specific Excel restoration was not applied because the v2 envelope requires caller-trusted input.",
                HtmlDiagnosticSeverity.Warning, HtmlConversionLossKind.Approximation,
                detail: "restoration=" + envelope.RestorationMode);
            ImportGenericDocument(semanticDocument, workbook, result, options, budget);
            return result;
        }

        if (!useSemantic) {
            ImportGenericDocument(semanticDocument, workbook, result, options, budget);
            return result;
        }

        if (envelope.IsPresent && envelope.IsLegacy) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticSchemaLegacy,
                "Legacy Excel semantic HTML without an explicit schema version was imported using version 1 compatibility rules.",
                HtmlDiagnosticSeverity.Info);
        }

        if (sheetSections.Count == 0) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticContentMissing,
                "No semantic Excel sheet sections were found.", HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Failure);
            workbook.AddWorksheet("Imported");
            return result;
        }

        ApplyFormulaTrustBoundary(document, trust, options, result);

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement section in sheetSections) {
            if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "Additional semantic worksheets were omitted because the shared import limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: containerLimit);
                break;
            }

            ExcelSheet sheet = workbook.AddWorksheet(GetUniqueSheetName(GetSheetName(section), usedNames));
            var importedFormulaCells = new HashSet<long>();
            result.Sheets++;
            ImportTable(section, sheet, result, options, budget, importedFormulaCells);
            if (options.ImportFormulas) {
                ImportFormulas(section, sheet, result, budget, importedFormulaCells);
            }

            if (options.ImportComments) {
                ImportComments(section, sheet, result, budget);
            }

            if (options.ImportImages || options.ImportChartInventory) {
                ImportDrawings(section, sheet, options, result, budget);
            }

            ApplySheetVisibility(section, sheet);
        }

        return result;
    }

    private static void ApplyFormulaTrustBoundary(
        IHtmlDocument document,
        HtmlInputTrust trust,
        HtmlToExcelOptions options,
        HtmlToExcelResult result) {
        if (!options.ImportFormulas
            || trust == HtmlInputTrust.Trusted
            || options.AllowUntrustedFormulas) {
            return;
        }

        bool containsFormulaMetadata = document.QuerySelectorAll("[data-officeimo-value-kind]")
            .Any(element => string.Equals(
                element.GetAttribute("data-officeimo-value-kind"),
                "formula",
                StringComparison.OrdinalIgnoreCase))
            || document.QuerySelector("section.officeimo-formulas li[data-officeimo-cell]") != null;
        options.ImportFormulas = false;
        if (!containsFormulaMetadata) {
            return;
        }

        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticRestorationTrustRequired,
            "Excel formulas were imported as visible text because formula restoration requires caller-trusted input or an explicit AllowUntrustedFormulas opt-in.",
            HtmlDiagnosticSeverity.Warning,
            HtmlConversionLossKind.Approximation);
    }

    private static ExcelDocument GetWorkbookOrThrow(HtmlToExcelResult result) {
        if (result.Succeeded) return result.Value;

        result.Value.Dispose();
        throw new HtmlConversionException(result.Report.Diagnostics);
    }

    private static void ApplySheetVisibility(IElement section, ExcelSheet sheet) {
        string? visibility = section.GetAttribute("data-officeimo-visibility");
        if (string.IsNullOrWhiteSpace(visibility)) {
            return;
        }

        if (visibility!.Equals("veryHidden", StringComparison.OrdinalIgnoreCase)) {
            sheet.SetVeryHidden(true);
        } else if (visibility.Equals("hidden", StringComparison.OrdinalIgnoreCase)) {
            sheet.SetHidden(true);
        }
    }

    private static void ImportTable(
        IElement section,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        HashSet<long> importedFormulaCells) {
        IElement? table = section.Children.FirstOrDefault(child => IsElement(child, "table"));
        if (table == null) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticBlockMissing,
                "Sheet '" + sheet.Name + "' did not contain a direct semantic table.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        if (!budget.TryReserveTable(out string tableLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Sheet '" + sheet.Name + "' table was omitted because the shared import limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: tableLimit);
            return;
        }

        ReadRangeOrigin(section.GetAttribute("data-officeimo-range"), out int firstRow, out int firstColumn);
        ImportTableGrid(table, sheet, result, options, budget, firstRow, firstColumn, importedFormulaCells, useSemanticValues: true);
    }

    private static void ReadRangeOrigin(string? range, out int row, out int column) {
        row = 1;
        column = 1;
        if (string.IsNullOrWhiteSpace(range)) {
            return;
        }

        string firstReference = range!.Split(':')[0].Trim();
        if (TryParseCellReference(firstReference, out int parsedRow, out int parsedColumn)) {
            row = parsedRow;
            column = parsedColumn;
        }
    }

    private static void ImportFormulas(
        IElement section,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlImportBudget budget,
        HashSet<long> importedFormulaCells) {
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-formulas li[data-officeimo-cell]")) {
            string? reference = item.GetAttribute("data-officeimo-cell");
            if (!TryParseCellReference(reference, out int row, out int column)) {
                AddInvalidCellCoordinateDiagnostic(result, reference, "formula");
                continue;
            }

            long cellKey = GetImportCellKey(row, column);
            if (importedFormulaCells.Contains(cellKey)) {
                continue;
            }

            string formula = item.QuerySelector("code")?.TextContent ?? string.Empty;
            if (formula.Length == 0) {
                continue;
            }

            string annotationLimit = string.Empty;
            if (!IsWithinExcelFieldLimit(formula, budget, ExcelFormulaCharacterLimit, "ExcelFormulaCharacterLimit", out string metadataLimit)
                || !budget.TryReserveAnnotation(out annotationLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "A formula was omitted because the shared semantic metadata limit was reached.",
                    lossKind: HtmlConversionLossKind.Omission,
                    detail: metadataLimit.Length > 0 ? metadataLimit : annotationLimit);
                continue;
            }

            sheet.CellFormula(row, column, formula);
            importedFormulaCells.Add(cellKey);
            result.Formulas++;
        }
    }

    private static bool SetCellValue(
        ExcelSheet sheet,
        int row,
        int column,
        IElement cell,
        string fallbackText,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        HashSet<long>? importedFormulaCells,
        bool useSemanticValues) {
        string? kind = cell.GetAttribute("data-officeimo-value-kind");
        string? rawValue = cell.GetAttribute("data-officeimo-value");
        if (!useSemanticValues || string.IsNullOrWhiteSpace(kind) || rawValue == null) {
            return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
        }

        bool isFormula = kind!.Equals("formula", StringComparison.OrdinalIgnoreCase);
        if (isFormula && options.ImportFormulas) {
            // The table cell is the canonical formula source for current envelopes. Remember the
            // coordinate even when a budget rejects it so the compatibility inventory cannot make
            // a second, conflicting decision for the same cell.
            importedFormulaCells?.Add(GetImportCellKey(row, column));
        }

        if (!budget.IsMetadataWithinLimit(rawValue, out string metadataLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "Cell " + BuildCellReference(row, column) + " semantic value exceeded the shared metadata limit and was imported from visible text.",
                lossKind: HtmlConversionLossKind.Approximation, detail: metadataLimit);
            return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
        }

        if (kind!.Equals("number", StringComparison.OrdinalIgnoreCase)) {
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                && !double.IsNaN(number)
                && !double.IsInfinity(number)) {
                sheet.CellValue(row, column, number);
                return true;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic number value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (kind.Equals("boolean", StringComparison.OrdinalIgnoreCase)) {
            if (rawValue.Equals("1", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, true);
                return true;
            }

            if (rawValue.Equals("0", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("false", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, false);
                return true;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic boolean value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (kind.Equals("text", StringComparison.OrdinalIgnoreCase)) {
            return TrySetCellTextValue(sheet, row, column, rawValue, result, budget);
        } else if (kind.Equals("date-time", StringComparison.OrdinalIgnoreCase)) {
            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime dateTime)) {
                sheet.CellValue(row, column, dateTime);
                return true;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic date/time value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (isFormula) {
            if (!options.ImportFormulas) {
                return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
            }

            string formulaLimit = string.Empty;
            string annotationLimit = string.Empty;
            if (!IsWithinExcelFieldLimit(rawValue, budget, ExcelFormulaCharacterLimit, "ExcelFormulaCharacterLimit", out formulaLimit)
                || !budget.TryReserveAnnotation(out annotationLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "Cell " + BuildCellReference(row, column) + " formula was omitted because a semantic or native formula limit was reached.",
                    lossKind: HtmlConversionLossKind.Omission, detail: formulaLimit.Length > 0 ? formulaLimit : annotationLimit);
                return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
            }

            sheet.CellFormula(row, column, rawValue);
            result.Formulas++;
            return true;
        } else if (kind.Equals("error", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellError(row, column, rawValue);
            return true;
        }

        return TrySetCellTextValue(sheet, row, column, fallbackText, result, budget);
    }

    private static bool IsSemanticEmptyCell(IElement cell) =>
        string.Equals(cell.GetAttribute("data-officeimo-empty"), "true", StringComparison.OrdinalIgnoreCase);

    private static void ImportComments(IElement section, ExcelSheet sheet, HtmlToExcelResult result, HtmlImportBudget budget) {
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-comments li[data-officeimo-cell]")) {
            string? reference = item.GetAttribute("data-officeimo-cell");
            if (!TryParseCellReference(reference, out int row, out int column)) {
                AddInvalidCellCoordinateDiagnostic(result, reference, "comment");
                continue;
            }

            string text = item.QuerySelector("p")?.TextContent?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            string annotationLimit = string.Empty;
            if (!IsWithinExcelFieldLimit(text, budget, ExcelCellTextCharacterLimit, "ExcelCellTextCharacterLimit", out string metadataLimit)
                || !budget.TryReserveAnnotation(out annotationLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "A cell comment was omitted because the shared semantic metadata limit was reached.",
                    lossKind: HtmlConversionLossKind.Omission,
                    detail: metadataLimit.Length > 0 ? metadataLimit : annotationLimit);
                continue;
            }

            sheet.SetComment(row, column, text, ReadAuthor(item));
            result.Comments++;
        }
    }

    private static void ImportDrawings(IElement section, ExcelSheet sheet, HtmlToExcelOptions options, HtmlToExcelResult result, HtmlImportBudget budget) {
        var drawings = new List<ExcelDrawingImportItem>();
        int fallbackOrder = 0;
        if (options.ImportImages) {
            foreach (IElement item in section.QuerySelectorAll("section.officeimo-images li")) {
                drawings.Add(new ExcelDrawingImportItem(item, ExcelDrawingImportKind.Image, ReadOptionalIntAttribute(item, "data-officeimo-layer-index"), fallbackOrder++));
            }
        }

        if (options.ImportChartInventory) {
            foreach (IElement item in section.QuerySelectorAll("section.officeimo-charts li")) {
                drawings.Add(new ExcelDrawingImportItem(item, ExcelDrawingImportKind.Chart, ReadOptionalIntAttribute(item, "data-officeimo-layer-index"), fallbackOrder++));
            }
        }

        string range = section.GetAttribute("data-officeimo-range") ?? sheet.GetUsedRangeA1();
        int chartIndex = 0;
        foreach (ExcelDrawingImportItem drawing in drawings.OrderBy(item => item.LayerIndex ?? item.FallbackOrder).ThenBy(item => item.FallbackOrder)) {
            if (drawing.Kind == ExcelDrawingImportKind.Image) {
                ImportImage(drawing.Element, sheet, result, budget);
            } else {
                ImportChart(drawing.Element, sheet, result, budget, range, ref chartIndex);
            }
        }
    }

    private static void ImportImage(IElement item, ExcelSheet sheet, HtmlToExcelResult result, HtmlImportBudget budget) {
        IElement? image = item.QuerySelector("img[src]");
        if (image == null || !HtmlImageDataUri.TryParse(image.GetAttribute("src"), out HtmlImageDataUri dataUri)) {
            return;
        }

        if (!budget.IsImageWithinLimit(dataUri, out string imageLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An embedded worksheet image was omitted because the shared import limit was reached.",
                lossKind: HtmlConversionLossKind.Omission,
                detail: imageLimit);
            return;
        }

        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "Image inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' could not be decoded.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        if (!budget.TryReserveImageWithShape(dataUri, out imageLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "An embedded worksheet image was omitted because the shared import limit was reached.",
                lossKind: HtmlConversionLossKind.Omission,
                detail: imageLimit);
            return;
        }

        ReadImagePlacement(item, budget, result, out int row, out int column, out int width, out int height, out int offsetX, out int offsetY);
        string name = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        string description = NormalizeText(item.QuerySelector("p")?.TextContent);
        if (description.Length == 0) {
            description = NormalizeText(image.GetAttribute("alt"));
        }

        ExcelImage importedImage;
        if (IsAbsoluteImageAnchor(item) && TryReadIntAttribute(item, "data-officeimo-x", out int xPixels) && TryReadIntAttribute(item, "data-officeimo-y", out int yPixels)) {
            int maxGeometry = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
            xPixels = NormalizeImportInt(xPixels, 0, -maxGeometry, maxGeometry, budget, result, "image x position");
            yPixels = NormalizeImportInt(yPixels, 0, -maxGeometry, maxGeometry, budget, result, "image y position");
            importedImage = sheet.AddImageAbsolute(xPixels, yPixels, bytes, dataUri.MediaType, width, height, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        } else if (IsAbsoluteImageAnchor(item)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Image inventory item '" + (name.Length == 0 ? "Image" : name) + "' used an absolute anchor without semantic x/y coordinates and was restored to its fallback cell anchor.", lossKind: HtmlConversionLossKind.Approximation);
            importedImage = sheet.AddImage(row, column, bytes, dataUri.MediaType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        } else if (IsTwoCellImageAnchor(item)) {
            importedImage = AddTwoCellImage(item, sheet, result, budget, bytes, dataUri.MediaType, row, column, width, height, offsetX, offsetY, name, description);
        } else {
            importedImage = sheet.AddImage(row, column, bytes, dataUri.MediaType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        }

        ApplyImageTransforms(item, importedImage, budget, result);
        result.Images++;
    }

    private static ExcelImage AddTwoCellImage(
        IElement item,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlImportBudget budget,
        byte[] bytes,
        string contentType,
        int row,
        int column,
        int width,
        int height,
        int offsetX,
        int offsetY,
        string name,
        string description) {
        if (TryReadIntAttribute(item, "data-officeimo-to-row", out int toRow)
            && TryReadIntAttribute(item, "data-officeimo-to-column", out int toColumn)) {
            toRow = NormalizeImportInt(toRow, Math.Min(A1.MaxRows, row + 1), 1, A1.MaxRows, budget, result, "image ending row");
            toColumn = NormalizeImportInt(toColumn, Math.Min(A1.MaxColumns, column + 1), 1, A1.MaxColumns, budget, result, "image ending column");
            int endRow = Math.Max(row, toRow - 1);
            int endColumn = Math.Max(column, toColumn - 1);
            int maxGeometry = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
            int endOffsetX = NormalizeImportInt(ReadOptionalIntAttribute(item, "data-officeimo-to-offset-x") ?? 0, 0, 0, maxGeometry, budget, result, "image ending x offset");
            int endOffsetY = NormalizeImportInt(ReadOptionalIntAttribute(item, "data-officeimo-to-offset-y") ?? 0, 0, 0, maxGeometry, budget, result, "image ending y offset");
            ExcelImage importedImage = sheet.AddImageToRange(
                BuildRangeReference(row, column, endRow, endColumn),
                bytes,
                contentType,
                offsetX,
                offsetY,
                endOffsetX,
                endOffsetY,
                name: name.Length == 0 ? null : name,
                altText: description.Length == 0 ? null : description);
            if (toRow <= row || toColumn <= column) {
                importedImage.SetTwoCellEndingMarker(toRow, toColumn, endOffsetX, endOffsetY);
            }

            return importedImage;
        }

        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "Image inventory item '" + (name.Length == 0 ? "Image" : name) + "' used a two-cell anchor without semantic ending marker coordinates and was restored to its fallback cell anchor.", lossKind: HtmlConversionLossKind.Approximation);
        return sheet.AddImage(row, column, bytes, contentType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
    }

    private static void ImportChart(IElement item, ExcelSheet sheet, HtmlToExcelResult result, HtmlImportBudget budget, string range, ref int chartIndex) {
        string title = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        bool hasSemanticData = TryReadChartData(item, out ExcelChartData? chartData) && chartData != null;
        bool hasUsableRange = !string.IsNullOrWhiteSpace(range) && !string.Equals(range, "A1", StringComparison.OrdinalIgnoreCase);
        if (!hasSemanticData && !hasUsableRange) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                "Chart inventory item '" + title + "' on sheet '" + sheet.Name + "' did not contain semantic chart data and no usable table range was available.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        ReadChartDimensions(item, out int seriesCount, out int categoryCount);
        ExcelChartType type = ReadExcelChartType(item);
        ReadChartPlacement(item, chartIndex, budget, result, out int row, out int column, out int width, out int height);
        if (!budget.TryReserveChartWithShape(seriesCount, categoryCount, out HtmlImportBudgetReservation reservation, out string chartLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "Chart inventory item '" + title + "' was omitted because the shared chart limit was reached.",
                lossKind: HtmlConversionLossKind.Omission,
                detail: chartLimit);
            return;
        }

        using (reservation) {
            try {
                if (hasSemanticData) {
                    sheet.AddChart(chartData!, row: row, column: column, widthPixels: width, heightPixels: height, type: type, title: title.Length == 0 ? null : title);
                } else {
                    sheet.AddChartFromRange(range, row: row, column: column, widthPixels: width, heightPixels: height, type: type, title: title.Length == 0 ? null : title);
                }

                reservation.Commit();
                result.Charts++;
                chartIndex++;
            } catch (Exception ex) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ArtifactCreationFailed,
                    "Chart inventory item '" + title + "' could not be restored as a native chart: " + ex.Message,
                    lossKind: HtmlConversionLossKind.Omission, detail: ex.GetType().Name);
            }
        }
    }

    private static void ReadChartDimensions(IElement item, out int series, out int categories) {
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            series = 1;
            categories = 1;
            return;
        }

        List<IElement> rows = table.QuerySelectorAll("tr").ToList();
        series = Math.Max(1, rows.Count - 1);
        categories = rows.Count == 0
            ? 1
            : Math.Max(1, rows[0].Children.Count(child => IsElement(child, "th") || IsElement(child, "td")) - 1);
    }

    private static bool TryReadChartData(IElement item, out ExcelChartData? chartData) {
        chartData = null;
        IElement? table = item.QuerySelector("table.officeimo-chart-data");
        if (table == null) {
            return false;
        }

        List<IElement> rows = table.QuerySelectorAll("tr").ToList();
        if (rows.Count < 2) {
            return false;
        }

        List<string> categories = rows[0]
            .Children
            .Where(child => IsElement(child, "th") || IsElement(child, "td"))
            .Skip(1)
            .Select(cell => NormalizeText(cell.TextContent))
            .ToList();
        if (categories.Count == 0) {
            return false;
        }

        var series = new List<ExcelChartSeries>();
        foreach (IElement row in rows.Skip(1)) {
            List<IElement> cells = row.Children.Where(child => IsElement(child, "th") || IsElement(child, "td")).ToList();
            if (cells.Count < 2) {
                return false;
            }

            string name = NormalizeText(cells[0].TextContent);
            bool hasXValues = cells.Skip(1).Any(cell => cell.GetAttribute("data-officeimo-x") != null);
            if (!hasXValues && cells.Count != categories.Count + 1) {
                return false;
            }

            int pointCount = hasXValues ? cells.Count - 1 : categories.Count;
            var values = new double[pointCount];
            var xValues = new double[pointCount];
            for (int i = 0; i < pointCount; i++) {
                IElement valueCell = cells[i + 1];
                if (!double.TryParse(NormalizeText(valueCell.TextContent), NumberStyles.Float, CultureInfo.InvariantCulture, out values[i])
                    || double.IsNaN(values[i])
                    || double.IsInfinity(values[i])) {
                    return false;
                }

                if (hasXValues) {
                    string? rawXValue = valueCell.GetAttribute("data-officeimo-x");
                    if (rawXValue == null
                        || !double.TryParse(rawXValue, NumberStyles.Float, CultureInfo.InvariantCulture, out xValues[i])
                        || double.IsNaN(xValues[i])
                        || double.IsInfinity(xValues[i])) {
                        return false;
                    }
                }
            }

            ExcelChartType? chartType = null;
            string? rawChartType = row.GetAttribute("data-officeimo-chart-type");
            if (!string.IsNullOrWhiteSpace(rawChartType) &&
                Enum.TryParse(rawChartType, ignoreCase: true, out ExcelChartType parsedChartType)) {
                chartType = parsedChartType;
            }

            series.Add(hasXValues
                ? new ExcelChartSeries(name, values, xValues, chartType)
                : new ExcelChartSeries(name, values, chartType));
        }

        if (series.Count == 0) {
            return false;
        }

        chartData = new ExcelChartData(categories, series);
        return true;
    }

    private static ExcelChartType ReadExcelChartType(IElement item) {
        string? chartTypeAttribute = item.GetAttribute("data-officeimo-chart-type");
        if (!string.IsNullOrWhiteSpace(chartTypeAttribute) &&
            Enum.TryParse(chartTypeAttribute, ignoreCase: true, out ExcelChartType attributeType)) {
            return attributeType;
        }

        string meta = string.Join(" ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        const string marker = "Type:";
        int index = meta.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (index >= 0) {
            string value = meta.Substring(index + marker.Length).Split(new[] { ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;
            if (Enum.TryParse(value, ignoreCase: true, out ExcelChartType type)) {
                return type;
            }
        }

        return ExcelChartType.ColumnClustered;
    }

    private static void ReadChartPlacement(IElement item, int chartIndex, HtmlImportBudget budget, HtmlToExcelResult result, out int row, out int column, out int width, out int height) {
        row = 1 + chartIndex * 12;
        column = 6;
        width = 320;
        height = 180;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        foreach (string part in meta.Split(';')) {
            string value = part.Trim();
            if (value.StartsWith("Cell:", StringComparison.OrdinalIgnoreCase)) {
                string[] pieces = value.Substring("Cell:".Length).Split(',');
                if (pieces.Length == 2) {
                    _ = int.TryParse(pieces[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
                    _ = int.TryParse(pieces[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out column);
                }
            } else if (value.StartsWith("Size:", StringComparison.OrdinalIgnoreCase)) {
                string[] pieces = value.Substring("Size:".Length).Split('x');
                if (pieces.Length == 2) {
                    _ = int.TryParse(pieces[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out width);
                    _ = int.TryParse(pieces[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out height);
                }
            }
        }

        row = Math.Max(1, row);
        column = Math.Max(1, column);
        width = Math.Max(1, width);
        height = Math.Max(1, height);
        if (TryReadIntAttribute(item, "data-officeimo-row", out int attributeRow)) row = Math.Max(1, attributeRow);
        if (TryReadIntAttribute(item, "data-officeimo-column", out int attributeColumn)) column = Math.Max(1, attributeColumn);
        if (TryReadIntAttribute(item, "data-officeimo-width", out int attributeWidth)) width = Math.Max(1, attributeWidth);
        if (TryReadIntAttribute(item, "data-officeimo-height", out int attributeHeight)) height = Math.Max(1, attributeHeight);
        row = NormalizeImportInt(row, 1, 1, A1.MaxRows, budget, result, "chart row");
        column = NormalizeImportInt(column, 6, 1, A1.MaxColumns, budget, result, "chart column");
        int maxGeometry = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
        width = NormalizeImportInt(width, 320, 1, maxGeometry, budget, result, "chart width");
        height = NormalizeImportInt(height, 180, 1, maxGeometry, budget, result, "chart height");
    }

    private static void ReadImagePlacement(IElement item, HtmlImportBudget budget, HtmlToExcelResult result, out int row, out int column, out int width, out int height, out int offsetX, out int offsetY) {
        row = 1;
        column = 1;
        width = 96;
        height = 32;
        offsetX = 0;
        offsetY = 0;
        string meta = string.Join("; ", item.QuerySelectorAll(".officeimo-feature-meta").Select(element => element.TextContent));
        foreach (string part in meta.Split(';')) {
            string value = part.Trim();
            if (value.StartsWith("Cell:", StringComparison.OrdinalIgnoreCase)) {
                string[] pieces = value.Substring("Cell:".Length).Split(',');
                if (pieces.Length == 2) {
                    _ = int.TryParse(pieces[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out row);
                    _ = int.TryParse(pieces[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out column);
                }
            } else if (value.StartsWith("Size:", StringComparison.OrdinalIgnoreCase)) {
                string[] pieces = value.Substring("Size:".Length).Split('x');
                if (pieces.Length == 2) {
                    _ = int.TryParse(pieces[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out width);
                    _ = int.TryParse(pieces[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out height);
                }
            } else if (value.StartsWith("Offset:", StringComparison.OrdinalIgnoreCase)) {
                string[] pieces = value.Substring("Offset:".Length).Split(',');
                if (pieces.Length == 2) {
                    _ = int.TryParse(pieces[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out offsetX);
                    _ = int.TryParse(pieces[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out offsetY);
                }
            }
        }

        row = Math.Max(1, row);
        column = Math.Max(1, column);
        width = Math.Max(1, width);
        height = Math.Max(1, height);
        offsetX = Math.Max(0, offsetX);
        offsetY = Math.Max(0, offsetY);
        if (TryReadIntAttribute(item, "data-officeimo-row", out int attributeRow)) row = Math.Max(1, attributeRow);
        if (TryReadIntAttribute(item, "data-officeimo-column", out int attributeColumn)) column = Math.Max(1, attributeColumn);
        if (TryReadIntAttribute(item, "data-officeimo-width", out int attributeWidth)) width = Math.Max(1, attributeWidth);
        if (TryReadIntAttribute(item, "data-officeimo-height", out int attributeHeight)) height = Math.Max(1, attributeHeight);
        if (TryReadIntAttribute(item, "data-officeimo-offset-x", out int attributeOffsetX)) offsetX = Math.Max(0, attributeOffsetX);
        if (TryReadIntAttribute(item, "data-officeimo-offset-y", out int attributeOffsetY)) offsetY = Math.Max(0, attributeOffsetY);
        row = NormalizeImportInt(row, 1, 1, A1.MaxRows, budget, result, "image row");
        column = NormalizeImportInt(column, 1, 1, A1.MaxColumns, budget, result, "image column");
        int maxGeometry = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
        width = NormalizeImportInt(width, 96, 1, maxGeometry, budget, result, "image width");
        height = NormalizeImportInt(height, 32, 1, maxGeometry, budget, result, "image height");
        offsetX = NormalizeImportInt(offsetX, 0, 0, maxGeometry, budget, result, "image x offset");
        offsetY = NormalizeImportInt(offsetY, 0, 0, maxGeometry, budget, result, "image y offset");
    }

    private static void ApplyImageTransforms(IElement item, ExcelImage image, HtmlImportBudget budget, HtmlToExcelResult result) {
        if (TryReadDoubleAttribute(item, "data-officeimo-rotation", out double rotation)) {
            if (budget.TryNormalizeRange(rotation, 0D, -budget.Limits.MaxAbsoluteGeometry, budget.Limits.MaxAbsoluteGeometry, out double normalizedRotation)) {
                image.SetRotation(normalizedRotation);
            } else {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                    "Invalid image rotation used the neutral fallback.", lossKind: HtmlConversionLossKind.Approximation);
            }
        }

        bool hasHorizontalFlip = TryReadBoolAttribute(item, "data-officeimo-flip-horizontal", out bool horizontalFlip);
        bool hasVerticalFlip = TryReadBoolAttribute(item, "data-officeimo-flip-vertical", out bool verticalFlip);
        if (hasHorizontalFlip || hasVerticalFlip) {
            image.SetFlip(hasHorizontalFlip && horizontalFlip, hasVerticalFlip && verticalFlip);
        }

        double left = NormalizeCrop(item, "data-officeimo-crop-left", budget, result);
        double top = NormalizeCrop(item, "data-officeimo-crop-top", budget, result);
        double right = NormalizeCrop(item, "data-officeimo-crop-right", budget, result);
        double bottom = NormalizeCrop(item, "data-officeimo-crop-bottom", budget, result);
        if (left > 0D || top > 0D || right > 0D || bottom > 0D) {
            image.SetCropRatio(left, top, right, bottom);
        }
    }

    private static double NormalizeCrop(IElement item, string attributeName, HtmlImportBudget budget, HtmlToExcelResult result) {
        double value = ReadOptionalDoubleAttribute(item, attributeName) ?? 0D;
        if (budget.TryNormalizeRange(value, 0D, 0D, 1D, out double normalized)) return normalized;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "Invalid image crop metadata used the zero fallback.", lossKind: HtmlConversionLossKind.Approximation, source: attributeName);
        return 0D;
    }

    private static int NormalizeImportInt(
        int value,
        int fallback,
        int minimum,
        int maximum,
        HtmlImportBudget budget,
        HtmlToExcelResult result,
        string source) {
        if (value >= minimum && value <= maximum) return value;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "Invalid " + source + " metadata used a bounded fallback.",
            lossKind: HtmlConversionLossKind.Approximation,
            source: source,
            detail: "value=" + value.ToString(CultureInfo.InvariantCulture) + "; maximum=" + maximum.ToString(CultureInfo.InvariantCulture));
        return fallback;
    }

    private static bool IsAbsoluteImageAnchor(IElement item) =>
        string.Equals(item.GetAttribute("data-officeimo-anchor"), "absolute", StringComparison.OrdinalIgnoreCase);

    private static bool IsTwoCellImageAnchor(IElement item) =>
        string.Equals(item.GetAttribute("data-officeimo-anchor"), "twoCell", StringComparison.OrdinalIgnoreCase);

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

    private static string ReadAuthor(IElement item) {
        foreach (IElement meta in item.QuerySelectorAll(".officeimo-feature-meta")) {
            string text = NormalizeText(meta.TextContent);
            if (text.StartsWith("Author:", StringComparison.OrdinalIgnoreCase)) {
                return text.Substring("Author:".Length).Trim();
            }
        }

        return "OfficeIMO";
    }

    private static bool TryParseCellReference(string? value, out int row, out int column) {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        int index = 0;
        string reference = value!;
        while (index < reference.Length && char.IsLetter(reference[index])) {
            int letter = char.ToUpperInvariant(reference[index]) - 'A' + 1;
            if (letter < 1 || letter > 26 || column > (A1.MaxColumns - letter) / 26) {
                row = 0;
                column = 0;
                return false;
            }
            column = column * 26 + letter;
            index++;
        }

        for (int digitIndex = index; digitIndex < reference.Length; digitIndex++) {
            if (reference[digitIndex] < '0' || reference[digitIndex] > '9') return false;
        }

        return column > 0
            && column <= A1.MaxColumns
            && index < reference.Length
            && int.TryParse(reference.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row)
            && row > 0
            && row <= A1.MaxRows;
    }

    private static bool TrySetCellTextValue(
        ExcelSheet sheet,
        int row,
        int column,
        string text,
        HtmlToExcelResult result,
        HtmlImportBudget budget) {
        if (!IsWithinExcelFieldLimit(text, budget, ExcelCellTextCharacterLimit, "ExcelCellTextCharacterLimit", out string detail)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                "Cell " + BuildCellReference(row, column) + " text was omitted because a semantic or native Excel field limit was reached.",
                lossKind: HtmlConversionLossKind.Omission, detail: detail);
            return false;
        }

        sheet.CellValue(row, column, text);
        return true;
    }

    private static bool IsWithinExcelFieldLimit(
        string? value,
        HtmlImportBudget budget,
        int nativeLimit,
        string nativeLimitName,
        out string detail) {
        if (!budget.IsMetadataWithinLimit(value, out detail)) return false;
        int length = value?.Length ?? 0;
        if (length <= nativeLimit) return true;
        detail = nativeLimitName + ": Actual=" + length.ToString(CultureInfo.InvariantCulture)
            + "; Limit=" + nativeLimit.ToString(CultureInfo.InvariantCulture);
        return false;
    }

    private static void AddInvalidCellCoordinateDiagnostic(
        HtmlToExcelResult result,
        string? reference,
        string contentKind) {
        if (string.IsNullOrWhiteSpace(reference)) return;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
            "The " + contentKind + " coordinate '" + reference + "' was outside the Excel worksheet grid and was omitted.",
            lossKind: HtmlConversionLossKind.Omission,
            detail: "Rows=1-" + A1.MaxRows.ToString(CultureInfo.InvariantCulture)
                + "; Columns=1-" + A1.MaxColumns.ToString(CultureInfo.InvariantCulture));
    }

    private static string BuildCellReference(int row, int column) {
        var letters = new StringBuilder();
        int current = column;
        while (current > 0) {
            current--;
            letters.Insert(0, (char)('A' + current % 26));
            current /= 26;
        }

        return letters.Append(row.ToString(CultureInfo.InvariantCulture)).ToString();
    }

    private static string BuildRangeReference(int startRow, int startColumn, int endRow, int endColumn) =>
        BuildCellReference(startRow, startColumn) + ":" + BuildCellReference(endRow, endColumn);

    private static bool IsElement(IElement element, string name) =>
        string.Equals(element.LocalName, name, StringComparison.OrdinalIgnoreCase);

    private static string NormalizeText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : string.Join(" ", text!.Split((char[]?)null!, StringSplitOptions.RemoveEmptyEntries));

    private sealed class ExcelDrawingImportItem {
        internal ExcelDrawingImportItem(IElement element, ExcelDrawingImportKind kind, int? layerIndex, int fallbackOrder) {
            Element = element;
            Kind = kind;
            LayerIndex = layerIndex;
            FallbackOrder = fallbackOrder;
        }

        internal IElement Element { get; }

        internal ExcelDrawingImportKind Kind { get; }

        internal int? LayerIndex { get; }

        internal int FallbackOrder { get; }
    }

    private enum ExcelDrawingImportKind {
        Image,
        Chart
    }
}
