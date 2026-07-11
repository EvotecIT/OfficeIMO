using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Extension methods for importing semantic OfficeIMO Excel HTML.
/// </summary>
public static partial class HtmlExcelConverterExtensions {
    /// <summary>
    /// Imports semantic OfficeIMO Excel HTML into a native workbook.
    /// </summary>
    /// <example><code>using ExcelDocument workbook = html.ToExcelDocument();</code></example>
    public static ExcelDocument ToExcelDocument(this string html, HtmlToExcelOptions? options = null) {
        return GetWorkbookOrThrow(ToExcelDocumentResult(html, options));
    }

    /// <summary>
    /// Imports a prepared shared HTML conversion document into a native workbook without reparsing its adapter DOM.
    /// </summary>
    public static ExcelDocument ToExcelDocument(this HtmlConversionDocument document, HtmlToExcelOptions? options = null) {
        return GetWorkbookOrThrow(ToExcelDocumentResult(document, options));
    }

    /// <summary>
    /// Imports semantic OfficeIMO Excel HTML into a native workbook and returns import evidence.
    /// </summary>
    /// <example><code>HtmlToExcelResult result = html.ToExcelDocumentResult();</code></example>
    public static HtmlToExcelResult ToExcelDocumentResult(this string html, HtmlToExcelOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        return ImportDocument(HtmlDocumentParser.ParseDocument(html), options ?? new HtmlToExcelOptions());
    }

    /// <summary>
    /// Imports a prepared shared HTML conversion document and returns the workbook plus structured evidence.
    /// </summary>
    public static HtmlToExcelResult ToExcelDocumentResult(this HtmlConversionDocument document, HtmlToExcelOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ImportDocument(document.DocumentForConversion, options ?? new HtmlToExcelOptions());
    }

    private static HtmlToExcelResult ImportDocument(IHtmlDocument document, HtmlToExcelOptions options) {
        var stream = new MemoryStream();
        ExcelDocument workbook = ExcelDocument.Create(stream);
        var result = new HtmlToExcelResult(workbook);

        List<IElement> sheetSections = document.QuerySelectorAll("section.officeimo-sheet").ToList();
        if (sheetSections.Count == 0) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticContentMissing,
                "No semantic Excel sheet sections were found.", HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Failure);
            workbook.AddWorkSheet("Imported");
            return result;
        }

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement section in sheetSections) {
            ExcelSheet sheet = workbook.AddWorkSheet(GetUniqueSheetName(GetSheetName(section), usedNames));
            result.Sheets++;
            ImportTable(section, sheet, result, options);
            if (options.ImportFormulas) {
                ImportFormulas(section, sheet, result);
            }

            if (options.ImportComments) {
                ImportComments(section, sheet, result);
            }

            if (options.ImportImages || options.ImportChartInventory) {
                ImportDrawings(section, sheet, options, result);
            }

            ApplySheetVisibility(section, sheet);
        }

        return result;
    }

    private static ExcelDocument GetWorkbookOrThrow(HtmlToExcelResult result) {
        if (result.Succeeded) {
            return result.Workbook;
        }

        result.Workbook.Dispose();
        throw new HtmlConversionException(result.Diagnostics.Diagnostics);
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

    private static void ImportTable(IElement section, ExcelSheet sheet, HtmlToExcelResult result, HtmlToExcelOptions options) {
        IElement? table = section.Children.FirstOrDefault(child => IsElement(child, "table"));
        if (table == null) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticBlockMissing,
                "Sheet '" + sheet.Name + "' did not contain a direct semantic table.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        ReadRangeOrigin(section.GetAttribute("data-officeimo-range"), out int firstRow, out int firstColumn);
        ImportTableGrid(table, sheet, result, options, firstRow, firstColumn);
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

    private static void ImportFormulas(IElement section, ExcelSheet sheet, HtmlToExcelResult result) {
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-formulas li[data-officeimo-cell]")) {
            if (!TryParseCellReference(item.GetAttribute("data-officeimo-cell"), out int row, out int column)) {
                continue;
            }

            string formula = item.QuerySelector("code")?.TextContent ?? string.Empty;
            if (formula.Length == 0) {
                continue;
            }

            sheet.CellFormula(row, column, formula);
            result.Formulas++;
        }
    }

    private static void SetCellValue(ExcelSheet sheet, int row, int column, IElement cell, string fallbackText, HtmlToExcelResult result) {
        string? kind = cell.GetAttribute("data-officeimo-value-kind");
        string? rawValue = cell.GetAttribute("data-officeimo-value");
        if (string.IsNullOrWhiteSpace(kind) || rawValue == null) {
            sheet.CellValue(row, column, fallbackText);
            return;
        }

        if (kind!.Equals("number", StringComparison.OrdinalIgnoreCase)) {
            if (double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                sheet.CellValue(row, column, number);
                return;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic number value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (kind.Equals("boolean", StringComparison.OrdinalIgnoreCase)) {
            if (rawValue.Equals("1", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, true);
                return;
            }

            if (rawValue.Equals("0", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("false", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, false);
                return;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic boolean value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (kind.Equals("text", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellValue(row, column, rawValue);
            return;
        } else if (kind.Equals("date-time", StringComparison.OrdinalIgnoreCase)) {
            if (DateTime.TryParse(rawValue, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime dateTime)) {
                sheet.CellValue(row, column, dateTime);
                return;
            }

            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticValueInvalid,
                "Cell " + BuildCellReference(row, column) + " contained a semantic date/time value that could not be parsed and was imported as text.", lossKind: HtmlConversionLossKind.Approximation);
        } else if (kind.Equals("formula", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellFormula(row, column, rawValue);
            return;
        } else if (kind.Equals("error", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellError(row, column, rawValue);
            return;
        }

        sheet.CellValue(row, column, fallbackText);
    }

    private static bool IsSemanticEmptyCell(IElement cell) =>
        string.Equals(cell.GetAttribute("data-officeimo-empty"), "true", StringComparison.OrdinalIgnoreCase);

    private static void ImportComments(IElement section, ExcelSheet sheet, HtmlToExcelResult result) {
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-comments li[data-officeimo-cell]")) {
            if (!TryParseCellReference(item.GetAttribute("data-officeimo-cell"), out int row, out int column)) {
                continue;
            }

            string text = item.QuerySelector("p")?.TextContent?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            sheet.SetComment(row, column, text, ReadAuthor(item));
            result.Comments++;
        }
    }

    private static void ImportDrawings(IElement section, ExcelSheet sheet, HtmlToExcelOptions options, HtmlToExcelResult result) {
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
                ImportImage(drawing.Element, sheet, result);
            } else {
                ImportChart(drawing.Element, sheet, result, range, ref chartIndex);
            }
        }
    }

    private static void ImportImage(IElement item, ExcelSheet sheet, HtmlToExcelResult result) {
        IElement? image = item.QuerySelector("img[src]");
        if (image == null || !HtmlImageDataUri.TryParse(image.GetAttribute("src"), out HtmlImageDataUri dataUri)) {
            return;
        }

        if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "Image inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' could not be decoded.", lossKind: HtmlConversionLossKind.Omission);
            return;
        }

        ReadImagePlacement(item, out int row, out int column, out int width, out int height, out int offsetX, out int offsetY);
        string name = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        string description = NormalizeText(item.QuerySelector("p")?.TextContent);
        if (description.Length == 0) {
            description = NormalizeText(image.GetAttribute("alt"));
        }

        ExcelImage importedImage;
        if (IsAbsoluteImageAnchor(item) && TryReadIntAttribute(item, "data-officeimo-x", out int xPixels) && TryReadIntAttribute(item, "data-officeimo-y", out int yPixels)) {
            importedImage = sheet.AddImageAbsolute(xPixels, yPixels, bytes, dataUri.MediaType, width, height, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        } else if (IsAbsoluteImageAnchor(item)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                "Image inventory item '" + (name.Length == 0 ? "Image" : name) + "' used an absolute anchor without semantic x/y coordinates and was restored to its fallback cell anchor.", lossKind: HtmlConversionLossKind.Approximation);
            importedImage = sheet.AddImage(row, column, bytes, dataUri.MediaType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        } else if (IsTwoCellImageAnchor(item)) {
            importedImage = AddTwoCellImage(item, sheet, result, bytes, dataUri.MediaType, row, column, width, height, offsetX, offsetY, name, description);
        } else {
            importedImage = sheet.AddImage(row, column, bytes, dataUri.MediaType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
        }

        ApplyImageTransforms(item, importedImage);
        result.Images++;
    }

    private static ExcelImage AddTwoCellImage(
        IElement item,
        ExcelSheet sheet,
        HtmlToExcelResult result,
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
            int endRow = Math.Max(row, toRow - 1);
            int endColumn = Math.Max(column, toColumn - 1);
            int endOffsetX = Math.Max(0, ReadOptionalIntAttribute(item, "data-officeimo-to-offset-x") ?? 0);
            int endOffsetY = Math.Max(0, ReadOptionalIntAttribute(item, "data-officeimo-to-offset-y") ?? 0);
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

    private static void ImportChart(IElement item, ExcelSheet sheet, HtmlToExcelResult result, string range, ref int chartIndex) {
        string title = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
        ExcelChartType type = ReadExcelChartType(item);
        ReadChartPlacement(item, chartIndex, out int row, out int column, out int width, out int height);
        try {
            if (TryReadChartData(item, out ExcelChartData? chartData) && chartData != null) {
                sheet.AddChart(chartData, row: row, column: column, widthPixels: width, heightPixels: height, type: type, title: title.Length == 0 ? null : title);
            } else if (!string.IsNullOrWhiteSpace(range) && !string.Equals(range, "A1", StringComparison.OrdinalIgnoreCase)) {
                sheet.AddChartFromRange(range, row: row, column: column, widthPixels: width, heightPixels: height, type: type, title: title.Length == 0 ? null : title);
            } else {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentOmitted,
                    "Chart inventory item '" + title + "' on sheet '" + sheet.Name + "' did not contain semantic chart data and no usable table range was available.", lossKind: HtmlConversionLossKind.Omission);
                return;
            }

            result.Charts++;
            chartIndex++;
        } catch (Exception ex) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ArtifactCreationFailed,
                "Chart inventory item '" + title + "' could not be restored as a native chart: " + ex.Message,
                lossKind: HtmlConversionLossKind.Omission, detail: ex.GetType().Name);
        }
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
                if (!double.TryParse(NormalizeText(valueCell.TextContent), NumberStyles.Float, CultureInfo.InvariantCulture, out values[i])) {
                    return false;
                }

                if (hasXValues) {
                    string? rawXValue = valueCell.GetAttribute("data-officeimo-x");
                    if (rawXValue == null || !double.TryParse(rawXValue, NumberStyles.Float, CultureInfo.InvariantCulture, out xValues[i])) {
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

    private static string GetSheetName(IElement section) {
        string? name = section.GetAttribute("data-officeimo-sheet");
        if (!string.IsNullOrWhiteSpace(name)) {
            return name!.Trim();
        }

        return NormalizeText(section.QuerySelector("h2")?.TextContent) is { Length: > 0 } heading
            ? heading
            : "Sheet";
    }

    private static string GetUniqueSheetName(string name, HashSet<string> usedNames) {
        string baseName = SanitizeSheetName(name);
        string candidate = baseName;
        int suffix = 2;
        while (!usedNames.Add(candidate)) {
            string suffixText = " " + suffix.ToString(CultureInfo.InvariantCulture);
            int maxBaseLength = Math.Max(1, 31 - suffixText.Length);
            candidate = baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) + suffixText : baseName + suffixText;
            suffix++;
        }

        return candidate;
    }

    private static string SanitizeSheetName(string name) {
        string value = string.IsNullOrWhiteSpace(name) ? "Sheet" : name.Trim();
        foreach (char invalid in new[] { ':', '\\', '/', '?', '*', '[', ']' }) {
            value = value.Replace(invalid, '-');
        }

        return value.Length > 31 ? value.Substring(0, 31) : value;
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

    private static void ReadChartPlacement(IElement item, int chartIndex, out int row, out int column, out int width, out int height) {
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
    }

    private static void ReadImagePlacement(IElement item, out int row, out int column, out int width, out int height, out int offsetX, out int offsetY) {
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
    }

    private static void ApplyImageTransforms(IElement item, ExcelImage image) {
        if (TryReadDoubleAttribute(item, "data-officeimo-rotation", out double rotation)) {
            image.SetRotation(rotation);
        }

        bool hasHorizontalFlip = TryReadBoolAttribute(item, "data-officeimo-flip-horizontal", out bool horizontalFlip);
        bool hasVerticalFlip = TryReadBoolAttribute(item, "data-officeimo-flip-vertical", out bool verticalFlip);
        if (hasHorizontalFlip || hasVerticalFlip) {
            image.SetFlip(hasHorizontalFlip && horizontalFlip, hasVerticalFlip && verticalFlip);
        }

        double left = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-left") ?? 0D;
        double top = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-top") ?? 0D;
        double right = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-right") ?? 0D;
        double bottom = ReadOptionalDoubleAttribute(item, "data-officeimo-crop-bottom") ?? 0D;
        if (left > 0D || top > 0D || right > 0D || bottom > 0D) {
            image.SetCropRatio(left, top, right, bottom);
        }
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
            column = column * 26 + (char.ToUpperInvariant(reference[index]) - 'A' + 1);
            index++;
        }

        return column > 0
            && index < reference.Length
            && int.TryParse(reference.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row)
            && row > 0;
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
