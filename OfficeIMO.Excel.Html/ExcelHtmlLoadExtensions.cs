using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Extension methods for importing semantic OfficeIMO Excel HTML.
/// </summary>
public static class ExcelHtmlLoadExtensions {
    /// <summary>
    /// Imports semantic OfficeIMO Excel HTML into a native workbook.
    /// </summary>
    public static ExcelDocument LoadExcelFromHtml(this string html, ExcelHtmlLoadOptions? options = null) =>
        LoadExcelFromHtmlWithResult(html, options).Workbook;

    /// <summary>
    /// Imports semantic OfficeIMO Excel HTML into a native workbook and returns import evidence.
    /// </summary>
    public static ExcelHtmlLoadResult LoadExcelFromHtmlWithResult(this string html, ExcelHtmlLoadOptions? options = null) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        options ??= new ExcelHtmlLoadOptions();

        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var stream = new MemoryStream();
        ExcelDocument workbook = ExcelDocument.Create(stream);
        var result = new ExcelHtmlLoadResult(workbook);

        List<IElement> sheetSections = document.QuerySelectorAll("section.officeimo-sheet").ToList();
        if (sheetSections.Count == 0) {
            result.Diagnostics.Add("No semantic Excel sheet sections were found.");
            workbook.AddWorkSheet("Imported");
            return result;
        }

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IElement section in sheetSections) {
            ExcelSheet sheet = workbook.AddWorkSheet(GetUniqueSheetName(GetSheetName(section), usedNames));
            result.Sheets++;
            ImportTable(section, sheet, result);
            if (options.ImportFormulas) {
                ImportFormulas(section, sheet, result);
            }

            if (options.ImportComments) {
                ImportComments(section, sheet, result);
            }

            if (options.ImportImages) {
                ImportImages(section, sheet, result);
            }

            if (options.ImportChartInventory) {
                ImportCharts(section, sheet, result);
            }
        }

        return result;
    }

    private static void ImportTable(IElement section, ExcelSheet sheet, ExcelHtmlLoadResult result) {
        IElement? table = section.Children.FirstOrDefault(child => IsElement(child, "table"));
        if (table == null) {
            result.Diagnostics.Add("Sheet '" + sheet.Name + "' did not contain a direct semantic table.");
            return;
        }

        ReadRangeOrigin(section.GetAttribute("data-officeimo-range"), out int firstRow, out int firstColumn);
        int rowIndex = firstRow;
        foreach (IElement row in table.QuerySelectorAll("tr")) {
            int columnIndex = firstColumn;
            foreach (IElement cell in row.Children.Where(child => IsElement(child, "th") || IsElement(child, "td"))) {
                string text = NormalizeText(cell.TextContent);
                if (!IsSemanticEmptyCell(cell) && (text.Length > 0 || cell.GetAttribute("data-officeimo-value") != null)) {
                    SetCellValue(sheet, rowIndex, columnIndex, cell, text, result);
                    result.Cells++;
                }

                columnIndex++;
            }

            rowIndex++;
        }
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

    private static void ImportFormulas(IElement section, ExcelSheet sheet, ExcelHtmlLoadResult result) {
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

    private static void SetCellValue(ExcelSheet sheet, int row, int column, IElement cell, string fallbackText, ExcelHtmlLoadResult result) {
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

            result.Diagnostics.Add("Cell " + BuildCellReference(row, column) + " contained a semantic number value that could not be parsed and was imported as text.");
        } else if (kind.Equals("boolean", StringComparison.OrdinalIgnoreCase)) {
            if (rawValue.Equals("1", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, true);
                return;
            }

            if (rawValue.Equals("0", StringComparison.OrdinalIgnoreCase) || rawValue.Equals("false", StringComparison.OrdinalIgnoreCase)) {
                sheet.CellValue(row, column, false);
                return;
            }

            result.Diagnostics.Add("Cell " + BuildCellReference(row, column) + " contained a semantic boolean value that could not be parsed and was imported as text.");
        } else if (kind.Equals("text", StringComparison.OrdinalIgnoreCase)) {
            sheet.CellValue(row, column, rawValue);
            return;
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

    private static void ImportComments(IElement section, ExcelSheet sheet, ExcelHtmlLoadResult result) {
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

    private static void ImportImages(IElement section, ExcelSheet sheet, ExcelHtmlLoadResult result) {
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-images li")) {
            IElement? image = item.QuerySelector("img[src]");
            if (image == null || !HtmlImageDataUri.TryParse(image.GetAttribute("src"), out HtmlImageDataUri dataUri)) {
                continue;
            }

            if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                result.Diagnostics.Add("Image inventory item '" + NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent) + "' could not be decoded.");
                continue;
            }

            ReadImagePlacement(item, out int row, out int column, out int width, out int height, out int offsetX, out int offsetY);
            string name = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
            string description = NormalizeText(item.QuerySelector("p")?.TextContent);
            if (description.Length == 0) {
                description = NormalizeText(image.GetAttribute("alt"));
            }

            sheet.AddImage(row, column, bytes, dataUri.MediaType, width, height, offsetX, offsetY, name: name.Length == 0 ? null : name, altText: description.Length == 0 ? null : description);
            result.Images++;
        }
    }

    private static void ImportCharts(IElement section, ExcelSheet sheet, ExcelHtmlLoadResult result) {
        string range = section.GetAttribute("data-officeimo-range") ?? sheet.GetUsedRangeA1();
        int chartIndex = 0;
        foreach (IElement item in section.QuerySelectorAll("section.officeimo-charts li")) {
            string title = NormalizeText(item.QuerySelector(".officeimo-feature-label")?.TextContent);
            ExcelChartType type = ReadExcelChartType(item);
            try {
                if (TryReadChartData(item, out ExcelChartData? chartData) && chartData != null) {
                    sheet.AddChart(chartData, row: 1 + chartIndex * 12, column: 6, widthPixels: 320, heightPixels: 180, type: type, title: title.Length == 0 ? null : title);
                } else if (!string.IsNullOrWhiteSpace(range) && !string.Equals(range, "A1", StringComparison.OrdinalIgnoreCase)) {
                    sheet.AddChartFromRange(range, row: 1 + chartIndex * 12, column: 6, widthPixels: 320, heightPixels: 180, type: type, title: title.Length == 0 ? null : title);
                } else {
                    result.Diagnostics.Add("Chart inventory item '" + title + "' on sheet '" + sheet.Name + "' did not contain semantic chart data and no usable table range was available.");
                    continue;
                }

                result.Charts++;
                chartIndex++;
            } catch (Exception ex) {
                result.Diagnostics.Add("Chart inventory item '" + title + "' could not be restored as a native chart: " + ex.Message);
            }
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
            if (cells.Count != categories.Count + 1) {
                return false;
            }

            string name = NormalizeText(cells[0].TextContent);
            var values = new double[categories.Count];
            var xValues = new double[categories.Count];
            bool hasXValues = cells.Skip(1).Any(cell => cell.GetAttribute("data-officeimo-x") != null);
            for (int i = 0; i < categories.Count; i++) {
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

    private static bool IsElement(IElement element, string name) =>
        string.Equals(element.LocalName, name, StringComparison.OrdinalIgnoreCase);

    private static string NormalizeText(string? text) =>
        string.IsNullOrWhiteSpace(text) ? string.Empty : string.Join(" ", text!.Split((char[]?)null!, StringSplitOptions.RemoveEmptyEntries));
}
