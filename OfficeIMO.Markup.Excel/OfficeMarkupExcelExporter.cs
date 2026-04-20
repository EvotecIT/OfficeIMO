using OfficeIMO.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.Excel;

public sealed class OfficeMarkupExcelExporter {
    public void Export(OfficeMarkupDocument document, OfficeMarkupExcelExportOptions options) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (document.Profile != OfficeMarkupProfile.Workbook) {
            throw new InvalidOperationException("Excel export requires the Workbook OfficeIMO markup profile.");
        }

        if (string.IsNullOrWhiteSpace(options.OutputPath)) {
            throw new InvalidOperationException("Excel export requires an output path.");
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(options.OutputPath));
        if (!string.IsNullOrEmpty(directory)) {
            Directory.CreateDirectory(directory);
        }

        using var workbook = ExcelDocument.Create(options.OutputPath);
        var context = new WorkbookExportContext(workbook, options);

        foreach (var block in document.Blocks) {
            ExportBlock(context, block);
        }

        context.EnsureCurrentSheet();
        FinalizeWorkbook(context);
        workbook.Save(options.OutputPath, openExcel: false, options: new ExcelSaveOptions {
            SafePreflight = options.SafePreflight,
            ValidateOpenXml = options.ValidateOpenXml,
            SafeRepairDefinedNames = options.SafeRepairDefinedNames
        });
    }

    private static void ExportBlock(WorkbookExportContext context, OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupSheetBlock sheet:
                context.SetCurrentSheet(sheet.Name);
                break;
            case OfficeMarkupRangeBlock range:
                AddRange(context, range);
                break;
            case OfficeMarkupFormulaBlock formula:
                AddFormula(context, formula);
                break;
            case OfficeMarkupNamedTableBlock table:
                AddTable(context, table);
                break;
            case OfficeMarkupChartBlock chart:
                AddChart(context, chart);
                break;
            case OfficeMarkupFormattingBlock formatting:
                AddFormatting(context, formatting);
                break;
            case OfficeMarkupHeadingBlock heading when context.Options.IncludeMarkdownAsWorksheetText:
                AddWorksheetText(context, heading.Text, bold: true);
                break;
            case OfficeMarkupParagraphBlock paragraph when context.Options.IncludeMarkdownAsWorksheetText:
                AddWorksheetText(context, paragraph.Text, bold: false);
                break;
            case OfficeMarkupListBlock list when context.Options.IncludeMarkdownAsWorksheetText:
                foreach (var item in list.Items) {
                    AddWorksheetText(context, "- " + item.Text, bold: false);
                }

                break;
            case OfficeMarkupTableBlock table when context.Options.IncludeMarkdownAsWorksheetText:
                AddMarkdownTable(context, table);
                break;
        }
    }

    private static void AddRange(WorkbookExportContext context, OfficeMarkupRangeBlock range) {
        var (sheet, address) = ResolveTargetReference(context, range.Sheet, range.Address);
        var (row, column) = ParseCell(address, "range address");
        var currentRow = row;
        foreach (var valueRow in range.Values) {
            var currentColumn = column;
            foreach (var value in valueRow) {
                SetCell(sheet, currentRow, currentColumn, value);
                currentColumn++;
            }

            currentRow++;
        }

        context.Touch(sheet.Name, currentRow);
    }

    private static void AddFormula(WorkbookExportContext context, OfficeMarkupFormulaBlock formula) {
        var (sheet, cell) = ResolveTargetReference(context, formula.Sheet, formula.Cell);
        var (row, column) = ParseCell(cell, "formula cell");
        sheet.CellFormula(row, column, formula.Expression);
        context.Touch(sheet.Name, row + 1);
    }

    private static void AddTable(WorkbookExportContext context, OfficeMarkupNamedTableBlock table) {
        var (sheet, rangeAddress) = ResolveTargetReference(context, GetAttribute(table.Attributes, "sheet"), table.Range);
        var style = ParseTableStyle(GetAttribute(table.Attributes, "style"));
        sheet.AddTable(rangeAddress, table.HasHeader, table.Name, style);
        if (table.HasHeader && A1.TryParseRange(rangeAddress, out var startRow, out _, out _, out _)) {
            context.RegisterFreezeRow(sheet.Name, startRow);
        }
    }

    private static void AddChart(WorkbookExportContext context, OfficeMarkupChartBlock chart) {
        var (sheet, placementCell) = ResolveChartTarget(context, chart);
        var (row, column, width, height) = ResolveChartPlacement(context.Options, chart, placementCell);
        var type = ParseChartType(chart.ChartType);

        if (!string.IsNullOrWhiteSpace(chart.Source)) {
            var (sourceSheet, source) = ResolveChartSource(context, sheet, chart.Source!);
            if (source.IndexOf(':') >= 0) {
                var range = CreateChartDataRange(sourceSheet.Name, source);
                var excelChart = sheet.AddChart(range, row, column, width, height, type, title: chart.Title);
                ApplyChartStyle(context, excelChart, chart, range.SeriesCount);
            } else {
                var tableRange = sourceSheet.GetTableRange(source);
                if (string.IsNullOrWhiteSpace(tableRange)) {
                    throw new InvalidOperationException($"Table '{source}' was not found on sheet '{sourceSheet.Name}'.");
                }

                var range = CreateChartDataRange(sourceSheet.Name, tableRange!);
                var excelChart = sheet.AddChart(range, row, column, width, height, type, title: chart.Title);
                ApplyChartStyle(context, excelChart, chart, range.SeriesCount);
            }

            return;
        }

        if (chart.Data.Count > 1) {
            var data = ToChartData(chart);
            var excelChart = sheet.AddChart(data, row, column, width, height, type, chart.Title);
            ApplyChartStyle(context, excelChart, chart, data.Series.Count);
        }
    }

    private static void AddFormatting(WorkbookExportContext context, OfficeMarkupFormattingBlock formatting) {
        if (string.IsNullOrWhiteSpace(formatting.Target)) {
            return;
        }

        var (sheet, target) = ResolveTargetReference(context, GetAttribute(formatting.Attributes, "sheet"), formatting.Target);
        foreach (var (row, column) in EnumerateCells(target)) {
            if (!string.IsNullOrWhiteSpace(formatting.NumberFormat)) {
                sheet.FormatCell(row, column, formatting.NumberFormat!);
            }

            var fill = GetAttribute(formatting.Attributes, "fill") ?? GetAttribute(formatting.Attributes, "background");
            if (!string.IsNullOrWhiteSpace(fill)) {
                sheet.CellBackground(row, column, fill!);
            }

            var fontColor = GetAttribute(formatting.Attributes,
                "color", "font-color", "fontColor", "text-color", "textColor", "textcolor");
            if (!string.IsNullOrWhiteSpace(fontColor)) {
                sheet.CellFontColor(row, column, fontColor!);
            }

            var bold = GetAttribute(formatting.Attributes, "bold");
            if (!string.IsNullOrWhiteSpace(bold) && IsTruthy(bold)) {
                sheet.CellBold(row, column, true);
            }

            var italic = GetAttribute(formatting.Attributes, "italic");
            if (!string.IsNullOrWhiteSpace(italic) && IsTruthy(italic)) {
                sheet.CellItalic(row, column, true);
            }

            var underline = GetAttribute(formatting.Attributes, "underline");
            if (!string.IsNullOrWhiteSpace(underline) && IsTruthy(underline)) {
                sheet.CellUnderline(row, column, true);
            }

            var alignment = GetAttribute(formatting.Attributes,
                "align", "alignment", "horizontal-align", "horizontalAlign", "horizontalalignment", "text-align", "textAlign");
            if (TryParseHorizontalAlignment(alignment, out var horizontalAlignment)) {
                sheet.CellAlign(row, column, horizontalAlignment);
            }

            var verticalAlignment = GetAttribute(formatting.Attributes,
                "vertical-align", "verticalAlign", "verticalalignment", "valign");
            if (TryParseVerticalAlignment(verticalAlignment, out var parsedVerticalAlignment)) {
                sheet.CellVerticalAlign(row, column, parsedVerticalAlignment);
            }

            var wrap = GetAttribute(formatting.Attributes, "wrap", "wrap-text", "wrapText");
            if (!string.IsNullOrWhiteSpace(wrap) && IsTruthy(wrap)) {
                sheet.WrapCells(row, row, column);
            }

            var border = GetAttribute(formatting.Attributes, "border", "border-style", "borderStyle");
            if (TryParseBorderStyle(border, out var borderStyle)) {
                var borderColor = GetAttribute(formatting.Attributes, "border-color", "borderColor", "line-color", "lineColor");
                sheet.CellBorder(row, column, borderStyle, borderColor);
            }
        }
    }

    private static void AddWorksheetText(WorkbookExportContext context, string text, bool bold) {
        if (string.IsNullOrWhiteSpace(text)) {
            return;
        }

        var sheet = context.EnsureCurrentSheet();
        var row = context.NextTextRow(sheet.Name);
        sheet.CellValue(row, 1, text);
        if (bold) {
            sheet.CellBold(row, 1, true);
            if (context.Options.StyleMarkdownHeadings) {
                sheet.CellBackground(row, 1, "1F4E78");
                sheet.CellFontColor(row, 1, "FFFFFF");
            }
        }

        context.Touch(sheet.Name, row + 1);
    }

    private static void AddMarkdownTable(WorkbookExportContext context, OfficeMarkupTableBlock table) {
        var sheet = context.EnsureCurrentSheet();
        var row = context.NextTextRow(sheet.Name);
        var currentRow = row;

        if (table.Headers.Count > 0) {
            for (var index = 0; index < table.Headers.Count; index++) {
                sheet.CellValue(currentRow, index + 1, table.Headers[index]);
                sheet.CellBold(currentRow, index + 1, true);
            }

            currentRow++;
        }

        foreach (var valueRow in table.Rows) {
            for (var index = 0; index < valueRow.Count; index++) {
                SetCell(sheet, currentRow, index + 1, valueRow[index]);
            }

            currentRow++;
        }

        context.Touch(sheet.Name, currentRow);
    }

    private static void FinalizeWorkbook(WorkbookExportContext context) {
        foreach (var sheet in context.Sheets) {
            if (context.Options.FreezeTableHeaderRows && context.TryGetFreezeRow(sheet.Name, out var freezeRow)) {
                sheet.Freeze(topRows: freezeRow, leftCols: 0);
            }

            if (context.Options.HideGridlines) {
                sheet.SetGridlinesVisible(false);
            }

            if (context.Options.AutoFitColumns) {
                sheet.AutoFitColumns();
            }
        }
    }

    private static void ApplyChartStyle(WorkbookExportContext context, ExcelChart chart, OfficeMarkupChartBlock source, int seriesCount) {
        if (!context.Options.StyleCharts) {
            return;
        }

        var seriesColors = ResolveChartPalette(source);
        var normalizedType = Normalize(source.ChartType);

        for (var index = 0; index < seriesCount; index++) {
            var color = seriesColors[index % seriesColors.Count];
            if (normalizedType == "line") {
                chart.SetSeriesLineColor(index, color, widthPoints: 2.25);
            } else {
                chart.SetSeriesFillColor(index, color);
                chart.SetSeriesLineColor(index, color, widthPoints: 0.5);
            }
        }

        ApplyChartSemanticOptions(chart, source, normalizedType);
    }

    private static void ApplyChartSemanticOptions(ExcelChart chart, OfficeMarkupChartBlock source, string normalizedType) {
        var categoryTitle = GetAttribute(source.Attributes,
            "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle");
        if (!string.IsNullOrWhiteSpace(categoryTitle)) {
            chart.SetCategoryAxisTitle(categoryTitle!);
        }

        var valueTitle = GetAttribute(source.Attributes,
            "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle");
        if (!string.IsNullOrWhiteSpace(valueTitle)) {
            chart.SetValueAxisTitle(valueTitle!);
        }

        var categoryFormat = GetAttribute(source.Attributes,
            "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat");
        if (!string.IsNullOrWhiteSpace(categoryFormat)) {
            chart.SetCategoryAxisNumberFormat(categoryFormat!);
        }

        var valueFormat = GetAttribute(source.Attributes,
            "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat");
        if (!string.IsNullOrWhiteSpace(valueFormat)) {
            chart.SetValueAxisNumberFormat(valueFormat!);
        }

        ApplyLegendOptions(chart, source);
        ApplyDataLabelOptions(chart, source);
        ApplyGridlineOptions(chart, source, normalizedType);

        chart.SetTitleTextStyle(fontSizePoints: 14, bold: true, color: "172033", fontName: "Aptos Display");
        chart.SetLegendTextStyle(fontSizePoints: 9, color: "475569", fontName: "Aptos");
    }

    private static void ApplyLegendOptions(ExcelChart chart, OfficeMarkupChartBlock source) {
        var legend = GetAttribute(source.Attributes, "legend", "legend-position", "legendPosition");
        if (string.IsNullOrWhiteSpace(legend)) {
            return;
        }

        var normalized = Normalize(legend!);
        if (normalized == "false" || normalized == "none" || normalized == "hidden" || normalized == "off") {
            chart.HideLegend();
            return;
        }

        if (TryParseLegendPosition(normalized, out var position)) {
            chart.SetLegend(position);
        }
    }

    private static void ApplyDataLabelOptions(ExcelChart chart, OfficeMarkupChartBlock source) {
        var labels = GetAttribute(source.Attributes, "labels", "data-labels", "dataLabels");
        if (!IsTruthy(labels)) {
            return;
        }

        var labelPosition = GetAttribute(source.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
        var labelFormat = GetAttribute(source.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
        chart.SetDataLabels(
            showValue: true,
            showCategoryName: false,
            showSeriesName: false,
            showLegendKey: false,
            showPercent: false,
            position: TryParseDataLabelPosition(labelPosition, out var parsedPosition) ? parsedPosition : null,
            numberFormat: string.IsNullOrWhiteSpace(labelFormat) ? null : labelFormat,
            sourceLinked: false);
        chart.SetDataLabelTextStyle(fontSizePoints: 9, color: "334155", fontName: "Aptos");
    }

    private static void ApplyGridlineOptions(ExcelChart chart, OfficeMarkupChartBlock source, string normalizedType) {
        var allGridlines = GetAttribute(source.Attributes, "gridlines");
        var valueGridlines = GetAttribute(source.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? allGridlines;
        var categoryGridlines = GetAttribute(source.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");

        if (valueGridlines != null) {
            chart.SetValueAxisGridlines(showMajor: IsTruthy(valueGridlines), showMinor: false, lineColor: "CBD5E1", lineWidthPoints: 0.5);
        }

        if (categoryGridlines != null) {
            chart.SetCategoryAxisGridlines(showMajor: IsTruthy(categoryGridlines), showMinor: false, lineColor: "E2E8F0", lineWidthPoints: 0.5);
        }
    }

    private static IReadOnlyList<string> ResolveChartPalette(OfficeMarkupChartBlock chart) {
        if (chart.Attributes.TryGetValue("palette", out var palette) && !string.IsNullOrWhiteSpace(palette)) {
            var colors = palette.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(ToExcelColor)
                .Where(color => !string.IsNullOrWhiteSpace(color))
                .Cast<string>()
                .ToList();
            if (colors.Count > 0) {
                return colors;
            }
        }

        return new[] { "2563EB", "F97316", "10B981", "A855F7", "EF4444", "14B8A6" };
    }

    private static ExcelChartData ToChartData(OfficeMarkupChartBlock chart) {
        var header = chart.Data[0];
        if (header.Count < 2) {
            throw new InvalidOperationException("Chart data requires one category column and at least one value column.");
        }

        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        var series = new List<ExcelChartSeries>();
        for (var column = 1; column < header.Count; column++) {
            var values = new List<double>();
            foreach (var row in chart.Data.Skip(1)) {
                values.Add(row.Count > column && double.TryParse(row[column], NumberStyles.Any, CultureInfo.InvariantCulture, out var number) ? number : 0d);
            }

            series.Add(new ExcelChartSeries(header[column], values));
        }

        return new ExcelChartData(categories, series);
    }

    private static void SetCell(ExcelSheet sheet, int row, int column, string value) {
        if (value != null && value.StartsWith("=", StringComparison.Ordinal)) {
            sheet.CellFormula(row, column, value);
            return;
        }

        sheet.CellValue(row, column, ConvertValue(value ?? string.Empty));
    }

    private static object ConvertValue(string value) {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) {
            return number;
        }

        if (bool.TryParse(value, out var boolean)) {
            return boolean;
        }

        return value ?? string.Empty;
    }

    private static (ExcelSheet Sheet, string Reference) ResolveTargetReference(WorkbookExportContext context, string? explicitSheet, string reference) {
        if (TrySplitSheetQualifiedReference(reference, out var sheetName, out var localReference)) {
            return (context.GetSheet(sheetName, setCurrent: false), localReference);
        }

        return (context.GetSheet(explicitSheet, setCurrent: string.IsNullOrWhiteSpace(explicitSheet)), (reference ?? string.Empty).Trim());
    }

    private static (ExcelSheet Sheet, string Source) ResolveChartSource(WorkbookExportContext context, ExcelSheet fallbackSheet, string source) {
        if (TrySplitSheetQualifiedReference(source, out var sheetName, out var localReference)) {
            return (context.GetSheet(sheetName, setCurrent: false), localReference);
        }

        return (fallbackSheet, (source ?? string.Empty).Trim());
    }

    private static (ExcelSheet Sheet, string? PlacementCell) ResolveChartTarget(WorkbookExportContext context, OfficeMarkupChartBlock chart) {
        var cell = GetAttribute(chart.Attributes, "cell");
        if (TrySplitSheetQualifiedReference(cell, out var sheetName, out var localCell)) {
            return (context.GetSheet(sheetName, setCurrent: false), localCell);
        }

        return (context.GetSheet(chart.Sheet, setCurrent: string.IsNullOrWhiteSpace(chart.Sheet)), cell);
    }

    private static bool TrySplitSheetQualifiedReference(string? reference, out string sheetName, out string localReference) {
        sheetName = string.Empty;
        localReference = string.Empty;
        if (string.IsNullOrWhiteSpace(reference)) {
            return false;
        }

        var value = reference!.Trim();
        var bangIndex = value.LastIndexOf('!');
        if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
            return false;
        }

        sheetName = value.Substring(0, bangIndex).Trim().Trim('\'').Replace("''", "'");
        localReference = value.Substring(bangIndex + 1).Trim();
        return !string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(localReference);
    }

    private static IEnumerable<(int Row, int Column)> EnumerateCells(string target) {
        if (A1.TryParseRange(target, out var r1, out var c1, out var r2, out var c2)) {
            for (var row = r1; row <= r2; row++) {
                for (var column = c1; column <= c2; column++) {
                    yield return (row, column);
                }
            }

            yield break;
        }

        var cell = A1.ParseCellRef(target);
        if (cell.Row <= 0 || cell.Col <= 0) {
            throw new InvalidOperationException($"Invalid cell or range reference '{target}'.");
        }

        yield return (cell.Row, cell.Col);
    }

    private static (int Row, int Column) ParseCell(string value, string description) {
        var cell = A1.ParseCellRef(value);
        if (cell.Row <= 0 || cell.Col <= 0) {
            throw new InvalidOperationException($"Invalid {description} '{value}'.");
        }

        return (cell.Row, cell.Col);
    }

    private static (int Row, int Column, int Width, int Height) ResolveChartPlacement(OfficeMarkupExcelExportOptions options, OfficeMarkupChartBlock chart, string? placementCell) {
        var row = GetInt(chart.Attributes, "row")
            ?? GetInt(chart.Attributes, "y")
            ?? options.DefaultChartRow;
        var column = GetInt(chart.Attributes, "column")
            ?? GetInt(chart.Attributes, "col")
            ?? GetInt(chart.Attributes, "x")
            ?? options.DefaultChartColumn;
        var cell = placementCell ?? GetAttribute(chart.Attributes, "cell");
        if (!string.IsNullOrWhiteSpace(cell)) {
            var parsed = A1.ParseCellRef(cell!);
            if (parsed.Row > 0 && parsed.Col > 0) {
                row = parsed.Row;
                column = parsed.Col;
            }
        }

        var width = GetInt(chart.Attributes, "width")
            ?? GetInt(chart.Attributes, "w")
            ?? options.DefaultChartWidthPixels;
        var height = GetInt(chart.Attributes, "height")
            ?? GetInt(chart.Attributes, "h")
            ?? options.DefaultChartHeightPixels;

        return (Math.Max(1, row), Math.Max(1, column), Math.Max(1, width), Math.Max(1, height));
    }

    private static int? GetInt(IDictionary<string, string> attributes, string name) {
        var value = GetAttribute(attributes, name);
        if (string.IsNullOrWhiteSpace(value) || value!.IndexOf('%') >= 0) {
            return null;
        }

        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result) ? result : null;
    }

    private static string Normalize(string value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

    private static string? ToExcelColor(string? color) {
        if (string.IsNullOrWhiteSpace(color)) {
            return null;
        }

        color = color!.Trim();
        if (color.StartsWith("#", StringComparison.Ordinal)) {
            color = color.Substring(1);
        }

        return color.Length == 6 && color.All(IsHexDigit) ? color.ToUpperInvariant() : null;
    }

    private static bool IsHexDigit(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'a' && value <= 'f')
        || (value >= 'A' && value <= 'F');

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static string? GetAttribute(IDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value)) {
                return value;
            }
        }

        return null;
    }

    private static bool IsTruthy(string? value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);

    private static bool TryParseHorizontalAlignment(string? value, out DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues alignment) {
        switch (Normalize(value ?? string.Empty)) {
            case "general":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.General;
                return true;
            case "left":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Left;
                return true;
            case "center":
            case "centre":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center;
                return true;
            case "centercontinuous":
            case "centeracross":
            case "centeracrossselection":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.CenterContinuous;
                return true;
            case "right":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right;
                return true;
            case "fill":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Fill;
                return true;
            case "justify":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Justify;
                return true;
            case "distributed":
                alignment = DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Distributed;
                return true;
            default:
                alignment = default;
                return false;
        }
    }

    private static bool TryParseVerticalAlignment(string? value, out DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues alignment) {
        switch (Normalize(value ?? string.Empty)) {
            case "top":
                alignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Top;
                return true;
            case "middle":
            case "center":
            case "centre":
                alignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center;
                return true;
            case "bottom":
                alignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Bottom;
                return true;
            case "justify":
                alignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Justify;
                return true;
            case "distributed":
                alignment = DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Distributed;
                return true;
            default:
                alignment = default;
                return false;
        }
    }

    private static bool TryParseBorderStyle(string? value, out DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues style) {
        switch (Normalize(value ?? string.Empty)) {
            case "true":
            case "yes":
            case "on":
            case "1":
            case "thin":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin;
                return true;
            case "medium":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Medium;
                return true;
            case "thick":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thick;
                return true;
            case "dashed":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Dashed;
                return true;
            case "dotted":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Dotted;
                return true;
            case "double":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Double;
                return true;
            case "hair":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Hair;
                return true;
            case "dashdot":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.DashDot;
                return true;
            case "dashdotdot":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.DashDotDot;
                return true;
            case "mediumdashed":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.MediumDashed;
                return true;
            case "mediumdashdot":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.MediumDashDot;
                return true;
            case "mediumdashdotdot":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.MediumDashDotDot;
                return true;
            case "slantdashdot":
                style = DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.SlantDashDot;
                return true;
            default:
                style = default;
                return false;
        }
    }

    private static bool TryParseLegendPosition(string value, out C.LegendPositionValues position) {
        switch (Normalize(value)) {
            case "left":
                position = C.LegendPositionValues.Left;
                return true;
            case "right":
                position = C.LegendPositionValues.Right;
                return true;
            case "top":
                position = C.LegendPositionValues.Top;
                return true;
            case "bottom":
                position = C.LegendPositionValues.Bottom;
                return true;
            case "corner":
            case "topright":
                position = C.LegendPositionValues.TopRight;
                return true;
            default:
                position = default;
                return false;
        }
    }

    private static bool TryParseDataLabelPosition(string? value, out C.DataLabelPositionValues position) {
        switch (Normalize(value ?? string.Empty)) {
            case "bestfit":
                position = C.DataLabelPositionValues.BestFit;
                return true;
            case "bottom":
                position = C.DataLabelPositionValues.Bottom;
                return true;
            case "center":
                position = C.DataLabelPositionValues.Center;
                return true;
            case "insidebase":
                position = C.DataLabelPositionValues.InsideBase;
                return true;
            case "insideend":
                position = C.DataLabelPositionValues.InsideEnd;
                return true;
            case "left":
                position = C.DataLabelPositionValues.Left;
                return true;
            case "outsideend":
                position = C.DataLabelPositionValues.OutsideEnd;
                return true;
            case "right":
                position = C.DataLabelPositionValues.Right;
                return true;
            case "top":
                position = C.DataLabelPositionValues.Top;
                return true;
            default:
                position = default;
                return false;
        }
    }

    private static ExcelChartDataRange CreateChartDataRange(string sheetName, string source) {
        if (!A1.TryParseRange(source, out var r1, out var c1, out var r2, out var c2)) {
            throw new InvalidOperationException($"Invalid chart source range '{source}'.");
        }

        var categoryCount = r2 - r1;
        var seriesCount = c2 - c1;
        if (categoryCount <= 0 || seriesCount <= 0) {
            throw new InvalidOperationException("Chart source range must include a header row, a category column, and at least one value column.");
        }

        return new ExcelChartDataRange(sheetName, r1, c1, categoryCount, seriesCount, hasHeaderRow: true);
    }

    private static TableStyle ParseTableStyle(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return TableStyle.TableStyleMedium2;
        }

        if (Enum.TryParse<TableStyle>(value, true, out var direct)) {
            return direct;
        }

        var normalized = value!.Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty);
        if (!normalized.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase)) {
            normalized = "TableStyle" + normalized;
        }

        return Enum.TryParse<TableStyle>(normalized, true, out var parsed) ? parsed : TableStyle.TableStyleMedium2;
    }

    private static ExcelChartType ParseChartType(string value) {
        var normalized = (value ?? string.Empty).Trim().Replace("-", string.Empty).Replace("_", string.Empty).ToLowerInvariant();
        return normalized switch {
            "bar" or "barclustered" => ExcelChartType.BarClustered,
            "barstacked" or "stackedbar" => ExcelChartType.BarStacked,
            "columnstacked" or "stackedcolumn" => ExcelChartType.ColumnStacked,
            "line" => ExcelChartType.Line,
            "area" => ExcelChartType.Area,
            "pie" => ExcelChartType.Pie,
            "doughnut" or "donut" => ExcelChartType.Doughnut,
            "scatter" or "xy" => ExcelChartType.Scatter,
            "bubble" => ExcelChartType.Bubble,
            _ => ExcelChartType.ColumnClustered
        };
    }

    private sealed class WorkbookExportContext {
        private readonly ExcelDocument _workbook;
        private readonly Dictionary<string, ExcelSheet> _sheets = new Dictionary<string, ExcelSheet>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _nextRows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _freezeRows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private ExcelSheet? _currentSheet;

        public WorkbookExportContext(ExcelDocument workbook, OfficeMarkupExcelExportOptions options) {
            _workbook = workbook;
            Options = options;
        }

        public OfficeMarkupExcelExportOptions Options { get; }

        public IEnumerable<ExcelSheet> Sheets => _sheets.Values.Distinct();

        public void SetCurrentSheet(string name) {
            _currentSheet = GetSheet(name);
        }

        public ExcelSheet EnsureCurrentSheet() {
            _currentSheet ??= GetSheet(Options.DefaultSheetName);
            return _currentSheet;
        }

        public ExcelSheet GetSheet(string? name, bool setCurrent = true) {
            var sheetName = string.IsNullOrWhiteSpace(name) ? _currentSheet?.Name ?? Options.DefaultSheetName : name!.Trim();
            if (_sheets.TryGetValue(sheetName, out var sheet)) {
                if (setCurrent || _currentSheet == null) {
                    _currentSheet = sheet;
                }

                return sheet;
            }

            sheet = _workbook.AddWorkSheet(sheetName);
            _sheets[sheetName] = sheet;
            _sheets[sheet.Name] = sheet;
            _nextRows[sheet.Name] = 1;
            if (setCurrent || _currentSheet == null) {
                _currentSheet = sheet;
            }

            return sheet;
        }

        public int NextTextRow(string sheetName) =>
            _nextRows.TryGetValue(sheetName, out var row) ? row : 1;

        public void Touch(string sheetName, int nextRow) {
            if (!_nextRows.TryGetValue(sheetName, out var existing) || nextRow > existing) {
                _nextRows[sheetName] = nextRow;
            }
        }

        public void RegisterFreezeRow(string sheetName, int row) {
            if (row <= 0) {
                return;
            }

            if (!_freezeRows.TryGetValue(sheetName, out var existing) || row > existing) {
                _freezeRows[sheetName] = row;
            }
        }

        public bool TryGetFreezeRow(string sheetName, out int row) =>
            _freezeRows.TryGetValue(sheetName, out row);
    }
}
