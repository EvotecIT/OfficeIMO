using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void ImportGenericDocument(
        HtmlSemanticDocument document,
        ExcelDocument workbook,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget) {
        IReadOnlyList<HtmlSemanticBlock> tables = document.RootTables;
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (tables.Count > 0) {
            for (int index = 0; index < tables.Count; index++) {
                if (!budget.TryReserveSemanticContainer(out string containerLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Additional HTML tables were omitted because the shared worksheet limit was reached.",
                        HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: containerLimit);
                    break;
                }
                if (!budget.TryReserveTable(out string tableLimit)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Additional HTML tables were omitted because the shared table limit was reached.",
                        HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: tableLimit);
                    break;
                }

                string title = tables[index].Table?.Caption ?? "Table " + (index + 1);
                ExcelSheet sheet = workbook.AddWorksheet(GetUniqueSheetName(title, usedNames));
                result.Sheets++;
                ImportTableGrid(
                    tables[index].SourceElement,
                    sheet,
                    result,
                    options,
                    budget,
                    1,
                    1,
                    importedFormulaCells: null,
                    useSemanticValues: false);
                ApplySemanticTableFormatting(tables[index].Table, sheet, result, budget, 1, 1);
            }
        }

        bool hasNarrative = tables.Count == 0 || document.Sections
            .Any(section => section.Blocks.Any(block => IsSectionNarrativeBlock(section, block)));
        ExcelSheet? narrativeSheet = null;
        int row = 1;
        if (hasNarrative) {
            if (!budget.TryReserveSemanticContainer(out string textContainerLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "HTML text could not be imported because the shared worksheet limit was reached.",
                    HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: textContainerLimit);
            } else {
                narrativeSheet = workbook.AddWorksheet(GetUniqueSheetName("Imported", usedNames));
                result.Sheets++;
                int maxTableCells = budget.Limits.MaxTableCells;
                foreach (HtmlSemanticSection section in document.Sections) {
                    if (row > maxTableCells || row > A1.MaxRows) break;
                    bool sectionHasNarrative = section.Blocks.Any(block => IsSectionNarrativeBlock(section, block));
                    if (!sectionHasNarrative && tables.Count > 0) continue;
                    if (TrySetCellTextValue(narrativeSheet, row, 1, section.Title, result, budget)) {
                        narrativeSheet.CellAt(row, 1).SetBold();
                        row++;
                        result.Cells++;
                    }
                    foreach (HtmlSemanticBlock block in section.Blocks) {
                        if (!IsSectionNarrativeBlock(section, block)) continue;
                        if (row > maxTableCells || row > A1.MaxRows) {
                            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                                "Remaining HTML text blocks were omitted because the configured cell limit was reached.",
                                lossKind: HtmlConversionLossKind.Omission, detail: "limit=" + maxTableCells);
                            break;
                        }
                        if (TrySetCellTextValue(narrativeSheet, row, 1, block.Text, result, budget)) {
                            ApplySemanticCellFormatting(narrativeSheet, row, 1, block.Runs,
                                block.Kind == HtmlSemanticBlockKind.Heading, block.Style, result, budget);
                            row++;
                            result.Cells++;
                        }
                    }
                }
            }
        }

        if (options.ImportImages && workbook.Sheets.Count > 0) {
            ExcelSheet imageSheet = narrativeSheet ?? workbook.Sheets[0];
            int imageRow = narrativeSheet != null ? row : 2;
            if (narrativeSheet == null
                && A1.TryParseRange(imageSheet.GetUsedRangeA1(), out _, out _, out int lastRow, out _)) {
                imageRow = Math.Min(A1.MaxRows, lastRow + 2);
            }
            ImportGenericImages(document, imageSheet, result, budget, ref imageRow);
        }
    }

    private static bool IsGenericTextBlock(HtmlSemanticBlockKind kind) =>
        kind == HtmlSemanticBlockKind.Heading || kind == HtmlSemanticBlockKind.Paragraph
        || kind == HtmlSemanticBlockKind.Code || kind == HtmlSemanticBlockKind.Quote
        || kind == HtmlSemanticBlockKind.List || kind == HtmlSemanticBlockKind.Note;

    private static bool IsSectionNarrativeBlock(HtmlSemanticSection section, HtmlSemanticBlock block) =>
        IsGenericTextBlock(block.Kind) && block.Text.Length > 0
        && !(block.Kind == HtmlSemanticBlockKind.Heading
            && string.Equals(block.Text, section.Title, StringComparison.Ordinal));

    private static void ImportGenericImages(
        HtmlSemanticDocument document,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlImportBudget budget,
        ref int row) {
        foreach (HtmlSemanticResource resource in document.Resources.Where(item => item.Kind == HtmlResourceKind.Image)) {
            if (!HtmlImageDataUri.TryParse(resource.Source, out HtmlImageDataUri dataUri)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                    "A generic worksheet image was omitted because synchronous native import currently requires a bounded image data URI.",
                    lossKind: HtmlConversionLossKind.Omission, source: resource.Source);
                continue;
            }
            if (!budget.IsImageWithinLimit(dataUri, out string imageLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "An embedded generic worksheet image was omitted because the shared image limit was reached.",
                    lossKind: HtmlConversionLossKind.Omission, source: resource.Source, detail: imageLimit);
                continue;
            }
            if (!dataUri.TryDecodeBytes(out byte[] bytes)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                    "An embedded generic worksheet image could not be decoded.",
                    lossKind: HtmlConversionLossKind.Omission, source: resource.Source);
                continue;
            }
            if (!budget.TryReserveImageWithShape(dataUri, out imageLimit)) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "An embedded generic worksheet image was omitted because the shared image or drawing limit was reached.",
                    lossKind: HtmlConversionLossKind.Omission, source: resource.Source, detail: imageLimit);
                continue;
            }
            if (row > A1.MaxRows) break;
            int width = ReadGenericImageDimension(resource.WidthPixels, "width", 160, budget, result);
            int height = ReadGenericImageDimension(resource.HeightPixels, "height", 90, budget, result);
            sheet.AddImage(row, 1, bytes, dataUri.MediaType, width, height,
                name: null,
                altText: string.IsNullOrWhiteSpace(resource.AlternateText) ? null : resource.AlternateText);
            result.Images++;
            row = Math.Min(A1.MaxRows + 1, row + Math.Max(2, (height + 19) / 20 + 1));
        }
    }

    private static int ReadGenericImageDimension(
        double? pixels,
        string property,
        int fallback,
        HtmlImportBudget budget,
        HtmlToExcelResult result) {
        int value = fallback;
        if (pixels.HasValue && pixels.Value <= int.MaxValue) {
            value = (int)Math.Round(pixels.Value);
        }
        int maximum = (int)Math.Min(int.MaxValue, budget.Limits.MaxAbsoluteGeometry);
        return NormalizeImportInt(value, fallback, 1, maximum, budget, result, "generic image " + property);
    }

    private static void ApplySemanticTableFormatting(HtmlSemanticTable? table, ExcelSheet sheet,
        HtmlToExcelResult result, HtmlImportBudget budget, int firstRow, int firstColumn) {
        if (table == null) return;
        int row = firstRow;
        foreach (HtmlSemanticTableRow sourceRow in table.Rows) {
            int column = firstColumn;
            foreach (HtmlSemanticTableCell sourceCell in sourceRow.Cells) {
                ApplySemanticCellFormatting(sheet, row, column, sourceCell.Runs, sourceCell.IsHeader,
                    sourceCell.Style, result, budget);
                column += Math.Max(1, sourceCell.ColumnSpan);
            }
            row++;
        }
    }

    private static void ApplySemanticCellFormatting(
        ExcelSheet sheet,
        int row,
        int column,
        IReadOnlyList<HtmlSemanticRun> runs,
        bool isHeader,
        HtmlComputedStyle? style,
        HtmlToExcelResult result,
        HtmlImportBudget budget) {
        ExcelCell cell = sheet.CellAt(row, column);
        if (runs.Count > 0 && runs.Any(IsFormattedRun)) {
            string richText = string.Concat(runs.Select(run => run.Text));
            if (IsWithinExcelFieldLimit(richText, budget, ExcelCellTextCharacterLimit,
                    "ExcelCellTextCharacterLimit", out string detail)) {
                cell.SetRichText(runs.Select(ToExcelRun).ToArray());
            } else {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.SemanticMetadataLimitExceeded,
                    "Cell " + BuildCellReference(row, column) + " rich text formatting was omitted because the normalized runs exceeded a semantic or native Excel field limit.",
                    lossKind: HtmlConversionLossKind.Approximation, detail: detail);
            }
        }
        if (isHeader) cell.SetBold();
        string fontColor = NormalizeHexColor(style?.GetValue("color"));
        if (fontColor.Length > 0) cell.SetFontColor(fontColor);
        string fillColor = NormalizeHexColor(style?.GetValue("background-color"));
        if (fillColor.Length > 0) cell.SetFillColor(fillColor);

        string? hyperlink = runs.Select(run => run.Hyperlink).FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!string.IsNullOrWhiteSpace(hyperlink)) {
            sheet.SetHyperlinkReference(row, column, hyperlink!, style: false);
        }
    }

    private static bool IsFormattedRun(HtmlSemanticRun run) =>
        run.Bold || run.Italic || run.Underline || run.Strikethrough
        || run.Superscript || run.Subscript || !string.IsNullOrWhiteSpace(run.Hyperlink)
        || (run.Style?.Properties.Count ?? 0) > 0;

    private static ExcelRichTextRun ToExcelRun(HtmlSemanticRun source) {
        var run = new ExcelRichTextRun(source.Text) {
            Bold = source.Bold,
            Italic = source.Italic,
            Underline = source.Underline,
            Strikethrough = source.Strikethrough
        };
        string color = NormalizeHexColor(source.Style?.GetValue("color"));
        if (color.Length > 0) run.FontColor = color;
        string fontName = NormalizeFontName(source.Style?.GetValue("font-family"));
        if (fontName.Length > 0) run.FontName = fontName;
        if (TryParseCssPixels(source.Style?.GetValue("font-size"), out double pixels)) run.FontSize = pixels * 0.75D;
        return run;
    }

    private static string NormalizeHexColor(string? value) {
        string color = (value ?? string.Empty).Trim();
        if (color.Length == 7 && color[0] == '#') return color.Substring(1).ToUpperInvariant();
        if (color.Length == 4 && color[0] == '#') {
            return string.Concat(char.ToUpperInvariant(color[1]), char.ToUpperInvariant(color[1]),
                char.ToUpperInvariant(color[2]), char.ToUpperInvariant(color[2]),
                char.ToUpperInvariant(color[3]), char.ToUpperInvariant(color[3]));
        }
        return string.Empty;
    }

    private static string NormalizeFontName(string? value) =>
        (value ?? string.Empty).Split(',').FirstOrDefault()?.Trim().Trim('\'', '"') ?? string.Empty;

    private static bool TryParseCssPixels(string? value, out double pixels) {
        pixels = 0D;
        string text = (value ?? string.Empty).Trim();
        if (!text.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return false;
        return double.TryParse(text.Substring(0, text.Length - 2), System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out pixels) && pixels > 0D;
    }
}
