using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void ImportGenericDocument(
        IHtmlDocument document,
        ExcelDocument workbook,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget) {
        IReadOnlyList<IElement> tables = HtmlGenericDocumentProjector.SelectRootTables(document);
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

                string title = HtmlGenericDocumentProjector.GetTableTitle(document, tables[index], index + 1);
                ExcelSheet sheet = workbook.AddWorksheet(GetUniqueSheetName(title, usedNames));
                result.Sheets++;
                ImportTableGrid(
                    tables[index],
                    sheet,
                    result,
                    options,
                    budget,
                    1,
                    1,
                    importedFormulaCells: null,
                    useSemanticValues: false);
            }
            return;
        }

        if (!budget.TryReserveSemanticContainer(out string textContainerLimit)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                "HTML text could not be imported because the shared worksheet limit was reached.",
                HtmlDiagnosticSeverity.Error, HtmlConversionLossKind.Omission, detail: textContainerLimit);
            return;
        }

        ExcelSheet textSheet = workbook.AddWorksheet(GetUniqueSheetName("Imported", usedNames));
        result.Sheets++;
        int row = 1;
        int maxTableCells = budget.Limits.MaxTableCells;
        foreach (HtmlGenericSectionProjection section in HtmlGenericDocumentProjector.CreateSections(document)) {
            if (row > maxTableCells || row > A1.MaxRows) break;
            if (TrySetCellTextValue(textSheet, row, 1, section.Title, result, budget)) {
                row++;
                result.Cells++;
            }
            foreach (IElement block in HtmlGenericDocumentProjector.EnumerateBlocks(section)) {
                if (!HtmlGenericDocumentProjector.IsTextBlock(block)) continue;
                string text = HtmlGenericDocumentProjector.GetBlockText(block);
                if (text.Length == 0) continue;
                if (row > maxTableCells || row > A1.MaxRows) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Remaining HTML text blocks were omitted because the configured cell limit was reached.",
                        lossKind: HtmlConversionLossKind.Omission, detail: "limit=" + maxTableCells);
                    return;
                }
                if (TrySetCellTextValue(textSheet, row, 1, text, result, budget)) {
                    row++;
                    result.Cells++;
                }
            }
        }
    }
}
