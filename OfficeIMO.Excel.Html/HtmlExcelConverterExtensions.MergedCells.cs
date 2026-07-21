using AngleSharp.Dom;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static void ImportTableGrid(
        IElement table,
        ExcelSheet sheet,
        HtmlToExcelResult result,
        HtmlToExcelOptions options,
        HtmlImportBudget budget,
        int firstRow,
        int firstColumn,
        HashSet<long>? importedFormulaCells,
        bool useSemanticValues) {
        int maxTableCells = budget.Limits.MaxTableCells;

        var occupiedCells = new HashSet<long>();
        int rowOffset = 0;

        foreach (IElement row in EnumerateDirectTableRows(table)) {
            int rowIndex = firstRow + rowOffset;
            if (rowIndex > A1.MaxRows) {
                AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                    "HTML table rows exceeded the Excel worksheet row limit; remaining rows were skipped.", lossKind: HtmlConversionLossKind.Omission);
                break;
            }

            int columnIndex = firstColumn;
            foreach (IElement cell in row.Children.Where(IsTableCell)) {
                while (occupiedCells.Contains(GetImportCellKey(rowIndex, columnIndex))) {
                    columnIndex++;
                }

                if (columnIndex > A1.MaxColumns) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "HTML table columns exceeded the Excel worksheet column limit in row " + rowIndex.ToString(CultureInfo.InvariantCulture) + "; remaining cells in the row were skipped.", lossKind: HtmlConversionLossKind.Omission);
                    break;
                }

                int cellRow = rowIndex;
                int cellColumn = columnIndex;
                string? semanticReference = cell.GetAttribute("data-officeimo-cell");
                if (TryParseCellReference(semanticReference, out int semanticRow, out int semanticColumn)
                    && semanticRow <= A1.MaxRows
                    && semanticColumn <= A1.MaxColumns) {
                    cellRow = semanticRow;
                    cellColumn = semanticColumn;
                } else if (!string.IsNullOrWhiteSpace(semanticReference)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
                        "Cell coordinate '" + semanticReference + "' was outside the Excel worksheet grid and the table position was used instead.", lossKind: HtmlConversionLossKind.Approximation);
                }

                if (occupiedCells.Contains(GetImportCellKey(cellRow, cellColumn))) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " overlapped an earlier HTML table span and was moved to the next available column.", lossKind: HtmlConversionLossKind.Approximation);
                    cellRow = rowIndex;
                    cellColumn = columnIndex;
                }

                int rowSpan = ReadSpan(cell, "rowspan", cellRow, A1.MaxRows, cellRow, cellColumn, result);
                int columnSpan = ReadSpan(cell, "colspan", cellColumn, A1.MaxColumns, cellRow, cellColumn, result);
                long spanArea = (long)rowSpan * columnSpan;
                if (spanArea > maxTableCells - occupiedCells.Count) {
                    if (occupiedCells.Count >= maxTableCells) {
                        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                            "HTML table exceeded the configured MaxTableCells limit; remaining cells were skipped.", lossKind: HtmlConversionLossKind.Omission);
                        return;
                    }

                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " contained a span that exceeded the configured MaxTableCells limit; the span was ignored.", lossKind: HtmlConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                if (SpanOverlaps(occupiedCells, cellRow, cellColumn, rowSpan, columnSpan)) {
                    AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                        "Cell " + BuildCellReference(cellRow, cellColumn) + " contained an overlapping HTML table span; the span was ignored.", lossKind: HtmlConversionLossKind.Approximation);
                    rowSpan = 1;
                    columnSpan = 1;
                }

                ReserveSpan(occupiedCells, cellRow, cellColumn, rowSpan, columnSpan);

                string text = NormalizeText(cell.TextContent);
                if (!IsSemanticEmptyCell(cell) && (text.Length > 0 || cell.GetAttribute("data-officeimo-value") != null)) {
                    if (SetCellValue(sheet, cellRow, cellColumn, cell, text, result, options, budget, importedFormulaCells, useSemanticValues)) {
                        result.Cells++;
                    }
                }

                if (rowSpan > 1 || columnSpan > 1) {
                    sheet.MergeRange(BuildRangeReference(cellRow, cellColumn, cellRow + rowSpan - 1, cellColumn + columnSpan - 1));
                    result.MergedRanges++;
                }

                columnIndex = Math.Max(columnIndex, cellColumn + columnSpan);
            }

            rowOffset++;
        }
    }

    private static IEnumerable<IElement> EnumerateDirectTableRows(IElement table) {
        foreach (IElement child in table.Children) {
            if (IsElement(child, "tr")) {
                yield return child;
                continue;
            }

            if (!IsElement(child, "thead") && !IsElement(child, "tbody") && !IsElement(child, "tfoot")) {
                continue;
            }

            foreach (IElement row in child.Children.Where(candidate => IsElement(candidate, "tr"))) {
                yield return row;
            }
        }
    }

    private static bool HasDirectTableCells(IElement table) =>
        EnumerateDirectTableRows(table).Any(row => row.Children.Any(IsTableCell));

    private static bool IsTableCell(IElement element) => IsElement(element, "th") || IsElement(element, "td");

    private static int ReadSpan(
        IElement cell,
        string attributeName,
        int start,
        int maximum,
        int cellRow,
        int cellColumn,
        HtmlToExcelResult result) {
        string? rawValue = cell.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return 1;
        }

        if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int span)
            || span <= 0
            || start > maximum
            || span > maximum - start + 1) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TableSpanInvalid,
                "Cell " + BuildCellReference(cellRow, cellColumn) + " contained an invalid " + attributeName + " value; a span of 1 was used.", lossKind: HtmlConversionLossKind.Approximation);
            return 1;
        }

        return span;
    }

    private static bool SpanOverlaps(HashSet<long> occupiedCells, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                if (occupiedCells.Contains(GetImportCellKey(currentRow, currentColumn))) {
                    return true;
                }
            }
        }

        return false;
    }

    private static void ReserveSpan(HashSet<long> occupiedCells, int row, int column, int rowSpan, int columnSpan) {
        for (int currentRow = row; currentRow < row + rowSpan; currentRow++) {
            for (int currentColumn = column; currentColumn < column + columnSpan; currentColumn++) {
                occupiedCells.Add(GetImportCellKey(currentRow, currentColumn));
            }
        }
    }

    private static long GetImportCellKey(int row, int column) => ((long)row << 32) | (uint)column;
}
