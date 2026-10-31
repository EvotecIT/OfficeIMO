using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Names the pivot tables whose location holds no cell values. Excel stores a
        /// pivot's rendered output as ordinary cells; a pivot written without them only
        /// appears after a spreadsheet application refreshes it.
        /// </summary>
        internal IReadOnlyList<string> GetPivotTablesWithoutRenderedCells() {
            MaterializePendingDirectCellValues();
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var names = new List<string>();
                foreach (PivotTablePart part in _worksheetPart.PivotTableParts) {
                    PivotTableDefinition? definition = part.PivotTableDefinition;
                    string? name = definition?.Name?.Value;
                    if (string.IsNullOrEmpty(name) ||
                        !ExcelA1.TryParseRange(definition!.Location?.Reference?.Value ?? string.Empty, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                        continue;
                    }

                    bool hasValue = false;
                    for (int row = firstRow; row <= lastRow && !hasValue; row++) {
                        for (int column = firstColumn; column <= lastColumn && !hasValue; column++) {
                            Cell? cell = TryGetExistingCell(row, column);
                            hasValue = cell != null && (cell.CellValue != null || cell.InlineString != null);
                        }
                    }

                    if (!hasValue) names.Add(name!);
                }

                return names;
            });
        }
    }
}
