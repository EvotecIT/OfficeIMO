using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>Resolves omitted worksheet coordinates without modifying the DOM.</summary>
    internal static class ExcelWorksheetCoordinates {
        internal static int GetRowIndex(Row row, ref int previousRowIndex) {
            if (row.RowIndex?.Value is uint explicitIndex && explicitIndex > 0) {
                return previousRowIndex = checked((int)explicitIndex);
            }
            foreach (Cell cell in row.Elements<Cell>()) {
                if (ExcelA1.TryParseCellReferenceFast((cell.CellReference?.Value).AsSpan(), out int referencedRow, out _)) {
                    return previousRowIndex = referencedRow;
                }
            }
            return ++previousRowIndex;
        }

        internal static int GetColumnIndex(Cell cell, ref int nextColumnIndex) {
            string? reference = cell.CellReference?.Value;
            int columnIndex = string.IsNullOrEmpty(reference)
                ? nextColumnIndex
                : ExcelA1.ParseColumnIndexFromCellReferenceFast(reference);
            // An unparseable reference takes the next column, as a missing one does and as Excel reads it.
            if (columnIndex <= 0) columnIndex = nextColumnIndex;
            nextColumnIndex = columnIndex + 1;
            return columnIndex;
        }
    }
}
