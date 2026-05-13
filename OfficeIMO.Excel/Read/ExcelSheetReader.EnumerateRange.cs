using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range enumeration for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Enumerates non-empty cells within the given A1 range as typed values.
        /// </summary>
        public IEnumerable<CellValueInfo> EnumerateRange(string a1Range) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            foreach (var row in EnumerateWorksheetRows()) {
                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;
                    if (TryConvertCell(cell, out var value))
                        yield return new CellValueInfo(rIndex, cIndex, value);
                }
            }
        }
    }
}

