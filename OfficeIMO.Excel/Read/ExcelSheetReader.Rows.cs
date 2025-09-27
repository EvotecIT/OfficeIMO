using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Row-oriented readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Lazily reads each row within the A1 range as a typed object array.
        /// Values are converted using shared strings and styles (date detection).
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (e.g., "A1:C100").</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>Sequence of rows as object?[] with fixed width equal to the range width. Rows without any cells yield null.</returns>
        public IEnumerable<object?[]?> ReadRows(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) yield break;

            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null) yield break;

            // Build quick row lookup for the requested span
            var map = new Dictionary<int, Row>();
            foreach (var row in sheetData.Elements<Row>()) {
                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) break;
                map[ri] = row;
            }

            int width = c2 - c1 + 1;
            for (int r = r1; r <= r2; r++) {
                if (ct.IsCancellationRequested) yield break;

                if (!map.TryGetValue(r, out var row)) { yield return null; continue; }

                var arr = new object?[width];
                bool hasCells = false;

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is null) continue;
                    var (rr, cc) = A1.ParseCellRef(cell.CellReference.Value);
                    if (cc < c1 || cc > c2) continue;
                    var val = ConvertCell(cell);
                    arr[cc - c1] = val ?? arr[cc - c1];
                    hasCells = true;
                }

                yield return hasCells ? arr : null;
            }
        }
    }
}

