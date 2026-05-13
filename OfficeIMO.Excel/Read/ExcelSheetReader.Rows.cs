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
        /// <remarks>
        /// A <c>null</c> row is emitted when the worksheet row is missing from the requested range or when it
        /// contains no cells within the specified bounds. Consumers that require dense data can call
        /// <see cref="ReadRowsAs{T}(string, Func{object, T}?, CancellationToken)"/> which throws an
        /// <see cref="InvalidOperationException"/> when an empty worksheet row is encountered.
        /// </remarks>
        public IEnumerable<object?[]?> ReadRows(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) yield break;

            bool canCancel = ct.CanBeCanceled;
            int height = r2 - r1 + 1;
            int width = c2 - c1 + 1;
            if (height > DenseSnapshotCapacityLimit && RowsAreSortedWithinRange(r1, r2, ct)) {
                int nextRow = r1;
                foreach (var row in EnumerateWorksheetRows(ct)) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int ri = checked((int)row.RowIndex!.Value);
                    if (ri < r1) continue;
                    if (ri > r2) break;

                    while (nextRow < ri) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return null;
                        nextRow++;
                    }

                    yield return ReadRowValue(row, c1, c2, width, ct);
                    nextRow = ri + 1;
                }

                while (nextRow <= r2) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    yield return null;
                    nextRow++;
                }

                yield break;
            }

            var map = new Dictionary<int, Row>(GetSnapshotCapacity(height));
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) continue;
                map[ri] = row;
            }

            for (int r = r1; r <= r2; r++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (!map.TryGetValue(r, out var row)) { yield return null; continue; }

                yield return ReadRowValue(row, c1, c2, width, ct);
            }

            object?[]? ReadRowValue(Row row, int firstColumn, int lastColumn, int rowWidth, CancellationToken token) {
                var arr = new object?[rowWidth];
                bool hasCells = false;
                bool canCancelCell = token.CanBeCanceled;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancelCell) {
                        token.ThrowIfCancellationRequested();
                    }

                    int cc = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cc < firstColumn || cc > lastColumn) continue;
                    hasCells = true;
                    if (TryConvertCell(cell, out object? value)) {
                        arr[cc - firstColumn] = value ?? arr[cc - firstColumn];
                    }
                }

                return hasCells ? arr : null;
            }
        }
    }
}

