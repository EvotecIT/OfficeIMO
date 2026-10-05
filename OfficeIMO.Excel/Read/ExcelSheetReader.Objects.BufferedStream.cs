using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Bounded typed-row buffering for worksheets whose row order cannot support direct streaming.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        private IEnumerable<T> ReadObjectsStreamBufferedIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            Row? headerRow = null;
            var pendingRows = new Dictionary<int, Row>();
            foreach (Row row in EnumerateWorksheetRows(ct)) {
                ct.ThrowIfCancellationRequested();
                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex == r1) {
                    headerRow = row;
                } else if (rowIndex > r1 && rowIndex <= r2) {
                    AddPendingTypedRow(pendingRows, rowIndex, row);
                }
            }

            var bindings = headerRow == null
                ? CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols)
                : CreateTypedHeaderBindingsFromRow<T>(headerRow, a1Range, c1, c2, cols);
            int convertedCells = 0;
            for (int rowIndex = r1 + 1; rowIndex <= r2; rowIndex++) {
                ct.ThrowIfCancellationRequested();
                var target = new T();
                if (pendingRows.TryGetValue(rowIndex, out var row)) {
                    pendingRows.Remove(rowIndex);
                    FillTypedObjectFromRow(row, c1, c2, bindings, target, ct, ref convertedCells);
                }
                yield return target;
            }
        }
    }
}
