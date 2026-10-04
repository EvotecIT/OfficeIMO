#nullable enable

using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static IEnumerable<Row> EnumerateRowsWithCoordinates(IEnumerable<Row> rows, CancellationToken ct) {
            int previousRowIndex = 0;
            foreach (Row row in rows) {
                ct.ThrowIfCancellationRequested();
                int rowIndex = ExcelWorksheetCoordinates.GetRowIndex(row, ref previousRowIndex);
                if (row.RowIndex?.Value is uint explicitIndex && explicitIndex > 0) {
                    yield return row;
                } else {
                    // Loaded row fragments are detached. Clone attached DOM rows
                    // so reading never mutates a caller's live workbook model.
                    Row projectedRow = row.Parent == null ? row : (Row)row.CloneNode(true);
                    projectedRow.RowIndex = (uint)rowIndex;
                    yield return projectedRow;
                }
            }
        }

    }
}
