using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Column-oriented readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a single-column A1 range (e.g., "B2:B1000") as a typed sequence.
        /// </summary>
        public IEnumerable<object?> ReadColumn(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (c1 != c2) throw new ArgumentException("ReadColumn expects a single-column A1 range (e.g., 'B2:B100').", nameof(a1Range));

            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null) yield break;

            var rowMap = new Dictionary<int, Row>();
            foreach (var row in sheetData.Elements<Row>()) {
                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) break;
                rowMap[ri] = row;
            }

            for (int r = r1; r <= r2; r++) {
                if (ct.IsCancellationRequested) yield break;
                if (!rowMap.TryGetValue(r, out var row)) { yield return null; continue; }

                object? value = null;
                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is null) continue;
                    var (rr, cc) = A1.ParseCellRef(cell.CellReference.Value);
                    if (cc != c1) continue;
                    value = ConvertCell(cell);
                    break;
                }
                yield return value;
            }
        }
    }
}

