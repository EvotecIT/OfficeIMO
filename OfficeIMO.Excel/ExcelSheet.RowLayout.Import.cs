using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Applies native imported row dimensions in one ordered DOM pass and worksheet save.</summary>
        internal void SetImportedRowLayout(IReadOnlyDictionary<int, double> heights,
            IEnumerable<int> hiddenRows, CancellationToken cancellationToken) {
            var hidden = new HashSet<int>(hiddenRows);
            var indexes = new SortedSet<int>(heights.Keys.Concat(hidden));
            if (indexes.Count == 0) return;
            cancellationToken.ThrowIfCancellationRequested();
            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SheetData sheetData = GetOrCreateSheetData();
                Row? next = sheetData.GetFirstChild<Row>();
                Row? previous = null;
                foreach (int index in indexes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (index <= 0) continue;
                    while (next != null && (next.RowIndex == null || next.RowIndex.Value < (uint)index)) {
                        cancellationToken.ThrowIfCancellationRequested();
                        previous = next;
                        next = next.NextSibling<Row>();
                    }
                    Row row;
                    if (next?.RowIndex?.Value == (uint)index) {
                        row = next;
                        next = next.NextSibling<Row>();
                    } else {
                        row = new Row { RowIndex = (uint)index };
                        if (next == null) sheetData.Append(row);
                        // InsertBefore asks the SDK to find the prior sibling, which
                        // rescans a growing prefix. Our forward cursor already has it.
                        else if (previous == null) sheetData.PrependChild(row);
                        else sheetData.InsertAfter(row, previous);
                    }
                    previous = row;
                    if (heights.TryGetValue(index, out double height)) {
                        SetRowHeightCore(row, height, roundToHundredths: false);
                    }
                    if (hidden.Contains(index)) row.Hidden = true;
                }
                UpdateSheetFormat();
                WorksheetRoot.Save();
            });
        }
    }
}
