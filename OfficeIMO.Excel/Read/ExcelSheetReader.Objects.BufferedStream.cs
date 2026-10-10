using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading;
using System.Xml;

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
            CancellationToken ct,
            bool enforcePendingLimit = true) where T : new() {
            var pendingRows = new Dictionary<int, CellRaw?[]>();
            CellRaw?[]? headerCells = CanUseTypedObjectXmlReader()
                ? BufferTypedRowsXml(pendingRows, r1, c1, r2, c2, cols, ct, enforcePendingLimit)
                : BufferTypedRowsDom(pendingRows, r1, c1, r2, c2, cols, ct, enforcePendingLimit);
            var headerValues = new object?[cols];
            if (headerCells != null) {
                for (int column = 0; column < cols; column++) {
                    if (headerCells[column] is CellRaw cell) headerValues[column] = ConvertRaw(cell).TypedValue;
                }
            }
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, column => headerValues[column]?.ToString(), _opt.NormalizeHeaders);
            var bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
            headerCells = null;
            // Map each merged logical row once, after qualifying its final header.
            // An omitted cell retains its earlier value; a present blank replaces it
            // and follows the fresh object's normal default-value mapping.
            for (int rowIndex = r1 + 1; rowIndex <= r2; rowIndex++) {
                ct.ThrowIfCancellationRequested();
                var target = new T();
                if (pendingRows.TryGetValue(rowIndex, out CellRaw?[]? cells)) {
                    pendingRows.Remove(rowIndex);
                    for (int column = 0; column < cols; column++) {
                        if ((column & 1023) == 0) ct.ThrowIfCancellationRequested();
                        if (cells[column] is CellRaw cell && bindings[column] is { } binding)
                            TrySetRawCellForBinding(cell, binding, target);
                    }
                }
                yield return target;
            }
        }

        private CellRaw?[]? BufferTypedRowsXml(
            Dictionary<int, CellRaw?[]> pendingRows, int r1, int c1, int r2, int c2, int cols,
            CancellationToken ct, bool enforcePendingLimit) {
            CellRaw?[]? headerCells = null;
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            var worksheetRows = new WorksheetXmlRowSelector();
            int nextRowIndex = 1;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (!worksheetRows.IsRowElement(reader)) continue;
                int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text);
                if (rowIndex <= 0) rowIndex = ResolveImplicitXmlRowIndex(reader, nextRowIndex, ct);
                nextRowIndex = rowIndex + 1;
                if (rowIndex < r1 || rowIndex > r2) {
                    SkipXmlElement(reader, "row");
                    continue;
                }
                CellRaw?[] cells = rowIndex == r1
                    ? headerCells = new CellRaw?[cols]
                    : GetPendingTypedCells(pendingRows, rowIndex, cols, enforcePendingLimit);
                if (reader.IsEmptyElement) continue;
                int depth = reader.Depth;
                int nextColumnIndex = 1;
                int visitedNodes = 0;
                while (reader.Read()) {
                    if ((++visitedNodes & 1023) == 0) ct.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "row") break;
                    if (!SpreadsheetXmlContent.IsDirectChildElement(reader, depth, "c")) continue;
                    int columnIndex = GetXmlCellColumnIndex(reader, ref nextColumnIndex);
                    if (columnIndex < c1 || columnIndex > c2) {
                        SkipXmlElement(reader, "c");
                        continue;
                    }
                    cells[columnIndex - c1] = ReadXmlCellRaw(reader, rowIndex, columnIndex,
                        ParseXmlCellKind(ReadXmlCellTypeAttribute(reader)), readStyleIndex: true);
                }
            }
            return headerCells;
        }

        private CellRaw?[]? BufferTypedRowsDom(
            Dictionary<int, CellRaw?[]> pendingRows, int r1, int c1, int r2, int c2, int cols,
            CancellationToken ct, bool enforcePendingLimit) {
            CellRaw?[]? headerCells = null;
            int visitedCells = 0;
            foreach (Row row in EnumerateWorksheetRows(ct)) {
                ct.ThrowIfCancellationRequested();
                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1 || rowIndex > r2) continue;
                CellRaw?[] cells = rowIndex == r1
                    ? headerCells = new CellRaw?[cols]
                    : GetPendingTypedCells(pendingRows, rowIndex, cols, enforcePendingLimit);
                int nextColumnIndex = 1;
                foreach (Cell cell in row.Elements<Cell>()) {
                    if ((++visitedCells & 1023) == 0) ct.ThrowIfCancellationRequested();
                    int columnIndex = ExcelWorksheetCoordinates.GetColumnIndex(cell, ref nextColumnIndex);
                    if (columnIndex >= c1 && columnIndex <= c2)
                        cells[columnIndex - c1] = SnapshotCell(cell, rowIndex, columnIndex);
                }
            }
            return headerCells;
        }

        private CellRaw?[] GetPendingTypedCells(Dictionary<int, CellRaw?[]> pendingRows, int rowIndex, int cols, bool enforcePendingLimit) {
            if (pendingRows.TryGetValue(rowIndex, out CellRaw?[]? cells)) return cells;
            cells = new CellRaw?[cols];
            if (enforcePendingLimit) AddPendingTypedRow(pendingRows, rowIndex, cells);
            else pendingRows.Add(rowIndex, cells);
            return cells;
        }

    }
}
