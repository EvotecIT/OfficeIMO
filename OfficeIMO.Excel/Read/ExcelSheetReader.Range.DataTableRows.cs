using System.Data;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        // Object columns need no staged type inference. Keep the table's logical
        // row positions while reusing one input buffer, including sparse and repeated rows.
        private bool TryFillObjectDataTableRowsXml(
            DataTable dt, int r1, int c1, int r2, int c2, int dataRowCount,
            int cols, bool headersInFirstRow, CancellationToken ct, out bool requiresBuffering) {
            requiresBuffering = false;
            int startRow = headersInFirstRow ? 1 : 0;
            var values = new object?[cols];
            object[]? blankRow = null;
            bool canCancel = ct.CanBeCanceled;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                var worksheetRows = new WorksheetXmlRowSelector();
                int nextRowIndex = 1;
                dt.MinimumCapacity = Math.Max(dt.MinimumCapacity, dataRowCount);
                dt.BeginLoadData();
                try {
                    bool advanceReader = true;
                    while (!advanceReader || reader.Read()) {
                        advanceReader = true;
                        if (canCancel) ct.ThrowIfCancellationRequested();
                        if (!worksheetRows.IsRowElement(reader)) continue;

                        int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text);
                        if (rowIndex <= 0) rowIndex = ResolveImplicitXmlRowIndex(reader, nextRowIndex, ct);
                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            reader.Skip();
                            advanceReader = false;
                            continue;
                        }
                        if (headersInFirstRow && rowIndex == r1) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        int rowOffset = rowIndex - r1 - startRow;
                        if ((uint)rowOffset >= (uint)dataRowCount) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }
                        if (dt.Rows.Count == 0 && rowOffset > 0) {
                            // Starting later in the range can mean reversed rows. Staging
                            // avoids creating blank records and then replacing them all.
                            requiresBuffering = true;
                            return false;
                        }
                        while (dt.Rows.Count < rowOffset) {
                            if (canCancel && (dt.Rows.Count & 1023) == 0) ct.ThrowIfCancellationRequested();
                            blankRow ??= CreateDbNullRow(cols);
                            dt.Rows.Add(blankRow);
                        }

                        DataRow? existing = rowOffset < dt.Rows.Count ? dt.Rows[rowOffset] : null;
                        if (existing == null) {
                            Array.Clear(values, 0, values.Length);
                        } else {
                            // Later physical rows replace only the cells they contain.
                            for (int c = 0; c < cols; c++) values[c] = existing[c];
                        }
                        ReadXmlRowIntoDataTableBuffer(reader, c1, c2, cols, null, values, null, ct);
                        for (int c = 0; c < cols; c++) values[c] ??= DBNull.Value;
                        if (existing == null) {
                            dt.Rows.Add(values);
                        } else {
                            existing.ItemArray = values;
                        }
                    }
                    while (dt.Rows.Count < dataRowCount) {
                        if (canCancel && (dt.Rows.Count & 1023) == 0) ct.ThrowIfCancellationRequested();
                        blankRow ??= CreateDbNullRow(cols);
                        dt.Rows.Add(blankRow);
                    }
                } finally {
                    dt.EndLoadData();
                }
                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }
    }
}
