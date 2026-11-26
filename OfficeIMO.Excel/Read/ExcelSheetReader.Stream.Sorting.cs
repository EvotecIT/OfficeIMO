using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        /// <summary>Runs on a data reader's read that starts a switch to the general XML reader; tests use it to cancel that read.</summary>
        internal static readonly AsyncLocal<Action?> HandoverStartingForTests = new();

        /// <summary>
        /// Whether the worksheet's rows are in increasing order, and the actual range of its cells, read from the reader's
        /// own worksheet stream. Used when the validation scan collects no bounds, as for worksheets with tables.
        /// </summary>
        private bool ScanRowOrderAndCellRangeXml(CancellationToken ct, out string? cellRange) {
            cellRange = null;
            using Stream stream = OpenReaderWorksheetStream(ct);
            RewindWorksheetStream(stream);
            using XmlReader reader = OpenWorksheetXmlReader(stream);
            var rows = new WorksheetXmlRowSelector();
            var bounds = new WorksheetRangeAccumulator();
            bool sorted = true;
            int previousRow = 0;
            int rowDepth = -1;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (rowDepth >= 0 && reader.Depth <= rowDepth) {
                    bounds.EndRow();
                    rowDepth = -1;
                }
                if (rows.IsCellElement(reader)) {
                    bounds.AddCell(ReadXmlReferenceAttribute(reader).Text);
                    continue;
                }
                if (!rows.IsRowElement(reader)) continue;
                int explicitRow = bounds.BeginRow(ReadXmlReferenceAttribute(reader).Text);
                int row = explicitRow > 0 ? explicitRow : previousRow + 1;
                if (row <= previousRow) sorted = false;
                previousRow = row;
                if (reader.IsEmptyElement) bounds.EndRow();
                else rowDepth = reader.Depth;
            }
            if (rowDepth >= 0) bounds.EndRow();
            if (bounds.TryGetReference(out string range)) cellRange = range;
            return sorted;
        }

        private bool RowsAreSortedWithinRangeXmlFast(int firstRow, int lastRow, CancellationToken token) {
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                var worksheetRows = new WorksheetXmlRowSelector();
                bool canCancel = token.CanBeCanceled;
                bool hasPrevious = false;
                bool sawRowAfterRange = false;
                int previous = 0;
                int nextRowIndex = 1;

                bool advanceReader = true;
                while (!advanceReader || reader.Read()) {
                    advanceReader = true;
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    if (!worksheetRows.IsRowElement(reader)) {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(ReadXmlReferenceAttribute(reader).Text);
                    if (rowIndex <= 0) {
                        rowIndex = ResolveImplicitXmlRowIndex(reader, nextRowIndex, token);
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < firstRow || rowIndex > lastRow) {
                        if (rowIndex > lastRow) sawRowAfterRange = true;
                        reader.Skip();
                        advanceReader = false;
                        continue;
                    }

                    if (sawRowAfterRange) {
                        return false;
                    }

                    if (hasPrevious && rowIndex <= previous) {
                        return false;
                    }

                    previous = rowIndex;
                    hasPrevious = true;
                    reader.Skip();
                    advanceReader = false;
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

        private static bool RowsAreSortedWithinRange(SheetData data, int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;

            foreach (var row in EnumerateRowsWithCoordinates(data.Elements<Row>(), token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    sawRowAfterRange = true;
                    continue;
                }
                if (sawRowAfterRange) {
                    return false;
                }

                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
            }

            return true;
        }

        private bool RowsAreSortedWithinRange(int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;

            foreach (var row in EnumerateWorksheetRows(token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    sawRowAfterRange = true;
                    continue;
                }
                if (sawRowAfterRange) {
                    return false;
                }

                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
            }

            return true;
        }
    }
}
