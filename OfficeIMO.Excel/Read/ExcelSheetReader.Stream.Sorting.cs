using OfficeIMO.OpenXml.Internal;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
        /// <summary>Runs on a data reader's read that starts a switch to the general XML reader; tests use it to cancel that read.</summary>
        internal static readonly AsyncLocal<Action?> HandoverStartingForTests = new();

        /// <summary>Runs with the row count and text size of each row batch the background check hands over.</summary>
        internal static readonly AsyncLocal<Action<int, int>?> RowBatchFlushedForTests = new();

        /// <summary>
        /// Whether the worksheet's rows are in increasing order, the actual range of its cells, and the first row with a
        /// cell outside the given columns or below the given last row (0 when none), read from the reader's own worksheet
        /// stream. Rows without an index resolve as the XML replay resolves them, and their coordinates are published
        /// for it. Used when the validation scan collects no bounds, as for worksheets with tables.
        /// </summary>
        private bool ScanRowOrderAndCellRangeXml(CancellationToken ct, int firstColumn, int lastColumn, int lastRow,
            out string? cellRange, out int firstRowOutside) {
            cellRange = null;
            firstRowOutside = 0;
            using Stream stream = OpenReaderWorksheetStream(ct);
            RewindWorksheetStream(stream);
            using XmlReader reader = OpenWorksheetXmlReader(stream);
            var rows = new WorksheetXmlRowSelector();
            var bounds = new WorksheetRangeAccumulator();
            var coordinates = new ImplicitXmlRowIndexBuilder();
            int rowDepth = -1;
            void EndRow(ref int outside) {
                bounds.EndRow();
                coordinates.EndRow();
                if (outside == 0 && bounds.Exceeds(firstColumn, lastColumn, lastRow)) outside = coordinates.LastRowIndex;
            }
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (rowDepth >= 0 && reader.Depth <= rowDepth) {
                    EndRow(ref firstRowOutside);
                    rowDepth = -1;
                }
                if (rows.IsCellElement(reader)) {
                    ReadOnlySpan<char> reference = ReadXmlReferenceAttribute(reader).Text;
                    bounds.AddCell(reference);
                    if (coordinates.NeedsCellReference) coordinates.AddCell(reference);
                    continue;
                }
                if (!rows.IsRowElement(reader)) continue;
                coordinates.BeginRow(reader, bounds.BeginRow(ReadXmlReferenceAttribute(reader).Text));
                if (reader.IsEmptyElement) EndRow(ref firstRowOutside);
                else rowDepth = reader.Depth;
            }
            if (rowDepth >= 0) EndRow(ref firstRowOutside);
            if (bounds.TryGetReference(out string range)) cellRange = range;
            Interlocked.CompareExchange(ref _implicitXmlRowIndexes, coordinates.Indexes, null);
            return coordinates.RowsStrictlyIncreasing;
        }

        private bool RowsAreSortedWithinRangeXmlFast(int firstRow, int lastRow, CancellationToken token) {
            try {
                using var stream = OpenReaderWorksheetStream(default);
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

        private static bool RowsAreSortedWithinRange(OfficeXmlNode data, int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;

            foreach (var row in EnumerateRowsWithCoordinates(data, token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = row.Index;
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

                int rowIndex = row.Index;
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
