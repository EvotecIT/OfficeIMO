using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    internal sealed partial class ExcelSheetReader {
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
