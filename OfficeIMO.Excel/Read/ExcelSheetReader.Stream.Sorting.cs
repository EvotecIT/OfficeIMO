using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private bool RowsAreSortedWithinRangeXmlFast(int firstRow, int lastRow, CancellationToken token) {
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = token.CanBeCanceled;
                bool hasPrevious = false;
                bool sawRowAfterRange = false;
                int previous = 0;
                int nextRowIndex = 1;
                int rowCount = lastRow - firstRow + 1;
                var seenRows = CreateCompletedRowTracker(rowCount);

                while (reader.Read()) {
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < firstRow) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex > lastRow) {
                        if (seenRows.AllRowsSeen) {
                            return true;
                        }

                        sawRowAfterRange = true;
                        SkipXmlElement(reader, "row");
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
                    seenRows.MarkSeen(rowIndex - firstRow);
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
            int rowCount = lastRow - firstRow + 1;
            var seenRows = CreateCompletedRowTracker(rowCount);

            foreach (var row in data.Elements<Row>()) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    if (seenRows.AllRowsSeen) {
                        return true;
                    }

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
                seenRows.MarkSeen(rowIndex - firstRow);
            }

            return true;
        }

        private bool RowsAreSortedWithinRange(int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;
            int rowCount = lastRow - firstRow + 1;
            var seenRows = CreateCompletedRowTracker(rowCount);

            foreach (var row in EnumerateWorksheetRows(token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    if (seenRows.AllRowsSeen) {
                        return true;
                    }

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
                seenRows.MarkSeen(rowIndex - firstRow);
            }

            return true;
        }
    }
}