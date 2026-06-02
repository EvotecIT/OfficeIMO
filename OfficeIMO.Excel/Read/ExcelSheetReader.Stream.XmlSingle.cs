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
        private static bool ShouldUseOrderedBufferedXmlStream(int estimatedRows, int firstColumn, int lastColumn) {
            int width = lastColumn - firstColumn + 1;
            return width > 0
                && estimatedRows > 0
                && ((long)estimatedRows * width) <= OrderedBufferedRangeStreamCellLimit;
        }

        private bool CanUseRangeStreamXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }

        private bool CanAttemptRangeStreamXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && _canStreamWorksheetPart
                && _hasWorksheetPartStreamContent != false;
        }

        private bool TryReadSingleRangeChunkXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            CancellationToken ct,
            out RangeChunk? chunk) {
            chunk = null;
            int width = c2 - c1 + 1;
            object?[][]? rows = null;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    _hasWorksheetPartStreamContent = false;
                    return false;
                }

                _hasWorksheetPartStreamContent = true;
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int requestedRowCount = r2 - r1 + 1;
                var seenRows = CreateCompletedRowTracker(requestedRowCount);
                bool orderedRows = true;
                int orderedRowsSeen = 0;

                if (canCancel) {
                    while (reader.Read()) {
                        ct.ThrowIfCancellationRequested();

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == requestedRowCount : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        if (rows == null) {
                            rows = new object?[requestedRowCount][];
                            for (int row = 0; row < rows.Length; row++) {
                                rows[row] = new object?[width];
                            }
                        }

                        ReadXmlRowIntoChunk(reader, rows, rowIndex, r1, c1, c2, ct);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == requestedRowCount) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                } else {
                    while (reader.Read()) {
                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == requestedRowCount : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        if (rows == null) {
                            rows = new object?[requestedRowCount][];
                            for (int row = 0; row < rows.Length; row++) {
                                rows[row] = new object?[width];
                            }
                        }

                        ReadXmlRowIntoChunk(reader, rows, rowIndex, r1, c1, c2, CancellationToken.None);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == requestedRowCount) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                }

                if (rows != null) {
                    chunk = new RangeChunk(r1, rows.Length, c1, width, rows);
                }

                return true;
            } catch (XmlException) {
                chunk = null;
                return false;
            } catch (IOException) {
                chunk = null;
                return false;
            } catch (UnauthorizedAccessException) {
                chunk = null;
                return false;
            } catch (ObjectDisposedException) {
                chunk = null;
                return false;
            }
        }

    }
}