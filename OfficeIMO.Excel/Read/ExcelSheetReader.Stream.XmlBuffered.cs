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
        private IEnumerable<RangeChunk> ReadRangeStreamXmlFast(int r1, int c1, int r2, int c2, int chunkRows, CancellationToken ct) {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            int width = c2 - c1 + 1;
            int currentWindow = -1;
            int currentStartRow = 0;
            object?[][]? currentRows = null;
            int nextRowIndex = 1;
            int requestedRowCount = r2 - r1 + 1;
            var seenRows = CreateCompletedRowTracker(requestedRowCount);

            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                    continue;
                }

                int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                if (rowIndex <= 0) {
                    rowIndex = nextRowIndex;
                }

                nextRowIndex = rowIndex + 1;
                if (rowIndex < r1 || rowIndex > r2) {
                    if (rowIndex > r2 && seenRows.AllRowsSeen) {
                        break;
                    }

                    SkipXmlElement(reader, "row");
                    continue;
                }

                int window = (rowIndex - r1) / chunkRows;
                if (currentRows != null && window != currentWindow) {
                    yield return new RangeChunk(currentStartRow, currentRows.Length, c1, width, currentRows);
                    currentRows = null;
                }

                if (currentRows == null) {
                    currentWindow = window;
                    currentStartRow = r1 + (window * chunkRows);
                    int rowCount = Math.Min(chunkRows, r2 - currentStartRow + 1);
                    currentRows = new object?[rowCount][];
                    for (int i = 0; i < rowCount; i++) {
                        currentRows[i] = new object?[width];
                    }
                }

                ReadXmlRowIntoChunk(reader, currentRows, rowIndex, currentStartRow, c1, c2, ct);
                seenRows.MarkSeen(rowIndex - r1);
                if (seenRows.AllRowsSeen) {
                    break;
                }
            }

            if (currentRows != null) {
                yield return new RangeChunk(currentStartRow, currentRows.Length, c1, width, currentRows);
            }
        }

        private bool TryReadOrderedBufferedRangeStreamXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            int estimatedRows,
            CancellationToken ct,
            out RangeChunk[] chunks) {
            chunks = Array.Empty<RangeChunk>();
            int width = c2 - c1 + 1;
            int windowCount = ((estimatedRows - 1) / chunkRows) + 1;
            var chunksByWindow = new RangeChunk?[windowCount];
            int populatedWindows = 0;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(estimatedRows);
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1 || rowIndex > r2) {
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    int window = (rowIndex - r1) / chunkRows;
                    if ((uint)window >= (uint)chunksByWindow.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    var chunk = chunksByWindow[window];
                    if (chunk == null) {
                        int startRow = r1 + (window * chunkRows);
                        int rowCount = Math.Min(chunkRows, r2 - startRow + 1);
                        var rows = new object?[rowCount][];
                        for (int row = 0; row < rowCount; row++) {
                            rows[row] = new object?[width];
                        }

                        chunk = new RangeChunk(startRow, rowCount, c1, width, rows);
                        chunksByWindow[window] = chunk;
                        populatedWindows++;
                    }

                    ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, ct);
                    seenRows.MarkSeen(rowIndex - r1);
                    if (seenRows.AllRowsSeen) {
                        break;
                    }
                }

                if (populatedWindows == 0) {
                    return true;
                }

                int index = 0;
                chunks = new RangeChunk[populatedWindows];
                for (int window = 0; window < chunksByWindow.Length; window++) {
                    var chunk = chunksByWindow[window];
                    if (chunk != null) {
                        chunks[index++] = chunk;
                    }
                }

                return true;
            } catch (XmlException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (IOException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (UnauthorizedAccessException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (ObjectDisposedException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            }
        }

        private bool TryReadBufferedRangeStreamXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            int estimatedRows,
            CancellationToken ct,
            out RangeChunk[] chunks) {
            int width = c2 - c1 + 1;
            int windowCount = ((estimatedRows - 1) / chunkRows) + 1;
            chunks = new RangeChunk[windowCount];
            for (int window = 0; window < chunks.Length; window++) {
                int startRow = r1 + (window * chunkRows);
                int rowCount = Math.Min(chunkRows, r2 - startRow + 1);
                var rows = new object?[rowCount][];
                for (int row = 0; row < rowCount; row++) {
                    rows[row] = new object?[width];
                }

                chunks[window] = new RangeChunk(startRow, rowCount, c1, width, rows);
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(estimatedRows);
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
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == estimatedRows : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        int rowOffset = rowIndex - r1;
                        int window = orderedRows && rowOffset == orderedRowsSeen
                            ? orderedRowsSeen / chunkRows
                            : rowOffset / chunkRows;
                        if ((uint)window >= (uint)chunks.Length) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        var chunk = chunks[window];
                        ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, ct);
                        if (orderedRows && rowOffset == orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == estimatedRows) {
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

                        seenRows.MarkSeen(rowOffset);
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
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == estimatedRows : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        int rowOffset = rowIndex - r1;
                        int window = orderedRows && rowOffset == orderedRowsSeen
                            ? orderedRowsSeen / chunkRows
                            : rowOffset / chunkRows;
                        if ((uint)window >= (uint)chunks.Length) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        var chunk = chunks[window];
                        ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, CancellationToken.None);
                        if (orderedRows && rowOffset == orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == estimatedRows) {
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

                        seenRows.MarkSeen(rowOffset);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                }

                return true;
            } catch (XmlException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (IOException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (UnauthorizedAccessException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (ObjectDisposedException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            }
        }
    }
}