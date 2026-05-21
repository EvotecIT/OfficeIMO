using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Row-oriented readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Lazily reads each row within the A1 range as a typed object array.
        /// Values are converted using shared strings and styles (date detection).
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (e.g., "A1:C100").</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>Sequence of rows as object?[] with fixed width equal to the range width. Rows without any cells yield null.</returns>
        /// <remarks>
        /// A <c>null</c> row is emitted when the worksheet row is missing from the requested range or when it
        /// contains no cells within the specified bounds. Consumers that require dense data can call
        /// <see cref="ReadRowsAs{T}(string, Func{object, T}?, CancellationToken)"/> which throws an
        /// <see cref="InvalidOperationException"/> when an empty worksheet row is encountered.
        /// </remarks>
        public IEnumerable<object?[]?> ReadRows(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) yield break;

            bool canCancel = ct.CanBeCanceled;
            int height = r2 - r1 + 1;
            int width = c2 - c1 + 1;
            if (CanUseRowsXmlReader()) {
                if (ShouldUseOrderedBufferedXmlStream(height, c1, c2)
                    && TryReadRowsOrderedBufferedXmlFast(r1, c1, r2, c2, width, height, ct, out var orderedRows)) {
                    for (int r = 0; r < orderedRows.Length; r++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return orderedRows[r];
                    }

                    yield break;
                }

                if (height > DenseSnapshotCapacityLimit && RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                    foreach (var row in ReadRowsXmlFast(r1, c1, r2, c2, width, ct)) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return row;
                    }

                    yield break;
                }
            }

            if (height > DenseSnapshotCapacityLimit && RowsAreSortedWithinRange(r1, r2, ct)) {
                int nextRow = r1;
                foreach (var row in EnumerateWorksheetRows(ct)) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int ri = checked((int)row.RowIndex!.Value);
                    if (ri < r1) continue;
                    if (ri > r2) break;

                    while (nextRow < ri) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        yield return null;
                        nextRow++;
                    }

                    yield return ReadRowValue(row, c1, c2, width, ct);
                    nextRow = ri + 1;
                }

                while (nextRow <= r2) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    yield return null;
                    nextRow++;
                }

                yield break;
            }

            var map = new Dictionary<int, Row>(GetSnapshotCapacity(height));
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) continue;
                map[ri] = row;
            }

            for (int r = r1; r <= r2; r++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (!map.TryGetValue(r, out var row)) { yield return null; continue; }

                yield return ReadRowValue(row, c1, c2, width, ct);
            }

            object?[]? ReadRowValue(Row row, int firstColumn, int lastColumn, int rowWidth, CancellationToken token) {
                object?[]? arr = null;
                bool canCancelCell = token.CanBeCanceled;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancelCell) {
                        token.ThrowIfCancellationRequested();
                    }

                    int cc = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cc < firstColumn || cc > lastColumn) continue;
                    arr ??= new object?[rowWidth];
                    if (TryConvertCell(cell, out object? value)) {
                        arr[cc - firstColumn] = value ?? arr[cc - firstColumn];
                    }
                }

                return arr;
            }
        }

        private bool TryReadRowsOrderedBufferedXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int width,
            int height,
            CancellationToken ct,
            out object?[]?[] rows) {
            rows = new object?[height][];

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(height);
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

                    int rowOffset = rowIndex - r1;
                    if ((uint)rowOffset >= (uint)rows.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    rows[rowOffset] = ReadXmlRowValue(reader, c1, c2, width, ct);
                    seenRows.MarkSeen(rowOffset);
                }

                return true;
            } catch (XmlException) {
                rows = Array.Empty<object?[]>();
                return false;
            } catch (IOException) {
                rows = Array.Empty<object?[]>();
                return false;
            } catch (UnauthorizedAccessException) {
                rows = Array.Empty<object?[]>();
                return false;
            } catch (ObjectDisposedException) {
                rows = Array.Empty<object?[]>();
                return false;
            }
        }

        private IEnumerable<object?[]?> ReadRowsXmlFast(int r1, int c1, int r2, int c2, int width, CancellationToken ct) {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            int nextRowIndex = 1;
            int nextRequestedRow = r1;
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
                if (rowIndex < r1) {
                    SkipXmlElement(reader, "row");
                    continue;
                }

                if (rowIndex > r2) {
                    SkipXmlElement(reader, "row");
                    break;
                }

                while (nextRequestedRow < rowIndex) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    yield return null;
                    nextRequestedRow++;
                }

                yield return ReadXmlRowValue(reader, c1, c2, width, ct);
                nextRequestedRow = rowIndex + 1;
            }

            while (nextRequestedRow <= r2) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                yield return null;
                nextRequestedRow++;
            }
        }

        private object?[]? ReadXmlRowValue(XmlReader rowReader, int c1, int c2, int width, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return null;
            }

            object?[]? values = null;
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = width <= 64;
            ulong seenColumns = 0;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return values;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int offset = columnIndex - c1;
                if ((uint)offset >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                values ??= new object?[width];
                values[offset] = ReadXmlCellValue(rowReader);
                if (canTrackColumns && MarkRequestedColumnSeen(offset, width, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return values;
                }
            }

            return values;
        }

        private bool CanUseRowsXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }
    }
}

