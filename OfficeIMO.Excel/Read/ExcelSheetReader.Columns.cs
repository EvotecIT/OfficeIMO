using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Column-oriented readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Reads a single-column A1 range (e.g., "B2:B1000") as a typed sequence.
        /// </summary>
        public IEnumerable<object?> ReadColumn(string a1Range, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (c1 != c2) throw new ArgumentException("ReadColumn expects a single-column A1 range (e.g., 'B2:B100').", nameof(a1Range));

            bool canCancel = ct.CanBeCanceled;
            int height = r2 - r1 + 1;
            if (CanUseColumnXmlReader()) {
                if (height > DenseSnapshotCapacityLimit
                    && TryReadColumnAdaptiveBufferedXmlFast(r1, c1, r2, height, ct, out var adaptiveValues)) {
                    if (adaptiveValues.TryGetSparseValues(out var sparseOffsets, out var sparseValues)) {
                        int sparseIndex = 0;
                        if (canCancel) {
                            for (int i = 0; i < adaptiveValues.Count; i++) {
                                ct.ThrowIfCancellationRequested();

                                if (sparseIndex < sparseOffsets.Length && sparseOffsets[sparseIndex] == i) {
                                    yield return sparseValues[sparseIndex];
                                    sparseIndex++;
                                } else {
                                    yield return null;
                                }
                            }
                        } else {
                            for (int i = 0; i < adaptiveValues.Count; i++) {
                                if (sparseIndex < sparseOffsets.Length && sparseOffsets[sparseIndex] == i) {
                                    yield return sparseValues[sparseIndex];
                                    sparseIndex++;
                                } else {
                                    yield return null;
                                }
                            }
                        }
                    } else if (canCancel) {
                        for (int i = 0; i < adaptiveValues.Count; i++) {
                            ct.ThrowIfCancellationRequested();

                            yield return adaptiveValues.GetDenseValue(i);
                        }
                    } else {
                        for (int i = 0; i < adaptiveValues.Count; i++) {
                            yield return adaptiveValues.GetDenseValue(i);
                        }
                    }

                    yield break;
                }

                if (TryReadColumnXmlFast(r1, c1, r2, height, ct, out var xmlValues)) {
                    if (canCancel) {
                        for (int i = 0; i < height; i++) {
                            ct.ThrowIfCancellationRequested();

                            yield return xmlValues[i];
                        }
                    } else {
                        for (int i = 0; i < height; i++) {
                            yield return xmlValues[i];
                        }
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

                    yield return ReadColumnValue(row, c1, ct);
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

            var rowMap = new Dictionary<int, Row>(GetSnapshotCapacity(height));
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int ri = checked((int)row.RowIndex!.Value);
                if (ri < r1) continue;
                if (ri > r2) continue;
                rowMap[ri] = row;
            }

            for (int r = r1; r <= r2; r++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (!rowMap.TryGetValue(r, out var row)) { yield return null; continue; }

                yield return ReadColumnValue(row, c1, ct);
            }

            object? ReadColumnValue(Row row, int columnIndex, CancellationToken token) {
                bool canCancelCell = token.CanBeCanceled;
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancelCell) {
                        token.ThrowIfCancellationRequested();
                    }

                    int cc = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cc != columnIndex) continue;
                    return TryConvertCell(cell, out object? value) ? value : null;
                }

                return null;
            }
        }

        private bool TryReadColumnAdaptiveBufferedXmlFast(int r1, int columnIndex, int r2, int height, CancellationToken ct, out OrderedColumnBuffer values) {
            values = default;
            int sparseCapacity = Math.Min(height, SparseReadInitialBufferCapacity);
            var sparseOffsets = new List<int>(sparseCapacity);
            var sparseValues = new List<object?>(sparseCapacity);
            object?[]? denseValues = null;
            int populatedRows = 0;
            int previousSparseOffset = -1;
            bool sparseOffsetsSorted = true;
            HashSet<int>? sparseOffsetSet = null;

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
                    if ((uint)rowOffset >= (uint)height) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    seenRows.MarkSeen(rowOffset);
                    bool hasColumnCell = TryReadXmlColumnValue(reader, columnIndex, ct, out object? value);

                    if (denseValues != null) {
                        denseValues[rowOffset] = value;
                        continue;
                    }

                    if (sparseOffsetSet != null || rowOffset <= previousSparseOffset) {
                        sparseOffsetSet ??= new HashSet<int>(sparseOffsets);
                        if (rowOffset <= previousSparseOffset) {
                            sparseOffsetsSorted = false;
                        }

                        if (sparseOffsetSet.Contains(rowOffset)) {
                            ReplaceSparseColumnValue(sparseOffsets, sparseValues, rowOffset, value);
                            previousSparseOffset = Math.Max(previousSparseOffset, rowOffset);
                            continue;
                        }

                        if (!hasColumnCell || value == null) {
                            continue;
                        }

                        sparseOffsetSet.Add(rowOffset);
                    } else if (!hasColumnCell || value == null) {
                        continue;
                    }

                    int nextPopulatedRows = populatedRows + 1;
                    if (ShouldPromoteSparseColumnToDense(rowOffset, nextPopulatedRows)) {
                        denseValues = new object?[height];
                        for (int i = 0; i < sparseOffsets.Count; i++) {
                            denseValues[sparseOffsets[i]] = sparseValues[i];
                        }

                        sparseOffsets.Clear();
                        sparseValues.Clear();
                        denseValues[rowOffset] = value;
                        populatedRows = nextPopulatedRows;
                    } else {
                        previousSparseOffset = Math.Max(previousSparseOffset, rowOffset);
                        sparseOffsets.Add(rowOffset);
                        sparseValues.Add(value);
                        populatedRows = nextPopulatedRows;
                    }
                }

                values = denseValues != null
                    ? OrderedColumnBuffer.FromDense(denseValues)
                    : OrderedColumnBuffer.FromSparse(height, sparseOffsets, sparseValues, sparseOffsetsSorted);
                return true;
            } catch (XmlException) {
                values = default;
                return false;
            } catch (IOException) {
                values = default;
                return false;
            } catch (UnauthorizedAccessException) {
                values = default;
                return false;
            } catch (ObjectDisposedException) {
                values = default;
                return false;
            }
        }

        private static bool ShouldPromoteSparseColumnToDense(int rowOffset, int populatedRows) {
            return rowOffset >= 1024 && populatedRows * 4 > rowOffset + 1;
        }

        private static void ReplaceSparseColumnValue(List<int> offsets, List<object?> values, int rowOffset, object? value) {
            for (int i = offsets.Count - 1; i >= 0; i--) {
                if (offsets[i] == rowOffset) {
                    values[i] = value;
                    return;
                }
            }
        }

        private readonly struct OrderedColumnBuffer {
            private readonly object?[]? _denseValues;
            private readonly int[]? _sparseOffsets;
            private readonly object?[]? _sparseValues;

            private OrderedColumnBuffer(int count, object?[]? denseValues, int[]? sparseOffsets, object?[]? sparseValues) {
                Count = count;
                _denseValues = denseValues;
                _sparseOffsets = sparseOffsets;
                _sparseValues = sparseValues;
            }

            internal int Count { get; }

            internal static OrderedColumnBuffer FromDense(object?[] denseValues) {
                return new OrderedColumnBuffer(denseValues.Length, denseValues, null, null);
            }

            internal static OrderedColumnBuffer FromSparse(int count, List<int> sparseOffsets, List<object?> sparseValues, bool offsetsSorted) {
                if (sparseOffsets.Count == 0) {
                    return new OrderedColumnBuffer(count, null, Array.Empty<int>(), Array.Empty<object?>());
                }

                int[] offsets = sparseOffsets.ToArray();
                object?[] values = sparseValues.ToArray();
                if (!offsetsSorted) {
                    Array.Sort(offsets, values);
                }

                return new OrderedColumnBuffer(count, null, offsets, values);
            }

            internal bool TryGetSparseValues(out int[] offsets, out object?[] values) {
                if (_sparseOffsets != null && _sparseValues != null) {
                    offsets = _sparseOffsets;
                    values = _sparseValues;
                    return true;
                }

                offsets = Array.Empty<int>();
                values = Array.Empty<object?>();
                return false;
            }

            internal object? GetDenseValue(int offset) {
                return _denseValues![offset];
            }
        }

        private bool TryReadColumnXmlFast(int r1, int columnIndex, int r2, int height, CancellationToken ct, out object?[] values) {
            values = new object?[height];

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
                    if ((uint)rowOffset >= (uint)values.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (TryReadXmlColumnValue(reader, columnIndex, ct, out object? value)) {
                        values[rowOffset] = value;
                    }

                    seenRows.MarkSeen(rowOffset);
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

        private bool TryReadXmlColumnValue(XmlReader rowReader, int targetColumnIndex, CancellationToken ct, out object? value) {
            value = null;
            if (rowReader.IsEmptyElement) {
                return false;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return false;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex != targetColumnIndex) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                value = ReadXmlCellValue(rowReader);
                SkipXmlElementContent(rowReader, depth);
                return true;
            }

            return false;
        }

        private bool CanUseColumnXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }
    }
}

