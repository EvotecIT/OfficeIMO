using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Streams a rectangular range and maps each data row into an instance of T without materializing the full result set first.
        /// Header cells are matched to public writable properties on T by name (case-insensitive).
        /// Enumerate the returned sequence while the owning reader is still open.
        /// </summary>
        public IEnumerable<T> ReadObjectsStream<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(string a1Range, CancellationToken ct = default) where T : new() {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (rows <= 1 || cols == 0) {
                return Array.Empty<T>();
            }

            if (CanUseTypedObjectXmlReader()) {
                if (_opt.CellValueConverter == null
                    && _opt.TypeConverter == null) {
                    return ReadObjectsStreamUtf8OrXmlAdaptive<T>(a1Range, r1, c1, r2, c2, cols, ct);
                }

                if (RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                    return ReadObjectsStreamXmlFast<T>(a1Range, r1, c1, r2, c2, cols, ct);
                }
            }

            return ReadObjectsStreamIterator<T>(a1Range, r1, c1, r2, c2, cols, ct);
        }

        private bool CanUseTypedObjectXmlReader() {
            return CanStreamWorksheetPart();
        }

        private IEnumerable<T> ReadObjectsStreamXmlAdaptive<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            int nextRowIndex = 1;
            int nextDataRow = r1 + 1;
            Dictionary<int, T>? pendingRows = null;
            bool checkedSortedGap = false;
            bool useSortedGapStreaming = false;

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
                    if (rowIndex > r2 && nextDataRow > r2) {
                        break;
                    }

                    SkipXmlElement(reader, "row");
                    continue;
                }

                if (rowIndex == r1) {
                    object?[] headerValues = ReadXmlRowValues(reader, rowIndex, c1, c2, cols, ct);
                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
                    canTrackMappedColumns = TryGetMappedColumnMask(bindings, out mappedColumns);
                    continue;
                }

                if (bindings == null) {
                    foreach (var item in ReadObjectsStreamIterator<T>(a1Range, r1, c1, r2, c2, cols, ct)) {
                        yield return item;
                    }

                    yield break;
                }

                if (!canTrackMappedColumns) {
                    canTrackMappedColumns = TryGetMappedColumnMask(bindings, out mappedColumns);
                }

                var target = new T();
                ReadXmlRowIntoTypedObject(reader, rowIndex, c1, c2, bindings, canTrackMappedColumns, mappedColumns, target, ct);
                if (rowIndex == nextDataRow) {
                    yield return target;
                    nextDataRow++;

                    while (TryRemovePendingRow(pendingRows, nextDataRow, out var pending)) {
                        yield return pending;
                        nextDataRow++;
                    }
                } else if (rowIndex > nextDataRow) {
                    if (!checkedSortedGap) {
                        checkedSortedGap = true;
                        useSortedGapStreaming = RowsAreSortedWithinRangeXmlFast(r1, r2, ct);
                    }

                    if (useSortedGapStreaming) {
                        while (nextDataRow < rowIndex && nextDataRow <= r2) {
                            if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                                ct.ThrowIfCancellationRequested();
                            }

                            yield return new T();
                            nextDataRow++;
                        }

                        yield return target;
                        nextDataRow = rowIndex + 1;
                        continue;
                    }

                    pendingRows ??= new Dictionary<int, T>();
                    AddPendingTypedRow(pendingRows, rowIndex, target);
                }
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            while (nextDataRow <= r2) {
                if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (TryRemovePendingRow(pendingRows, nextDataRow, out var pending)) {
                    yield return pending;
                } else {
                    yield return new T();
                }

                nextDataRow++;
            }
        }

        private static bool TryRemovePendingRow<T>(Dictionary<int, T>? pendingRows, int rowIndex, out T pending) {
            if (pendingRows != null && pendingRows.TryGetValue(rowIndex, out pending!)) {
                pendingRows.Remove(rowIndex);
                return true;
            }

            pending = default!;
            return false;
        }

        private void AddPendingTypedRow<T>(Dictionary<int, T> pendingRows, int rowIndex, T row) {
            if (_opt.MaxPendingTypedRows <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxPendingTypedRows), "Maximum pending typed rows must be positive.");
            }

            if (!pendingRows.ContainsKey(rowIndex) && pendingRows.Count >= _opt.MaxPendingTypedRows) {
                throw new InvalidDataException(
                    $"Typed row streaming exceeded the configured out-of-order row buffer limit of {_opt.MaxPendingTypedRows.ToString(CultureInfo.InvariantCulture)}.");
            }

            pendingRows[rowIndex] = row;
        }

        private bool TryReadObjectsStreamOrderedXmlFast<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out T[] results) where T : new() {
            int dataRows = rows - 1;
            results = dataRows <= 0 ? Array.Empty<T>() : new T[dataRows];
            if (dataRows <= 0) {
                return true;
            }

            bool[]? assignedRows = null;
            int assignedRowCount = 0;
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            bool sawHeader = false;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
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
                        if (rowIndex > r2 && sawHeader && assignedRowCount == dataRows) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex == r1) {
                        if (bindings != null && !sawHeader) {
                            results = Array.Empty<T>();
                            return false;
                        }

                        object?[] headerValues = ReadXmlRowValues(reader, rowIndex, c1, c2, cols, ct);
                        var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                        bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
                        canTrackMappedColumns = TryGetMappedColumnMask(bindings, out mappedColumns);
                        sawHeader = true;
                        continue;
                    }

                    if (bindings == null) {
                        results = Array.Empty<T>();
                        return false;
                    }

                    int dataRowOffset = rowIndex - (r1 + 1);
                    if ((uint)dataRowOffset >= (uint)results.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    var target = new T();
                    ReadXmlRowIntoTypedObject(reader, rowIndex, c1, c2, bindings, canTrackMappedColumns, mappedColumns, target, ct);
                    results[dataRowOffset] = target;
                    if (assignedRows == null && dataRowOffset == assignedRowCount) {
                        assignedRowCount++;
                    } else {
                        assignedRows ??= CreateAssignedRowTracker(assignedRowCount, results.Length);
                        if (!assignedRows[dataRowOffset]) {
                            assignedRows[dataRowOffset] = true;
                            assignedRowCount++;
                        }
                    }
                }

                if (assignedRowCount != results.Length) {
                    if (assignedRows == null) {
                        for (int i = assignedRowCount; i < results.Length; i++) {
                            results[i] = new T();
                        }
                    } else {
                        for (int i = 0; i < results.Length; i++) {
                            if (!assignedRows[i]) {
                                results[i] = new T();
                            }
                        }
                    }
                }

                return true;
            } catch (XmlException) {
                results = Array.Empty<T>();
                return false;
            } catch (IOException) {
                results = Array.Empty<T>();
                return false;
            } catch (UnauthorizedAccessException) {
                results = Array.Empty<T>();
                return false;
            } catch (ObjectDisposedException) {
                results = Array.Empty<T>();
                return false;
            }
        }

        private static bool[] CreateAssignedRowTracker(int assignedDensePrefixLength, int rowCount) {
            var assignedRows = new bool[rowCount];
            for (int i = 0; i < assignedDensePrefixLength; i++) {
                assignedRows[i] = true;
            }

            return assignedRows;
        }

        private IEnumerable<T> ReadObjectsStreamXmlFast<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            int nextRowIndex = 1;
            int nextDataRow = r1 + 1;

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
                    break;
                }

                if (rowIndex == r1) {
                    object?[] headerValues = ReadXmlRowValues(reader, rowIndex, c1, c2, cols, ct);
                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    bindings = GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
                    canTrackMappedColumns = TryGetMappedColumnMask(bindings, out mappedColumns);
                    continue;
                }

                if (bindings == null) {
                    bindings = CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
                    canTrackMappedColumns = TryGetMappedColumnMask(bindings, out mappedColumns);
                }

                while (nextDataRow < rowIndex && nextDataRow <= r2) {
                    if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    yield return new T();
                    nextDataRow++;
                }

                if (rowIndex < nextDataRow) {
                    SkipXmlElement(reader, "row");
                    continue;
                }

                var target = new T();
                ReadXmlRowIntoTypedObject(reader, rowIndex, c1, c2, bindings, canTrackMappedColumns, mappedColumns, target, ct);
                yield return target;
                nextDataRow = rowIndex + 1;
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            while (nextDataRow <= r2) {
                if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                yield return new T();
                nextDataRow++;
            }
        }

    }
}
