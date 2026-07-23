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
        /// Reads a rectangular range and maps rows (excluding the header row) into instances of T.
        /// Header cells are matched to public writable properties on T by name (case-insensitive).
        /// </summary>
        public IEnumerable<T> ReadObjects<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) where T : new() {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            long cellCount = (long)rows * cols;
            if (_opt.MaxRangeCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxRangeCells), "Maximum dense range cell count must be positive.");
            }

            if (cellCount > _opt.MaxRangeCells) {
                throw new InvalidDataException(
                    $"Range '{a1Range}' contains {cellCount.ToString(CultureInfo.InvariantCulture)} cells, exceeding the configured limit of {_opt.MaxRangeCells.ToString(CultureInfo.InvariantCulture)}.");
            }

            if (rows <= 1 || cols == 0) return Array.Empty<T>();

            var policy = _opt.Execution;
            var requested = mode ?? policy.Mode;
            var decided = requested;
            int workload = rows * cols;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                if (CanUseAutomaticXmlReadFastPath(policy)) {
                    if (_opt.CellValueConverter == null
                        && _opt.TypeConverter == null
                        && ShouldAttemptUtf8Range(r1, r2)
                        && RangeReachesDeclaredWorksheetEnd(r2)) {
                        return ReadObjectsStreamUtf8OrXmlAdaptive<T>(a1Range, r1, c1, r2, c2, cols, ct).ToList();
                    }

                    if (ShouldUseOrderedBufferedXmlStream(rows, c1, c2)
                        && TryReadObjectsStreamOrderedXmlFast<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var orderedRows)) {
                        return orderedRows;
                    }

                    if (TryReadObjectsFromXmlMaterialized<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var automaticStreamResult)) {
                        return automaticStreamResult;
                    }

                    if (TryReadObjectsSequentialSinglePass<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var automaticSinglePassResult)) {
                        return automaticSinglePassResult;
                    }
                }

                decided = policy.Decide("ReadObjectsAs", workload);
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel
                && TryReadObjectsFromXmlMaterialized<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var streamResult)) {
                return streamResult;
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel
                && TryReadObjectsSequentialSinglePass<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var singlePassResult)) {
                return singlePassResult;
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel) {
                if (TryReadObjectsFromFastRange<T>(a1Range, r1, c1, rows, cols, ct, out var rangeResult)) {
                    return rangeResult;
                }

                if (TryReadObjectsSequentialSinglePass<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var fastResult)) {
                    return fastResult;
                }

                return ReadObjectsSequential<T>(a1Range, r1, c1, r2, c2, rows, cols, ct);
            }

            var rawCells = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadObjectsAs", decided, ct, workload);

            // Build property map from normalized, disambiguated headers so repeated
            // source headers remain addressable instead of colliding.
            var headerValues = new object?[cols];
            foreach (var cell in rawCells) {
                if (cell.Row != r1) {
                    continue;
                }

                int cc = cell.Col - c1;
                if ((uint)cc < (uint)cols) {
                    headerValues[cc] = cell.TypedValue;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);

            var headerBindings = GetTypedHeaderBindings<T>(headers, a1Range);
            var map = headerBindings.Bindings;
            bool canCancel = ct.CanBeCanceled;

            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            int dataRowCount = rows - 1;
            var result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                result.Add(new T());
            }

            bool hasCustomConverters = _opt.CellValueConverter != null || _opt.TypeConverter != null;
            for (int i = 0; i < rawCells.Count; i++) {
                if (canCancel && (i & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                var cell = rawCells[i];
                if (cell.Row <= r1 || cell.TypedValue is null) {
                    continue;
                }

                int rr = cell.Row - r1 - 1;
                int cc = cell.Col - c1;
                if ((uint)rr >= (uint)result.Count || (uint)cc >= (uint)cols) {
                    continue;
                }

                var binding = map[cc];
                if (binding == null) {
                    continue;
                }

                object? converted = TryChangeType(cell.TypedValue, binding, _opt.Culture);
                if (converted is null
                    && !hasCustomConverters
                    && ShouldRetryRawDateStyledNumericBinding(cell, binding)
                    && TryConvertRawForBinding(cell, binding, out object? rawConverted)) {
                    converted = rawConverted;
                }

                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (converted is not null || binding.IsNullable) {
                    binding.SetValue(result[rr], converted);
                }
            }

            return result;
        }

        private bool TryReadObjectsFromXmlMaterialized<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<T> result) where T : new() {
            result = [];

            if (!CanStreamWorksheetPart()) {
                return false;
            }

            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            for (int i = 0; i < dataRowCount; i++) {
                result.Add(new T());
            }

            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            int nextRowIndex = 1;
            bool sawRow = false;
            bool sawHeader = false;
            bool[]? assignedRows = null;
            int assignedRowCount = 0;

            while (reader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                    continue;
                }

                sawRow = true;
                int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                if (rowIndex <= 0) {
                    rowIndex = nextRowIndex;
                }

                nextRowIndex = rowIndex + 1;
                if (rowIndex < r1 || rowIndex > r2) {
                    if (rowIndex > r2 && sawHeader && assignedRowCount == dataRowCount) {
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
                    sawHeader = true;
                    continue;
                }

                if (bindings == null) {
                    result = [];
                    return false;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    SkipXmlElement(reader, "row");
                    continue;
                }

                ReadXmlRowIntoTypedObject(reader, rowIndex, c1, c2, bindings, canTrackMappedColumns, mappedColumns, result[resultIndex], ct);
                if (assignedRows == null && resultIndex == assignedRowCount) {
                    assignedRowCount++;
                } else {
                    assignedRows ??= CreateAssignedRowTracker(assignedRowCount, result.Count);
                    if (!assignedRows[resultIndex]) {
                        assignedRows[resultIndex] = true;
                        assignedRowCount++;
                    }
                }
            }

            if (!sawRow) {
                result = [];
                return false;
            }

            if (bindings == null) {
                bindings = CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            }

            return bindings != null;
        }

        private bool TryReadObjectsFromFastRange<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int rows,
            int cols,
            CancellationToken ct,
            out List<T> result) where T : new() {
            result = [];

            if (_opt.CellValueConverter != null
                || _opt.TypeConverter != null
                || !CanUseXmlFastReader()) {
                return false;
            }

            object?[,] values = ReadRange(a1Range, OfficeIMO.Excel.ExecutionMode.Sequential, ct);
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), _opt.NormalizeHeaders);
            var headerBindings = GetTypedHeaderBindings<T>(headers, a1Range);
            var bindings = headerBindings.Bindings;

            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            bool canCancel = ct.CanBeCanceled;
            for (int r = 1; r < rows; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                var target = new T();
                for (int c = 0; c < cols; c++) {
                    var binding = bindings[c];
                    if (binding == null) {
                        continue;
                    }

                    object? value = values[r, c];
                    if (value == null) {
                        if (binding.IsNullable) {
                            binding.SetValue(target, null);
                        }

                        continue;
                    }

                    if (_opt.TreatDatesUsingNumberFormat
                        && value is DateTime
                        && IsNumericBindingDestination(binding.BindingKind)) {
                        result = [];
                        return false;
                    }

                    object? converted = TryChangeType(value, binding, _opt.Culture);
                    if (converted is not null || binding.IsNullable) {
                        binding.SetValue(target, converted);
                    }
                }

                result.Add(target);
            }

            return true;
        }

    }
}
