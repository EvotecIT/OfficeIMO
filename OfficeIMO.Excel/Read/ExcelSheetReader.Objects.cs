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
            if (rows <= 1 || cols == 0) return Array.Empty<T>();

            var policy = _opt.Execution;
            var requested = mode ?? policy.Mode;
            var decided = requested;
            int workload = rows * cols;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                decided = policy.Decide("ReadObjectsAs", workload);
            }

            if (requested != OfficeIMO.Excel.ExecutionMode.Parallel
                && TryReadObjectsFromXmlMaterialized<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var streamResult)) {
                return streamResult;
            }

            if (requested != OfficeIMO.Excel.ExecutionMode.Parallel
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

            if (_opt.TypeConverter != null
                || _opt.CellValueConverter != null
                || _opt.Culture != CultureInfo.InvariantCulture
                || !_canStreamWorksheetPart) {
                return false;
            }

            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            for (int i = 0; i < dataRowCount; i++) {
                result.Add(new T());
            }

            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
            bool canCancel = ct.CanBeCanceled;
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            int nextRowIndex = 1;
            bool sawRow = false;

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
                    result = [];
                    return false;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    SkipXmlElement(reader, "row");
                    continue;
                }

                ReadXmlRowIntoTypedObject(reader, rowIndex, c1, c2, bindings, canTrackMappedColumns, mappedColumns, result[resultIndex], ct);
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

            if (_opt.TypeConverter == null
                && CanUseXmlFastReader()) {
                if (rows > BufferedRangeStreamRowLimit
                    && ShouldUseOrderedBufferedXmlStream(rows, c1, c2)
                    && TryReadObjectsStreamOrderedXmlFast<T>(a1Range, r1, c1, r2, c2, rows, cols, ct, out var orderedRows)) {
                    return orderedRows;
                }

                if (RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                    return ReadObjectsStreamXmlFast<T>(a1Range, r1, c1, r2, c2, cols, ct);
                }
            }

            return ReadObjectsStreamIterator<T>(a1Range, r1, c1, r2, c2, cols, ct);
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

            var assignedRows = new bool[dataRows];
            TypedPropertyBinding<T>?[]? bindings = null;
            bool canTrackMappedColumns = false;
            ulong mappedColumns = 0;
            bool sawHeader = false;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
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
                    assignedRows[dataRowOffset] = true;
                }

                for (int i = 0; i < results.Length; i++) {
                    if (!assignedRows[i]) {
                        results[i] = new T();
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
            var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
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

        private object?[] ReadXmlRowValues(XmlReader rowReader, int rowIndex, int c1, int c2, int cols, CancellationToken ct) {
            var values = new object?[cols];
            if (rowReader.IsEmptyElement) {
                return values;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
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
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                CellRaw raw = ReadXmlCellRaw(rowReader, rowIndex, columnIndex);
                values[columnIndex - c1] = ConvertRaw(raw).TypedValue;
                if (MarkRequestedColumnSeen(columnIndex - c1, cols, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return values;
                }
            }

            return values;
        }

        private void ReadXmlRowIntoTypedObject<T>(
            XmlReader rowReader,
            int rowIndex,
            int c1,
            int c2,
            TypedPropertyBinding<T>?[] bindings,
            bool canTrackMappedColumns,
            ulong mappedColumns,
            T target,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int convertedCells = 0;
            ulong seenMappedColumns = 0;
            while (rowReader.Read()) {
                if (canCancel && (++convertedCells & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                var binding = bindings[columnIndex - c1];
                if (binding == null) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                CellRaw raw = ReadXmlCellRaw(rowReader, rowIndex, columnIndex);
                TrySetRawCellForBinding(raw, binding, target);
                if (canTrackMappedColumns) {
                    seenMappedColumns |= 1UL << (columnIndex - c1);
                    if (seenMappedColumns == mappedColumns) {
                        SkipXmlElementContent(rowReader, depth, "row");
                        return;
                    }
                }
            }
        }

        private static bool TryGetMappedColumnMask<T>(TypedPropertyBinding<T>?[] bindings, out ulong mask) {
            mask = 0;
            for (int i = 0; i < bindings.Length; i++) {
                if (bindings[i] == null) {
                    continue;
                }

                if ((uint)i >= 64u) {
                    mask = 0;
                    return false;
                }

                mask |= 1UL << i;
            }

            return mask != 0;
        }

        private static int GetXmlCellColumnIndex(XmlReader cellReader, ref int nextColumnIndex) {
            string? reference = cellReader.GetAttribute("r");
            int columnIndex = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(reference);
            if (columnIndex <= 0) {
                columnIndex = string.IsNullOrEmpty(reference) ? nextColumnIndex : 0;
            }

            if (columnIndex > 0) {
                nextColumnIndex = columnIndex + 1;
            }

            return columnIndex;
        }

        private CellRaw ReadXmlCellRaw(XmlReader cellReader, int rowIndex, int columnIndex) {
            CellValues? typeHint = ParseXmlCellType(cellReader.GetAttribute("t"));
            uint? styleIndex = TryParseUInt(cellReader.GetAttribute("s"), out uint parsedStyle) ? parsedStyle : null;
            var raw = new CellRaw {
                Row = rowIndex,
                Col = columnIndex,
                TypeHint = typeHint,
                StyleIndex = styleIndex
            };

            if (cellReader.IsEmptyElement) {
                return raw;
            }

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasFormula = false;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        rawText = cellReader.ReadElementContentAsString();
                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        hasFormula = true;
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!_opt.UseCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth, "c");
                            raw.HasFormula = true;
                            raw.FormulaText = formulaText;
                            return raw;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        inlineText = ReadXmlInlineString(cellReader);
                        hasNode = true;
                        continue;
                    }
                }

                hasNode = cellReader.Read();
            }

            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            raw.HasFormula = hasFormula;
            raw.FormulaText = formulaText;
            raw.RawText = preferFormulaText ? null : rawText;
            raw.InlineText = preferFormulaText ? null : inlineText;
            return raw;
        }

        private static CellValues? ParseXmlCellType(string? type)
            => type switch {
                "b" => CellValues.Boolean,
                "d" => CellValues.Date,
                "inlineStr" => CellValues.InlineString,
                "n" => CellValues.Number,
                "s" => CellValues.SharedString,
                "str" => CellValues.String,
                _ => null
            };

        private static void SkipXmlElement(XmlReader reader, string localName) {
            if (reader.IsEmptyElement) {
                return;
            }

            int depth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                    return;
                }
            }
        }

        private static void SkipXmlElementContent(XmlReader reader, int depth, string localName) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                return;
            }

            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                    return;
                }
            }
        }

        private IEnumerable<T> ReadObjectsStreamIterator<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            Dictionary<int, Row>? pendingRows = null;
            int nextDataRow = r1 + 1;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (bindings == null) {
                    if (rowIndex == r1) {
                        bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);

                        while (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                            pendingRows.Remove(nextDataRow);
                            var pendingTarget = new T();
                            FillTypedObjectFromRow(pendingRow, c1, c2, bindings, pendingTarget, ct, ref convertedCells);
                            yield return pendingTarget;
                            nextDataRow++;
                        }

                        continue;
                    }

                    pendingRows ??= new Dictionary<int, Row>();
                    pendingRows[rowIndex] = row;
                    continue;
                }

                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex < nextDataRow) {
                    continue;
                }

                if (rowIndex > nextDataRow) {
                    pendingRows ??= new Dictionary<int, Row>();
                    pendingRows[rowIndex] = row;
                    continue;
                }

                var currentRow = row;
                while (true) {
                    var target = new T();
                    FillTypedObjectFromRow(currentRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                    nextDataRow++;

                    if (pendingRows == null || !pendingRows.TryGetValue(nextDataRow, out currentRow)) {
                        break;
                    }

                    pendingRows.Remove(nextDataRow);
                }
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);
            while (nextDataRow <= r2) {
                if (canCancel && ((nextDataRow - r1) & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (pendingRows != null && pendingRows.TryGetValue(nextDataRow, out var pendingRow)) {
                    pendingRows.Remove(nextDataRow);
                    var target = new T();
                    FillTypedObjectFromRow(pendingRow, c1, c2, bindings, target, ct, ref convertedCells);
                    yield return target;
                } else {
                    yield return new T();
                }

                nextDataRow++;
            }
        }

        private bool TryReadObjectsSequentialSinglePass<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<T> result) where T : new() {
            int dataRowCount = rows - 1;
            result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                result.Add(new T());
            }

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (rowIndex == r1) {
                    bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                    continue;
                }

                if (bindings == null) {
                    return false;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return bindings != null;
        }

        private IEnumerable<T> ReadObjectsSequential<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct) where T : new() {
            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            TypedPropertyBinding<T>?[]? bindings = null;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex != r1) {
                    continue;
                }

                bindings = CreateTypedHeaderBindingsFromRow<T>(row, a1Range, c1, c2, cols);
                break;
            }

            bindings ??= CreateTypedHeaderBindingsFromMissingRow<T>(a1Range, cols);

            int dataRowCount = rows - 1;
            var result = new List<T>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                result.Add(new T());
            }

            int convertedCells = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex <= r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                int resultIndex = rowIndex - r1 - 1;
                if ((uint)resultIndex >= (uint)result.Count) {
                    continue;
                }

                FillTypedObjectFromRow(row, c1, c2, bindings, result[resultIndex], ct, ref convertedCells);
            }

            return result;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            string a1Range,
            int c1,
            int c2,
            int cols) where T : new() {
            var headerValues = new object?[cols];
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                if (TryConvertCell(cell, out object? value)) {
                    headerValues[columnIndex - c1] = value;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private TypedPropertyBinding<T>?[] CreateTypedHeaderBindingsFromMissingRow<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string a1Range,
            int cols) where T : new() {
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, static _ => null, _opt.NormalizeHeaders);
            return GetTypedHeaderBindings<T>(headers, a1Range).Bindings;
        }

        private void FillTypedObjectFromRow<T>(
            DocumentFormat.OpenXml.Spreadsheet.Row row,
            int c1,
            int c2,
            TypedPropertyBinding<T>?[] bindings,
            T target,
            CancellationToken ct,
            ref int convertedCells) {
            bool canCancel = ct.CanBeCanceled;
            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                if (canCancel && (++convertedCells & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                if (columnIndex < c1 || columnIndex > c2) {
                    continue;
                }

                var binding = bindings[columnIndex - c1];
                if (binding == null) {
                    continue;
                }

                bool convertedSuccessfully = TryConvertCellForBinding(cell, binding, out object? converted);
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (convertedSuccessfully) {
                    binding.SetValue(target, converted);
                }
            }
        }

        private TypedHeaderBindingCache<T> GetTypedHeaderBindings<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
            string[] headers,
            string a1Range) where T : new() {
            var propertyMaps = TypedObjectBindingCache<T>.PropertyMaps;
            string typeName = propertyMaps.TypeName;
            foreach (string diagnostic in propertyMaps.Diagnostics) {
                _opt.Execution.ReportInfo(diagnostic);
            }

            var headerBindings = TypedObjectBindingCache<T>.GetHeaderBindings(headers);

            if (_opt.StrictTypedMapping) {
                var strictIssues = new List<string>(propertyMaps.Diagnostics);
                strictIssues.AddRange(headerBindings.UnmappedIssues);
                if (strictIssues.Count > 0) {
                    throw new InvalidOperationException(
                        $"Typed mapping for '{typeName}' is strict and could not resolve all headers in range '{a1Range}'. " +
                        string.Join(" ", strictIssues));
                }
            }

            return headerBindings;
        }

        private sealed class TypedPropertyBinding<TTarget> {
            internal TypedPropertyBinding(
                PropertyInfo property,
                Type propertyType,
                Type destinationType,
                TypedBindingKind bindingKind,
                bool isNullable,
                bool needsDateStyleConversion,
                Action<TTarget, object?> setValue,
                Action<TTarget, string?>? setString,
                Action<TTarget, int>? setInt32,
                Action<TTarget, long>? setInt64,
                Action<TTarget, double>? setDouble,
                Action<TTarget, decimal>? setDecimal,
                Action<TTarget, bool>? setBoolean,
                Action<TTarget, DateTime>? setDateTime,
                Func<object, CultureInfo, object?> convertValue) {
                Property = property;
                PropertyType = propertyType;
                DestinationType = destinationType;
                BindingKind = bindingKind;
                IsNullable = isNullable;
                NeedsDateStyleConversion = needsDateStyleConversion;
                SetValue = setValue;
                SetString = setString;
                SetInt32 = setInt32;
                SetInt64 = setInt64;
                SetDouble = setDouble;
                SetDecimal = setDecimal;
                SetBoolean = setBoolean;
                SetDateTime = setDateTime;
                ConvertValue = convertValue;
            }

            internal PropertyInfo Property { get; }
            internal Type PropertyType { get; }
            internal Type DestinationType { get; }
            internal TypedBindingKind BindingKind { get; }
            internal bool IsNullable { get; }
            internal bool NeedsDateStyleConversion { get; }
            internal Action<TTarget, object?> SetValue { get; }
            internal Action<TTarget, string?>? SetString { get; }
            internal Action<TTarget, int>? SetInt32 { get; }
            internal Action<TTarget, long>? SetInt64 { get; }
            internal Action<TTarget, double>? SetDouble { get; }
            internal Action<TTarget, decimal>? SetDecimal { get; }
            internal Action<TTarget, bool>? SetBoolean { get; }
            internal Action<TTarget, DateTime>? SetDateTime { get; }
            internal Func<object, CultureInfo, object?> ConvertValue { get; }
        }

        private enum TypedBindingKind {
            Other,
            String,
            Int32,
            Int64,
            Double,
            Decimal,
            Boolean,
            DateTime
        }

        private sealed class TypedPropertyMapCache {
            internal TypedPropertyMapCache(
                string typeName,
                Dictionary<string, PropertyInfo> exactProperties,
                Dictionary<string, PropertyInfo> exactAliases,
                Dictionary<string, PropertyInfo> canonicalProperties,
                Dictionary<string, PropertyInfo> canonicalAliases,
                IReadOnlyList<string> diagnostics) {
                TypeName = typeName;
                ExactProperties = exactProperties;
                ExactAliases = exactAliases;
                CanonicalProperties = canonicalProperties;
                CanonicalAliases = canonicalAliases;
                Diagnostics = diagnostics;
            }

            internal string TypeName { get; }
            internal Dictionary<string, PropertyInfo> ExactProperties { get; }
            internal Dictionary<string, PropertyInfo> ExactAliases { get; }
            internal Dictionary<string, PropertyInfo> CanonicalProperties { get; }
            internal Dictionary<string, PropertyInfo> CanonicalAliases { get; }
            internal IReadOnlyList<string> Diagnostics { get; }
        }

        private sealed class TypedHeaderBindingCache<TTarget> where TTarget : new() {
            internal TypedHeaderBindingCache(TypedPropertyBinding<TTarget>?[] bindings, IReadOnlyList<string> unmappedIssues) {
                Bindings = bindings;
                UnmappedIssues = unmappedIssues;
            }

            internal TypedPropertyBinding<TTarget>?[] Bindings { get; }
            internal IReadOnlyList<string> UnmappedIssues { get; }
        }

        private static class TypedObjectBindingCache<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] TTarget> where TTarget : new() {
            private const int HeaderBindingCacheLimit = 64;

            internal static readonly PropertyInfo[] WritableProperties = typeof(TTarget)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => p.CanWrite)
                .ToArray();

            internal static readonly Dictionary<PropertyInfo, TypedPropertyBinding<TTarget>> Bindings = WritableProperties
                .ToDictionary(prop => prop, CreateBinding);

            internal static readonly TypedPropertyMapCache PropertyMaps = CreatePropertyMaps();

            private static readonly ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>> HeaderBindings = new ConcurrentDictionary<string, TypedHeaderBindingCache<TTarget>>(StringComparer.Ordinal);

            internal static TypedHeaderBindingCache<TTarget> GetHeaderBindings(string[] headers) {
                string key = CreateHeaderBindingKey(headers);
                if (HeaderBindings.TryGetValue(key, out var cached)) {
                    return cached;
                }

                var created = CreateHeaderBindings(headers);
                if (HeaderBindings.Count < HeaderBindingCacheLimit) {
                    return HeaderBindings.GetOrAdd(key, created);
                }

                return created;
            }

            private static TypedPropertyBinding<TTarget> CreateBinding(PropertyInfo property) {
                var nullable = Nullable.GetUnderlyingType(property.PropertyType);
                var destinationType = nullable ?? property.PropertyType;
                var bindingKind = GetBindingKind(destinationType);
                return new TypedPropertyBinding<TTarget>(
                    property,
                    property.PropertyType,
                    destinationType,
                    bindingKind,
                    !property.PropertyType.IsValueType || nullable != null,
                    NeedsDateStyleConversion(destinationType),
                    CreateSetter(property),
                    bindingKind == TypedBindingKind.String ? CreateTypedSetter<string?>(property) : null,
                    bindingKind == TypedBindingKind.Int32 ? CreateTypedSetter<int>(property) : null,
                    bindingKind == TypedBindingKind.Int64 ? CreateTypedSetter<long>(property) : null,
                    bindingKind == TypedBindingKind.Double ? CreateTypedSetter<double>(property) : null,
                    bindingKind == TypedBindingKind.Decimal ? CreateTypedSetter<decimal>(property) : null,
                    bindingKind == TypedBindingKind.Boolean ? CreateTypedSetter<bool>(property) : null,
                    bindingKind == TypedBindingKind.DateTime ? CreateTypedSetter<DateTime>(property) : null,
                    CreateConverter(destinationType));
            }

            private static TypedBindingKind GetBindingKind(Type destinationType) {
                if (destinationType == typeof(string)) {
                    return TypedBindingKind.String;
                }

                if (destinationType == typeof(int)) {
                    return TypedBindingKind.Int32;
                }

                if (destinationType == typeof(long)) {
                    return TypedBindingKind.Int64;
                }

                if (destinationType == typeof(double)) {
                    return TypedBindingKind.Double;
                }

                if (destinationType == typeof(decimal)) {
                    return TypedBindingKind.Decimal;
                }

                if (destinationType == typeof(bool)) {
                    return TypedBindingKind.Boolean;
                }

                if (destinationType == typeof(DateTime)) {
                    return TypedBindingKind.DateTime;
                }

                return TypedBindingKind.Other;
            }

            private static bool NeedsDateStyleConversion(Type destinationType) {
                return destinationType == typeof(DateTime) || destinationType == typeof(string);
            }

            private static Func<object, CultureInfo, object?> CreateConverter(Type destinationType) {
                if (destinationType == typeof(string)) {
                    return static (value, culture) => value as string ?? Convert.ToString(value, culture);
                }

                if (destinationType == typeof(bool)) {
                    return static (value, culture) => {
                        try {
                            if (value is bool boolValue) return boolValue;
                            return Convert.ToBoolean(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(int)) {
                    return static (value, culture) => {
                        try {
                            if (value is int intValue) return intValue;
                            if (value is double doubleValue
                                && doubleValue >= int.MinValue
                                && doubleValue <= int.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (int)doubleValue;
                            }

                            return Convert.ToInt32(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(long)) {
                    return static (value, culture) => {
                        try {
                            if (value is long longValue) return longValue;
                            if (value is double doubleValue
                                && doubleValue >= long.MinValue
                                && doubleValue <= long.MaxValue
                                && Math.Truncate(doubleValue) == doubleValue) {
                                return (long)doubleValue;
                            }

                            return Convert.ToInt64(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(double)) {
                    return static (value, culture) => {
                        try {
                            if (value is double doubleValue) return doubleValue;
                            return Convert.ToDouble(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(decimal)) {
                    return static (value, culture) => {
                        try {
                            if (value is decimal decimalValue) return decimalValue;
                            return Convert.ToDecimal(value, culture);
                        } catch {
                            return null;
                        }
                    };
                }

                if (destinationType == typeof(DateTime)) {
                    return static (value, culture) => {
                        try {
                            if (value is DateTime dt) return dt;
                            if (value is double oa) return DateTime.FromOADate(oa);
                            if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                            return null;
                        } catch {
                            return null;
                        }
                    };
                }

                return (value, culture) => ConvertToDestinationType(value, destinationType, culture);
            }

            private static Action<TTarget, object?> CreateSetter(PropertyInfo property) {
                try {
                    var target = Expression.Parameter(typeof(TTarget), "target");
                    var value = Expression.Parameter(typeof(object), "value");
                    var converted = Expression.Convert(value, property.PropertyType);
                    var body = Expression.Assign(Expression.Property(target, property), converted);
                    return Expression.Lambda<Action<TTarget, object?>>(body, target, value).Compile();
                } catch {
                    return (target, value) => property.SetValue(target, value);
                }
            }

            private static Action<TTarget, TValue>? CreateTypedSetter<TValue>(PropertyInfo property) {
                try {
                    var target = Expression.Parameter(typeof(TTarget), "target");
                    var value = Expression.Parameter(typeof(TValue), "value");
                    var converted = Expression.Convert(value, property.PropertyType);
                    var body = Expression.Assign(Expression.Property(target, property), converted);
                    return Expression.Lambda<Action<TTarget, TValue>>(body, target, value).Compile();
                } catch {
                    return null;
                }
            }

            private static TypedPropertyMapCache CreatePropertyMaps() {
                string typeName = typeof(TTarget).Name;
                var diagnostics = new List<string>();

                return new TypedPropertyMapCache(
                    typeName,
                    BuildPropertyMap(WritableProperties, prop => new[] { prop.Name }, diagnostics, typeName, "exact property"),
                    BuildPropertyMap(WritableProperties, GetPropertyAliases, diagnostics, typeName, "explicit alias"),
                    BuildPropertyMap(WritableProperties, prop => new[] { CanonicalizeMemberName(prop.Name) }, diagnostics, typeName, "friendly property"),
                    BuildPropertyMap(WritableProperties, prop => GetPropertyAliases(prop).Select(CanonicalizeMemberName), diagnostics, typeName, "friendly alias"),
                    diagnostics.ToArray());
            }

            private static TypedHeaderBindingCache<TTarget> CreateHeaderBindings(string[] headers) {
                var map = new TypedPropertyBinding<TTarget>?[headers.Length];
                var assignedProps = new HashSet<PropertyInfo>();

                // Exact property matches win first so alias/friendly fallback does not steal
                // a property from a later exact-name column.
                for (int c = 0; c < headers.Length; c++) {
                    if (PropertyMaps.ExactProperties.TryGetValue(headers[c], out var pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                // Explicit aliases come next (DisplayName/DataMember/ExcelColumn).
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    if (PropertyMaps.ExactAliases.TryGetValue(headers[c], out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string canonicalHeader = CanonicalizeMemberName(headers[c]);
                    if (canonicalHeader.Length == 0) {
                        continue;
                    }

                    if (PropertyMaps.CanonicalProperties.TryGetValue(canonicalHeader, out var pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                        continue;
                    }

                    if (PropertyMaps.CanonicalAliases.TryGetValue(canonicalHeader, out pi) && !assignedProps.Contains(pi)) {
                        map[c] = Bindings[pi];
                        assignedProps.Add(pi);
                    }
                }

                var unmappedIssues = new List<string>();
                for (int c = 0; c < headers.Length; c++) {
                    if (map[c] != null) {
                        continue;
                    }

                    string header = headers[c];
                    if (header.Length == 0) {
                        continue;
                    }

                    unmappedIssues.Add($"[TypedRead UnmappedHeader] Type='{PropertyMaps.TypeName}', header='{header}', column={c + 1}.");
                }

                return new TypedHeaderBindingCache<TTarget>(map, unmappedIssues.ToArray());
            }

            private static string CreateHeaderBindingKey(string[] headers) {
                var builder = new StringBuilder(headers.Length * 16);
                for (int i = 0; i < headers.Length; i++) {
                    string header = headers[i] ?? string.Empty;
                    builder.Append(header.Length.ToString(CultureInfo.InvariantCulture));
                    builder.Append(':');
                    builder.Append(header);
                    builder.Append('|');
                }

                return builder.ToString();
            }
        }

        private object? TryChangeType<TTarget>(object value, TypedPropertyBinding<TTarget> binding, CultureInfo culture) {
            if (value == null) return null;
            var srcType = value.GetType();
            if (binding.PropertyType.IsAssignableFrom(srcType)) return value;

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, binding.DestinationType, culture);
                if (ok) return v;
            }

            return binding.ConvertValue(value, culture);
        }

        private bool TryConvertCellForBinding<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (_opt.CellValueConverter != null || _opt.TypeConverter != null) {
                object? value = ConvertCell(cell);
                if (value is null) {
                    return binding.IsNullable;
                }

                converted = TryChangeType(value, binding, _opt.Culture);
                return converted is not null || binding.IsNullable;
            }

            CellValues? typeHint = cell.DataType?.Value;
            bool hasFormula = cell.CellFormula is not null;
            string? formulaText = hasFormula ? ExtractFormulaText(cell) : null;
            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            string? rawText = preferFormulaText ? null : ExtractRawText(cell);
            string? inlineText = preferFormulaText ? null : ExtractInlineString(cell, typeHint);

            if (rawText == null && inlineText == null && formulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (hasFormula && (!_opt.UseCachedFormulaResult || rawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = formulaText ?? rawText ?? inlineText;
                    return converted is not null || binding.IsNullable;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (!string.IsNullOrEmpty(inlineText)
                && ReturnBindingConversion(TryConvertStringForBinding(inlineText, binding, out converted), binding, converted)) {
                return true;
            }

            if (typeHint == CellValues.SharedString) {
                string? text = rawText;
                if (TryParseSharedStringIndex(rawText, out int sstIndex)) {
                    text = _sst.Get(sstIndex);
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(text, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Boolean && rawText != null) {
                if (ReturnBindingConversion(TryConvertBooleanForBinding(rawText == "1", binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.String || typeHint == CellValues.InlineString) {
                if (ReturnBindingConversion(TryConvertStringForBinding(rawText ?? inlineText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (typeHint == CellValues.Date && rawText != null) {
                if (DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)
                    && ReturnBindingConversion(TryConvertDateTimeForBinding(dt, binding, out converted), binding, converted)) {
                    return true;
                }

                if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                    return true;
                }

                return TryConvertCellForBindingFallback(cell, binding, out converted);
            }

            if (rawText == null) {
                return binding.IsNullable;
            }

            uint? styleIndex = null;
            if (_opt.TreatDatesUsingNumberFormat && binding.NeedsDateStyleConversion) {
                styleIndex = cell.StyleIndex?.Value;
                if (styleIndex is not null && _styles.IsDateLike(styleIndex.Value)) {
                    if ((TryParseInvariantDoubleFast(rawText, out var oa)
                            || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))
                        && ReturnBindingConversion(TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted), binding, converted)) {
                        return true;
                    }

                    if (ReturnBindingConversion(TryConvertStringForBinding(rawText, binding, out converted), binding, converted)) {
                        return true;
                    }

                    return TryConvertCellForBindingFallback(cell, binding, out converted);
                }
            }

            if (ReturnBindingConversion(TryConvertNumericTextForBinding(rawText, binding, out converted), binding, converted)) {
                return true;
            }

            return TryConvertCellForBindingFallback(cell, binding, out converted);
        }

        private bool TryConvertCellForBindingFallback<TTarget>(
            DocumentFormat.OpenXml.Spreadsheet.Cell cell,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            var raw = SnapshotCell(cell);
            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null && !CellHasExplicitBlank(cell)) {
                return binding.IsNullable;
            }

            if (TryConvertRawForBinding(raw, binding, out converted)) {
                return converted is not null || binding.IsNullable;
            }

            object? typedValue = ConvertRaw(raw).TypedValue;
            if (typedValue is null) {
                return binding.IsNullable;
            }

            converted = TryChangeType(typedValue, binding, _opt.Culture);
            return converted is not null || binding.IsNullable;
        }

        private static bool ReturnBindingConversion<TTarget>(
            bool convertedByFastPath,
            TypedPropertyBinding<TTarget> binding,
            object? converted) {
            return convertedByFastPath && (converted is not null || binding.IsNullable);
        }

        private bool TryConvertRawForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (raw.HasFormula && (!_opt.UseCachedFormulaResult || raw.RawText == null)) {
                if (binding.DestinationType == typeof(string)) {
                    converted = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                    return true;
                }

                return false;
            }

            if (!string.IsNullOrEmpty(raw.InlineText)) {
                return TryConvertStringForBinding(raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                if (!TryParseSharedStringIndex(raw.RawText, out int sstIndex)) {
                    return TryConvertStringForBinding(raw.RawText, binding, out converted);
                }

                return TryConvertStringForBinding(_sst.Get(sstIndex), binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean && raw.RawText != null) {
                return TryConvertBooleanForBinding(raw.RawText == "1", binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return TryConvertStringForBinding(raw.RawText ?? raw.InlineText, binding, out converted);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date && raw.RawText != null) {
                if (DateTime.TryParse(raw.RawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                    return TryConvertDateTimeForBinding(dt, binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            if (raw.RawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && raw.StyleIndex is not null
                && _styles.IsDateLike(raw.StyleIndex.Value)) {
                if (TryParseInvariantDoubleFast(raw.RawText, out var oa)
                    || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    return TryConvertDateTimeForBinding(DateTime.FromOADate(oa), binding, out converted);
                }

                return TryConvertStringForBinding(raw.RawText, binding, out converted);
            }

            return TryConvertNumericTextForBinding(raw.RawText, binding, out converted);
        }

        private bool TryConvertRawCellForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;

            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null) {
                return binding.IsNullable;
            }

            if (TryConvertRawForBinding(raw, binding, out converted)) {
                return converted is not null || binding.IsNullable;
            }

            object? typedValue = ConvertRaw(raw).TypedValue;
            if (typedValue is null) {
                return binding.IsNullable;
            }

            converted = TryChangeType(typedValue, binding, _opt.Culture);
            return converted is not null || binding.IsNullable;
        }

        private bool TrySetRawCellForBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (raw.RawText == null && raw.InlineText == null && raw.FormulaText == null) {
                if (binding.IsNullable) {
                    binding.SetValue(target, null);
                    return true;
                }

                return false;
            }

            if (raw.HasFormula && (!_opt.UseCachedFormulaResult || raw.RawText == null)) {
                if (binding.BindingKind == TypedBindingKind.String) {
                    string? formulaValue = raw.FormulaText ?? raw.RawText ?? raw.InlineText;
                    SetStringBinding(binding, target, formulaValue);
                    return formulaValue is not null || binding.IsNullable;
                }

                return false;
            }

            if (!string.IsNullOrEmpty(raw.InlineText)) {
                if (TrySetStringTextBinding(raw.InlineText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                string? text = TryParseSharedStringIndex(raw.RawText, out int sstIndex)
                    ? _sst.Get(sstIndex)
                    : raw.RawText;
                if (TrySetStringTextBinding(text, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean && raw.RawText != null) {
                bool boolValue = raw.RawText == "1";
                if (binding.SetBoolean != null && binding.BindingKind == TypedBindingKind.Boolean) {
                    binding.SetBoolean(target, boolValue);
                    return true;
                }

                if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                    binding.SetString(target, boolValue.ToString());
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                if (TrySetStringTextBinding(raw.RawText ?? raw.InlineText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.TypeHint == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date && raw.RawText != null) {
                if (binding.SetDateTime != null
                    && DateTime.TryParse(raw.RawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dateValue)) {
                    binding.SetDateTime(target, dateValue);
                    return true;
                }

                if (TrySetStringTextBinding(raw.RawText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (raw.RawText == null) {
                return false;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && raw.StyleIndex is not null
                && _styles.IsDateLike(raw.StyleIndex.Value)) {
                if (TryParseInvariantDoubleFast(raw.RawText, out var oa)
                    || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    DateTime dateValue = DateTime.FromOADate(oa);
                    if (binding.SetDateTime != null && binding.BindingKind == TypedBindingKind.DateTime) {
                        binding.SetDateTime(target, dateValue);
                        return true;
                    }

                    if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                        binding.SetString(target, dateValue.ToString(_opt.Culture));
                        return true;
                    }
                }

                if (TrySetStringTextBinding(raw.RawText, binding, target)) {
                    return true;
                }

                return TrySetRawCellForBindingFallback(raw, binding, target);
            }

            if (TrySetNumericTextBinding(raw.RawText, binding, target)) {
                return true;
            }

            return TrySetRawCellForBindingFallback(raw, binding, target);
        }

        private bool TrySetRawCellForBindingFallback<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (TryConvertRawCellForBinding(raw, binding, out object? converted)) {
                binding.SetValue(target, converted);
                return true;
            }

            return false;
        }

        private bool TrySetStringTextBinding<TTarget>(
            string? text,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            if (text == null) {
                if (binding.IsNullable) {
                    binding.SetValue(target, null);
                    return true;
                }

                return false;
            }

            if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                binding.SetString(target, text);
                return true;
            }

            if (binding.SetBoolean != null
                && binding.BindingKind == TypedBindingKind.Boolean
                && bool.TryParse(text, out bool boolValue)) {
                binding.SetBoolean(target, boolValue);
                return true;
            }

            if (binding.SetDateTime != null
                && binding.BindingKind == TypedBindingKind.DateTime
                && DateTime.TryParse(text, _opt.Culture, DateTimeStyles.AssumeLocal, out var dateValue)) {
                binding.SetDateTime(target, dateValue);
                return true;
            }

            return TrySetNumericTextBinding(text, binding, target);
        }

        private bool TrySetNumericTextBinding<TTarget>(
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            switch (binding.BindingKind) {
                case TypedBindingKind.Int32: {
                    if (binding.SetInt32 == null) {
                        return false;
                    }

                    if (int.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out int intValue)) {
                        binding.SetInt32(target, intValue);
                        return true;
                    }

                    if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        binding.SetInt32(target, (int)doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Int64: {
                    if (binding.SetInt64 == null) {
                        return false;
                    }

                    if (long.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out long longValue)) {
                        binding.SetInt64(target, longValue);
                        return true;
                    }

                    if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                        && doubleValue >= long.MinValue
                        && doubleValue <= long.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        binding.SetInt64(target, (long)doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Double: {
                    if (binding.SetDouble == null) {
                        return false;
                    }

                    if ((_opt.Culture != CultureInfo.InvariantCulture
                            && double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue))
                        || TryParseInvariantDoubleFast(rawText, out doubleValue)
                        || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue)) {
                        binding.SetDouble(target, doubleValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Decimal: {
                    if (binding.SetDecimal == null) {
                        return false;
                    }

                    if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out decimal decimalValue)
                        || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out decimalValue)) {
                        binding.SetDecimal(target, decimalValue);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.Boolean: {
                    if (binding.SetBoolean == null) {
                        return false;
                    }

                    if (rawText == "1") {
                        binding.SetBoolean(target, true);
                        return true;
                    }

                    if (rawText == "0") {
                        binding.SetBoolean(target, false);
                        return true;
                    }

                    return false;
                }

                case TypedBindingKind.String: {
                    if (binding.SetString == null) {
                        return false;
                    }

                    binding.SetString(target, rawText);
                    return true;
                }

                default:
                    return false;
            }
        }

        private static void SetStringBinding<TTarget>(
            TypedPropertyBinding<TTarget> binding,
            TTarget target,
            string? value) {
            if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                binding.SetString(target, value);
            } else {
                binding.SetValue(target, value);
            }
        }

        private bool ShouldRetryRawDateStyledNumericBinding<TTarget>(
            CellRaw raw,
            TypedPropertyBinding<TTarget> binding) {
            if (!_opt.TreatDatesUsingNumberFormat
                || binding.NeedsDateStyleConversion
                || !IsNumericBindingDestination(binding.BindingKind)
                || raw.RawText == null
                || raw.StyleIndex is null
                || !_styles.IsDateLike(raw.StyleIndex.Value)) {
                return false;
            }

            return TryParseInvariantDoubleFast(raw.RawText, out _)
                || double.TryParse(raw.RawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out _);
        }

        private static bool IsNumericBindingDestination(TypedBindingKind bindingKind) {
            return bindingKind == TypedBindingKind.Int32
                || bindingKind == TypedBindingKind.Int64
                || bindingKind == TypedBindingKind.Double
                || bindingKind == TypedBindingKind.Decimal;
        }

        private bool TryConvertNumericTextForBinding<TTarget>(
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;

            if (destinationType == typeof(int)) {
                if (int.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out int intValue)) {
                    converted = intValue;
                    return true;
                }

                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                    && doubleValue >= int.MinValue
                    && doubleValue <= int.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (int)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(long)) {
                if (long.TryParse(rawText, NumberStyles.Integer, _opt.Culture, out long longValue)) {
                    converted = longValue;
                    return true;
                }

                if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue)
                    && doubleValue >= long.MinValue
                    && doubleValue <= long.MaxValue
                    && Math.Truncate(doubleValue) == doubleValue) {
                    converted = (long)doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(double)) {
                if ((_opt.Culture != CultureInfo.InvariantCulture
                        && double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out double doubleValue))
                    || TryParseInvariantDoubleFast(rawText, out doubleValue)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue)) {
                    converted = doubleValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(decimal)) {
                if (decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out decimal decimalValue)
                    || decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out decimalValue)) {
                    converted = decimalValue;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(bool)) {
                if (rawText == "1") {
                    converted = true;
                    return true;
                }

                if (rawText == "0") {
                    converted = false;
                    return true;
                }

                return false;
            }

            if (destinationType == typeof(string)) {
                converted = rawText;
                return true;
            }

            return false;
        }

        private bool TryConvertStringForBinding<TTarget>(
            string? text,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            if (text == null) {
                return binding.IsNullable;
            }

            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(string)) {
                converted = text;
                return true;
            }

            if (destinationType == typeof(bool) && bool.TryParse(text, out bool boolValue)) {
                converted = boolValue;
                return true;
            }

            if (destinationType == typeof(DateTime)
                && DateTime.TryParse(text, _opt.Culture, DateTimeStyles.AssumeLocal, out var dt)) {
                converted = dt;
                return true;
            }

            return TryConvertNumericTextForBinding(text, binding, out converted);
        }

        private static bool TryConvertBooleanForBinding<TTarget>(
            bool value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(bool)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString();
                return true;
            }

            return false;
        }

        private bool TryConvertDateTimeForBinding<TTarget>(
            DateTime value,
            TypedPropertyBinding<TTarget> binding,
            out object? converted) {
            converted = null;
            Type destinationType = binding.DestinationType;
            if (destinationType == typeof(DateTime)) {
                converted = value;
                return true;
            }

            if (destinationType == typeof(string)) {
                converted = value.ToString(_opt.Culture);
                return true;
            }

            return false;
        }

        private static Dictionary<string, PropertyInfo> BuildPropertyMap(
            IEnumerable<PropertyInfo> props,
            Func<PropertyInfo, IEnumerable<string>> candidateFactory,
            ICollection<string> diagnostics,
            string typeName,
            string mappingKind) {
            var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            var ambiguous = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ambiguousProps = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var prop in props) {
                foreach (string rawCandidate in candidateFactory(prop)) {
                    if (string.IsNullOrWhiteSpace(rawCandidate)) {
                        continue;
                    }

                    string candidate = rawCandidate;
                    if (candidate.Length == 0 || ambiguous.Contains(candidate)) {
                        continue;
                    }

                    if (map.TryGetValue(candidate, out var existing) && existing != prop) {
                        map.Remove(candidate);
                        ambiguous.Add(candidate);
                        if (!ambiguousProps.TryGetValue(candidate, out var propNames)) {
                            propNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { existing.Name };
                            ambiguousProps[candidate] = propNames;
                        }

                        propNames.Add(prop.Name);
                        continue;
                    }

                    if (ambiguousProps.TryGetValue(candidate, out var existingAmbiguousNames)) {
                        existingAmbiguousNames.Add(prop.Name);
                        continue;
                    }

                    map[candidate] = prop;
                }
            }

            foreach (var pair in ambiguousProps.OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)) {
                diagnostics.Add(
                    $"[TypedRead AmbiguousMapping] Type='{typeName}', match='{mappingKind}', header='{pair.Key}', properties='{string.Join(", ", pair.Value.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}'.");
            }

            return map;
        }

        private static IEnumerable<string> GetPropertyAliases(PropertyInfo propertyInfo) {
            var yielded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void YieldIfUnique(string? candidate, List<string> buffer) {
                if (!string.IsNullOrWhiteSpace(candidate)) {
                    string text = candidate!;
                    if (yielded.Add(text)) {
                        buffer.Add(text);
                    }
                }
            }

            var aliases = new List<string>();

            var displayName = propertyInfo.GetCustomAttribute<DisplayNameAttribute>(inherit: true);
            YieldIfUnique(displayName?.DisplayName, aliases);

            var dataMember = propertyInfo.GetCustomAttribute<DataMemberAttribute>(inherit: true);
            YieldIfUnique(dataMember?.Name, aliases);

            var excelColumn = propertyInfo.GetCustomAttribute<ExcelColumnAttribute>(inherit: true);
            if (excelColumn != null) {
                YieldIfUnique(excelColumn.Name, aliases);
                foreach (string alias in excelColumn.Aliases) {
                    YieldIfUnique(alias, aliases);
                }
            }

            return aliases;
        }

        private static string CanonicalizeMemberName(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            string text = value ?? string.Empty;
            var builder = new StringBuilder(text.Length);
            foreach (char character in text) {
                if (char.IsLetterOrDigit(character)) {
                    builder.Append(char.ToUpperInvariant(character));
                }
            }

            return builder.ToString();
        }

        private object? TryChangeType(object value, Type targetType, CultureInfo culture) {
            if (value == null) return null;
            var srcType = value.GetType();
            if (targetType.IsAssignableFrom(srcType)) return value;

            var nullable = Nullable.GetUnderlyingType(targetType);
            var destType = nullable ?? targetType;

            var hook = _opt.TypeConverter;
            if (hook != null) {
                var (ok, v) = hook(value, destType, culture);
                if (ok) return v;
            }

            return ConvertToDestinationType(value, destType, culture);
        }

        private static object? ConvertToDestinationType(object value, Type destType, CultureInfo culture) {
            try {
                if (destType == typeof(string)) {
                    return value as string ?? Convert.ToString(value, culture);
                }

                if (destType == typeof(bool)) {
                    if (value is bool boolValue) return boolValue;
                    return Convert.ToBoolean(value, culture);
                }

                if (destType == typeof(int)) {
                    if (value is int intValue) return intValue;
                    if (value is double doubleValue
                        && doubleValue >= int.MinValue
                        && doubleValue <= int.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (int)doubleValue;
                    }

                    return Convert.ToInt32(value, culture);
                }

                if (destType == typeof(long)) {
                    if (value is long longValue) return longValue;
                    if (value is double doubleValue
                        && doubleValue >= long.MinValue
                        && doubleValue <= long.MaxValue
                        && Math.Truncate(doubleValue) == doubleValue) {
                        return (long)doubleValue;
                    }

                    return Convert.ToInt64(value, culture);
                }

                if (destType == typeof(double)) {
                    if (value is double doubleValue) return doubleValue;
                    return Convert.ToDouble(value, culture);
                }

                if (destType == typeof(decimal)) {
                    if (value is decimal decimalValue) return decimalValue;
                    return Convert.ToDecimal(value, culture);
                }

                if (destType == typeof(DateTime)) {
                    if (value is DateTime dt) return dt;
                    if (value is double oa) return DateTime.FromOADate(oa);
                    if (DateTime.TryParse(Convert.ToString(value, culture), culture, DateTimeStyles.AssumeLocal, out var parsed)) return parsed;
                    return null;
                }

                return Convert.ChangeType(value, destType, culture);
            } catch {
                return null;
            }
        }
    }
}
