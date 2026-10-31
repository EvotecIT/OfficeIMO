using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.IO;
using System.Threading;
using OpenXmlCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    // Header-based helpers for addressing cells and columns by header name.
    public partial class ExcelSheet {
        private ref Dictionary<string, int>? _headerMapCache => ref Persistence.ReadCaches.HeaderMap;
        private ref string? _headerMapSourceA1 => ref Persistence.ReadCaches.HeaderMapSourceA1;
        private ref int _headerMapHeaderRowIndex => ref Persistence.ReadCaches.HeaderRowIndex;
        private ref int _headerMapHeaderCellCount => ref Persistence.ReadCaches.HeaderCellCount;
        private ref int _headerMapFirstColumnIndex => ref Persistence.ReadCaches.FirstColumnIndex;
        private ref int _headerMapLastColumnIndex => ref Persistence.ReadCaches.LastColumnIndex;
        private ref bool _headerMapNormalize => ref Persistence.ReadCaches.NormalizeHeaders;
        private ref bool _headerMapCachePopulated => ref Persistence.ReadCaches.HeaderMapPopulated;
        private object _headerMapLock => Persistence.ReadCaches.HeaderMapLock;
        private static readonly ExcelReaderOptions DefaultHeaderReadOptions = new ExcelReaderOptions();

        /// <summary>
        /// Builds or returns a cached case-insensitive map of header name to 1-based column index using the first row of UsedRange.
        /// Cache reuse respects the used range, header conversion options and named-style replacements.
        /// </summary>
        public Dictionary<string, int> GetHeaderMap(ExcelReaderOptions? options = null) {
            var opt = options ?? DefaultHeaderReadOptions;
            return new Dictionary<string, int>(GetHeaderMapCached(opt), StringComparer.OrdinalIgnoreCase);
        }

        private Dictionary<string, int> GetHeaderMapCached(ExcelReaderOptions opt) {
            opt.CancellationToken.ThrowIfCancellationRequested();
            if (Volatile.Read(ref _headerMapCachePopulated)) {
                lock (_headerMapLock) {
                    if (_headerMapCache != null && HeaderMapOptionsMatch(opt)) {
                        if (HeaderMapCacheCanReturnWithoutRebuild()) {
                            return _headerMapCache;
                        }

                        ClearHeaderMapCacheUnsafe();
                    }
                }
            }

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            if (!_hasWorksheetMutations
                && TryBuildHeaderMapFromReader(opt, out var readerMap, out var readerUsedRange, out int readerHeaderRow, out int readerFirstColumn, out int readerLastColumn, out int readerHeaderCellCount)) {
                lock (_headerMapLock) {
                    if (_headerMapCache != null && string.Equals(_headerMapSourceA1, readerUsedRange, StringComparison.Ordinal) && HeaderMapOptionsMatch(opt)) {
                        return _headerMapCache;
                    }

                    _headerMapCache = readerMap;
                    _headerMapSourceA1 = readerUsedRange;
                    _headerMapHeaderRowIndex = readerHeaderRow;
                    _headerMapHeaderCellCount = readerHeaderCellCount;
                    _headerMapFirstColumnIndex = readerFirstColumn;
                    _headerMapLastColumnIndex = readerLastColumn;
                    CaptureHeaderMapOptions(opt);
                    Volatile.Write(ref _headerMapCachePopulated, true);
                    return _headerMapCache;
                }
            }

            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            var a1Used = reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
            lock (_headerMapLock) {
                if (_headerMapCache != null && string.Equals(_headerMapSourceA1, a1Used, StringComparison.Ordinal) && HeaderMapOptionsMatch(opt)) {
                    return _headerMapCache;
                }
                var (r1, c1, _, c2) = ExcelA1.ParseRange(a1Used);
                if (TryBuildHeaderMapFromWorksheetDom(r1, c1, c2, opt, out var directMap)) {
                    _headerMapCache = directMap;
                    _headerMapSourceA1 = a1Used;
                    _headerMapHeaderRowIndex = r1;
                    _headerMapHeaderCellCount = CountHeaderRowCells(r1);
                    _headerMapFirstColumnIndex = c1;
                    _headerMapLastColumnIndex = c2;
                    CaptureHeaderMapOptions(opt);
                    Volatile.Write(ref _headerMapCachePopulated, true);
                    return _headerMapCache;
                }

                using var rdr = _excelDocument.CreateReader(opt);
                var sh = rdr.GetSheet(this.Name);
                string headerRange = ExcelA1.CellReference(r1, c1) + ":" + ExcelA1.CellReference(r1, c2);
                var values = sh.ReadRange(headerRange);

                if (values.GetLength(0) == 0 || values.GetLength(1) == 0) {
                    var empty = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    _headerMapCache = empty;
                    _headerMapSourceA1 = a1Used;
                    _headerMapHeaderRowIndex = r1;
                    _headerMapHeaderCellCount = CountHeaderRowCells(r1);
                    _headerMapFirstColumnIndex = c1;
                    _headerMapLastColumnIndex = c2;
                    CaptureHeaderMapOptions(opt);
                    Volatile.Write(ref _headerMapCachePopulated, true);
                    return _headerMapCache;
                }

                int cols = values.GetLength(1);
                var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                bool anyHeader = false;
                for (int c = 0; c < cols; c++) {
                    if (!string.IsNullOrEmpty(ExcelHeaderNameHelper.NormalizeHeader(values[0, c]?.ToString(), opt.NormalizeHeaders))) {
                        anyHeader = true;
                        break;
                    }
                }

                if (!anyHeader) {
                    _headerMapCache = map;
                    _headerMapSourceA1 = a1Used;
                    _headerMapHeaderRowIndex = r1;
                    _headerMapHeaderCellCount = CountHeaderRowCells(r1);
                    _headerMapFirstColumnIndex = c1;
                    _headerMapLastColumnIndex = c2;
                    CaptureHeaderMapOptions(opt);
                    Volatile.Write(ref _headerMapCachePopulated, true);
                    return _headerMapCache;
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), opt.NormalizeHeaders);

                for (int c = 0; c < cols; c++) {
                    map[headers[c]] = c1 + c;
                }
                _headerMapCache = map;
                _headerMapSourceA1 = a1Used;
                _headerMapHeaderRowIndex = r1;
                _headerMapHeaderCellCount = CountHeaderRowCells(r1);
                _headerMapFirstColumnIndex = c1;
                _headerMapLastColumnIndex = c2;
                CaptureHeaderMapOptions(opt);
                Volatile.Write(ref _headerMapCachePopulated, true);
                return _headerMapCache;
            }
        }

        private bool TryBuildHeaderMapFromReader(
            ExcelReaderOptions opt,
            out Dictionary<string, int> map,
            out string usedRange,
            out int headerRowIndex,
            out int firstColumnIndex,
            out int lastColumnIndex,
            out int headerCellCount) {
            map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            usedRange = string.Empty;
            headerRowIndex = 0;
            firstColumnIndex = 0;
            lastColumnIndex = 0;
            headerCellCount = 0;

            try {
                using var rdr = _excelDocument.CreateReader(opt);
                var sh = rdr.GetSheet(this.Name);
                usedRange = sh.GetUsedRangeA1();
                var (r1, c1, _, c2) = ExcelA1.ParseRange(usedRange);
                string headerRange = ExcelA1.CellReference(r1, c1) + ":" + ExcelA1.CellReference(r1, c2);
                object?[,] values = sh.ReadRange(headerRange);

                headerRowIndex = r1;
                firstColumnIndex = c1;
                lastColumnIndex = c2;

                if (values.GetLength(0) == 0 || values.GetLength(1) == 0) {
                    return true;
                }

                int cols = values.GetLength(1);
                headerCellCount = CountHeaderValues(values, cols);
                PopulateHeaderMapFromValues(map, values, cols, c1, opt.NormalizeHeaders);
                return true;
            } catch (ObjectDisposedException) {
                return false;
            } catch (ArgumentException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (InvalidOperationException) {
                return false;
            }
        }

        private static int CountHeaderValues(object?[,] values, int cols) {
            int count = 0;
            for (int c = 0; c < cols; c++) {
                if (values[0, c] != null) {
                    count++;
                }
            }

            return count;
        }

        private static void PopulateHeaderMapFromValues(Dictionary<string, int> map, object?[,] values, int cols, int firstColumn, bool normalizeHeaders) {
            bool anyHeader = false;
            for (int c = 0; c < cols; c++) {
                if (!string.IsNullOrEmpty(ExcelHeaderNameHelper.NormalizeHeader(values[0, c]?.ToString(), normalizeHeaders))) {
                    anyHeader = true;
                    break;
                }
            }

            if (!anyHeader) {
                return;
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), normalizeHeaders);
            for (int c = 0; c < cols; c++) {
                map[headers[c]] = firstColumn + c;
            }
        }

        private static bool HeaderMapOptionsReusable(ExcelReaderOptions options) =>
            options.CellValueConverter == null && ReferenceEquals(options.Culture, CultureInfo.InvariantCulture)
            && !options.NumericAsDecimal && options.UseCachedFormulaResult;

        private bool HeaderMapOptionsMatch(ExcelReaderOptions options) =>
            Persistence.ReadCaches.HeaderOptionsReusable && HeaderMapOptionsReusable(options)
            && _headerMapNormalize == options.NormalizeHeaders
            && Persistence.ReadCaches.TreatHeaderDates == options.TreatDatesUsingNumberFormat;

        private void CaptureHeaderMapOptions(ExcelReaderOptions options) {
            _headerMapNormalize = options.NormalizeHeaders;
            Persistence.ReadCaches.TreatHeaderDates = options.TreatDatesUsingNumberFormat;
            Persistence.ReadCaches.HeaderOptionsReusable = HeaderMapOptionsReusable(options);
            CellFormatCacheRevision? revision = GetCurrentCellFormatCacheRevision();
            Persistence.ReadCaches.HeaderStyleRevision = revision;
            Persistence.ReadCaches.HeaderStyleRevisionNumber = revision?.Number ?? 0;
        }

        private bool HeaderMapCacheCanReturnWithoutRebuild() {
            CellFormatCacheRevision? revision = GetCurrentCellFormatCacheRevision();
            if (!ReferenceEquals(revision, Persistence.ReadCaches.HeaderStyleRevision)
                || (revision?.Number ?? 0) != Persistence.ReadCaches.HeaderStyleRevisionNumber) return false;
            int headerRowIndex = Volatile.Read(ref _headerMapHeaderRowIndex);
            if (headerRowIndex <= 0) {
                return false;
            }

            if (CountHeaderRowCells(headerRowIndex) != Volatile.Read(ref _headerMapHeaderCellCount)) {
                return false;
            }

            return true;
        }

        private int CountHeaderRowCells(int headerRowIndex) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return 0;
            }

            int inferredRow = 0;
            foreach (var row in sheetData.Elements<Row>()) {
                int rowIndex = ExcelWorksheetCoordinates.GetRowIndex(row, ref inferredRow);

                if (rowIndex == headerRowIndex) {
                    return row.Elements<Cell>().Count();
                }
            }

            return 0;
        }

        private bool TryBuildHeaderMapFromWorksheetDom(int headerRowIndex, int firstColumn, int lastColumn, ExcelReaderOptions options, out Dictionary<string, int> map) {
            map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (options.CellValueConverter != null
                || options.Culture != CultureInfo.InvariantCulture
                || options.NumericAsDecimal
                || !options.UseCachedFormulaResult) {
                return false;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return true;
            }

            Row? headerRow = null;
            int inferredRow = 0;
            foreach (var row in sheetData.Elements<Row>()) {
                int rowIndex = ExcelWorksheetCoordinates.GetRowIndex(row, ref inferredRow);

                if (rowIndex == headerRowIndex) {
                    headerRow = row;
                    break;
                }
            }

            int columnCount = lastColumn - firstColumn + 1;
            var headerValues = new object?[columnCount];
            if (headerRow != null) {
                int nextColumnIndex = 1;
                var headerCells = new Cell?[columnCount];
                int maxSharedStringIndex = -1;
                foreach (var cell in headerRow.Elements<Cell>()) {
                    int columnIndex = ExcelWorksheetCoordinates.GetColumnIndex(cell, ref nextColumnIndex);

                    if (columnIndex < firstColumn || columnIndex > lastColumn) {
                        continue;
                    }

                    int offset = columnIndex - firstColumn;
                    headerCells[offset] = cell;
                    if (cell.DataType?.Value == OpenXmlCellValues.SharedString
                        && TryParseSharedStringIndex(cell.CellValue?.InnerText, out int sharedStringIndex)
                        && sharedStringIndex > maxSharedStringIndex) {
                        maxSharedStringIndex = sharedStringIndex;
                    }
                }

                if (options.TreatDatesUsingNumberFormat && HeaderCellsNeedReaderDateConversion(headerCells)) {
                    return false;
                }

                List<string>? sharedStrings = maxSharedStringIndex >= 0
                    ? LoadSharedStringTextsFromDom(maxSharedStringIndex)
                    : null;
                for (int c = 0; c < headerCells.Length; c++) {
                    if (headerCells[c] != null) {
                        headerValues[c] = ConvertHeaderCellFromDom(headerCells[c]!, sharedStrings);
                    }
                }
            }

            bool anyHeader = false;
            for (int c = 0; c < columnCount; c++) {
                if (!string.IsNullOrEmpty(ExcelHeaderNameHelper.NormalizeHeader(headerValues[c]?.ToString(), options.NormalizeHeaders))) {
                    anyHeader = true;
                    break;
                }
            }

            if (!anyHeader) {
                return true;
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(columnCount, c => headerValues[c]?.ToString(), options.NormalizeHeaders);
            for (int c = 0; c < columnCount; c++) {
                map[headers[c]] = firstColumn + c;
            }

            return true;
        }

        private bool HeaderCellsNeedReaderDateConversion(Cell?[] headerCells) {
            StylesCache? styles = null;
            for (int i = 0; i < headerCells.Length; i++) {
                Cell? cell = headerCells[i];
                if (cell?.StyleIndex == null || !HeaderCellCanUseDateStyle(cell)) {
                    continue;
                }

                styles ??= StylesCache.Build(_spreadSheetDocument);
                if (styles.HasDateStyles && styles.IsDateLike(cell.StyleIndex.Value)) {
                    return true;
                }
            }

            return false;
        }

        private static bool HeaderCellCanUseDateStyle(Cell cell) {
            OpenXmlCellValues? type = cell.DataType?.Value;
            return type == null || type == OpenXmlCellValues.Number;
        }

        private object? ConvertHeaderCellFromDom(Cell cell, List<string>? sharedStrings) {
            OpenXmlCellValues? type = cell.DataType?.Value;
            string? rawText = cell.CellValue?.InnerText;
            if (type == OpenXmlCellValues.SharedString) {
                if (!TryParseSharedStringIndex(rawText, out int index)) {
                    return rawText;
                }

                return sharedStrings != null && (uint)index < (uint)sharedStrings.Count ? sharedStrings[index] : rawText;
            }

            if (type == OpenXmlCellValues.InlineString || cell.InlineString != null) {
                if (cell.InlineString?.Text?.Text != null) {
                    return cell.InlineString.Text.Text;
                }

                return cell.InlineString?.HasChildren == true ? SharedStringCache.GetRunText(cell.InlineString) : rawText;
            }

            if (type == OpenXmlCellValues.Boolean && rawText != null) {
                return rawText == "1";
            }

            return rawText;
        }

        private List<string> LoadSharedStringTextsFromDom(int maxIndex) {
            var table = _spreadSheetDocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            if (table == null || maxIndex < 0) {
                return new List<string>();
            }

            int capacity = maxIndex + 1;
            if (table.Count?.Value is uint declaredCount && declaredCount < (uint)capacity) {
                capacity = (int)declaredCount;
            }

            var values = new List<string>(capacity);
            foreach (var item in table.Elements<SharedStringItem>()) {
                values.Add(item.Text?.Text ?? (item.HasChildren ? SharedStringCache.GetRunText(item) : string.Empty));
                if (values.Count > maxIndex) {
                    break;
                }
            }

            return values;
        }

        private static bool TryParseSharedStringIndex(string? rawText, out int index) {
            index = 0;
            if (string.IsNullOrEmpty(rawText)) {
                return false;
            }

            string text = rawText!;
            int parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U || parsed > (int.MaxValue - digit) / 10) {
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
            }

            index = parsed;
            return true;
        }

        /// <summary>
        /// Tries to resolve a 1-based column index for a given header.
        /// Returns <c>false</c> without throwing when the header cannot be found.
        /// </summary>
        public bool TryGetColumnIndexByHeader(string header, out int columnIndex, ExcelReaderOptions? options = null) {
            if (string.IsNullOrWhiteSpace(header)) {
                columnIndex = 0;
                return false;
            }

            var opt = options ?? DefaultHeaderReadOptions;
            var map = GetHeaderMapCached(opt);
            string normalizedHeader = ExcelHeaderNameHelper.NormalizeHeader(header, opt.NormalizeHeaders);
            return map.TryGetValue(normalizedHeader, out columnIndex);
        }

        /// <summary>
        /// Sets a cell value in the specified row by resolving the column using the header name.
        /// Does nothing when the header cannot be found.
        /// </summary>
        public void SetByHeader(int rowIndex, string header, object? value, ExcelReaderOptions? options = null) {
            if (rowIndex <= 0) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (!TryGetColumnIndexByHeader(header, out var col, options))
                return;

            WriteLockConditional(() => CellValueCore(rowIndex, col, value ?? string.Empty));
        }

        /// <summary>
        /// Clears the cached header map.
        /// </summary>
        public void ClearHeaderCache() {
            ClearHeaderCacheCore(markWorksheetMutation: true, invalidateHeaderMap: true);
        }

        private void ClearHeaderCacheForCellMutation(int rowIndex, int columnIndex = 0) {
            _hasCellValueDomWrites = true;

            if (_hasWorksheetMutations
                && _requiresSavePreparation
                && !Volatile.Read(ref _headerMapCachePopulated)
                && Volatile.Read(ref _cellTextSharedStringCache) == null
                && !Volatile.Read(ref _findFirstCacheHasValue)) {
                if (_isBatchOperation) {
                    _batchHasCellMutations = true;
                    return;
                }

                if (!_excelDocument.IsPackageDirtyWithoutPendingSaveCandidate) {
                    _excelDocument.MarkPackageDirty();
                }

                return;
            }

            ClearHeaderCacheCore(
                markWorksheetMutation: true,
                invalidateHeaderMap: HeaderMapMayBeAffectedByCellMutation(rowIndex, columnIndex));
        }

        private bool HeaderMapMayBeAffectedByCellMutation(int rowIndex, int columnIndex) {
            if (!Volatile.Read(ref _headerMapCachePopulated)) {
                return false;
            }

            int headerRowIndex = Volatile.Read(ref _headerMapHeaderRowIndex);
            if (headerRowIndex <= 0 || rowIndex <= headerRowIndex) {
                return true;
            }

            if (columnIndex > 0) {
                int firstColumnIndex = Volatile.Read(ref _headerMapFirstColumnIndex);
                int lastColumnIndex = Volatile.Read(ref _headerMapLastColumnIndex);
                if (firstColumnIndex <= 0 || lastColumnIndex <= 0 || columnIndex < firstColumnIndex || columnIndex > lastColumnIndex) {
                    return true;
                }
            }

            return false;
        }

        private void ClearHeaderCacheCore(bool markWorksheetMutation, bool invalidateHeaderMap) {
            if (markWorksheetMutation) {
                _hasWorksheetMutations = true;
                MarkRequiresSavePreparation();
                ClearCellTextSharedStringCache();
            }

            if (!invalidateHeaderMap) {
                return;
            }

            if (!Volatile.Read(ref _headerMapCachePopulated)) {
                return;
            }

            lock (_headerMapLock) {
                ClearHeaderMapCacheUnsafe();
            }
        }

        private void ClearHeaderMapCacheUnsafe() {
            _headerMapCache = null;
            _headerMapSourceA1 = null;
            _headerMapHeaderRowIndex = 0;
            _headerMapHeaderCellCount = 0;
            _headerMapFirstColumnIndex = 0;
            _headerMapLastColumnIndex = 0;
            Volatile.Write(ref _headerMapCachePopulated, false);
        }

        /// <summary>
        /// Forces rebuilding the header map for the current UsedRange and options.
        /// </summary>
        public void RefreshHeaderCache(ExcelReaderOptions? options = null) {
            ClearHeaderCacheCore(markWorksheetMutation: false, invalidateHeaderMap: true);
            GetHeaderMap(options);
        }
    }
}
