using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OpenXmlCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Header-based helpers for addressing cells and columns by header name.
    /// </summary>
    public partial class ExcelSheet {
        private Dictionary<string, int>? _headerMapCache;
        private string? _headerMapSourceA1;
        private bool _headerMapNormalize;
        private readonly object _headerMapLock = new object();
        private static readonly ExcelReadOptions DefaultHeaderReadOptions = new ExcelReadOptions();

        /// <summary>
        /// Builds or returns a cached case-insensitive map of header name to 1-based column index using the first row of UsedRange.
        /// Cache is keyed by UsedRange A1 and NormalizeHeaders option.
        /// </summary>
        public Dictionary<string, int> GetHeaderMap(ExcelReadOptions? options = null) {
            var opt = options ?? DefaultHeaderReadOptions;
            return new Dictionary<string, int>(GetHeaderMapCached(opt), StringComparer.OrdinalIgnoreCase);
        }

        private Dictionary<string, int> GetHeaderMapCached(ExcelReadOptions opt) {
            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            var a1Used = reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
            lock (_headerMapLock) {
                if (_headerMapCache != null && string.Equals(_headerMapSourceA1, a1Used, StringComparison.Ordinal) && _headerMapNormalize == opt.NormalizeHeaders) {
                    return _headerMapCache;
                }
                var (r1, c1, _, c2) = A1.ParseRange(a1Used);
                if (TryBuildHeaderMapFromWorksheetDom(r1, c1, c2, opt, out var directMap)) {
                    _headerMapCache = directMap;
                    _headerMapSourceA1 = a1Used;
                    _headerMapNormalize = opt.NormalizeHeaders;
                    return _headerMapCache;
                }

                using var rdr = _excelDocument.CreateReader(opt);
                var sh = rdr.GetSheet(this.Name);
                string headerRange = A1.CellReference(r1, c1) + ":" + A1.CellReference(r1, c2);
                var values = sh.ReadRange(headerRange);

                if (values.GetLength(0) == 0 || values.GetLength(1) == 0) {
                    var empty = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    _headerMapCache = empty;
                    _headerMapSourceA1 = a1Used;
                    _headerMapNormalize = opt.NormalizeHeaders;
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
                    _headerMapNormalize = opt.NormalizeHeaders;
                    return _headerMapCache;
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), opt.NormalizeHeaders);

                for (int c = 0; c < cols; c++) {
                    map[headers[c]] = c1 + c;
                }
                _headerMapCache = map;
                _headerMapSourceA1 = a1Used;
                _headerMapNormalize = opt.NormalizeHeaders;
                return _headerMapCache;
            }
        }

        private bool TryBuildHeaderMapFromWorksheetDom(int headerRowIndex, int firstColumn, int lastColumn, ExcelReadOptions options, out Dictionary<string, int> map) {
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
                int rowIndex;
                if (row.RowIndex != null) {
                    rowIndex = checked((int)row.RowIndex.Value);
                    inferredRow = rowIndex;
                } else {
                    rowIndex = ++inferredRow;
                }

                if (rowIndex < headerRowIndex) {
                    continue;
                }

                if (rowIndex == headerRowIndex) {
                    headerRow = row;
                }

                break;
            }

            int columnCount = lastColumn - firstColumn + 1;
            var headerValues = new object?[columnCount];
            if (headerRow != null) {
                int nextColumnIndex = 1;
                var headerCells = new Cell?[columnCount];
                int maxSharedStringIndex = -1;
                foreach (var cell in headerRow.Elements<Cell>()) {
                    int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (columnIndex <= 0) {
                        columnIndex = string.IsNullOrEmpty(cell.CellReference?.Value) ? nextColumnIndex : 0;
                    }

                    if (columnIndex > 0) {
                        nextColumnIndex = columnIndex + 1;
                    }

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
        public bool TryGetColumnIndexByHeader(string header, out int columnIndex, ExcelReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(header)) {
                columnIndex = 0;
                return false;
            }

            var opt = options ?? DefaultHeaderReadOptions;
            var map = GetHeaderMapCached(opt);
            return map.TryGetValue(header, out columnIndex);
        }

        /// <summary>
        /// Sets a cell value in the specified row by resolving the column using the header name.
        /// Does nothing when the header cannot be found.
        /// </summary>
        public void SetByHeader(int rowIndex, string header, object? value, ExcelReadOptions? options = null) {
            if (rowIndex <= 0) throw new ArgumentOutOfRangeException(nameof(rowIndex));
            if (!TryGetColumnIndexByHeader(header, out var col, options))
                return;
            if (value is null)
                CellValue(rowIndex, col, string.Empty);
            else
                CellValue(rowIndex, col, value);
        }

        /// <summary>
        /// Clears the cached header map.
        /// </summary>
        public void ClearHeaderCache() {
            _hasWorksheetMutations = true;
            MarkRequiresSavePreparation();
            ClearCellTextSharedStringCache();
            lock (_headerMapLock) {
                _headerMapCache = null;
                _headerMapSourceA1 = null;
            }
        }

        /// <summary>
        /// Forces rebuilding the header map for the current UsedRange and options.
        /// </summary>
        public void RefreshHeaderCache(ExcelReadOptions? options = null) {
            GetHeaderMap(options);
        }
    }
}
