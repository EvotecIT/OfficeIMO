using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static DateTime DefaultDateTimeOffsetWriteStrategy(DateTimeOffset value) => value.LocalDateTime;

        /// <summary>
        /// Writes a DataSet directly to an XLSX package, using one worksheet and one Excel table per DataTable.
        /// This path is intended for export workloads where the caller does not need to keep editing the workbook object.
        /// </summary>
        public static IReadOnlyList<ExcelDataSetImportResult> WriteDataSet(
            Stream stream,
            DataSet dataSet,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeHeaders = true,
            bool includeAutoFilter = true,
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));
            if (dataSet.Tables.Count == 0) throw new ArgumentException("The DataSet must contain at least one DataTable.", nameof(dataSet));

            var model = DirectDataSetWorkbookModel.Create(dataSet, createTables: true, tableStyle, includeHeaders, includeAutoFilter, autoFit: false, DefaultDateTimeOffsetWriteStrategy, ct);
            if (stream.CanSeek) {
                PrepareDestinationStreamForWrite(stream);
            }

            DirectDataSetWorkbookWriter.Write(stream, model, ct);
            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            return model.Results;
        }

        private void RegisterDirectDataSetSaveCandidate(
            DataSet dataSet,
            bool createTables,
            TableStyle tableStyle,
            bool includeHeaders,
            bool includeAutoFilter,
            bool autoFit,
            IReadOnlyList<ExcelDataSetImportResult> results) {
            ClearDirectDataSetSaveCandidate();

            try {
                var model = DirectDataSetWorkbookModel.Create(
                    dataSet,
                    createTables,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    results);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(dataSet, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: true);
            } catch {
                ClearDirectDataSetSaveCandidate();
            }
        }

        private bool TryRegisterDeferredDirectDataSetImport(
            DataSet dataSet,
            bool createTables,
            TableStyle tableStyle,
            bool includeHeaders,
            bool includeAutoFilter,
            bool autoFit,
            CancellationToken ct,
            out IReadOnlyList<ExcelDataSetImportResult> results) {
            results = Array.Empty<ExcelDataSetImportResult>();
            ClearDirectDataSetSaveCandidate();

            try {
                var model = DirectDataSetWorkbookModel.Create(
                    dataSet,
                    createTables,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    ct,
                    snapshotTables: true);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(dataSet, model, MaterializeDeferredDataSetImport, isDeferred: true, subscribeToSourceChanges: false);
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                results = model.Results;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        private void ClearDirectDataSetSaveCandidate() {
            var candidate = _directDataSetSaveCandidate;
            if (candidate == null) {
                return;
            }

            _directDataSetSaveCandidate = null;
            candidate.Dispose();
        }

        private void MaterializeDeferredDataSetImport() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                return;
            }

            _directDataSetSaveCandidate = null;
            candidate.Dispose();

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(candidate.Model);
            } finally {
                _materializingDeferredDataSetImport = false;
            }
        }

        private void MaterializeDirectDataSetModel(DirectDataSetWorkbookModel model) {
            foreach (var sheetModel in model.Sheets) {
                ExcelSheet sheet = AddWorkSheet(sheetModel.SheetName, SheetNameValidationMode.Strict);
                if (sheetModel.Range.Length == 0) {
                    continue;
                }

                if (sheetModel.HasTable) {
                    sheet.InsertDataTableAsTable(
                        sheetModel.Table.ToDataTable(),
                        includeHeaders: sheetModel.IncludeHeaders,
                        tableName: sheetModel.TableName,
                        style: sheetModel.TableStyle,
                        includeAutoFilter: sheetModel.IncludeAutoFilter);
                } else {
                    sheet.InsertDataTable(
                        sheetModel.Table.ToDataTable(),
                        includeHeaders: sheetModel.IncludeHeaders);
                }

                if (sheetModel.AutoFitColumns && sheetModel.Table.ColumnCount > 0) {
                    sheet.AutoFitColumnsFor(Enumerable.Range(1, sheetModel.Table.ColumnCount));
                }
            }
        }

        private bool TryWriteDirectDataSetPackage(
            Stream destination,
            ExcelSaveOptions? options,
            bool updateDocumentState,
            CancellationToken ct,
            out string? skipReason) {
            skipReason = null;

            if (destination == null || !destination.CanWrite) {
                skipReason = "Destination stream must be writable.";
                return false;
            }

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (options?.SafePreflight == true || options?.SafeRepairDefinedNames == true) {
                skipReason = "Save preflight options require the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid) {
                skipReason = "No valid direct DataSet save candidate is available.";
                ClearDirectDataSetSaveCandidate();
                return false;
            }

            ct.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            DirectDataSetWorkbookWriter.Write(destination, candidate.Model, ct);
            try { destination.Flush(); } catch (NotSupportedException) { }
            if (destination.CanSeek) {
                destination.Seek(0, SeekOrigin.Begin);
            }

            if (updateDocumentState) {
                _packageDirty = false;
                _packagePropertiesDirty = false;
                _requiresSavePreflight = false;
                _unchangedPackageBytes = null;
                _packageContentTypesKnownNormalized = true;
                _simplePackageContentKnown = true;
            }

            return true;
        }

        private bool TrySaveDirectDataSetPackageToFile(string targetPath, ExcelSaveOptions? options, CancellationToken ct, out string? skipReason) {
            skipReason = null;
            var temporaryPath = CreateTemporarySavePath(targetPath);
            byte[]? packageBytes = null;

            try {
                using (var fs = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None)) {
                    if (!TryWriteDirectDataSetPackage(fs, options, updateDocumentState: false, ct, out skipReason)) {
                        return false;
                    }
                }

                packageBytes = File.ReadAllBytes(temporaryPath);

                try { _spreadSheetDocument.Dispose(); } catch { }
                ReplaceTargetFile(temporaryPath, targetPath);
                temporaryPath = string.Empty;
                ClearDirectDataSetSaveCandidate();
                ReloadFromBytes(packageBytes, simplePackageContentKnown: true);

                FilePath = targetPath;
                LastSaveDiagnostics = ExcelSaveDiagnostics.DirectDataSetPackage();
                return true;
            } catch (Exception ex) {
                skipReason = "Direct DataSet package writer failed: " + ex.Message;
                if (packageBytes != null) {
                    try {
                        ClearDirectDataSetSaveCandidate();
                        ReloadFromBytes(packageBytes, simplePackageContentKnown: true);
                    } catch {
                    }
                }

                return false;
            } finally {
                DeleteFileIfExists(temporaryPath);
            }
        }

        private sealed class DirectDataSetWorkbookModel {
            private DirectDataSetWorkbookModel(
                IReadOnlyList<DirectDataSetSheetModel> sheets,
                IReadOnlyList<ExcelDataSetImportResult> results,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                Sheets = sheets;
                Results = results;
                DateTimeOffsetWriteStrategy = dateTimeOffsetWriteStrategy;
            }

            internal IReadOnlyList<DirectDataSetSheetModel> Sheets { get; }

            internal IReadOnlyList<ExcelDataSetImportResult> Results { get; }

            internal Func<DateTimeOffset, DateTime> DateTimeOffsetWriteStrategy { get; }

            internal static DirectDataSetWorkbookModel Create(
                DataSet dataSet,
                bool createTables,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool autoFit,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                IReadOnlyList<ExcelDataSetImportResult>? importResults = null,
                bool snapshotTables = false) {
                var sheets = new List<DirectDataSetSheetModel>();
                var results = new List<ExcelDataSetImportResult>();
                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int index = 1;
                foreach (DataTable table in dataSet.Tables) {
                    ct.ThrowIfCancellationRequested();
                    var tableModel = snapshotTables
                        ? DirectDataSetTableModel.Snapshot(table, ct)
                        : DirectDataSetTableModel.Reference(table);
                    string requestedName = string.IsNullOrWhiteSpace(table.TableName)
                        ? "Table" + index.ToString(CultureInfo.InvariantCulture)
                        : table.TableName;
                    ExcelDataSetImportResult? importResult = importResults != null && index <= importResults.Count
                        ? importResults[index - 1]
                        : null;
                    string sheetName = importResult?.SheetName ?? GetUniqueSheetName(SanitizeSheetName(requestedName), usedSheetNames);
                    usedSheetNames.Add(sheetName);
                    string tableName = importResult?.TableName ?? GetUniqueName(SanitizeTableName(requestedName), usedTableNames, 255);
                    usedTableNames.Add(tableName);
                    int rowCount = tableModel.RowCount + (includeHeaders ? 1 : 0);
                    ValidateWorksheetBounds(tableModel, rowCount, requestedName);
                    string range = importResult?.Range ?? (tableModel.ColumnCount == 0 || rowCount == 0
                        ? string.Empty
                        : "A1:" + A1.CellReference(rowCount, tableModel.ColumnCount));

                    bool hasTable = createTables && range.Length > 0;
                    double[]? columnWidths = autoFit && tableModel.ColumnCount > 0
                        ? tableModel.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct)
                        : null;
                    var sheet = new DirectDataSetSheetModel(index, sheetName, hasTable ? tableName : null, range, tableModel, tableStyle, includeHeaders, includeAutoFilter, hasTable, autoFit, columnWidths);
                    sheets.Add(sheet);
                    results.Add(new ExcelDataSetImportResult(sheetName, hasTable ? tableName : null, range, tableModel.RowCount, tableModel.ColumnCount));
                    index++;
                }

                return new DirectDataSetWorkbookModel(sheets, results, dateTimeOffsetWriteStrategy ?? DefaultDateTimeOffsetWriteStrategy);
            }

            private static string GetUniqueSheetName(string baseName, HashSet<string> used) {
                string trimmed = baseName;
                if (string.IsNullOrWhiteSpace(trimmed)) {
                    int defaultIndex = 1;
                    string defaultCandidate = "Sheet1";
                    while (used.Contains(defaultCandidate)) {
                        defaultIndex++;
                        defaultCandidate = "Sheet" + defaultIndex.ToString(CultureInfo.InvariantCulture);
                    }

                    return defaultCandidate;
                }

                if (trimmed.Length > 31) {
                    trimmed = trimmed.Substring(0, 31);
                }

                if (!used.Contains(trimmed)) {
                    return trimmed;
                }

                int suffix = 2;
                while (true) {
                    string suffixText = " (" + suffix.ToString(CultureInfo.InvariantCulture) + ")";
                    int prefixLength = Math.Max(1, 31 - suffixText.Length);
                    string candidate = trimmed.Length > prefixLength
                        ? trimmed.Substring(0, prefixLength) + suffixText
                        : trimmed + suffixText;
                    if (!used.Contains(candidate)) {
                        return candidate;
                    }

                    suffix++;
                }
            }

            private static void ValidateWorksheetBounds(DirectDataSetTableModel table, int rowCount, string requestedName) {
                if (table.ColumnCount > A1.MaxColumns) {
                    throw new ArgumentException($"DataTable '{requestedName}' has {table.ColumnCount.ToString(CultureInfo.InvariantCulture)} columns, exceeding Excel's maximum of {A1.MaxColumns.ToString(CultureInfo.InvariantCulture)} columns.", nameof(table));
                }

                if (rowCount > A1.MaxRows) {
                    throw new ArgumentException($"DataTable '{requestedName}' has {rowCount.ToString(CultureInfo.InvariantCulture)} worksheet rows including headers, exceeding Excel's maximum of {A1.MaxRows.ToString(CultureInfo.InvariantCulture)} rows.", nameof(table));
                }
            }

            private static string GetUniqueName(string baseName, HashSet<string> used, int maxLength) {
                string trimmed = string.IsNullOrWhiteSpace(baseName) ? "Table" : baseName;
                if (trimmed.Length > maxLength) {
                    trimmed = trimmed.Substring(0, maxLength);
                }

                if (used.Add(trimmed)) {
                    return trimmed;
                }

                int suffix = 2;
                while (true) {
                    string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                    int prefixLength = Math.Max(1, maxLength - suffixText.Length);
                    string candidate = trimmed.Length > prefixLength
                        ? trimmed.Substring(0, prefixLength) + suffixText
                        : trimmed + suffixText;
                    if (used.Add(candidate)) {
                        return candidate;
                    }

                    suffix++;
                }
            }

            private static string SanitizeSheetName(string name) {
                string baseName = (name ?? string.Empty).Trim();
                baseName = baseName.Trim('\'', ' ');
                var builder = new StringBuilder(baseName.Length);
                foreach (char ch in baseName) {
                    builder.Append(ch is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? '_' : ch);
                }

                string value = _multipleUnderscoresRegex.Replace(builder.ToString().Trim(), "_");
                return value.Trim('_');
            }

            private static string SanitizeTableName(string name) {
                var builder = new StringBuilder(name.Length + 1);
                foreach (char ch in name) {
                    builder.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
                }

                string value = builder.ToString();
                if (string.IsNullOrWhiteSpace(value)) {
                    value = "Table";
                }

                if (!char.IsLetter(value[0]) && value[0] != '_') {
                    value = "_" + value;
                }

                return value;
            }
        }

        private sealed class DirectDataSetSheetModel {
            internal DirectDataSetSheetModel(
                int index,
                string sheetName,
                string? tableName,
                string range,
                DirectDataSetTableModel table,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool hasTable,
                bool autoFitColumns,
                double[]? columnWidths) {
                Index = index;
                SheetName = sheetName;
                TableName = tableName;
                Range = range;
                Table = table;
                TableStyle = tableStyle;
                IncludeHeaders = includeHeaders;
                IncludeAutoFilter = includeAutoFilter;
                HasTable = hasTable;
                AutoFitColumns = autoFitColumns;
                ColumnWidths = columnWidths;
            }

            internal int Index { get; }

            internal string SheetName { get; }

            internal string? TableName { get; }

            internal string Range { get; }

            internal DirectDataSetTableModel Table { get; }

            internal TableStyle TableStyle { get; }

            internal bool IncludeHeaders { get; }

            internal bool IncludeAutoFilter { get; }

            internal bool HasTable { get; }

            internal bool AutoFitColumns { get; }

            internal double[]? ColumnWidths { get; }
        }

        private sealed class DirectDataSetTableModel {
            private readonly DataTable? _sourceTable;
            private readonly DirectDataSetColumnModel[]? _columns;
            private readonly object?[][]? _rows;

            private DirectDataSetTableModel(DataTable sourceTable) {
                _sourceTable = sourceTable;
                _columns = CreateColumns(sourceTable);
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, object?[][] rows) {
                _columns = columns;
                _rows = rows;
            }

            internal static DirectDataSetTableModel Reference(DataTable table) => new DirectDataSetTableModel(table);

            internal static DirectDataSetTableModel Snapshot(DataTable table, CancellationToken ct) {
                var columns = CreateColumns(table);

                var rows = new object?[table.Rows.Count][];
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                    ct.ThrowIfCancellationRequested();
                    DataRow row = table.Rows[rowIndex];
                    var values = new object?[columns.Length];
                    for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
                        object? value = row[columnIndex];
                        values[columnIndex] = value == DBNull.Value ? null : value;
                    }

                    rows[rowIndex] = values;
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            private static DirectDataSetColumnModel[] CreateColumns(DataTable table) {
                var columns = new DirectDataSetColumnModel[table.Columns.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(table.Columns[i].ColumnName, table.Columns[i].DataType);
                }

                return columns;
            }

            internal int ColumnCount => _columns!.Length;

            internal int RowCount => _sourceTable?.Rows.Count ?? _rows!.Length;

            internal string GetColumnName(int index) => _columns![index].Name;

            internal Type GetColumnType(int index) => _columns![index].DataType;

            internal DirectDataSetRowModel GetRow(int rowIndex) {
                if (_sourceTable != null) {
                    return DirectDataSetRowModel.FromDataRow(_sourceTable.Rows[rowIndex]);
                }

                return DirectDataSetRowModel.FromSnapshot(_rows![rowIndex]);
            }

            internal double[] CalculateColumnWidths(bool includeHeaders, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, CancellationToken ct) {
                int columnCount = ColumnCount;
                var widths = new double[columnCount];
                if (columnCount == 0) {
                    return widths;
                }

                if (includeHeaders) {
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(GetColumnName(columnIndex)));
                    }
                }

                int rowCount = RowCount;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    ct.ThrowIfCancellationRequested();
                    var row = GetRow(rowIndex);
                    for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                        widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(row.GetValue(columnIndex), dateTimeOffsetWriteStrategy));
                    }
                }

                return widths;
            }

            private static double EstimateAutoFitWidth(string text) {
                if (string.IsNullOrEmpty(text)) {
                    return 0D;
                }

                int maxLineLength = 0;
                int currentLineLength = 0;
                for (int i = 0; i < text.Length; i++) {
                    char current = text[i];
                    if (current == '\r' || current == '\n') {
                        if (currentLineLength > maxLineLength) {
                            maxLineLength = currentLineLength;
                        }

                        currentLineLength = 0;
                        if (current == '\r' && i + 1 < text.Length && text[i + 1] == '\n') {
                            i++;
                        }
                    } else {
                        currentLineLength++;
                    }
                }

                if (currentLineLength > maxLineLength) {
                    maxLineLength = currentLineLength;
                }

                if (maxLineLength == 0) {
                    return 0D;
                }

                return Math.Min(255D, Math.Max(1D, maxLineLength + 2D));
            }

            private static double EstimateAutoFitWidth(object? value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                switch (value) {
                    case null:
                    case DBNull:
                        return 0D;
                    case string stringValue:
                        return EstimateAutoFitWidth(stringValue);
                    case bool boolValue:
                        return EstimateAutoFitWidthFromLength(boolValue ? 4 : 5);
                    case DateTime dateTime:
                        _ = dateTime;
                        return EstimateAutoFitWidthFromLength(16);
                    case DateTimeOffset dateTimeOffset:
                        try {
                            _ = dateTimeOffsetWriteStrategy(dateTimeOffset);
                            return EstimateAutoFitWidthFromLength(16);
                        } catch (ArgumentException) {
                            return EstimateAutoFitWidth(dateTimeOffset.ToString("o", CultureInfo.InvariantCulture));
                        } catch (OverflowException) {
                            return EstimateAutoFitWidth(dateTimeOffset.ToString("o", CultureInfo.InvariantCulture));
                        }
                    case TimeSpan timeSpan:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(timeSpan));
                    case double doubleValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(doubleValue));
                    case float floatValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(floatValue));
                    case decimal decimalValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(decimalValue));
                    case sbyte sbyteValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(sbyteValue));
                    case byte byteValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(byteValue));
                    case short shortValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(shortValue));
                    case ushort ushortValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ushortValue));
                    case int intValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(intValue));
                    case uint uintValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(uintValue));
                    case long longValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(longValue));
                    case ulong ulongValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ulongValue));
#if NET6_0_OR_GREATER
                    case DateOnly dateOnly:
                        _ = dateOnly;
                        return EstimateAutoFitWidthFromLength(10);
                    case TimeOnly timeOnly:
                        _ = timeOnly;
                        return EstimateAutoFitWidthFromLength(8);
#endif
                    case IFormattable formattable:
                        return EstimateAutoFitWidth(formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty);
                    default:
                        return EstimateAutoFitWidth(value.ToString() ?? string.Empty);
                }
            }

            private static double EstimateAutoFitWidthFromLength(int length) {
                if (length <= 0) {
                    return 0D;
                }

                return Math.Min(255D, Math.Max(1D, length + 2D));
            }

            private static int CountFormattedCharacters(double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(float value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(decimal value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[64];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(TimeSpan value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, "c", CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString("c", CultureInfo.InvariantCulture).Length;
            }

            private static int CountSignedIntegerCharacters(long value) {
                if (value < 0) {
                    ulong magnitude = (ulong)(-(value + 1)) + 1UL;
                    return 1 + CountUnsignedIntegerCharacters(magnitude);
                }

                return CountUnsignedIntegerCharacters((ulong)value);
            }

            private static int CountUnsignedIntegerCharacters(ulong value) {
                int count = 1;
                while (value >= 10UL) {
                    value /= 10UL;
                    count++;
                }

                return count;
            }

            internal DataTable ToDataTable() {
                if (_sourceTable != null) {
                    return _sourceTable;
                }

                var table = new DataTable { Locale = CultureInfo.InvariantCulture };
                foreach (var column in _columns!) {
                    table.Columns.Add(column.Name, column.DataType);
                }

                table.BeginLoadData();
                try {
                    foreach (var row in _rows!) {
                        var values = new object?[row.Length];
                        for (int i = 0; i < row.Length; i++) {
                            values[i] = row[i] ?? DBNull.Value;
                        }

                        table.Rows.Add(values);
                    }
                } finally {
                    table.EndLoadData();
                }

                return table;
            }
        }

        private readonly struct DirectDataSetRowModel {
            private readonly DataRow? _sourceRow;
            private readonly object?[]? _snapshotValues;

            private DirectDataSetRowModel(DataRow? sourceRow, object?[]? snapshotValues) {
                _sourceRow = sourceRow;
                _snapshotValues = snapshotValues;
            }

            internal static DirectDataSetRowModel FromDataRow(DataRow row) => new DirectDataSetRowModel(row, null);

            internal static DirectDataSetRowModel FromSnapshot(object?[] values) => new DirectDataSetRowModel(null, values);

            internal object? GetValue(int columnIndex) {
                object? value = _sourceRow != null
                    ? _sourceRow[columnIndex]
                    : _snapshotValues![columnIndex];
                return value == DBNull.Value ? null : value;
            }
        }

        private sealed class DirectDataSetColumnModel {
            internal DirectDataSetColumnModel(string name, Type dataType) {
                Name = name;
                DataType = dataType;
            }

            internal string Name { get; }

            internal Type DataType { get; }
        }

        private sealed class DirectDataSetSaveCandidate : IDisposable {
            private readonly DataSet _dataSet;
            private readonly Action _invalidate;
            private bool _disposed;

            internal DirectDataSetSaveCandidate(DataSet dataSet, DirectDataSetWorkbookModel model, Action invalidate, bool isDeferred, bool subscribeToSourceChanges) {
                _dataSet = dataSet;
                Model = model;
                _invalidate = invalidate;
                IsDeferred = isDeferred;
                if (subscribeToSourceChanges) {
                    Subscribe(dataSet);
                }
            }

            internal DirectDataSetWorkbookModel Model { get; }

            internal bool IsDeferred { get; }

            internal bool IsValid { get; private set; } = true;

            private void Subscribe(DataSet dataSet) {
                dataSet.Tables.CollectionChanged += OnCollectionChanged;
                foreach (DataTable table in dataSet.Tables) {
                    Subscribe(table);
                }
            }

            private void Subscribe(DataTable table) {
                table.Columns.CollectionChanged += OnCollectionChanged;
                table.RowChanged += OnDataChanged;
                table.RowChanging += OnDataChanging;
                table.RowDeleted += OnDataChanged;
                table.RowDeleting += OnDataChanging;
                table.ColumnChanged += OnColumnChanged;
                table.ColumnChanging += OnColumnChanging;
                table.TableCleared += OnDataChanged;
                table.TableClearing += OnDataChanging;
            }

            private void Unsubscribe(DataSet dataSet) {
                dataSet.Tables.CollectionChanged -= OnCollectionChanged;
                foreach (DataTable table in dataSet.Tables) {
                    Unsubscribe(table);
                }
            }

            private void Unsubscribe(DataTable table) {
                table.Columns.CollectionChanged -= OnCollectionChanged;
                table.RowChanged -= OnDataChanged;
                table.RowChanging -= OnDataChanging;
                table.RowDeleted -= OnDataChanged;
                table.RowDeleting -= OnDataChanging;
                table.ColumnChanged -= OnColumnChanged;
                table.ColumnChanging -= OnColumnChanging;
                table.TableCleared -= OnDataChanged;
                table.TableClearing -= OnDataChanging;
            }

            private void OnCollectionChanged(object? sender, CollectionChangeEventArgs e) => Invalidate();

            private void OnDataChanged(object sender, DataRowChangeEventArgs e) => Invalidate();

            private void OnDataChanging(object sender, DataRowChangeEventArgs e) => Invalidate();

            private void OnDataChanged(object sender, DataTableClearEventArgs e) => Invalidate();

            private void OnDataChanging(object sender, DataTableClearEventArgs e) => Invalidate();

            private void OnColumnChanged(object sender, DataColumnChangeEventArgs e) => Invalidate();

            private void OnColumnChanging(object sender, DataColumnChangeEventArgs e) => Invalidate();

            private void Invalidate() {
                if (!IsValid) {
                    return;
                }

                IsValid = false;
                _invalidate();
            }

            public void Dispose() {
                if (_disposed) {
                    return;
                }

                _disposed = true;
                try {
                    Unsubscribe(_dataSet);
                } catch {
                }
            }
        }
    }
}
