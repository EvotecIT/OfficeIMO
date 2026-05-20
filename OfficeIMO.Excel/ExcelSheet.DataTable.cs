using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string DataTableDateTimeNumberFormat = "yyyy-mm-dd hh:mm";
        private const string DataTableTimeSpanNumberFormat = "[h]:mm:ss";

        /// <summary>
        /// Inserts a DataTable into the worksheet starting at the specified cell.
        /// Uses the batch CellValues compute/apply model with SharedString and Style planners.
        /// </summary>
        /// <param name="table">Source DataTable.</param>
        /// <param name="startRow">1-based start row.</param>
        /// <param name="startColumn">1-based start column.</param>
        /// <param name="includeHeaders">Whether to write column headers.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        public void InsertDataTable(DataTable table, int startRow = 1, int startColumn = 1, bool includeHeaders = true,
            ExecutionMode? mode = null, CancellationToken ct = default) {
            InsertDataTableCore(table, startRow, startColumn, includeHeaders, mode, ct, copyDirectSaveTable: true);
        }

        internal void InsertOwnedDataTable(DataTable table, int startRow = 1, int startColumn = 1, bool includeHeaders = true,
            ExecutionMode? mode = null, CancellationToken ct = default, bool registerDirectSaveCandidate = true) {
            InsertDataTableCore(table, startRow, startColumn, includeHeaders, mode, ct, copyDirectSaveTable: false, registerDirectSaveCandidate);
        }

        private void InsertDataTableCore(DataTable table, int startRow, int startColumn, bool includeHeaders,
            ExecutionMode? mode, CancellationToken ct, bool copyDirectSaveTable, bool registerDirectSaveCandidate = true) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));

            bool canRegisterDirectSave = registerDirectSaveCandidate
                && !_excelDocument.IsMaterializingDeferredDataSetImport
                && mode != ExecutionMode.Parallel
                && CanRegisterDirectTabularSaveCandidate(startRow, startColumn, table.Columns.Count);

            if (mode != ExecutionMode.Parallel && TryInsertDataTableByAppendingRows(table, startRow, startColumn, includeHeaders, ct)) {
                RegisterDirectDataTableSaveCandidateIfPossible(table, startRow, startColumn, includeHeaders, canRegisterDirectSave, copyDirectSaveTable);
                return;
            }

            // Prepare a flat list of cells and optional number formats
            var cells = new List<(int Row, int Col, object? Val, string? NumFmt)>(
                (table.Rows.Count + (includeHeaders ? 1 : 0)) * Math.Max(1, table.Columns.Count));

            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < table.Columns.Count; c++) {
                    cells.Add((row, startColumn + c, table.Columns[c].ColumnName, null));
                }
                row++;
            }

            foreach (DataRow dr in table.Rows) {
                for (int c = 0; c < table.Columns.Count; c++) {
                    var col = table.Columns[c];
                    object? value = dr.IsNull(c) ? null : dr[c];
                    string? fmt = GetDataTableNumberFormat(col.DataType, value);
                    cells.Add((row, startColumn + c, value, fmt));
                }
                row++;
            }

            if (cells.Count == 0) return;

            // Prepared buffers for compute/apply
            var prepared = new (int Row, int Col, CellValue Val, EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> Type)[cells.Count];
            var wrapFlags = new bool[cells.Count];
            var ssPlanner = new SharedStringPlanner();
            var stylePlanner = new StylePlanner();

            ExecuteWithPolicy(
                opName: "InsertDataTable",
                itemCount: cells.Count,
                overrideMode: mode,
                sequentialCore: () => {
                    for (int i = 0; i < cells.Count; i++) {
                        var (r, c, v, fmt) = cells[i];
                        // Direct cell write path
                        CellValueCore(r, c, v);
                        if (!string.IsNullOrWhiteSpace(fmt)) {
                            // Apply number format using existing API
                            FormatCell(r, c, fmt!);
                        }
                    }
                },
                computeParallel: () => {
                    Parallel.For(0, cells.Count, new ParallelOptions {
                        CancellationToken = ct,
                        MaxDegreeOfParallelism = EffectiveExecution.MaxDegreeOfParallelism ?? -1
                    }, i => {
                        var (r, c, obj, fmt) = cells[i];
                        var (val, type) = CoerceForCellNoDom(obj, ssPlanner);
                        val ??= new CellValue(string.Empty);
                        type ??= new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
                        if (type.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString && val.Text is string raw) {
                            if (raw.Contains("\n") || raw.Contains("\r"))
                                wrapFlags[i] = true;
                        }
                        if (!string.IsNullOrWhiteSpace(fmt))
                            stylePlanner.NoteNumberFormat(fmt);
                        prepared[i] = (r, c, val, type);
                    });
                },
                applySequential: () => {
                    // Apply planners
                    ssPlanner.ApplyAndFixup(prepared, _excelDocument);
                    stylePlanner.ApplyTo(_excelDocument);

                    var writer = new BatchCellWriter(this);
                    for (int i = 0; i < prepared.Length; i++) {
                        var p = prepared[i];
                        var cell = writer.GetOrCreateCell(p.Row, p.Col);
                        cell.CellValue = p.Val;
                        cell.DataType = p.Type;
                        if (wrapFlags[i])
                            ApplyWrapText(cell);

                        var fmt = cells[i].NumFmt;
                        if (!string.IsNullOrWhiteSpace(fmt) && stylePlanner.TryGetCellFormatIndex(fmt, out uint idx)) {
                            cell.StyleIndex = idx;
                        }
                    }

                    ClearHeaderCache();
                },
                ct: ct
            );

            RegisterDirectDataTableSaveCandidateIfPossible(table, startRow, startColumn, includeHeaders, canRegisterDirectSave, copyDirectSaveTable);
        }

        private void RegisterDirectDataTableSaveCandidateIfPossible(DataTable table, int startRow, int startColumn, bool includeHeaders, bool canRegisterDirectSave, bool copyDirectSaveTable) {
            if (!canRegisterDirectSave) {
                return;
            }

            string range = BuildDataTableInsertedRange(table, startRow, startColumn, includeHeaders);
            if (range.Length == 0) {
                return;
            }

            _excelDocument.RegisterDirectTabularSaveCandidate(this, table, includeHeaders, range, copyTable: copyDirectSaveTable);
        }

        private static string BuildDataTableInsertedRange(DataTable table, int startRow, int startColumn, bool includeHeaders) {
            int rowsCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (table.Columns.Count == 0 || rowsCount == 0) {
                return string.Empty;
            }

            return A1.CellReference(startRow, startColumn) + ":" +
                A1.CellReference(startRow + rowsCount - 1, startColumn + table.Columns.Count - 1);
        }

        private bool TryInsertDataTableByAppendingRows(DataTable table, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
            int columnCount = table.Columns.Count;
            int rowsCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (columnCount == 0 || rowsCount == 0) {
                return true;
            }

            if (startColumn + columnCount - 1 > A1.MaxColumns || startRow + rowsCount - 1 > A1.MaxRows) {
                return false;
            }

            bool applied = false;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryInsertDataTableByAppendingRowsCore(table, startRow, startColumn, includeHeaders, ct));
            return applied;
        }

        private bool TryInsertDataTableByAppendingRowsCore(DataTable table, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
            var sheetData = GetOrCreateSheetData();
            int minExistingRow = int.MaxValue;
            int minExistingColumn = int.MaxValue;
            int maxExistingRow = 0;
            int maxExistingColumn = 0;
            foreach (var existingRow in sheetData.Elements<Row>()) {
                if (existingRow.RowIndex == null) {
                    return false;
                }

                int existingRowIndex = checked((int)(existingRow.RowIndex?.Value ?? 0U));
                if (existingRowIndex >= startRow) {
                    return false;
                }

                if (existingRowIndex <= 0 || !existingRow.HasChildren) {
                    continue;
                }

                foreach (var existingCell in existingRow.Elements<Cell>()) {
                    int existingColumnIndex = 0;
                    string? reference = existingCell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(reference)) {
                        existingColumnIndex = A1.ParseColumnIndexFromCellReference(reference!);
                    }

                    if (existingColumnIndex <= 0) {
                        continue;
                    }

                    if (existingRowIndex < minExistingRow) minExistingRow = existingRowIndex;
                    if (existingRowIndex > maxExistingRow) maxExistingRow = existingRowIndex;
                    if (existingColumnIndex < minExistingColumn) minExistingColumn = existingColumnIndex;
                    if (existingColumnIndex > maxExistingColumn) maxExistingColumn = existingColumnIndex;
                }
            }

            int columnCount = table.Columns.Count;
            var columnNames = new string[startColumn + columnCount];
            for (int column = startColumn; column < startColumn + columnCount; column++) {
                columnNames[column] = GetColumnName(column);
            }

            string?[] numberFormats = BuildDataTableNumberFormats(table);
            var stylePlanner = new StylePlanner();
            bool hasObjectColumn = false;
            foreach (DataColumn column in table.Columns) {
                if (column.DataType == typeof(object)) {
                    hasObjectColumn = true;
                    break;
                }
            }

            foreach (string? numberFormat in numberFormats) {
                stylePlanner.NoteNumberFormat(numberFormat);
            }

            if (hasObjectColumn) {
                stylePlanner.NoteNumberFormat(DataTableDateTimeNumberFormat);
                stylePlanner.NoteNumberFormat(DataTableTimeSpanNumberFormat);
            }

            stylePlanner.ApplyTo(_excelDocument);
            var styleIndexes = new uint?[numberFormats.Length];
            for (int i = 0; i < numberFormats.Length; i++) {
                if (stylePlanner.TryGetCellFormatIndex(numberFormats[i], out uint styleIndex)) {
                    styleIndexes[i] = styleIndex;
                }
            }

            uint? objectDateTimeStyleIndex = null;
            uint? objectTimeSpanStyleIndex = null;
            if (hasObjectColumn) {
                if (stylePlanner.TryGetCellFormatIndex(DataTableDateTimeNumberFormat, out uint dateTimeStyleIndex)) {
                    objectDateTimeStyleIndex = dateTimeStyleIndex;
                }

                if (stylePlanner.TryGetCellFormatIndex(DataTableTimeSpanNumberFormat, out uint timeSpanStyleIndex)) {
                    objectTimeSpanStyleIndex = timeSpanStyleIndex;
                }
            }

            int cellCount = (table.Rows.Count + (includeHeaders ? 1 : 0)) * columnCount;
            bool useDirectStringCells = cellCount >= 4096 && columnCount > 1;
            Dictionary<string, int>? sharedStringIndexes = null;
            var appendedRows = new List<OpenXmlElement>(Math.Max(1, table.Rows.Count + (includeHeaders ? 1 : 0)));
            int rowIndex = startRow;

            if (includeHeaders) {
                appendedRows.Add(CreateDataTableHeaderRow(rowIndex++, startColumn, columnNames, table, useDirectStringCells, ref sharedStringIndexes, ct));
            }

            foreach (DataRow dataRow in table.Rows) {
                ct.ThrowIfCancellationRequested();
                appendedRows.Add(CreateDataTableValueRow(rowIndex++, startColumn, columnNames, dataRow, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, ct));
            }

            sheetData.Append(appendedRows);
            ClearHeaderCacheForPreparedAppend();
            int lastRow = startRow + table.Rows.Count + (includeHeaders ? 1 : 0) - 1;
            int lastColumn = startColumn + columnCount - 1;
            int dimensionMinRow = minExistingRow == int.MaxValue ? startRow : Math.Min(minExistingRow, startRow);
            int dimensionMinColumn = minExistingColumn == int.MaxValue ? startColumn : Math.Min(minExistingColumn, startColumn);
            int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
            int dimensionMaxColumn = Math.Max(maxExistingColumn, lastColumn);
            SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
            _requiresSavePreparation = false;
            return true;
        }

        private Row CreateDataTableHeaderRow(
            int rowIndex,
            int startColumn,
            IReadOnlyList<string> columnNames,
            DataTable table,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            CancellationToken ct) {
            string rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
            var cells = new List<OpenXmlElement>(table.Columns.Count);
            for (int offset = 0; offset < table.Columns.Count; offset++) {
                ct.ThrowIfCancellationRequested();
                int column = startColumn + offset;
                var (cellValue, cellType) = CoerceDataTableAppendValue(table.Columns[offset].ColumnName, useDirectStringCells, ref sharedStringIndexes);
                var cell = new Cell {
                    CellReference = columnNames[column] + rowReference,
                    CellValue = cellValue,
                    DataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType)
                };

                cells.Add(cell);
            }

            var row = new Row { RowIndex = (uint)rowIndex };
            row.Append(cells);
            return row;
        }

        private Row CreateDataTableValueRow(
            int rowIndex,
            int startColumn,
            IReadOnlyList<string> columnNames,
            DataRow dataRow,
            IReadOnlyList<uint?> styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            CancellationToken ct) {
            string rowReference = rowIndex.ToString(CultureInfo.InvariantCulture);
            int columnCount = dataRow.Table.Columns.Count;
            var cells = new List<OpenXmlElement>(columnCount);
            for (int offset = 0; offset < columnCount; offset++) {
                ct.ThrowIfCancellationRequested();
                object? value = dataRow.IsNull(offset) ? null : dataRow[offset];
                int column = startColumn + offset;
                var (cellValue, cellType) = CoerceDataTableAppendValue(value, useDirectStringCells, ref sharedStringIndexes);
                var cell = new Cell {
                    CellReference = columnNames[column] + rowReference,
                    CellValue = cellValue,
                    DataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType)
                };

                if (offset < styleIndexes.Count && styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                cells.Add(cell);
            }

            var row = new Row { RowIndex = (uint)rowIndex };
            row.Append(cells);
            return row;
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceDataTableAppendValue(
            object? value,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            var indexes = sharedStringIndexes;
            CellValue HandleString(string text) {
                return useDirectStringCells
                    ? CreatePlainAppendStringValue(text)
                    : CreatePlainAppendSharedStringValue(text, ref indexes);
            }

            var (cellValue, cellType) = CoerceValueHelper.Coerce(
                value,
                HandleString,
                _excelDocument.DateTimeOffsetWriteStrategy);
            sharedStringIndexes = indexes;

            if (useDirectStringCells && cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                cellType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
            }

            return (cellValue, cellType);
        }

        private static string?[] BuildDataTableNumberFormats(DataTable table) {
            var formats = new string?[table.Columns.Count];
            for (int i = 0; i < table.Columns.Count; i++) {
                formats[i] = GetDataTableNumberFormat(table.Columns[i].DataType, value: null);
            }

            return formats;
        }

        private static string? GetDataTableNumberFormat(Type type, object? value) {
            if (type == typeof(DateTime) || type == typeof(DateTimeOffset) || value is DateTime || value is DateTimeOffset) {
                return DataTableDateTimeNumberFormat;
            }

            if (type == typeof(TimeSpan) || value is TimeSpan) {
                return DataTableTimeSpanNumberFormat;
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly) || value is DateOnly) {
                return DataTableDateTimeNumberFormat;
            }

            if (type == typeof(TimeOnly) || value is TimeOnly) {
                return DataTableTimeSpanNumberFormat;
            }
#endif

            return null;
        }

        private static bool TryGetObjectDataTableValueStyleIndex(object? value, uint? dateTimeStyleIndex, uint? timeSpanStyleIndex, out uint styleIndex) {
            styleIndex = 0U;
            if (value is DateTime || value is DateTimeOffset
#if NET6_0_OR_GREATER
                || value is DateOnly
#endif
                ) {
                if (dateTimeStyleIndex.HasValue) {
                    styleIndex = dateTimeStyleIndex.Value;
                    return true;
                }

                return false;
            }

            if (value is TimeSpan
#if NET6_0_OR_GREATER
                || value is TimeOnly
#endif
                ) {
                if (timeSpanStyleIndex.HasValue) {
                    styleIndex = timeSpanStyleIndex.Value;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Inserts a DataTable and immediately creates an Excel Table over the written range.
        /// Returns the A1-style range of the created table.
        /// </summary>
        public string InsertDataTableAsTable(
            DataTable table,
            int startRow = 1,
            int startColumn = 1,
            bool includeHeaders = true,
            string? tableName = null,
            TableStyle style = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = true,
            ExecutionMode? mode = null,
            CancellationToken ct = default) {
            if (table == null) throw new ArgumentNullException(nameof(table));

            bool canRegisterDirectSave = !_excelDocument.IsMaterializingDeferredDataSetImport
                && mode != ExecutionMode.Parallel
                && CanRegisterDirectTabularSaveCandidate(startRow, startColumn, table.Columns.Count);

            int rowsCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (table.Columns.Count == 0 || rowsCount == 0) {
                return string.Empty;
            }

            int colsCount = table.Columns.Count;
            string startRef = A1.CellReference(startRow, startColumn);
            string endRef = A1.CellReference(startRow + rowsCount - 1, startColumn + colsCount - 1);
            string range = startRef + ":" + endRef;

            InsertDataTableCore(
                table,
                startRow,
                startColumn,
                includeHeaders,
                mode,
                ct,
                copyDirectSaveTable: true,
                registerDirectSaveCandidate: false);

            string[]? headerNames = includeHeaders
                ? table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray()
                : null;

            // Create the Table with optional AutoFilter and style
            string actualTableName = AddTableAndGetName(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter, ensureRangeCellsExist: false, headerNames: headerNames, deferPartSave: canRegisterDirectSave, skipExistingTableScan: canRegisterDirectSave);
            if (canRegisterDirectSave) {
                DataTable directSaveTable = includeHeaders
                    ? table
                    : CreateHeaderlessDirectSaveTable(table);
                _excelDocument.RegisterDirectTabularSaveCandidate(
                    this,
                    directSaveTable,
                    includeHeaders,
                    range,
                    actualTableName,
                    createTable: true,
                    style,
                    includeAutoFilter,
                    autoFit: false,
                    copyTable: includeHeaders);
            }

            return range;
        }

        private static DataTable CreateHeaderlessDirectSaveTable(DataTable source) {
            var table = new DataTable(source.TableName) {
                Locale = CultureInfo.InvariantCulture
            };

            for (int i = 0; i < source.Columns.Count; i++) {
                table.Columns.Add("Column" + (i + 1).ToString(CultureInfo.InvariantCulture), source.Columns[i].DataType);
            }

            table.BeginLoadData();
            try {
                foreach (DataRow sourceRow in source.Rows) {
                    var row = table.NewRow();
                    for (int i = 0; i < source.Columns.Count; i++) {
                        row[i] = sourceRow.IsNull(i) ? DBNull.Value : sourceRow[i];
                    }

                    table.Rows.Add(row);
                }
            } finally {
                table.EndLoadData();
            }

            return table;
        }

        /// <summary>
        /// Appends rows from a <see cref="DataTable"/> to an existing Excel table and expands the table range.
        /// </summary>
        /// <param name="dataTable">Source DataTable containing rows to append.</param>
        /// <param name="tableName">Existing table name or display name.</param>
        /// <param name="matchColumnsByHeader">When true, DataTable columns are matched to table columns by header text. When false, columns are appended by position.</param>
        /// <param name="mode">Optional execution mode override.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>The updated A1 range of the table.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="dataTable"/> or <paramref name="tableName"/> is null.</exception>
        /// <exception cref="ArgumentException">Thrown when the source columns cannot be mapped to the existing table.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the table cannot be found or cannot be safely expanded.</exception>
        public string AppendDataTableToTable(
            DataTable dataTable,
            string tableName,
            bool matchColumnsByHeader = true,
            ExecutionMode? mode = null,
            CancellationToken ct = default) {
            if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));
            if (tableName == null) throw new ArgumentNullException(nameof(tableName));
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentException("Table name cannot be empty.", nameof(tableName));

            var tableDefinitionPart = FindTableDefinitionPart(tableName);
            var table = tableDefinitionPart?.Table;
            if (table == null) {
                throw new InvalidOperationException($"Table '{tableName}' was not found on worksheet '{Name}'.");
            }

            string? currentRange = table.Reference?.Value;
            if (string.IsNullOrWhiteSpace(currentRange) || !A1.TryParseRange(currentRange!, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
                throw new InvalidOperationException($"Table '{tableName}' does not have a valid range.");
            }

            if (HasActiveTotalsRow(table)) {
                throw new InvalidOperationException($"Table '{tableName}' has a totals row. Appending before totals rows is not supported yet.");
            }

            var tableColumnNames = table.TableColumns?.Elements<TableColumn>()
                .Select(column => column.Name?.Value ?? string.Empty)
                .ToList() ?? new List<string>();
            int tableColumnCount = endColumn - startColumn + 1;
            if (tableColumnNames.Count != tableColumnCount) {
                throw new InvalidOperationException($"Table '{tableName}' column metadata does not match its range.");
            }

            bool hasHeaderRow = (table.HeaderRowCount?.Value ?? 1U) > 0U;
            bool useHeaderMapping = matchColumnsByHeader && ShouldMapAppendColumnsByHeader(dataTable, tableColumnNames, hasHeaderRow);
            DataTable appendTable = BuildAppendDataTable(dataTable, tableColumnNames, useHeaderMapping);
            if (appendTable.Rows.Count == 0) {
                return currentRange!;
            }

            int appendStartRow = endRow + 1;
            int appendEndRow = endRow + appendTable.Rows.Count;
            if (appendEndRow > A1.MaxRows) {
                throw new InvalidOperationException($"Appending {appendTable.Rows.Count} rows would exceed the Excel row limit.");
            }

            EnsureAppendTargetIsEmpty(appendStartRow, appendEndRow, startColumn, endColumn, tableName);

            InsertDataTable(appendTable, appendStartRow, startColumn, includeHeaders: false, mode, ct);

            string updatedRange = A1.CellReference(startRow, startColumn) + ":" + A1.CellReference(appendEndRow, endColumn);
            WriteLock(() => {
                table.Reference = updatedRange;
                var autoFilter = table.GetFirstChild<AutoFilter>();
                if (autoFilter != null) {
                    autoFilter.Reference = updatedRange;
                }

                table.Save();
                WorksheetRoot.Save();
            });

            return updatedRange;
        }

        private TableDefinitionPart? FindTableDefinitionPart(string tableName) {
            return _worksheetPart.TableDefinitionParts
                .FirstOrDefault(part => {
                    var table = part.Table;
                    if (table == null) {
                        return false;
                    }

                    string? name = table.Name?.Value;
                    string? displayName = table.DisplayName?.Value;
                    return string.Equals(name, tableName, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(displayName, tableName, StringComparison.OrdinalIgnoreCase);
                });
        }

        private static DataTable BuildAppendDataTable(DataTable source, IReadOnlyList<string> tableColumnNames, bool matchColumnsByHeader) {
            if (source.Columns.Count != tableColumnNames.Count) {
                throw new ArgumentException($"Source table has {source.Columns.Count} columns, but the Excel table has {tableColumnNames.Count} columns.", nameof(source));
            }

            if (!matchColumnsByHeader) {
                return source;
            }

            var sourceColumns = new Dictionary<string, DataColumn>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in source.Columns) {
                if (sourceColumns.ContainsKey(column.ColumnName)) {
                    throw new ArgumentException($"Source table contains duplicate column '{column.ColumnName}'.", nameof(source));
                }

                sourceColumns.Add(column.ColumnName, column);
            }

            var orderedColumns = new DataColumn[tableColumnNames.Count];
            var matchedSourceNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < tableColumnNames.Count; i++) {
                string tableColumnName = tableColumnNames[i];
                if (!sourceColumns.TryGetValue(tableColumnName, out DataColumn? sourceColumn)) {
                    throw new ArgumentException($"Source table is missing column '{tableColumnName}'.", nameof(source));
                }

                orderedColumns[i] = sourceColumn;
                matchedSourceNames.Add(sourceColumn.ColumnName);
            }

            foreach (DataColumn column in source.Columns) {
                if (!matchedSourceNames.Contains(column.ColumnName)) {
                    throw new ArgumentException($"Source table column '{column.ColumnName}' does not exist in the Excel table.", nameof(source));
                }
            }

            var ordered = new DataTable(source.TableName);
            foreach (DataColumn sourceColumn in orderedColumns) {
                ordered.Columns.Add(sourceColumn.ColumnName, sourceColumn.DataType);
            }

            foreach (DataRow sourceRow in source.Rows) {
                DataRow row = ordered.NewRow();
                for (int i = 0; i < orderedColumns.Length; i++) {
                    row[i] = sourceRow[orderedColumns[i]];
                }

                ordered.Rows.Add(row);
            }

            return ordered;
        }

        private static bool ShouldMapAppendColumnsByHeader(DataTable source, IReadOnlyList<string> tableColumnNames, bool hasHeaderRow) {
            if (hasHeaderRow) {
                return true;
            }

            if (SourceContainsTableColumns(source, tableColumnNames)) {
                return true;
            }

            return !HasDefaultHeaderlessColumnNames(tableColumnNames);
        }

        private static bool SourceContainsTableColumns(DataTable source, IReadOnlyList<string> tableColumnNames) {
            var sourceColumnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in source.Columns) {
                sourceColumnNames.Add(column.ColumnName);
            }

            foreach (string tableColumnName in tableColumnNames) {
                if (!sourceColumnNames.Contains(tableColumnName)) {
                    return false;
                }
            }

            return true;
        }

        private static bool HasDefaultHeaderlessColumnNames(IReadOnlyList<string> tableColumnNames) {
            for (int i = 0; i < tableColumnNames.Count; i++) {
                if (!string.Equals(tableColumnNames[i], "Column" + (i + 1), StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
            }

            return tableColumnNames.Count > 0;
        }

        private static bool HasActiveTotalsRow(Table table) {
            uint? totalsRowCount = table.TotalsRowCount?.Value;
            if (totalsRowCount.HasValue) {
                return totalsRowCount.Value > 0U;
            }

            return table.TotalsRowShown?.Value == true;
        }

        private void EnsureAppendTargetIsEmpty(int startRow, int endRow, int startColumn, int endColumn, string tableName) {
            if (startRow > endRow) {
                return;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return;
            }

            foreach (Row rowElement in sheetData.Elements<Row>()) {
                if (rowElement.RowIndex == null) {
                    continue;
                }

                int rowIndex = (int)rowElement.RowIndex.Value;
                if (rowIndex < startRow) {
                    continue;
                }

                if (rowIndex > endRow) {
                    break;
                }

                foreach (Cell cell in rowElement.Elements<Cell>()) {
                    string? reference = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(reference)) {
                        continue;
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReference(reference!);
                    if (columnIndex < startColumn || columnIndex > endColumn) {
                        continue;
                    }

                    if (CellHasContent(cell)) {
                        throw new InvalidOperationException($"Cannot append to table '{tableName}' because cell {reference} already contains data.");
                    }
                }
            }
        }

        private static bool CellHasContent(Cell cell) {
            if (cell.CellFormula != null) {
                return true;
            }

            if (cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text)) {
                return true;
            }

            if (cell.InlineString != null) {
                return true;
            }

            return false;
        }
    }
}
