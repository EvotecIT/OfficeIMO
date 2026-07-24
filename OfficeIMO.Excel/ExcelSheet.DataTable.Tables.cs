using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal string InsertTabularRowSourceAsTableForDeferredMaterialization(
            IExcelSheetTabularRowSource source,
            int startRow = 1,
            int startColumn = 1,
            bool includeHeaders = true,
            string? tableName = null,
            TableStyle style = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = true,
            CancellationToken ct = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));

            int rowsCount = source.RowCount + (includeHeaders ? 1 : 0);
            int columnCount = source.ColumnCount;
            if (columnCount == 0 || rowsCount == 0) {
                return string.Empty;
            }

            string startRef = A1.CellReference(startRow, startColumn);
            string endRef = A1.CellReference(startRow + rowsCount - 1, startColumn + columnCount - 1);
            string range = startRef + ":" + endRef;

            if (!TryInsertTabularRowSourceForDeferredMaterialization(source, startRow, startColumn, includeHeaders, ct)) {
                return string.Empty;
            }

            string[]? headerNames = null;
            if (includeHeaders) {
                headerNames = new string[columnCount];
                for (int i = 0; i < headerNames.Length; i++) {
                    headerNames[i] = source.GetColumnName(i);
                }
            }

            AddTableAndGetName(
                range,
                includeHeaders,
                tableName ?? string.Empty,
                style,
                includeAutoFilter,
                ensureRangeCellsExist: false,
                headerNames: headerNames,
                deferPartSave: true,
                skipExistingTableScan: true);

            return range;
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

            if (canRegisterDirectSave
                && TryInsertDataTableAsDeferredDirectSave(
                    table,
                    startRow,
                    startColumn,
                    includeHeaders,
                    copyDirectSaveTable: true,
                    createTable: true,
                    tableName,
                    style,
                    includeAutoFilter,
                    ct)) {
                return range;
            }

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

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                _excelDocument.MaterializeDeferredDataSetImport();
            }

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

            int tableColumnCount = endColumn - startColumn + 1;
            var tableColumnNames = GetTableColumnNames(table.TableColumns, tableColumnCount);
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

            var sourceColumns = new Dictionary<string, int>(source.Columns.Count, StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in source.Columns) {
                if (sourceColumns.ContainsKey(column.ColumnName)) {
                    throw new ArgumentException($"Source table contains duplicate column '{column.ColumnName}'.", nameof(source));
                }

                sourceColumns.Add(column.ColumnName, column.Ordinal);
            }

            var orderedColumnIndexes = new int[tableColumnNames.Count];
            var matchedSourceColumns = new bool[source.Columns.Count];
            for (int i = 0; i < tableColumnNames.Count; i++) {
                string tableColumnName = tableColumnNames[i];
                if (!sourceColumns.TryGetValue(tableColumnName, out int sourceColumnIndex)) {
                    throw new ArgumentException($"Source table is missing column '{tableColumnName}'.", nameof(source));
                }

                orderedColumnIndexes[i] = sourceColumnIndex;
                matchedSourceColumns[sourceColumnIndex] = true;
            }

            foreach (DataColumn column in source.Columns) {
                if (!matchedSourceColumns[column.Ordinal]) {
                    throw new ArgumentException($"Source table column '{column.ColumnName}' does not exist in the Excel table.", nameof(source));
                }
            }

            var ordered = new DataTable(source.TableName);
            for (int i = 0; i < orderedColumnIndexes.Length; i++) {
                DataColumn sourceColumn = source.Columns[orderedColumnIndexes[i]];
                ordered.Columns.Add(sourceColumn.ColumnName, sourceColumn.DataType);
            }

            foreach (DataRow sourceRow in source.Rows) {
                DataRow row = ordered.NewRow();
                for (int i = 0; i < orderedColumnIndexes.Length; i++) {
                    row[i] = sourceRow[orderedColumnIndexes[i]];
                }

                ordered.Rows.Add(row);
            }

            return ordered;
        }

        private static List<string> GetTableColumnNames(TableColumns? tableColumns, int capacity) {
            if (tableColumns == null) {
                return new List<string>();
            }

            var names = new List<string>(Math.Max(0, capacity));
            foreach (TableColumn column in tableColumns.Elements<TableColumn>()) {
                names.Add(column.Name?.Value ?? string.Empty);
            }

            return names;
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
            foreach (Row rowElement in sheetData?.Elements<Row>() ?? Enumerable.Empty<Row>()) {
                int ordinalColumnIndex = 0;
                foreach (Cell cell in rowElement.Elements<Cell>()) {
                    ordinalColumnIndex++;
                    string? reference = cell.CellReference?.Value;
                    int rowIndex;
                    int columnIndex;
                    if (!string.IsNullOrEmpty(reference)) {
                        (rowIndex, columnIndex) = A1.ParseCellRef(reference!);
                    } else {
                        uint rawRowIndex = rowElement.RowIndex?.Value ?? 0U;
                        rowIndex = rawRowIndex <= A1.MaxRows ? (int)rawRowIndex : 0;
                        columnIndex = ordinalColumnIndex;
                    }

                    if (rowIndex < startRow || rowIndex > endRow
                        || columnIndex < startColumn || columnIndex > endColumn) {
                        continue;
                    }

                    string location = string.IsNullOrEmpty(reference)
                        ? A1.ColumnIndexToLetters(columnIndex) + rowIndex.ToString(CultureInfo.InvariantCulture)
                        : reference!;
                    throw new InvalidOperationException($"Cannot append to table '{tableName}' because cell {location} already contains worksheet data or formatting.");
                }
            }

            EnsureAppendTargetHasNoRangeMetadata(startRow, endRow, startColumn, endColumn, tableName);
        }

        private void EnsureAppendTargetHasNoRangeMetadata(int startRow, int endRow, int startColumn, int endColumn, string tableName) {
            foreach (Hyperlink hyperlink in WorksheetRoot.Descendants<Hyperlink>()) {
                if (AppendTargetIntersectsReference(hyperlink.Reference?.Value, startRow, endRow, startColumn, endColumn)) {
                    throw new InvalidOperationException($"Cannot append to table '{tableName}' because the target range contains a hyperlink.");
                }
            }

            foreach (MergeCell mergeCell in WorksheetRoot.Descendants<MergeCell>()) {
                if (AppendTargetIntersectsReference(mergeCell.Reference?.Value, startRow, endRow, startColumn, endColumn)) {
                    throw new InvalidOperationException($"Cannot append to table '{tableName}' because the target range intersects a merged cell.");
                }
            }

            foreach (DataValidation validation in WorksheetRoot.Descendants<DataValidation>()) {
                if (AppendTargetIntersectsReferences(validation.SequenceOfReferences?.InnerText, startRow, endRow, startColumn, endColumn)) {
                    throw new InvalidOperationException($"Cannot append to table '{tableName}' because the target range contains data validation.");
                }
            }

            foreach (ConditionalFormatting formatting in WorksheetRoot.Descendants<ConditionalFormatting>()) {
                if (AppendTargetIntersectsReferences(formatting.SequenceOfReferences?.InnerText, startRow, endRow, startColumn, endColumn)) {
                    throw new InvalidOperationException($"Cannot append to table '{tableName}' because the target range contains conditional formatting.");
                }
            }

            var comments = _worksheetPart.WorksheetCommentsPart?.Comments?.CommentList;
            if (comments != null) {
                foreach (Comment comment in comments.Elements<Comment>()) {
                    if (AppendTargetIntersectsReference(comment.Reference?.Value, startRow, endRow, startColumn, endColumn)) {
                        throw new InvalidOperationException($"Cannot append to table '{tableName}' because the target range contains a comment.");
                    }
                }
            }
        }

        private static bool AppendTargetIntersectsReferences(string? references, int startRow, int endRow, int startColumn, int endColumn) {
            if (string.IsNullOrWhiteSpace(references)) {
                return false;
            }

            string[] tokens = references!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            for (int index = 0; index < tokens.Length; index++) {
                if (AppendTargetIntersectsReference(tokens[index], startRow, endRow, startColumn, endColumn)) {
                    return true;
                }
            }

            return false;
        }

        private static bool AppendTargetIntersectsReference(string? reference, int startRow, int endRow, int startColumn, int endColumn) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            string normalized = reference!.Replace("$", string.Empty).Trim();
            int rangeStartRow;
            int rangeStartColumn;
            int rangeEndRow;
            int rangeEndColumn;
            if (normalized.IndexOf(':') >= 0) {
                if (!A1.TryParseRange(normalized, out rangeStartRow, out rangeStartColumn, out rangeEndRow, out rangeEndColumn)) {
                    return false;
                }
            } else {
                (rangeStartRow, rangeStartColumn) = A1.ParseCellRef(normalized);
                rangeEndRow = rangeStartRow;
                rangeEndColumn = rangeStartColumn;
            }

            return rangeStartRow > 0
                && rangeStartColumn > 0
                && rangeStartRow <= endRow
                && rangeEndRow >= startRow
                && rangeStartColumn <= endColumn
                && rangeEndColumn >= startColumn;
        }

    }
}
