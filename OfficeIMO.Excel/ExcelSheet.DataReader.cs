using System.Data;
using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int DirectDataReaderSaveCandidateRowLimit = 65536;

        /// <summary>
        /// Streams rows from an <see cref="IDataReader"/> (including provider-owned DbDataReader implementations) into the worksheet and optionally creates an Excel table.
        /// The caller owns the connection, command, query, and provider.
        /// </summary>
        /// <param name="reader">Open data reader positioned before the first row.</param>
        /// <param name="startRow">1-based start row.</param>
        /// <param name="startColumn">1-based start column.</param>
        /// <param name="includeHeaders">Write field names as the first row.</param>
        /// <param name="tableName">Optional Excel table name.</param>
        /// <param name="style">Excel table style to use when <paramref name="createTable"/> is true.</param>
        /// <param name="includeAutoFilter">Include table AutoFilter dropdowns when creating a table.</param>
        /// <param name="createTable">Create an Excel table over the imported range.</param>
        /// <param name="autoFit">Auto-fit imported columns after rows are written.</param>
        /// <param name="ct">Cancellation token.</param>
        /// <returns>A1 range occupied by the imported reader data.</returns>
        public string InsertDataReader(
            IDataReader reader,
            int startRow = 1,
            int startColumn = 1,
            bool includeHeaders = true,
            string? tableName = null,
            TableStyle style = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = true,
            bool createTable = true,
            bool autoFit = false,
            CancellationToken ct = default) {
            if (reader == null) throw new ArgumentNullException(nameof(reader));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));
            if (reader.FieldCount < 1) throw new ArgumentException("Data reader must expose at least one field.", nameof(reader));

            string[] headers = BuildReaderHeaders(reader);
            Type[] fieldTypes = BuildReaderFieldTypes(reader);
            bool canRegisterDirectSave = !_excelDocument.IsMaterializingDeferredDataSetImport
                && CanRegisterDirectTabularSaveCandidate(startRow, startColumn, headers.Length);

            if (canRegisterDirectSave
                && TryInsertDataReaderAsDeferredDirectSave(
                    reader,
                    headers,
                    fieldTypes,
                    startRow,
                    startColumn,
                    includeHeaders,
                    tableName,
                    style,
                    includeAutoFilter,
                    createTable,
                    autoFit,
                    ct,
                    out string deferredRange)) {
                return deferredRange;
            }

            _excelDocument.MaterializeDeferredDataSetImport();

            if (TryInsertDataReaderByAppendingRows(
                    reader,
                    startRow,
                    startColumn,
                    headers,
                    fieldTypes,
                    includeHeaders,
                    canRegisterDirectSave,
                    ct,
                    out int appendedDataRows,
                    out List<object?[]>? appendedDirectRows)) {
                int appendedOccupiedRows = appendedDataRows + (includeHeaders ? 1 : 0);
                if (appendedOccupiedRows == 0) {
                    return string.Empty;
                }

                string appendedRange = A1.CellReference(startRow, startColumn) + ":" +
                    A1.CellReference(startRow + appendedOccupiedRows - 1, startColumn + headers.Length - 1);

                string? appendedTableName = null;
                if (createTable) {
                    string[]? headerNames = includeHeaders ? headers : null;
                    appendedTableName = AddTableAndGetName(
                        appendedRange,
                        includeHeaders,
                        tableName ?? string.Empty,
                        style,
                        includeAutoFilter,
                        headerNames: headerNames,
                        ensureRangeCellsExist: false,
                        deferPartSave: canRegisterDirectSave,
                        skipExistingTableScan: canRegisterDirectSave);
                }

                if (autoFit) {
                    AutoFitColumnsFor(Enumerable.Range(startColumn, headers.Length));
                }

                RegisterDirectDataReaderSaveCandidateIfPossible(
                    appendedDirectRows,
                    headers,
                    fieldTypes,
                    includeHeaders,
                    appendedRange,
                    appendedTableName,
                    createTable,
                    style,
                    includeAutoFilter,
                    autoFit,
                    canRegisterDirectSave);

                return appendedRange;
            }

            List<object?[]>? directRows = canRegisterDirectSave ? new List<object?[]>() : null;

            int row = startRow;
            if (includeHeaders) {
                for (int i = 0; i < headers.Length; i++) {
                    CellValue(row, startColumn + i, headers[i]);
                }

                row++;
            }

            int dataRows = 0;
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                object?[]? directRow = directRows != null ? new object?[headers.Length] : null;
                for (int i = 0; i < headers.Length; i++) {
                    bool isDbNull = reader.IsDBNull(i);
                    object? value = isDbNull ? null : reader.GetValue(i);
                    if (directRow != null) {
                        directRow[i] = value;
                    }

                    int column = startColumn + i;
                    CellValue(row, column, value);

                    string? numberFormat = GetReaderNumberFormat(fieldTypes[i], value);
                    if (!string.IsNullOrWhiteSpace(numberFormat)) {
                        FormatCell(row, column, numberFormat!);
                    }
                }

                row++;
                dataRows++;
                if (directRows != null) {
                    if (directRows.Count < DirectDataReaderSaveCandidateRowLimit) {
                        directRows.Add(directRow!);
                    } else {
                        directRows = null;
                    }
                }
            }

            int occupiedRows = dataRows + (includeHeaders ? 1 : 0);
            if (occupiedRows == 0) {
                return string.Empty;
            }

            string range = A1.CellReference(startRow, startColumn) + ":" +
                A1.CellReference(startRow + occupiedRows - 1, startColumn + headers.Length - 1);

            string? actualTableName = null;
            if (createTable) {
                string[]? headerNames = includeHeaders ? headers : null;
                actualTableName = AddTableAndGetName(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter, headerNames: headerNames);
            }

            if (autoFit) {
                AutoFitColumnsFor(Enumerable.Range(startColumn, headers.Length));
            }

            RegisterDirectDataReaderSaveCandidateIfPossible(
                directRows,
                headers,
                fieldTypes,
                includeHeaders,
                range,
                actualTableName,
                createTable,
                style,
                includeAutoFilter,
                autoFit,
                canRegisterDirectSave);

            return range;
        }

        private bool TryInsertDataReaderAsDeferredDirectSave(
            IDataReader reader,
            IReadOnlyList<string> headers,
            IReadOnlyList<Type> fieldTypes,
            int startRow,
            int startColumn,
            bool includeHeaders,
            string? tableName,
            TableStyle style,
            bool includeAutoFilter,
            bool createTable,
            bool autoFit,
            CancellationToken ct,
            out string range) {
            range = string.Empty;
            int columnCount = headers.Count;
            var rows = new List<object?[]>();
            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (startRow + rows.Count + (includeHeaders ? 1 : 0) > A1.MaxRows) {
                    throw new InvalidOperationException("Data reader import exceeds the maximum worksheet row count.");
                }

                var values = new object?[columnCount];
                for (int offset = 0; offset < columnCount; offset++) {
                    values[offset] = reader.IsDBNull(offset) ? null : reader.GetValue(offset);
                }

                rows.Add(values);
            }

            int occupiedRows = rows.Count + (includeHeaders ? 1 : 0);
            if (occupiedRows == 0) {
                return true;
            }

            range = A1.CellReference(startRow, startColumn) + ":" +
                A1.CellReference(startRow + occupiedRows - 1, startColumn + columnCount - 1);

            string[] columnNames = BuildDirectReaderColumnNames(headers, includeHeaders);
            Type[] columnTypes = BuildDirectReaderColumnTypes(fieldTypes);
            if (_excelDocument.RegisterDeferredDirectTabularSaveCandidate(
                this,
                "ReaderData",
                columnNames,
                columnTypes,
                rows.ToArray(),
                includeHeaders,
                range,
                createTable ? tableName : null,
                createTable,
                style,
                includeAutoFilter,
                autoFit)) {
                return true;
            }

            _excelDocument.MaterializeDeferredDataSetImport();
            InsertBufferedDataReaderRows(
                headers,
                fieldTypes,
                rows,
                startRow,
                startColumn,
                includeHeaders,
                tableName,
                style,
                includeAutoFilter,
                createTable,
                autoFit,
                range);
            return true;
        }

        private void InsertBufferedDataReaderRows(
            IReadOnlyList<string> headers,
            IReadOnlyList<Type> fieldTypes,
            IReadOnlyList<object?[]> rows,
            int startRow,
            int startColumn,
            bool includeHeaders,
            string? tableName,
            TableStyle style,
            bool includeAutoFilter,
            bool createTable,
            bool autoFit,
            string range) {
            int row = startRow;
            if (includeHeaders) {
                for (int i = 0; i < headers.Count; i++) {
                    CellValue(row, startColumn + i, headers[i]);
                }

                row++;
            }

            foreach (object?[] values in rows) {
                for (int i = 0; i < headers.Count; i++) {
                    object? value = values[i];
                    int column = startColumn + i;
                    CellValue(row, column, value);

                    string? numberFormat = GetReaderNumberFormat(fieldTypes[i], value);
                    if (!string.IsNullOrWhiteSpace(numberFormat)) {
                        FormatCell(row, column, numberFormat!);
                    }
                }

                row++;
            }

            if (createTable && range.Length != 0) {
                string[]? headerNames = includeHeaders ? headers.ToArray() : null;
                AddTableAndGetName(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter, headerNames: headerNames);
            }

            if (autoFit) {
                AutoFitColumnsFor(Enumerable.Range(startColumn, headers.Count));
            }
        }

        private bool TryInsertDataReaderByAppendingRows(
            IDataReader reader,
            int startRow,
            int startColumn,
            IReadOnlyList<string> headers,
            IReadOnlyList<Type> fieldTypes,
            bool includeHeaders,
            bool collectDirectRows,
            CancellationToken ct,
            out int dataRows,
            out List<object?[]>? directRows) {
            dataRows = 0;
            directRows = collectDirectRows ? new List<object?[]>() : null;

            int columnCount = headers.Count;
            if (columnCount == 0 || startColumn + columnCount - 1 > A1.MaxColumns) {
                return false;
            }

            bool applied = false;
            int capturedDataRows = 0;
            List<object?[]>? capturedDirectRows = directRows;
            System.Threading.ReaderWriterLockSlim? lck = _excelDocument._lock;
            if (lck == null) {
                try { lck = _excelDocument.EnsureLock(); } catch { lck = null; }
            }

            Locking.ExecuteWrite(lck, () => applied = TryInsertDataReaderByAppendingRowsCore(
                reader,
                startRow,
                startColumn,
                headers,
                fieldTypes,
                includeHeaders,
                collectDirectRows,
                ct,
                out capturedDataRows,
                out capturedDirectRows));

            dataRows = capturedDataRows;
            directRows = capturedDirectRows;
            return applied;
        }

        private bool TryInsertDataReaderByAppendingRowsCore(
            IDataReader reader,
            int startRow,
            int startColumn,
            IReadOnlyList<string> headers,
            IReadOnlyList<Type> fieldTypes,
            bool includeHeaders,
            bool collectDirectRows,
            CancellationToken ct,
            out int dataRows,
            out List<object?[]>? directRows) {
            dataRows = 0;
            directRows = collectDirectRows ? new List<object?[]>() : null;

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

            int columnCount = headers.Count;
            var columnNames = new string[startColumn + columnCount];
            for (int column = startColumn; column < startColumn + columnCount; column++) {
                columnNames[column] = GetColumnName(column);
            }

            string?[] numberFormats = BuildReaderNumberFormats(fieldTypes);
            var stylePlanner = new StylePlanner();
            bool hasObjectColumn = false;
            foreach (Type fieldType in fieldTypes) {
                if ((Nullable.GetUnderlyingType(fieldType) ?? fieldType) == typeof(object)) {
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

            Dictionary<string, int>? sharedStringIndexes = null;
            bool useDirectStringCells = collectDirectRows && columnCount > 1;
            var appendedRows = new List<OpenXmlElement>();
            int rowIndex = startRow;

            if (includeHeaders) {
                appendedRows.Add(CreateDataReaderHeaderRow(rowIndex++, startColumn, columnNames, headers, useDirectStringCells, ref sharedStringIndexes, ct));
            }

            while (reader.Read()) {
                ct.ThrowIfCancellationRequested();
                if (rowIndex > A1.MaxRows) {
                    throw new InvalidOperationException("Data reader import exceeds the maximum worksheet row count.");
                }

                object?[]? directRow = directRows != null ? new object?[columnCount] : null;
                object?[] values = directRow ?? new object?[columnCount];
                for (int offset = 0; offset < columnCount; offset++) {
                    bool isDbNull = reader.IsDBNull(offset);
                    object? value = isDbNull ? null : reader.GetValue(offset);
                    values[offset] = value;
                }

                appendedRows.Add(CreateDataReaderValueRow(rowIndex++, startColumn, columnNames, values, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, ct));
                dataRows++;

                if (directRows != null) {
                    if (directRows.Count < DirectDataReaderSaveCandidateRowLimit) {
                        directRows.Add(directRow!);
                    } else {
                        directRows = null;
                    }
                }
            }

            if (appendedRows.Count > 0) {
                sheetData.Append(appendedRows);
                ClearHeaderCacheForPreparedAppend();
                int lastRow = startRow + appendedRows.Count - 1;
                int lastColumn = startColumn + columnCount - 1;
                int dimensionMinRow = minExistingRow == int.MaxValue ? startRow : Math.Min(minExistingRow, startRow);
                int dimensionMinColumn = minExistingColumn == int.MaxValue ? startColumn : Math.Min(minExistingColumn, startColumn);
                int dimensionMaxRow = Math.Max(maxExistingRow, lastRow);
                int dimensionMaxColumn = Math.Max(maxExistingColumn, lastColumn);
                SetSheetDimensionReference(dimensionMinRow, dimensionMinColumn, dimensionMaxRow, dimensionMaxColumn);
                _requiresSavePreparation = false;
            }

            return true;
        }

        private Row CreateDataReaderHeaderRow(
            int rowIndex,
            int startColumn,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<string> headers,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < headers.Count; offset++) {
                ct.ThrowIfCancellationRequested();
                int column = startColumn + offset;
                var (cellValue, cellType) = CoerceDataTableAppendValue(headers[offset], useDirectStringCells, ref sharedStringIndexes);
                row.Append(new Cell {
                    CellReference = columnNames[column] + rowReference,
                    CellValue = cellValue,
                    DataType = new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType)
                });
            }

            return row;
        }

        private Row CreateDataReaderValueRow(
            int rowIndex,
            int startColumn,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<object?> values,
            IReadOnlyList<uint?> styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < values.Count; offset++) {
                ct.ThrowIfCancellationRequested();
                object? value = values[offset];
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

                row.Append(cell);
            }

            return row;
        }

        private void RegisterDirectDataReaderSaveCandidateIfPossible(
            List<object?[]>? directRows,
            IReadOnlyList<string> headers,
            IReadOnlyList<Type> fieldTypes,
            bool includeHeaders,
            string range,
            string? tableName,
            bool createTable,
            TableStyle style,
            bool includeAutoFilter,
            bool autoFit,
            bool canRegisterDirectSave) {
            if (!canRegisterDirectSave || directRows == null || range.Length == 0) {
                return;
            }

            string[] columnNames = BuildDirectReaderColumnNames(headers, includeHeaders);
            Type[] columnTypes = BuildDirectReaderColumnTypes(fieldTypes);

            _excelDocument.RegisterDirectTabularSaveCandidate(
                this,
                "ReaderData",
                columnNames,
                columnTypes,
                directRows.ToArray(),
                includeHeaders,
                range,
                tableName,
                createTable,
                style,
                includeAutoFilter,
                autoFit);
        }

        private static string[] BuildReaderHeaders(IDataReader reader) {
            var headers = new List<string>(reader.FieldCount);
            for (int i = 0; i < reader.FieldCount; i++) {
                string name;
                try {
                    name = reader.GetName(i);
                } catch (Exception) {
                    name = string.Empty;
                }

                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Column" + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                headers.Add(name);
            }

            EnsureUniqueReaderHeaders(headers);
            return headers.ToArray();
        }

        private static Type[] BuildReaderFieldTypes(IDataReader reader) {
            var types = new Type[reader.FieldCount];
            for (int i = 0; i < reader.FieldCount; i++) {
                try {
                    types[i] = reader.GetFieldType(i) ?? typeof(object);
                } catch (Exception) {
                    types[i] = typeof(object);
                }
            }

            return types;
        }

        private static string[] BuildDirectReaderColumnNames(IReadOnlyList<string> headers, bool includeHeaders) {
            if (includeHeaders) {
                return headers.ToArray();
            }

            return Enumerable.Range(1, headers.Count)
                .Select(index => "Column" + index.ToString(CultureInfo.InvariantCulture))
                .ToArray();
        }

        private static Type[] BuildDirectReaderColumnTypes(IReadOnlyList<Type> fieldTypes) {
            Type[] columnTypes = new Type[fieldTypes.Count];
            for (int i = 0; i < fieldTypes.Count; i++) {
                Type fieldType = fieldTypes[i];
                if (fieldType == typeof(DBNull) || fieldType == typeof(void)) {
                    fieldType = typeof(object);
                }

                columnTypes[i] = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
            }

            return columnTypes;
        }

        private static string?[] BuildReaderNumberFormats(IReadOnlyList<Type> fieldTypes) {
            var formats = new string?[fieldTypes.Count];
            for (int i = 0; i < fieldTypes.Count; i++) {
                formats[i] = GetReaderNumberFormat(fieldTypes[i], value: null);
            }

            return formats;
        }

        private static void EnsureUniqueReaderHeaders(IList<string> headers) {
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++) {
                string baseName = string.IsNullOrWhiteSpace(headers[i])
                    ? "Column" + (i + 1).ToString(CultureInfo.InvariantCulture)
                    : headers[i].Trim();
                if (!seen.TryGetValue(baseName, out int count)) {
                    seen[baseName] = 1;
                    headers[i] = baseName;
                    continue;
                }

                count++;
                seen[baseName] = count;
                headers[i] = baseName + " (" + count.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        private static string? GetReaderNumberFormat(Type fieldType, object? value) {
            Type type = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
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
    }
}
