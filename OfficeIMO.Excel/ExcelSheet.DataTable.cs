using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    internal interface IExcelSheetTabularRowSource {
        int ColumnCount { get; }

        int RowCount { get; }

        string GetColumnName(int index);

        Type GetColumnType(int index);

        object? GetValue(int rowIndex, int columnIndex);

        bool TryGetBufferedRow(int rowIndex, out object?[]? values);

        bool TryGetFlatValues(out object?[] values, out int columnCount);
    }

    public partial class ExcelSheet {
        private const string DataTableDateTimeNumberFormat = "yyyy-mm-dd hh:mm";
        private const string DataTableTimeSpanNumberFormat = "[h]:mm:ss";
        private static readonly EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> DataTableStringCellType = new(DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
        private static readonly EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> DataTableSharedStringCellType = new(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
        private static readonly EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> DataTableNumberCellType = new(DocumentFormat.OpenXml.Spreadsheet.CellValues.Number);
        private static readonly EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> DataTableBooleanCellType = new(DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean);

        private enum TabularAppendColumnKind {
            General,
            String,
            Double,
            Float,
            Decimal,
            SignedInteger,
            UnsignedInteger,
            Boolean,
            DateTime,
            DateTimeOffset,
#if NET6_0_OR_GREATER
            DateOnly,
            TimeOnly,
#endif
            TimeSpan
        }

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

            if (canRegisterDirectSave
                && TryInsertDataTableAsDeferredDirectSave(
                    table,
                    startRow,
                    startColumn,
                    includeHeaders,
                    copyDirectSaveTable,
                    createTable: false,
                    tableName: null,
                    style: TableStyle.TableStyleMedium2,
                    includeAutoFilter: false,
                    ct)) {
                return;
            }

            _excelDocument.MaterializeDeferredDataSetImport();

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
                        type ??= GetCachedDataTableCellType(DocumentFormat.OpenXml.Spreadsheet.CellValues.String);
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

        private bool TryInsertDataTableAsDeferredDirectSave(
            DataTable table,
            int startRow,
            int startColumn,
            bool includeHeaders,
            bool copyDirectSaveTable,
            bool createTable,
            string? tableName,
            TableStyle style,
            bool includeAutoFilter,
            CancellationToken ct) {
            string range = BuildDataTableInsertedRange(table, startRow, startColumn, includeHeaders);
            if (range.Length == 0) {
                return true;
            }

            DataTable directSaveTable = table;
            bool directSaveIncludesHeaders = includeHeaders;
            if (createTable && !includeHeaders) {
                directSaveTable = CreateHeaderlessDirectSaveTable(table);
                directSaveIncludesHeaders = false;
                copyDirectSaveTable = false;
            }

            ct.ThrowIfCancellationRequested();
            return _excelDocument.RegisterDeferredDirectTabularSaveCandidate(
                this,
                directSaveTable,
                directSaveIncludesHeaders,
                range,
                tableName,
                createTable,
                style,
                includeAutoFilter,
                autoFit: false,
                copyTable: copyDirectSaveTable);
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
            string[] columnReferencePrefixes = BuildColumnReferencePrefixes(startColumn, columnCount);

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

            var columnKinds = BuildDataTableAppendColumnKinds(table);
            int cellCount = (table.Rows.Count + (includeHeaders ? 1 : 0)) * columnCount;
            bool useDirectStringCells = false;
            Dictionary<string, int>? sharedStringIndexes = null;
            var appendedRows = new List<OpenXmlElement>(Math.Max(1, table.Rows.Count + (includeHeaders ? 1 : 0)));
            int rowIndex = startRow;
            bool canCancel = ct.CanBeCanceled;

            if (includeHeaders) {
                appendedRows.Add(CreateDataTableHeaderRow(rowIndex++, columnReferencePrefixes, table, useDirectStringCells, ref sharedStringIndexes, canCancel, ct));
            }

            foreach (DataRow dataRow in table.Rows) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                appendedRows.Add(canCancel
                    ? CreateDataTableValueRow(rowIndex++, columnReferencePrefixes, dataRow, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct)
                    : CreateDataTableValueRow(rowIndex++, columnReferencePrefixes, dataRow, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes));
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

        internal bool TryInsertTabularRowSourceForDeferredMaterialization(IExcelSheetTabularRowSource source, int startRow = 1, int startColumn = 1, bool includeHeaders = true, CancellationToken ct = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));

            int columnCount = source.ColumnCount;
            int rowsCount = source.RowCount + (includeHeaders ? 1 : 0);
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

            Locking.ExecuteWrite(lck, () => applied = TryInsertTabularRowSourceByAppendingRowsCore(source, startRow, startColumn, includeHeaders, ct));
            return applied;
        }

        private bool TryInsertTabularRowSourceByAppendingRowsCore(IExcelSheetTabularRowSource source, int startRow, int startColumn, bool includeHeaders, CancellationToken ct) {
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

            int columnCount = source.ColumnCount;
            string[] columnReferencePrefixes = BuildColumnReferencePrefixes(startColumn, columnCount);

            string?[] numberFormats = BuildTabularRowSourceNumberFormats(source);
            var stylePlanner = new StylePlanner();
            bool hasObjectColumn = false;
            for (int i = 0; i < columnCount; i++) {
                if (source.GetColumnType(i) == typeof(object)) {
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

            var columnKinds = BuildTabularAppendColumnKinds(source);
            int rowCount = source.RowCount;
            int cellCount = (rowCount + (includeHeaders ? 1 : 0)) * columnCount;
            bool useDirectStringCells = cellCount >= 4096 && columnCount > 1;
            Dictionary<string, int>? sharedStringIndexes = null;
            var appendedRows = new List<OpenXmlElement>(Math.Max(1, rowCount + (includeHeaders ? 1 : 0)));
            int rowIndex = startRow;
            bool canCancel = ct.CanBeCanceled;
            object?[]? flatValues = null;
            bool useFlatValues = source.TryGetFlatValues(out var sourceFlatValues, out int flatColumnCount) && flatColumnCount == columnCount;
            if (useFlatValues) {
                flatValues = sourceFlatValues;
            }

            if (includeHeaders) {
                appendedRows.Add(CreateTabularRowSourceHeaderRow(rowIndex++, columnReferencePrefixes, source, useDirectStringCells, ref sharedStringIndexes, canCancel, ct));
            }

            for (int sourceRowIndex = 0; sourceRowIndex < rowCount; sourceRowIndex++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (flatValues != null && !canCancel) {
                    appendedRows.Add(CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, flatValues, sourceRowIndex * columnCount, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes));
                } else if (flatValues != null) {
                    appendedRows.Add(CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, flatValues, sourceRowIndex * columnCount, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct));
                } else if (source.TryGetBufferedRow(sourceRowIndex, out var rowValues) && rowValues != null && !canCancel) {
                    appendedRows.Add(CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, rowValues, 0, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes));
                } else if (rowValues != null) {
                    appendedRows.Add(CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, rowValues, 0, columnCount, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct));
                } else {
                    appendedRows.Add(CreateTabularRowSourceValueRow(rowIndex++, columnReferencePrefixes, source, sourceRowIndex, columnKinds, styleIndexes, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, useDirectStringCells, ref sharedStringIndexes, canCancel, ct));
                }
            }

            sheetData.Append(appendedRows);
            ClearHeaderCacheForPreparedAppend();
            int lastRow = startRow + rowCount + (includeHeaders ? 1 : 0) - 1;
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
            string[] columnReferencePrefixes,
            DataTable table,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < table.Columns.Count; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var (cellValue, cellType) = CoerceDataTableAppendValue(table.Columns[offset].ColumnName, useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceHeaderRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            IExcelSheetTabularRowSource source,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < source.ColumnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var (cellValue, cellType) = CoerceDataTableAppendValue(source.GetColumnName(offset), useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                row.Append(cell);
            }

            return row;
        }

        private Row CreateDataTableValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            DataRow dataRow,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = dataRow.Table.Columns.Count;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = dataRow[offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateDataTableValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            DataRow dataRow,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = dataRow.Table.Columns.Count;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                object? value = dataRow[offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            object?[] values,
            int valueOffset,
            int columnCount,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                object? value = values[valueOffset + offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            object?[] values,
            int valueOffset,
            int columnCount,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = values[valueOffset + offset];
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private Row CreateTabularRowSourceValueRow(
            int rowIndex,
            string[] columnReferencePrefixes,
            IExcelSheetTabularRowSource source,
            int sourceRowIndex,
            TabularAppendColumnKind[] columnKinds,
            uint?[] styleIndexes,
            uint? objectDateTimeStyleIndex,
            uint? objectTimeSpanStyleIndex,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes,
            bool canCancel,
            CancellationToken ct) {
            string rowReference = InvariantNumberText.Get(rowIndex);
            int columnCount = source.ColumnCount;
            bool hasObjectValueStyles = objectDateTimeStyleIndex.HasValue || objectTimeSpanStyleIndex.HasValue;
            var row = new Row { RowIndex = (uint)rowIndex };
            for (int offset = 0; offset < columnCount; offset++) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                object? value = source.GetValue(sourceRowIndex, offset);
                if (value == DBNull.Value) {
                    value = null;
                }

                var (cellValue, cellType) = CoerceTabularAppendValue(value, columnKinds[offset], useDirectStringCells, ref sharedStringIndexes);
                var cell = CreateTabularAppendCell(columnReferencePrefixes[offset] + rowReference, cellValue, cellType);

                if (styleIndexes[offset] is uint styleIndex) {
                    cell.StyleIndex = styleIndex;
                } else if (hasObjectValueStyles && TryGetObjectDataTableValueStyleIndex(value, objectDateTimeStyleIndex, objectTimeSpanStyleIndex, out uint objectValueStyleIndex)) {
                    cell.StyleIndex = objectValueStyleIndex;
                }

                row.Append(cell);
            }

            return row;
        }

        private static string[] BuildColumnReferencePrefixes(int startColumn, int columnCount) {
            var columnReferencePrefixes = new string[columnCount];
            for (int offset = 0; offset < columnReferencePrefixes.Length; offset++) {
                columnReferencePrefixes[offset] = GetColumnName(startColumn + offset);
            }

            return columnReferencePrefixes;
        }

        private static EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> GetCachedDataTableCellType(DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) {
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String) return DataTableStringCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) return DataTableSharedStringCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) return DataTableNumberCellType;
            if (cellType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) return DataTableBooleanCellType;
            return new EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues>(cellType);
        }

        private static Cell CreateTabularAppendCell(
            string cellReference,
            CellValue cellValue,
            DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) {
            var cell = new Cell {
                CellReference = cellReference,
                CellValue = cellValue
            };

            if (cellType != DocumentFormat.OpenXml.Spreadsheet.CellValues.Number) {
                cell.DataType = GetCachedDataTableCellType(cellType);
            }

            return cell;
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceTabularAppendValue(
            object? value,
            TabularAppendColumnKind columnKind,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            switch (columnKind) {
                case TabularAppendColumnKind.String:
                    if (value is string text) {
                        return CoerceTabularAppendStringValue(text, useDirectStringCells, ref sharedStringIndexes);
                    }
                    break;
                case TabularAppendColumnKind.Double:
                    if (value is double doubleValue) return CoerceValueHelper.HandleNumber(doubleValue);
                    break;
                case TabularAppendColumnKind.Float:
                    if (value is float floatValue) return CoerceValueHelper.HandleNumber(floatValue);
                    break;
                case TabularAppendColumnKind.Decimal:
                    if (value is decimal decimalValue) return CoerceValueHelper.HandleDecimal(decimalValue);
                    break;
                case TabularAppendColumnKind.SignedInteger:
                    if (value is int intValue) return CoerceValueHelper.HandleSignedInteger(intValue);
                    if (value is long longValue) return CoerceValueHelper.HandleSignedInteger(longValue);
                    if (value is short shortValue) return CoerceValueHelper.HandleSignedInteger(shortValue);
                    if (value is sbyte sbyteValue) return CoerceValueHelper.HandleSignedInteger(sbyteValue);
                    break;
                case TabularAppendColumnKind.UnsignedInteger:
                    if (value is uint uintValue) return CoerceValueHelper.HandleUnsignedInteger(uintValue);
                    if (value is ulong ulongValue) return CoerceValueHelper.HandleUnsignedInteger(ulongValue);
                    if (value is ushort ushortValue) return CoerceValueHelper.HandleUnsignedInteger(ushortValue);
                    if (value is byte byteValue) return CoerceValueHelper.HandleUnsignedInteger(byteValue);
                    break;
                case TabularAppendColumnKind.Boolean:
                    if (value is bool boolValue) return CoerceValueHelper.HandleBoolean(boolValue);
                    break;
                case TabularAppendColumnKind.DateTime:
                    if (value is DateTime dateTimeValue) return CoerceValueHelper.HandleNumber(dateTimeValue.ToOADate());
                    break;
                case TabularAppendColumnKind.DateTimeOffset:
                    break;
#if NET6_0_OR_GREATER
                case TabularAppendColumnKind.DateOnly:
                    if (value is DateOnly dateOnlyValue) return CoerceValueHelper.HandleNumber(dateOnlyValue.ToDateTime(TimeOnly.MinValue).ToOADate());
                    break;
                case TabularAppendColumnKind.TimeOnly:
                    if (value is TimeOnly timeOnlyValue) return CoerceValueHelper.HandleNumber(timeOnlyValue.ToTimeSpan().TotalDays);
                    break;
#endif
                case TabularAppendColumnKind.TimeSpan:
                    if (value is TimeSpan timeSpanValue) return CoerceValueHelper.HandleNumber(timeSpanValue.TotalDays);
                    break;
            }

            return CoerceDataTableAppendValue(value, useDirectStringCells, ref sharedStringIndexes);
        }

        private (CellValue cellValue, DocumentFormat.OpenXml.Spreadsheet.CellValues cellType) CoerceTabularAppendStringValue(
            string text,
            bool useDirectStringCells,
            ref Dictionary<string, int>? sharedStringIndexes) {
            return useDirectStringCells
                ? (CreatePlainAppendStringValue(text), DocumentFormat.OpenXml.Spreadsheet.CellValues.String)
                : (CreatePlainAppendSharedStringValue(text, ref sharedStringIndexes), DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString);
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

        private static string?[] BuildTabularRowSourceNumberFormats(IExcelSheetTabularRowSource source) {
            var formats = new string?[source.ColumnCount];
            for (int i = 0; i < source.ColumnCount; i++) {
                formats[i] = GetDataTableNumberFormat(source.GetColumnType(i), value: null);
            }

            return formats;
        }

        private static TabularAppendColumnKind[] BuildDataTableAppendColumnKinds(DataTable table) {
            var kinds = new TabularAppendColumnKind[table.Columns.Count];
            for (int i = 0; i < kinds.Length; i++) {
                kinds[i] = GetTabularAppendColumnKind(table.Columns[i].DataType);
            }

            return kinds;
        }

        private static TabularAppendColumnKind[] BuildTabularAppendColumnKinds(IExcelSheetTabularRowSource source) {
            var kinds = new TabularAppendColumnKind[source.ColumnCount];
            for (int i = 0; i < kinds.Length; i++) {
                kinds[i] = GetTabularAppendColumnKind(source.GetColumnType(i));
            }

            return kinds;
        }

        private static TabularAppendColumnKind GetTabularAppendColumnKind(Type type) {
            if (type == typeof(string)) return TabularAppendColumnKind.String;
            if (type == typeof(double)) return TabularAppendColumnKind.Double;
            if (type == typeof(float)) return TabularAppendColumnKind.Float;
            if (type == typeof(decimal)) return TabularAppendColumnKind.Decimal;
            if (type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(sbyte)) return TabularAppendColumnKind.SignedInteger;
            if (type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort) || type == typeof(byte)) return TabularAppendColumnKind.UnsignedInteger;
            if (type == typeof(bool)) return TabularAppendColumnKind.Boolean;
            if (type == typeof(DateTime)) return TabularAppendColumnKind.DateTime;
            if (type == typeof(DateTimeOffset)) return TabularAppendColumnKind.DateTimeOffset;
#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly)) return TabularAppendColumnKind.DateOnly;
            if (type == typeof(TimeOnly)) return TabularAppendColumnKind.TimeOnly;
#endif
            if (type == typeof(TimeSpan)) return TabularAppendColumnKind.TimeSpan;
            return TabularAppendColumnKind.General;
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
