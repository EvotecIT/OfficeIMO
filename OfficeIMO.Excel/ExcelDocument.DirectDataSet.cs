using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly DataSet DirectTabularSnapshotOwner = new DataSet("DirectTabularExport") { Locale = CultureInfo.InvariantCulture };

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
            CancellationToken ct = default,
            ExcelDateSystem dateSystem = ExcelDateSystem.NineteenHundred) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));
            if (dataSet.Tables.Count == 0) throw new ArgumentException("The DataSet must contain at least one DataTable.", nameof(dataSet));

            var model = DirectDataSetWorkbookModel.Create(dataSet, createTables: true, tableStyle, includeHeaders, includeAutoFilter, autoFit: false, DefaultDateTimeOffsetWriteStrategy, ct, omitBlankCells: true, dateSystem: dateSystem);
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
                    results,
                    dateSystem: DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(dataSet, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: true);
                _directDataSetMetadataSourceSheet = null;
            } catch {
                ClearDirectDataSetSaveCandidate();
            }
        }

        internal void RegisterDirectTabularSaveCandidate(
            ExcelSheet sheet,
            DataTable table,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false,
            bool copyTable = true) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!ReferenceEquals(sheet.Document, this)) {
                return;
            }

            ClearDirectDataSetSaveCandidate();

            try {
                bool shouldCopyTable = copyTable || table.DataSet != null;
                if (shouldCopyTable) {
                    string requestedName = string.IsNullOrWhiteSpace(table.TableName) ? sheet.Name : table.TableName;
                    var tableModel = DirectDataSetTableModel.Snapshot(table, CancellationToken.None);
                    var model = DirectDataSetWorkbookModel.CreateSingle(
                        sheet.Name,
                        requestedName,
                        createTable ? tableName : null,
                        range,
                        tableModel,
                        createTable,
                        tableStyle,
                        includeHeaders,
                        includeAutoFilter,
                        autoFit,
                        _dateTimeOffsetWriteStrategy,
                        CancellationToken.None,
                        dateSystem: DateSystem);
                    _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: false);
                    _directDataSetMetadataSourceSheet = sheet;
                } else {
                    var dataSet = new DataSet("ObjectExport") {
                        Locale = CultureInfo.InvariantCulture
                    };
                    if (string.IsNullOrWhiteSpace(table.TableName)) {
                        table.TableName = sheet.Name;
                    }

                    dataSet.Tables.Add(table);
                    var results = new[] {
                        new ExcelDataSetImportResult(sheet.Name, createTable ? tableName : null, range, table.Rows.Count, table.Columns.Count)
                    };
                    var model = DirectDataSetWorkbookModel.Create(
                        dataSet,
                        createTables: createTable,
                        tableStyle,
                        includeHeaders,
                        includeAutoFilter,
                        autoFit,
                        _dateTimeOffsetWriteStrategy,
                        CancellationToken.None,
                        results,
                        dateSystem: DateSystem);
                    _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(dataSet, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: false);
                    _directDataSetMetadataSourceSheet = null;
                }
            } catch {
                ClearDirectDataSetSaveCandidate();
            }
        }

        internal bool RegisterDeferredDirectTabularSaveCandidate(
            ExcelSheet sheet,
            DataTable table,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false,
            bool copyTable = false) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(table.TableName) ? sheet.Name : table.TableName;
                var tableModel = copyTable || table.DataSet != null
                    ? DirectDataSetTableModel.Snapshot(table, CancellationToken.None)
                    : DirectDataSetTableModel.Reference(table);
                var model = DirectDataSetWorkbookModel.CreateSingle(
                    sheet.Name,
                    requestedName,
                    createTable ? tableName : null,
                    range,
                    tableModel,
                    createTable,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    dateSystem: DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, MaterializeDeferredDataSetImport, isDeferred: true, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = sheet;
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        internal bool RegisterDeferredDirectTabularSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            IReadOnlyList<object?[]> rows,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false,
            bool useCellValueNumberFormats = false,
            bool replacingPendingDirectCellValues = false) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration(replacingPendingDirectCellValues)) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromRows(columnNames, columnTypes, rows);
                var model = DirectDataSetWorkbookModel.CreateSingle(
                    sheet.Name,
                    requestedName,
                    createTable ? tableName : null,
                    range,
                    tableModel,
                    createTable,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    useCellValueNumberFormats,
                    DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, MaterializeDeferredDataSetImport, isDeferred: true, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = sheet;
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        internal bool RegisterDeferredDirectCellValuesSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            object?[] values,
            int columnCount,
            int rowCount,
            bool valuesMatchColumnTypes,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false,
            bool useCellValueNumberFormats = false,
            bool replacingPendingDirectCellValues = false) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration(replacingPendingDirectCellValues)) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromCellValues(columnNames, columnTypes, values, columnCount, rowCount, valuesMatchColumnTypes);
                var model = DirectDataSetWorkbookModel.CreateSingle(
                    sheet.Name,
                    requestedName,
                    createTable ? tableName : null,
                    range,
                    tableModel,
                    createTable,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    useCellValueNumberFormats,
                    DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, MaterializeDeferredDataSetImport, isDeferred: true, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = sheet;
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        internal void RegisterDirectTabularSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            IReadOnlyList<object?[]> rows,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (rows == null) throw new ArgumentNullException(nameof(rows));
            if (!ReferenceEquals(sheet.Document, this)) {
                return;
            }

            ClearDirectDataSetSaveCandidate();

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromRows(columnNames, columnTypes, rows);
                var model = DirectDataSetWorkbookModel.CreateSingle(
                    sheet.Name,
                    requestedName,
                    createTable ? tableName : null,
                    range,
                    tableModel,
                    createTable,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    dateSystem: DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = sheet;
            } catch {
                ClearDirectDataSetSaveCandidate();
            }
        }

        internal void RegisterDirectCellValuesSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            object?[] values,
            int columnCount,
            int rowCount,
            bool valuesMatchColumnTypes,
            bool includeHeaders,
            string range,
            string? tableName = null,
            bool createTable = false,
            TableStyle tableStyle = TableStyle.TableStyleMedium2,
            bool includeAutoFilter = false,
            bool autoFit = false) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (columnTypes == null) throw new ArgumentNullException(nameof(columnTypes));
            if (values == null) throw new ArgumentNullException(nameof(values));
            if (!ReferenceEquals(sheet.Document, this)) {
                return;
            }

            ClearDirectDataSetSaveCandidate();

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromCellValues(columnNames, columnTypes, values, columnCount, rowCount, valuesMatchColumnTypes);
                var model = DirectDataSetWorkbookModel.CreateSingle(
                    sheet.Name,
                    requestedName,
                    createTable ? tableName : null,
                    range,
                    tableModel,
                    createTable,
                    tableStyle,
                    includeHeaders,
                    includeAutoFilter,
                    autoFit,
                    _dateTimeOffsetWriteStrategy,
                    CancellationToken.None,
                    dateSystem: DateSystem);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(DirectTabularSnapshotOwner, model, ClearDirectDataSetSaveCandidate, isDeferred: false, subscribeToSourceChanges: false);
                _directDataSetMetadataSourceSheet = sheet;
            } catch {
                ClearDirectDataSetSaveCandidate();
            }
        }

        internal bool TryPromoteDirectTabularSaveCandidateToTable(
            ExcelSheet sheet,
            string range,
            string tableName,
            bool includeHeaders,
            TableStyle tableStyle,
            bool includeAutoFilter) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (sheetModel.HasTable
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || !string.Equals(sheetModel.Range, range, StringComparison.OrdinalIgnoreCase)
                || sheetModel.IncludeHeaders != includeHeaders) {
                return false;
            }

            var promotedModel = candidate.Model.WithTable(
                sheet.Name,
                tableName,
                includeHeaders,
                tableStyle,
                includeAutoFilter,
                _dateTimeOffsetWriteStrategy,
                CancellationToken.None);
            _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                candidate.Owner,
                promotedModel,
                candidate.InvalidateCallback,
                candidate.IsDeferred,
                candidate.SubscribesToSourceChanges);
            _directDataSetMetadataSourceSheet = sheet;
            candidate.Dispose();

            return _directDataSetSaveCandidate != null && _directDataSetSaveCandidate.IsValid;
        }

        internal bool TryGetDirectTabularSaveCandidateHeaders(
            ExcelSheet sheet,
            string range,
            bool includeHeaders,
            out IReadOnlyList<string>? headers) {
            headers = null;
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (sheetModel.HasTable
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || !string.Equals(sheetModel.Range, range, StringComparison.OrdinalIgnoreCase)
                || sheetModel.IncludeHeaders != includeHeaders
                || sheetModel.Table.ColumnCount <= 0) {
                return false;
            }

            var candidateHeaders = new string[sheetModel.Table.ColumnCount];
            for (int i = 0; i < candidateHeaders.Length; i++) {
                candidateHeaders[i] = sheetModel.Table.GetColumnName(i);
            }

            headers = candidateHeaders;
            return true;
        }

    }
}
