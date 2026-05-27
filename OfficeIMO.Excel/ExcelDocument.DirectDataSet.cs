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
            CancellationToken ct = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The destination stream must be writable.", nameof(stream));
            if (dataSet == null) throw new ArgumentNullException(nameof(dataSet));
            if (dataSet.Tables.Count == 0) throw new ArgumentException("The DataSet must contain at least one DataTable.", nameof(dataSet));

            var model = DirectDataSetWorkbookModel.Create(dataSet, createTables: true, tableStyle, includeHeaders, includeAutoFilter, autoFit: false, DefaultDateTimeOffsetWriteStrategy, ct, omitBlankCells: true);
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
                        CancellationToken.None);
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
                        results);
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
                    CancellationToken.None);
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
                    useCellValueNumberFormats);
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
                var tableModel = DirectDataSetTableModel.FromCellValues(columnNames, columnTypes, values, columnCount, rowCount);
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
                    useCellValueNumberFormats);
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
                    CancellationToken.None);
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
                var tableModel = DirectDataSetTableModel.FromCellValues(columnNames, columnTypes, values, columnCount, rowCount);
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
                    CancellationToken.None);
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

        internal bool TryApplyDirectWorksheetAutoFilterMetadata(ExcelSheet sheet, string range) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException(nameof(range));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            return TryUpdateDeferredDirectWorksheetMetadata(
                sheet,
                requireWorksheetTable: false,
                metadata => (metadata ?? DirectWorksheetMetadata.Empty).WithAutoFilterXml("<autoFilter xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ref=\"" + EscapeXmlAttribute(range) + "\"/>"));
        }

        internal bool TryEnableDirectTableAutoFilterMetadata(ExcelSheet sheet, string range) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrEmpty(range)) throw new ArgumentNullException(nameof(range));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || !candidate.IsDeferred || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!sheetModel.HasTable
                || !sheetModel.IncludeHeaders
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || !string.Equals(sheetModel.Range, range, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (sheetModel.IncludeAutoFilter) {
                return true;
            }

            var model = candidate.Model.WithTableAutoFilter(sheet.Name, includeAutoFilter: true);
            _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                candidate.Owner,
                model,
                candidate.InvalidateCallback,
                candidate.IsDeferred,
                candidate.SubscribesToSourceChanges);
            _directDataSetMetadataSourceSheet = sheet;
            candidate.Dispose();
            _packageDirty = true;
            _unchangedPackageBytes = null;
            _requiresSavePreflight = false;
            return _directDataSetSaveCandidate != null && _directDataSetSaveCandidate.IsValid;
        }

        internal bool TryApplyDirectWorksheetFreezeMetadata(ExcelSheet sheet, int topRows, int leftCols) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (topRows < 0) throw new ArgumentOutOfRangeException(nameof(topRows));
            if (leftCols < 0) throw new ArgumentOutOfRangeException(nameof(leftCols));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            return TryUpdateDeferredDirectWorksheetMetadata(
                sheet,
                requireWorksheetTable: null,
                metadata => {
                    var baseMetadata = metadata ?? DirectWorksheetMetadata.Empty;
                    string? sheetViewsXml = topRows == 0 && leftCols == 0
                        ? null
                        : CreateFrozenSheetViewsXml(topRows, leftCols, baseMetadata.SheetViewsXml);
                    return baseMetadata.WithSheetViewsXml(sheetViewsXml);
                },
                mergeCapturedMetadata: true);
        }

        internal bool ShouldMaterializeDeferredDirectTabularSaveCandidateForTable(ExcelSheet sheet, string range, bool includeHeaders) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || !candidate.IsDeferred) {
                return false;
            }

            if (candidate.Model.Sheets.Count != 1) {
                return true;
            }

            var sheetModel = candidate.Model.Sheets[0];
            return sheetModel.HasTable
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || !string.Equals(sheetModel.Range, range, StringComparison.OrdinalIgnoreCase)
                || sheetModel.IncludeHeaders != includeHeaders;
        }

        private bool TryUpdateDeferredDirectWorksheetMetadata(
            ExcelSheet sheet,
            bool? requireWorksheetTable,
            Func<DirectWorksheetMetadata?, DirectWorksheetMetadata?> updateMetadata,
            bool mergeCapturedMetadata = false) {
            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || !candidate.IsDeferred || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                return false;
            }

            if (requireWorksheetTable.HasValue && sheetModel.HasTable != requireWorksheetTable.Value) {
                return false;
            }

            DirectWorksheetMetadata? baseMetadata = sheetModel.Metadata;
            if (mergeCapturedMetadata) {
                if (!TryCaptureDirectWorksheetMetadata(sheetModel, out DirectWorksheetMetadata? capturedMetadata, out _)) {
                    return false;
                }

                baseMetadata = MergeDirectWorksheetMetadata(baseMetadata, capturedMetadata);
            }

            DirectWorksheetMetadata? updatedMetadata = updateMetadata(baseMetadata);
            if (ReferenceEquals(updatedMetadata, sheetModel.Metadata)) {
                return true;
            }

            var metadata = new DirectWorksheetMetadata?[candidate.Model.Sheets.Count];
            metadata[0] = updatedMetadata?.IsEmpty == true ? null : updatedMetadata;
            var model = candidate.Model.WithWorksheetMetadata(metadata);
            _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                candidate.Owner,
                model,
                candidate.InvalidateCallback,
                candidate.IsDeferred,
                candidate.SubscribesToSourceChanges);
            _directDataSetMetadataSourceSheet = sheet;
            candidate.Dispose();
            _packageDirty = true;
            _unchangedPackageBytes = null;
            _requiresSavePreflight = false;
            return _directDataSetSaveCandidate != null && _directDataSetSaveCandidate.IsValid;
        }

        private static string CreateFrozenSheetViewsXml(int topRows, int leftCols, string? existingSheetViewsXml) {
            if (string.IsNullOrEmpty(existingSheetViewsXml)) {
                return CreateFrozenSheetViewsXml(topRows, leftCols);
            }

            string existingXml = existingSheetViewsXml!;
            var sheetViews = new SheetViews(existingXml);
            SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
            if (sheetView == null) {
                sheetView = new SheetView { WorkbookViewId = 0U };
                sheetViews.Append(sheetView);
            } else if (sheetView.WorkbookViewId == null) {
                sheetView.WorkbookViewId = 0U;
            }

            ApplyFrozenPaneToSheetView(sheetView, topRows, leftCols);
            return sheetViews.OuterXml;
        }

        private static string CreateFrozenSheetViewsXml(int topRows, int leftCols) {
            string topLeftCell = A1.CellReference(topRows + 1, leftCols + 1);
            string escapedTopLeftCell = EscapeXmlAttribute(topLeftCell);
            var builder = new StringBuilder(256);
            builder.Append("<sheetViews xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetView workbookViewId=\"0\"><pane");
            if (leftCols > 0) {
                builder.Append(" xSplit=\"");
                builder.Append(leftCols.ToString(CultureInfo.InvariantCulture));
                builder.Append('"');
            }

            if (topRows > 0) {
                builder.Append(" ySplit=\"");
                builder.Append(topRows.ToString(CultureInfo.InvariantCulture));
                builder.Append('"');
            }

            builder.Append(" topLeftCell=\"");
            builder.Append(escapedTopLeftCell);
            builder.Append("\" activePane=\"");
            string activePane = topRows > 0 && leftCols > 0
                ? "bottomRight"
                : topRows > 0
                    ? "bottomLeft"
                    : "topRight";
            builder.Append(activePane);
            builder.Append("\" state=\"frozen\"/>");

            if (topRows > 0 && leftCols > 0) {
                AppendFrozenSelectionXml(builder, "topRight", escapedTopLeftCell);
                AppendFrozenSelectionXml(builder, "bottomLeft", escapedTopLeftCell);
                AppendFrozenSelectionXml(builder, "bottomRight", escapedTopLeftCell);
            } else if (topRows > 0) {
                AppendFrozenSelectionXml(builder, "bottomLeft", escapedTopLeftCell);
            } else {
                AppendFrozenSelectionXml(builder, "topRight", escapedTopLeftCell);
            }

            builder.Append("<selection activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>");
            return builder.ToString();
        }

        private static void ApplyFrozenPaneToSheetView(SheetView sheetView, int topRows, int leftCols) {
            sheetView.RemoveAllChildren<Pane>();
            sheetView.RemoveAllChildren<Selection>();

            string topLeftCell = A1.CellReference(topRows + 1, leftCols + 1);
            var pane = new Pane {
                State = PaneStateValues.Frozen,
                TopLeftCell = topLeftCell
            };
            if (leftCols > 0) {
                pane.HorizontalSplit = leftCols;
            }

            if (topRows > 0) {
                pane.VerticalSplit = topRows;
            }

            if (topRows > 0 && leftCols > 0) {
                pane.ActivePane = PaneValues.BottomRight;
                sheetView.Append(pane);
                AppendFrozenSelection(sheetView, PaneValues.TopRight, topLeftCell);
                AppendFrozenSelection(sheetView, PaneValues.BottomLeft, topLeftCell);
                AppendFrozenSelection(sheetView, PaneValues.BottomRight, topLeftCell);
            } else if (topRows > 0) {
                pane.ActivePane = PaneValues.BottomLeft;
                sheetView.Append(pane);
                AppendFrozenSelection(sheetView, PaneValues.BottomLeft, topLeftCell);
            } else {
                pane.ActivePane = PaneValues.TopRight;
                sheetView.Append(pane);
                AppendFrozenSelection(sheetView, PaneValues.TopRight, topLeftCell);
            }

            sheetView.Append(new Selection {
                ActiveCell = "A1",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            });
        }

        private static void AppendFrozenSelection(SheetView sheetView, PaneValues pane, string topLeftCell) {
            sheetView.Append(new Selection {
                Pane = pane,
                ActiveCell = topLeftCell,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = topLeftCell }
            });
        }

        private static void AppendFrozenSelectionXml(StringBuilder builder, string pane, string escapedTopLeftCell) {
            builder.Append("<selection pane=\"");
            builder.Append(pane);
            builder.Append("\" activeCell=\"");
            builder.Append(escapedTopLeftCell);
            builder.Append("\" sqref=\"");
            builder.Append(escapedTopLeftCell);
            builder.Append("\"/>");
        }

        private static string EscapeXmlAttribute(string value) {
            if (value.IndexOfAny(['&', '<', '>', '"', '\'']) < 0) {
                return value;
            }

            var builder = new StringBuilder(value.Length + 8);
            foreach (char ch in value) {
                builder.Append(ch switch {
                    '&' => "&amp;",
                    '<' => "&lt;",
                    '>' => "&gt;",
                    '"' => "&quot;",
                    '\'' => "&apos;",
                    _ => ch.ToString()
                });
            }

            return builder.ToString();
        }

        internal bool TryEnableDirectTabularSaveCandidateAutoFit(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || sheetModel.Table.ColumnCount <= 0) {
                return false;
            }

            if (sheetModel.AutoFitColumns) {
                return true;
            }

            try {
                var model = candidate.Model.WithAutoFitColumns(sheet.Name, _dateTimeOffsetWriteStrategy, CancellationToken.None);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                    candidate.Owner,
                    model,
                    candidate.InvalidateCallback,
                    candidate.IsDeferred,
                    candidate.SubscribesToSourceChanges);
                _directDataSetMetadataSourceSheet = sheet;
                candidate.Dispose();
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        internal bool RegisterDeferredDirectDictionaryRowsSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows,
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
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromDictionaries(columnNames, columnTypes, rows);
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
                    CancellationToken.None);
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

        internal bool RegisterDeferredDirectExactDictionaryRowsSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            IReadOnlyList<Dictionary<string, object?>> rows,
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
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromExactDictionaries(columnNames, columnTypes, rows);
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
                    CancellationToken.None);
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

        internal bool RegisterDeferredDirectLegacyDictionaryRowsSaveCandidate(
            ExcelSheet sheet,
            string tableNameForModel,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<Type> columnTypes,
            IReadOnlyList<System.Collections.IDictionary> rows,
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
                return false;
            }

            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

            try {
                string requestedName = string.IsNullOrWhiteSpace(tableNameForModel) ? sheet.Name : tableNameForModel;
                var tableModel = DirectDataSetTableModel.FromLegacyDictionaries(columnNames, columnTypes, rows);
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
                    CancellationToken.None);
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

        internal bool TryEnableDirectTabularSaveCandidateAutoFit(ExcelSheet sheet, IReadOnlyList<int> columnIndexes) {
            if (columnIndexes == null || columnIndexes.Count == 0) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!ReferenceEquals(sheet.Document, this)
                || !string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || sheetModel.Table.ColumnCount <= 0) {
                return false;
            }

            bool allColumnsRequested = columnIndexes.Count == sheetModel.Table.ColumnCount;
            for (int i = 0; i < columnIndexes.Count; i++) {
                int columnIndex = columnIndexes[i];
                if (columnIndex <= 0 || columnIndex > sheetModel.Table.ColumnCount) {
                    return false;
                }

                allColumnsRequested &= columnIndex == i + 1;
            }

            if (allColumnsRequested) {
                return TryEnableDirectTabularSaveCandidateAutoFit(sheet);
            }

            if (sheetModel.AutoFitColumns) {
                return true;
            }

            try {
                var model = candidate.Model.WithAutoFitColumns(sheet.Name, columnIndexes, _dateTimeOffsetWriteStrategy, CancellationToken.None);
                _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                    candidate.Owner,
                    model,
                    candidate.InvalidateCallback,
                    candidate.IsDeferred,
                    candidate.SubscribesToSourceChanges);
                _directDataSetMetadataSourceSheet = sheet;
                candidate.Dispose();
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                ClearDirectDataSetSaveCandidate();
                return false;
            }
        }

        internal bool TrySetDirectTabularSaveCandidateColumnNumberFormat(ExcelSheet sheet, int columnIndex, string numberFormat) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(numberFormat)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            var sourceModel = candidate?.IsValid == true ? candidate.Model : _materializedDirectDataSetFastSaveModel;
            if (sourceModel == null || !ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var sheetModel = sourceModel.Sheets.FirstOrDefault(item => string.Equals(item.SheetName, sheet.Name, StringComparison.Ordinal));
            if (sheetModel == null
                || columnIndex <= 0
                || columnIndex > sheetModel.Table.ColumnCount) {
                return false;
            }

            try {
                var model = sourceModel.WithColumnNumberFormat(sheet.Name, columnIndex, numberFormat);
                if (candidate != null && candidate.IsValid) {
                    _directDataSetSaveCandidate = new DirectDataSetSaveCandidate(
                        candidate.Owner,
                        model,
                        candidate.InvalidateCallback,
                        candidate.IsDeferred,
                        candidate.SubscribesToSourceChanges);
                    candidate.Dispose();
                } else {
                    _materializedDirectDataSetFastSaveModel = model;
                    _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
                }

                if (_materializedDirectDataSetFastSaveModel != null) {
                    _materializedDirectDataSetFastSaveModel = _materializedDirectDataSetFastSaveModel.WithColumnNumberFormat(sheet.Name, columnIndex, numberFormat);
                    _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
                }

                _directDataSetMetadataSourceSheet = sheet;
                _packageDirty = true;
                _unchangedPackageBytes = null;
                _requiresSavePreflight = false;
                return true;
            } catch {
                if (candidate != null) {
                    ClearDirectDataSetSaveCandidate();
                }

                return false;
            }
        }

        internal bool TryGetDirectTabularSaveCandidateColumnByHeader(
            ExcelSheet sheet,
            string header,
            bool includeHeader,
            ExcelReadOptions? options,
            out int columnIndex,
            out int startRow,
            out int endRow) {
            columnIndex = 0;
            startRow = 0;
            endRow = -1;
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(header) || !ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            var sourceModel = candidate?.IsValid == true ? candidate.Model : _materializedDirectDataSetFastSaveModel;
            if (sourceModel == null) {
                return false;
            }

            var sheetModel = sourceModel.Sheets.FirstOrDefault(item => string.Equals(item.SheetName, sheet.Name, StringComparison.Ordinal));
            if (sheetModel == null || !sheetModel.IncludeHeaders) {
                return false;
            }

            bool normalizeHeaders = options?.NormalizeHeaders ?? true;
            string normalizedHeader = ExcelHeaderNameHelper.NormalizeHeader(header, normalizeHeaders);
            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(
                sheetModel.Table.ColumnCount,
                column => sheetModel.Table.GetColumnName(column),
                normalizeHeaders);
            for (int i = 0; i < headers.Length; i++) {
                if (!string.Equals(headers[i], normalizedHeader, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                columnIndex = i + 1;
                startRow = includeHeader ? 1 : 2;
                endRow = sheetModel.Table.RowCount + 1;
                return startRow <= endRow;
            }

            return false;
        }


        internal bool TryGetDirectTabularSaveCandidateColumnCount(ExcelSheet sheet, out int columnCount) {
            columnCount = 0;
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                return false;
            }

            columnCount = sheetModel.Table.ColumnCount;
            return columnCount > 0;
        }

        internal bool TryGetDeferredDirectTabularPivotSource(
            ExcelSheet sheet,
            int startRow,
            int startColumn,
            int endRow,
            int endColumn,
            out IExcelSheetTabularRowSource? source) {
            source = null;
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate != null
                && candidate.IsValid
                && candidate.IsDeferred
                && TryGetDirectTabularPivotSource(candidate.Model, sheet, startRow, startColumn, endRow, endColumn, out source)) {
                return true;
            }

            var materializedModel = _materializedDirectDataSetFastSaveModel;
            return materializedModel != null
                   && TryGetDirectTabularPivotSource(materializedModel, sheet, startRow, startColumn, endRow, endColumn, out source);
        }

        private static bool TryGetDirectTabularPivotSource(
            DirectDataSetWorkbookModel model,
            ExcelSheet sheet,
            int startRow,
            int startColumn,
            int endRow,
            int endColumn,
            out IExcelSheetTabularRowSource? source) {
            source = null;
            foreach (var sheetModel in model.Sheets) {
                if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)) {
                    continue;
                }

                if (!sheetModel.IncludeHeaders || startRow != 1 || startColumn <= 0) {
                    return false;
                }

                int rowCountWithHeader = sheetModel.Table.RowCount + 1;
                if (endRow > rowCountWithHeader || endColumn > sheetModel.Table.ColumnCount) {
                    return false;
                }

                source = sheetModel.Table;
                return true;
            }

            return false;
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
            if (!TryBeginDeferredDirectSaveCandidateRegistration()) {
                return false;
            }

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
                _directDataSetMetadataSourceSheet = null;
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
                _directDataSetMetadataSourceSheet = null;
                return;
            }

            _directDataSetSaveCandidate = null;
            _directDataSetMetadataSourceSheet = null;
            candidate.Dispose();
        }

        private bool TryBeginDeferredDirectSaveCandidateRegistration(bool replacingPendingDirectCellValues = false) {
            var candidate = _directDataSetSaveCandidate;
            if (candidate != null) {
                if (candidate.IsDeferred && candidate.IsValid) {
                    MaterializeDeferredDataSetImport();
                    return false;
                }

                ClearDirectDataSetSaveCandidate();
            }

            if (_pendingDirectCellValueSheet != null && !replacingPendingDirectCellValues) {
                MaterializePendingDirectCellValueSheetIfNeeded();
                return false;
            }

            return true;
        }

        internal bool TryReservePendingDirectCellValueSheet(ExcelSheet sheet) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                return false;
            }

            if (_pendingDirectCellValueSheet == null) {
                _pendingDirectCellValueSheet = sheet;
                return true;
            }

            return ReferenceEquals(_pendingDirectCellValueSheet, sheet);
        }

        internal void ClearPendingDirectCellValueSheet(ExcelSheet sheet) {
            if (ReferenceEquals(_pendingDirectCellValueSheet, sheet)) {
                _pendingDirectCellValueSheet = null;
            }
        }

        private void MaterializePendingDirectCellValueSheetIfNeeded() {
            var sheet = _pendingDirectCellValueSheet;
            if (sheet == null) {
                return;
            }

            _pendingDirectCellValueSheet = null;
            sheet.MaterializePendingDirectCellValues();
        }

        private void PromotePendingDirectCellValueSheetIfPossible() {
            var sheet = _pendingDirectCellValueSheet;
            if (sheet == null) {
                return;
            }

            if (!sheet.TryPromotePendingDirectCellValuesToSaveCandidate()) {
                MaterializePendingDirectCellValueSheetIfNeeded();
            }
        }

        internal void MaterializeDeferredDataSetImport() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

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

        internal bool HasDeferredDirectDataSetImport
            => !_materializingDeferredDataSetImport
               && _directDataSetSaveCandidate?.IsDeferred == true;

        internal bool HasPendingDirectCellValues => _pendingDirectCellValueSheet != null;

        private void MaterializeDirectDataSetModel(DirectDataSetWorkbookModel model) {
            foreach (var sheetModel in model.Sheets) {
                ExcelSheet sheet = TryGetExistingSheet(sheetModel.SheetName)
                    ?? AddWorkSheet(sheetModel.SheetName, SheetNameValidationMode.Strict);
                if (sheetModel.Range.Length == 0) {
                    continue;
                }

                DirectWorksheetMetadata? preservedMetadata = null;
                if (TryCaptureDirectWorksheetMetadata(sheetModel, out DirectWorksheetMetadata? sheetMetadata, out _, allowDrawings: true)) {
                    preservedMetadata = MergeDirectWorksheetMetadata(sheetModel.Metadata, sheetMetadata);
                }

                ResetWorksheetForDirectDataSetMaterialization(sheet.WorksheetPart);
                using var noLock = sheet.BeginNoLock();
                if (sheetModel.HasTable) {
                    string materializedRange = sheet.InsertTabularRowSourceAsTableForDeferredMaterialization(
                        sheetModel.Table,
                        includeHeaders: sheetModel.IncludeHeaders,
                        tableName: sheetModel.TableName,
                        style: sheetModel.TableStyle,
                        includeAutoFilter: sheetModel.IncludeAutoFilter);
                    if (materializedRange.Length == 0) {
                        sheet.InsertDataTableAsTable(
                            sheetModel.Table.ToDataTable(),
                            includeHeaders: sheetModel.IncludeHeaders,
                            tableName: sheetModel.TableName,
                            style: sheetModel.TableStyle,
                            includeAutoFilter: sheetModel.IncludeAutoFilter);
                    }
                } else {
                    if (!sheet.TryInsertTabularRowSourceForDeferredMaterialization(
                        sheetModel.Table,
                        includeHeaders: sheetModel.IncludeHeaders)) {
                        sheet.InsertDataTable(
                            sheetModel.Table.ToDataTable(),
                            includeHeaders: sheetModel.IncludeHeaders);
                    }
                }

                if (sheetModel.ColumnWidths is { Length: > 0 } columnWidths) {
                    sheet.ApplyAutoFitColumnWidthsForDeferredMaterialization(columnWidths);
                } else if (sheetModel.AutoFitColumns && sheetModel.Table.ColumnCount > 0) {
                    sheet.AutoFitColumnsFor(Enumerable.Range(1, sheetModel.Table.ColumnCount));
                }

                ApplyDirectMaterializedColumnNumberFormats(sheet, sheetModel);
                ApplyCapturedDirectWorksheetMetadata(sheet.WorksheetPart.Worksheet!, preservedMetadata);
            }
        }

        private static void ApplyDirectMaterializedColumnNumberFormats(ExcelSheet sheet, DirectDataSetSheetModel sheetModel) {
            var formats = sheetModel.ColumnNumberFormats;
            if (formats == null || formats.Count == 0 || sheetModel.Table.RowCount == 0) {
                return;
            }

            int startRow = sheetModel.IncludeHeaders ? 2 : 1;
            int endRow = startRow + sheetModel.Table.RowCount - 1;
            for (int i = 0; i < formats.Count && i < sheetModel.Table.ColumnCount; i++) {
                string? numberFormat = formats[i];
                if (string.IsNullOrWhiteSpace(numberFormat)) {
                    continue;
                }

                string column = A1.ColumnIndexToLetters(i + 1);
                sheet.FormatRange(
                    column + startRow.ToString(CultureInfo.InvariantCulture) + ":" + column + endRow.ToString(CultureInfo.InvariantCulture),
                    numberFormat!);
            }
        }

        private static void ApplyCapturedDirectWorksheetMetadata(Worksheet worksheet, DirectWorksheetMetadata? metadata) {
            if (metadata == null || metadata.IsEmpty) {
                return;
            }

            if (!string.IsNullOrEmpty(metadata.SheetPropertiesXml)) {
                InsertWorksheetMetadataElement(worksheet, new SheetProperties(metadata.SheetPropertiesXml!), typeof(SheetDimension), typeof(SheetViews), typeof(SheetFormatProperties), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.SheetViewsXml)) {
                InsertWorksheetMetadataElement(worksheet, new SheetViews(metadata.SheetViewsXml!), typeof(SheetFormatProperties), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.SheetFormatPropertiesXml)) {
                InsertWorksheetMetadataElement(worksheet, CreateElementWithAttributes<SheetFormatProperties>(metadata.SheetFormatPropertiesXml!), typeof(Columns), typeof(SheetData));
            }

            if (!string.IsNullOrEmpty(metadata.AutoFilterXml)) {
                InsertWorksheetMetadataElement(worksheet, new AutoFilter(metadata.AutoFilterXml!), typeof(DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting), typeof(DataValidations), typeof(TableParts));
            }

            foreach (var conditionalFormattingXml in metadata.ConditionalFormattingXml) {
                InsertWorksheetMetadataElement(worksheet, new DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting(conditionalFormattingXml), false, typeof(DataValidations), typeof(TableParts));
            }

            if (!string.IsNullOrEmpty(metadata.DataValidationsXml)) {
                InsertWorksheetMetadataElement(worksheet, new DataValidations(metadata.DataValidationsXml!), typeof(TableParts));
            }

            foreach (var xml in metadata.PostDataValidationXml) {
                var element = CreatePostDataValidationElement(xml);
                if (element != null) {
                    InsertWorksheetMetadataElement(worksheet, element, typeof(TableParts));
                }
            }

            if (!string.IsNullOrEmpty(metadata.DrawingXml)) {
                InsertWorksheetMetadataElement(worksheet, CreateElementWithAttributes<DocumentFormat.OpenXml.Spreadsheet.Drawing>(metadata.DrawingXml!), typeof(TableParts));
            }

            ApplyCapturedDirectOverlayCells(worksheet, metadata.OverlayCells);
        }

        private static void ApplyCapturedDirectOverlayCells(Worksheet worksheet, IReadOnlyList<DirectOverlayCell> overlayCells) {
            if (overlayCells.Count == 0) {
                return;
            }

            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
            foreach (var overlayCell in overlayCells.OrderBy(static cell => cell.Row).ThenBy(static cell => cell.Column)) {
                if (overlayCell.IsDeleted) {
                    continue;
                }

                Row row = GetOrCreateDirectOverlayRow(sheetData, overlayCell.Row);
                Cell cell = GetOrCreateDirectOverlayCell(row, overlayCell.Row, overlayCell.Column);
                cell.StyleIndex = overlayCell.StyleIndex.HasValue ? overlayCell.StyleIndex.Value : null;
                ApplyCapturedDirectOverlayCellValue(cell, overlayCell.Value);
            }
        }

        private static Row GetOrCreateDirectOverlayRow(SheetData sheetData, int rowIndex) {
            Row? insertAfter = null;
            foreach (Row row in sheetData.Elements<Row>()) {
                uint currentIndex = row.RowIndex?.Value ?? 0U;
                if (currentIndex == (uint)rowIndex) {
                    return row;
                }

                if (currentIndex > (uint)rowIndex) {
                    break;
                }

                insertAfter = row;
            }

            var created = new Row { RowIndex = (uint)rowIndex };
            if (insertAfter == null) {
                var first = sheetData.Elements<Row>().FirstOrDefault();
                if (first == null) {
                    sheetData.Append(created);
                } else {
                    sheetData.InsertBefore(created, first);
                }
            } else if (insertAfter.NextSibling<Row>() == null) {
                sheetData.Append(created);
            } else {
                sheetData.InsertAfter(created, insertAfter);
            }

            return created;
        }

        private static Cell GetOrCreateDirectOverlayCell(Row row, int rowIndex, int columnIndex) {
            string reference = A1.CellReference(rowIndex, columnIndex);
            Cell? insertAfter = null;
            foreach (Cell cell in row.Elements<Cell>()) {
                if (string.Equals(cell.CellReference?.Value, reference, StringComparison.Ordinal)) {
                    return cell;
                }

                if (cell.CellReference?.Value is string currentReference
                    && currentReference.Length > 0
                    && GetDirectOverlayColumnIndex(currentReference) > columnIndex) {
                    break;
                }

                insertAfter = cell;
            }

            var created = new Cell { CellReference = reference };
            if (insertAfter == null) {
                var first = row.Elements<Cell>().FirstOrDefault();
                if (first == null) {
                    row.Append(created);
                } else {
                    row.InsertBefore(created, first);
                }
            } else if (insertAfter.NextSibling<Cell>() == null) {
                row.Append(created);
            } else {
                row.InsertAfter(created, insertAfter);
            }

            return created;
        }

        private static void ApplyCapturedDirectOverlayCellValue(Cell cell, object? value) {
            cell.CellFormula = null;
            cell.InlineString = null;

            switch (value) {
                case null:
                case DBNull _:
                    cell.CellValue = new CellValue(string.Empty);
                    cell.DataType = CellValues.String;
                    break;
                case DirectFormulaCellValue formula:
                    cell.CellFormula = !string.IsNullOrEmpty(formula.FormulaXml)
                        ? CreateCellFormulaFromXml(formula.FormulaXml!)
                        : new CellFormula(formula.Formula);
                    cell.CellValue = formula.CachedValue != null ? new CellValue(formula.CachedValue) : null;
                    cell.DataType = null;
                    break;
                case DirectTypedCellValue typed:
                    cell.CellValue = typed.Value != null ? new CellValue(typed.Value) : null;
                    cell.DataType = GetDirectTypedCellDataType(typed.DataType);
                    cell.InlineString = !string.IsNullOrEmpty(typed.InlineStringXml)
                        ? CreateInlineStringFromXml(typed.InlineStringXml!)
                        : null;
                    break;
                case bool boolean:
                    cell.CellValue = new CellValue(boolean ? "1" : "0");
                    cell.DataType = CellValues.Boolean;
                    break;
                case byte number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case sbyte number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case short number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case ushort number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case int number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case uint number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case long number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case ulong number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case float number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case double number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case decimal number:
                    ApplyCapturedDirectOverlayNumber(cell, number);
                    break;
                case DateTime dateTime:
                    ApplyCapturedDirectOverlayNumber(cell, dateTime.ToOADate());
                    break;
                default:
                    cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    cell.DataType = CellValues.String;
                    break;
            }
        }

        private static void ApplyCapturedDirectOverlayNumber<T>(Cell cell, T value) where T : IFormattable {
            cell.CellValue = new CellValue(value.ToString(null, CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }

        private static int GetDirectOverlayColumnIndex(string cellReference) {
            int column = 0;
            for (int i = 0; i < cellReference.Length; i++) {
                char ch = cellReference[i];
                if (ch >= 'A' && ch <= 'Z') {
                    column = checked((column * 26) + ch - 'A' + 1);
                } else if (ch >= 'a' && ch <= 'z') {
                    column = checked((column * 26) + ch - 'a' + 1);
                } else {
                    break;
                }
            }

            return column;
        }

        internal void MaterializeDeferredDataSetImportPreservingFastSaveModel() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                return;
            }

            DirectDataSetWorkbookModel? fastSaveModel = null;
            if (TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel? packageModel, out _, allowDrawings: true)) {
                fastSaveModel = packageModel;
            }

            _directDataSetSaveCandidate = null;
            candidate.Dispose();

            _materializingDeferredDataSetImport = true;
            try {
                MaterializeDirectDataSetModel(candidate.Model);
                if (fastSaveModel != null) {
                    _materializedDirectDataSetFastSaveModel = fastSaveModel;
                    _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
                }
            } finally {
                _materializingDeferredDataSetImport = false;
            }
        }

        internal void PreserveDeferredDataSetFastSaveModelAndClearCandidate() {
            if (_materializingDeferredDataSetImport) {
                return;
            }

            MaterializePendingDirectCellValueSheetIfNeeded();

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsDeferred) {
                ClearDirectDataSetSaveCandidate();
                return;
            }

            if (!CanCreateDirectPackageModel(candidate.Model, out _, allowDrawings: true)) {
                return;
            }

            _materializedDirectDataSetFastSaveModel = candidate.Model;
            _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
            _directDataSetSaveCandidate = null;
            candidate.Dispose();
        }

        private static void InsertWorksheetMetadataElement(Worksheet worksheet, DocumentFormat.OpenXml.OpenXmlElement element, params Type[] beforeTypes) {
            InsertWorksheetMetadataElement(worksheet, element, removeExistingSameType: true, beforeTypes);
        }

        private static void InsertWorksheetMetadataElement(Worksheet worksheet, DocumentFormat.OpenXml.OpenXmlElement element, bool removeExistingSameType, params Type[] beforeTypes) {
            if (removeExistingSameType) {
                foreach (var existing in worksheet.ChildElements.Where(child => child.GetType() == element.GetType()).ToList()) {
                    worksheet.RemoveChild(existing);
                }
            }

            foreach (var child in worksheet.ChildElements) {
                for (int i = 0; i < beforeTypes.Length; i++) {
                    if (beforeTypes[i].IsInstanceOfType(child)) {
                        worksheet.InsertBefore(element, child);
                        return;
                    }
                }
            }

            worksheet.Append(element);
        }

        private static DocumentFormat.OpenXml.OpenXmlElement? CreatePostDataValidationElement(string xml) {
            return GetXmlRootLocalName(xml) switch {
                "printOptions" => CreateElementWithAttributes<PrintOptions>(xml),
                "pageMargins" => CreateElementWithAttributes<PageMargins>(xml),
                "pageSetup" => CreateElementWithAttributes<PageSetup>(xml),
                "headerFooter" => new HeaderFooter(xml),
                "rowBreaks" => new RowBreaks(xml),
                "colBreaks" => new ColumnBreaks(xml),
                "cellWatches" => new CellWatches(xml),
                "ignoredErrors" => new DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors(xml),
                _ => null
            };
        }

        private static T CreateElementWithAttributes<T>(string xml) where T : DocumentFormat.OpenXml.OpenXmlElement, new() {
            var element = new T();
            using var reader = System.Xml.XmlReader.Create(new StringReader(xml), new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });

            if (!reader.Read() || reader.NodeType != System.Xml.XmlNodeType.Element) {
                return element;
            }

            if (reader.HasAttributes) {
                while (reader.MoveToNextAttribute()) {
                    if (reader.Prefix == "xmlns" || string.Equals(reader.Name, "xmlns", StringComparison.Ordinal)) {
                        continue;
                    }

                    element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        reader.Prefix,
                        reader.LocalName,
                        reader.NamespaceURI,
                        reader.Value));
                }
            }

            return element;
        }

        private static CellFormula CreateCellFormulaFromXml(string xml) {
            var formula = new CellFormula();
            using var reader = System.Xml.XmlReader.Create(new StringReader(xml), new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            });

            if (!reader.Read() || reader.NodeType != System.Xml.XmlNodeType.Element) {
                return formula;
            }

            if (reader.HasAttributes) {
                while (reader.MoveToNextAttribute()) {
                    if (reader.Prefix == "xmlns" || string.Equals(reader.Name, "xmlns", StringComparison.Ordinal)) {
                        continue;
                    }

                    formula.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(
                        reader.Prefix,
                        reader.LocalName,
                        reader.NamespaceURI,
                        reader.Value));
                }

                reader.MoveToElement();
            }

            formula.Text = reader.IsEmptyElement ? string.Empty : reader.ReadElementContentAsString();
            return formula;
        }

        private static InlineString CreateInlineStringFromXml(string xml) {
            try {
                return new InlineString(xml);
            } catch (ArgumentException) {
                return new InlineString();
            }
        }

        private static CellValues GetDirectTypedCellDataType(string dataType) {
            return dataType switch {
                "b" => CellValues.Boolean,
                "d" => CellValues.Date,
                "e" => CellValues.Error,
                "inlineStr" => CellValues.InlineString,
                "n" => CellValues.Number,
                "s" => CellValues.SharedString,
                "str" => CellValues.String,
                _ => CellValues.String
            };
        }

        private static string GetXmlRootLocalName(string xml) {
            if (string.IsNullOrWhiteSpace(xml)) {
                return string.Empty;
            }

            int start = 0;
            while (start < xml.Length && char.IsWhiteSpace(xml[start])) start++;
            if (start >= xml.Length || xml[start] != '<') {
                return string.Empty;
            }

            start++;
            if (start < xml.Length && xml[start] == '/') start++;
            int end = start;
            while (end < xml.Length && !char.IsWhiteSpace(xml[end]) && xml[end] != '>' && xml[end] != '/') end++;
            if (end <= start) {
                return string.Empty;
            }

            string qualifiedName = xml.Substring(start, end - start);
            int separator = qualifiedName.IndexOf(':');
            return separator >= 0 ? qualifiedName.Substring(separator + 1) : qualifiedName;
        }

        private ExcelSheet? TryGetExistingSheet(string sheetName) {
            Sheet? sheetElement = null;
            var sheets = WorkbookRoot.Sheets;
            if (sheets != null) {
                foreach (var candidate in sheets.Elements<Sheet>()) {
                    if (string.Equals(candidate.Name?.Value, sheetName, StringComparison.Ordinal)) {
                        sheetElement = candidate;
                        break;
                    }
                }
            }

            if (sheetElement?.Id == null) {
                return null;
            }

            if (WorkbookPartRoot.GetPartById(sheetElement.Id!) is not WorksheetPart) {
                return null;
            }

            return new ExcelSheet(this, _spreadSheetDocument, sheetElement);
        }

        private void ResetWorksheetForDirectDataSetMaterialization(WorksheetPart worksheetPart) {
            foreach (var tablePart in worksheetPart.TableDefinitionParts.ToList()) {
                string? tableName = tablePart.Table?.Name?.Value;
                if (!string.IsNullOrWhiteSpace(tableName)) {
                    RemoveReservedTableName(tableName!);
                }

                worksheetPart.DeletePart(tablePart);
            }

            worksheetPart.Worksheet = new Worksheet(new SheetData());
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

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save policy requires the standard package finalization path.";
                return false;
            }

            if (_materializedDirectDataSetFastSaveModel != null) {
                skipReason = "A materialized direct DataSet fast-save model requires the extended package writer.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (WorkbookRoot.DefinedNames?.Elements<DocumentFormat.OpenXml.Spreadsheet.DefinedName>().Any() == true) {
                skipReason = "Workbook defined names require the standard package finalization path.";
                return false;
            }

            PromotePendingDirectCellValueSheetIfPossible();

            if (HasWorkbookContentOutsideDirectDataSetImport(allowSheets: true)) {
                skipReason = "Workbook-level metadata requires the standard package finalization path.";
                return false;
            }

            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid) {
                skipReason = "No valid direct DataSet save candidate is available.";
                ClearDirectDataSetSaveCandidate();
                return false;
            }

            if (!TryCreateDirectPackageModel(candidate.Model, out DirectDataSetWorkbookModel? packageModel, out skipReason)) {
                return false;
            }

            if (ct.CanBeCanceled) {
                ct.ThrowIfCancellationRequested();
            }

            PrepareDestinationStreamForWrite(destination);
            DirectDataSetWorkbookWriter.Write(destination, packageModel, ct);
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

        private bool TryCreateDirectPackageModel(DirectDataSetWorkbookModel sourceModel, out DirectDataSetWorkbookModel model, out string? skipReason, bool allowDrawings = false) {
            DirectWorksheetMetadata?[]? metadata = null;
            for (int i = 0; i < sourceModel.Sheets.Count; i++) {
                var sheetModel = sourceModel.Sheets[i];
                if (!TryCaptureDirectWorksheetMetadata(sheetModel, out DirectWorksheetMetadata? sheetMetadata, out skipReason, allowDrawings)) {
                    model = sourceModel;
                    return false;
                }

                sheetMetadata = MergeDirectWorksheetMetadata(sheetModel.Metadata, sheetMetadata);
                if (sheetMetadata?.IsEmpty == false) {
                    metadata ??= new DirectWorksheetMetadata?[sourceModel.Sheets.Count];
                    metadata[i] = sheetMetadata;
                }
            }

            model = metadata != null ? sourceModel.WithWorksheetMetadata(metadata) : sourceModel;
            skipReason = null;
            return true;
        }

        private bool CanCreateDirectPackageModel(DirectDataSetWorkbookModel sourceModel, out string? skipReason, bool allowDrawings = false) {
            for (int i = 0; i < sourceModel.Sheets.Count; i++) {
                if (!CanCaptureDirectWorksheetMetadata(sourceModel.Sheets[i], out skipReason, allowDrawings)) {
                    return false;
                }
            }

            skipReason = null;
            return true;
        }

        private bool TryRefreshMaterializedDirectDataSetFastSaveModel(out string? skipReason) {
            skipReason = null;
            var model = _materializedDirectDataSetFastSaveModel;
            if (model == null) {
                return true;
            }

            if (!TryCreateDirectPackageModel(model, out DirectDataSetWorkbookModel refreshedModel, out skipReason, allowDrawings: true)) {
                return false;
            }

            _materializedDirectDataSetFastSaveModel = refreshedModel;
            return true;
        }

        private static DirectWorksheetMetadata? MergeDirectWorksheetMetadata(DirectWorksheetMetadata? existing, DirectWorksheetMetadata? captured) {
            if (existing == null || existing.IsEmpty) {
                return captured?.IsEmpty == true ? null : captured;
            }

            if (captured == null || captured.IsEmpty) {
                return existing;
            }

            return new DirectWorksheetMetadata(
                existing.SheetPropertiesXml ?? captured.SheetPropertiesXml,
                existing.SheetViewsXml ?? captured.SheetViewsXml,
                existing.SheetFormatPropertiesXml ?? captured.SheetFormatPropertiesXml,
                existing.AutoFilterXml ?? captured.AutoFilterXml,
                CombineMetadataXmlLists(existing.ConditionalFormattingXml, captured.ConditionalFormattingXml),
                existing.DataValidationsXml ?? captured.DataValidationsXml,
                existing.DrawingXml ?? captured.DrawingXml,
                CombineMetadataXmlLists(existing.PostDataValidationXml, captured.PostDataValidationXml),
                CombineOverlayCells(existing.OverlayCells, captured.OverlayCells));
        }

        private static IReadOnlyList<string> CombineMetadataXmlLists(IReadOnlyList<string> first, IReadOnlyList<string> second) {
            if (first.Count == 0) return second;
            if (second.Count == 0) return first;

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var combined = new List<string>(first.Count + second.Count);
            for (int i = 0; i < first.Count; i++) {
                if (seen.Add(first[i])) {
                    combined.Add(first[i]);
                }
            }

            for (int i = 0; i < second.Count; i++) {
                if (seen.Add(second[i])) {
                    combined.Add(second[i]);
                }
            }

            return combined;
        }

        private static IReadOnlyList<DirectOverlayCell> CombineOverlayCells(IReadOnlyList<DirectOverlayCell> first, IReadOnlyList<DirectOverlayCell> second) {
            if (first.Count == 0) return second;
            if (second.Count == 0) return first;

            var combined = new Dictionary<(int Row, int Column), DirectOverlayCell>();
            for (int i = 0; i < first.Count; i++) {
                combined[(first[i].Row, first[i].Column)] = first[i];
            }

            for (int i = 0; i < second.Count; i++) {
                combined[(second[i].Row, second[i].Column)] = second[i];
            }

            return combined.Values
                .Where(static cell => !cell.IsDeleted)
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray();
        }

        private bool TryCaptureDirectWorksheetMetadata(DirectDataSetSheetModel sheetModel, out DirectWorksheetMetadata? metadata, out string? skipReason, bool allowDrawings = false) {
            metadata = null;
            skipReason = null;

            ExcelSheet? sheet = null;
            var metadataSourceSheet = _directDataSetMetadataSourceSheet;
            if (metadataSourceSheet != null
                && ReferenceEquals(metadataSourceSheet.Document, this)
                && string.Equals(metadataSourceSheet.Name, sheetModel.SheetName, StringComparison.Ordinal)) {
                sheet = metadataSourceSheet;
            }

            sheet ??= TryGetExistingSheet(sheetModel.SheetName);
            if (sheet == null) {
                return true;
            }

            var worksheetPart = sheet.DeferredMetadataWorksheetPart;
            if (worksheetPart.DrawingsPart != null && !allowDrawings) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            if (worksheetPart.HyperlinkRelationships.Any()) {
                skipReason = "Worksheet contains hyperlink relationships.";
                return false;
            }

            foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                if (!sheetModel.HasTable) {
                    skipReason = "Worksheet contains table metadata outside the direct table model.";
                    return false;
                }

                var tableAutoFilter = tableDefinitionPart.Table?.Elements<AutoFilter>().FirstOrDefault();
                if (tableAutoFilter != null && tableAutoFilter.HasChildren) {
                    skipReason = "Worksheet contains table AutoFilter criteria outside the direct table model.";
                    return false;
                }
            }

            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null) {
                return true;
            }

            string? sheetPropertiesXml = null;
            string? sheetViewsXml = null;
            string? sheetFormatPropertiesXml = null;
            string? autoFilterXml = null;
            string? dataValidationsXml = null;
            string? drawingXml = null;
            IReadOnlyList<DirectOverlayCell> overlayCells = Array.Empty<DirectOverlayCell>();
            List<string>? conditionalFormattingXml = null;
            List<string>? postDataValidationXml = null;
            foreach (var child in worksheet.ChildElements) {
                switch (child) {
                    case SheetProperties sheetProperties when sheetPropertiesXml == null:
                        sheetPropertiesXml = sheetProperties.OuterXml;
                        break;
                    case SheetDimension:
                        break;
                    case SheetData sheetData:
                        overlayCells = CaptureDirectWorksheetOverlayCells(sheet, sheetModel, sheetData, _spreadSheetDocument.WorkbookPart?.WorkbookStylesPart?.Stylesheet);
                        break;
                    case SheetViews sheetViews when sheetViewsXml == null:
                        sheetViewsXml = sheetViews.OuterXml;
                        break;
                    case SheetFormatProperties sheetFormatProperties when sheetFormatPropertiesXml == null:
                        sheetFormatPropertiesXml = sheetFormatProperties.OuterXml;
                        break;
                    case Columns when sheetModel.ColumnWidths is { Length: > 0 }:
                        break;
                    case Columns:
                        skipReason = "Worksheet contains column metadata outside the direct DataSet column width model.";
                        return false;
                    case AutoFilter autoFilter when autoFilterXml == null:
                        autoFilterXml = autoFilter.OuterXml;
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting conditionalFormatting:
                        conditionalFormattingXml ??= new List<string>();
                        conditionalFormattingXml.Add(conditionalFormatting.OuterXml);
                        break;
                    case DataValidations dataValidations when dataValidationsXml == null:
                        dataValidationsXml = dataValidations.OuterXml;
                        break;
                    case PrintOptions:
                    case PageMargins:
                    case PageSetup:
                    case HeaderFooter:
                    case RowBreaks:
                    case ColumnBreaks:
                    case CellWatches:
                    case DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors:
                        postDataValidationXml ??= new List<string>();
                        postDataValidationXml.Add(child.OuterXml);
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.Drawing drawing when allowDrawings && drawingXml == null:
                        drawingXml = drawing.OuterXml;
                        break;
                    case TableParts when sheetModel.HasTable:
                        break;
                    default:
                        skipReason = "Worksheet contains unsupported element '" + child.LocalName + "' for the direct DataSet package writer.";
                        return false;
                }
            }

            if (sheetPropertiesXml == null
                && sheetViewsXml == null
                && sheetFormatPropertiesXml == null
                && autoFilterXml == null
                && dataValidationsXml == null
                && drawingXml == null
                && overlayCells.Count == 0
                && (conditionalFormattingXml == null || conditionalFormattingXml.Count == 0)
                && (postDataValidationXml == null || postDataValidationXml.Count == 0)) {
                return true;
            }

            metadata = new DirectWorksheetMetadata(
                sheetPropertiesXml,
                sheetViewsXml,
                sheetFormatPropertiesXml,
                autoFilterXml,
                conditionalFormattingXml?.ToArray() ?? Array.Empty<string>(),
                dataValidationsXml,
                drawingXml,
                postDataValidationXml?.ToArray() ?? Array.Empty<string>(),
                overlayCells);
            return true;
        }

        private bool CanCaptureDirectWorksheetMetadata(DirectDataSetSheetModel sheetModel, out string? skipReason, bool allowDrawings = false) {
            skipReason = null;

            ExcelSheet? sheet = null;
            var metadataSourceSheet = _directDataSetMetadataSourceSheet;
            if (metadataSourceSheet != null
                && ReferenceEquals(metadataSourceSheet.Document, this)
                && string.Equals(metadataSourceSheet.Name, sheetModel.SheetName, StringComparison.Ordinal)) {
                sheet = metadataSourceSheet;
            }

            sheet ??= TryGetExistingSheet(sheetModel.SheetName);
            if (sheet == null) {
                return true;
            }

            var worksheetPart = sheet.DeferredMetadataWorksheetPart;
            if (worksheetPart.DrawingsPart != null && !allowDrawings) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            if (worksheetPart.HyperlinkRelationships.Any()) {
                skipReason = "Worksheet contains hyperlink relationships.";
                return false;
            }

            foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                if (!sheetModel.HasTable) {
                    skipReason = "Worksheet contains table metadata outside the direct table model.";
                    return false;
                }

                var tableAutoFilter = tableDefinitionPart.Table?.Elements<AutoFilter>().FirstOrDefault();
                if (tableAutoFilter != null && tableAutoFilter.HasChildren) {
                    skipReason = "Worksheet contains table AutoFilter criteria outside the direct table model.";
                    return false;
                }
            }

            var worksheet = worksheetPart.Worksheet;
            if (worksheet == null) {
                return true;
            }

            foreach (var child in worksheet.ChildElements) {
                switch (child) {
                    case SheetProperties:
                    case SheetDimension:
                    case SheetData:
                    case SheetViews:
                    case SheetFormatProperties:
                    case AutoFilter:
                    case DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting:
                    case DataValidations:
                    case PrintOptions:
                    case PageMargins:
                    case PageSetup:
                    case HeaderFooter:
                    case RowBreaks:
                    case ColumnBreaks:
                    case CellWatches:
                    case DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors:
                        break;
                    case Columns when sheetModel.ColumnWidths is { Length: > 0 }:
                        break;
                    case Columns:
                        skipReason = "Worksheet contains column metadata outside the direct DataSet column width model.";
                        return false;
                    case DocumentFormat.OpenXml.Spreadsheet.Drawing when allowDrawings:
                        break;
                    case TableParts when sheetModel.HasTable:
                        break;
                    default:
                        skipReason = "Worksheet contains unsupported element '" + child.LocalName + "' for the direct DataSet package writer.";
                        return false;
                }
            }

            return true;
        }

        private static IReadOnlyList<DirectOverlayCell> CaptureDirectWorksheetOverlayCells(ExcelSheet sheet, DirectDataSetSheetModel sheetModel, SheetData sheetData, Stylesheet? stylesheet) {
            int directLastRow = sheetModel.Table.RowCount + (sheetModel.IncludeHeaders ? 1 : 0);
            List<DirectOverlayCell>? cells = null;
            int nextRowIndex = 1;
            foreach (var row in sheetData.Elements<Row>()) {
                int rowIndex = row.RowIndex?.Value is uint explicitRow ? checked((int)explicitRow) : nextRowIndex;
                nextRowIndex = checked(rowIndex + 1);
                int nextColumnIndex = 1;
                foreach (var cell in row.Elements<Cell>()) {
                    if (!TryGetCellCoordinates(cell, rowIndex, nextColumnIndex, out int cellRow, out int cellColumn)) {
                        continue;
                    }

                    nextColumnIndex = checked(cellColumn + 1);
                    if (cellColumn <= 0 || (cellRow <= directLastRow && cellColumn <= sheetModel.Table.ColumnCount)) {
                        continue;
                    }

                    object? value = ReadDirectOverlayCellValue(sheet, cell);
                    if (value == null || value == DBNull.Value) {
                        cells ??= new List<DirectOverlayCell>();
                        cells.Add(new DirectOverlayCell(cellRow, cellColumn, null, null, null, isDeleted: true));
                        continue;
                    }

                    cells ??= new List<DirectOverlayCell>();
                    cells.Add(new DirectOverlayCell(cellRow, cellColumn, value, cell.StyleIndex?.Value, ResolveDirectOverlayNumberFormat(stylesheet, cell)));
                }
            }

            return cells ?? (IReadOnlyList<DirectOverlayCell>)Array.Empty<DirectOverlayCell>();
        }

        private static string? ResolveDirectOverlayNumberFormat(Stylesheet? stylesheet, Cell cell) {
            if (cell.StyleIndex?.Value is not uint styleIndex) {
                return null;
            }

            var cellFormat = stylesheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
            if (cellFormat?.NumberFormatId?.Value is not uint numberFormatId || numberFormatId == 0U) {
                return null;
            }

            string? customFormat = stylesheet?.NumberingFormats?.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
                ?.FormatCode
                ?.Value;
            return customFormat ?? ResolveBuiltInNumberFormatCode(numberFormatId);
        }

        private static string? ResolveBuiltInNumberFormatCode(uint numberFormatId) {
            return numberFormatId switch {
                1U => "0",
                2U => "0.00",
                3U => "#,##0",
                4U => "#,##0.00",
                9U => "0%",
                10U => "0.00%",
                11U => "0.00E+00",
                12U => "# ?/?",
                13U => "# ??/??",
                14U => "mm-dd-yy",
                15U => "d-mmm-yy",
                16U => "d-mmm",
                17U => "mmm-yy",
                18U => "h:mm AM/PM",
                19U => "h:mm:ss AM/PM",
                20U => "h:mm",
                21U => "h:mm:ss",
                22U => "m/d/yy h:mm",
                37U => "#,##0 ;(#,##0)",
                38U => "#,##0 ;[Red](#,##0)",
                39U => "#,##0.00;(#,##0.00)",
                40U => "#,##0.00;[Red](#,##0.00)",
                45U => "mm:ss",
                46U => "[h]:mm:ss",
                47U => "mmss.0",
                48U => "##0.0E+0",
                49U => "@",
                _ => null
            };
        }

        private static bool TryGetCellCoordinates(Cell cell, int fallbackRow, int fallbackColumn, out int row, out int column) {
            row = 0;
            column = 0;
            string? reference = cell.CellReference?.Value;
            if (!string.IsNullOrWhiteSpace(reference)) {
                try {
                    (row, column) = A1.ParseCellRef(reference!);
                    return row > 0 && column > 0;
                } catch {
                    return false;
                }
            }

            row = fallbackRow;
            column = fallbackColumn;
            return row > 0 && column > 0;
        }

        private static object? ReadDirectOverlayCellValue(ExcelSheet sheet, Cell cell) {
            if (cell.CellFormula != null) {
                return new DirectFormulaCellValue(cell.CellFormula.Text ?? string.Empty, cell.CellFormula.OuterXml, cell.CellValue?.Text);
            }

            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;
            if (dataType == CellValues.Boolean) {
                return string.Equals(text, "1", StringComparison.Ordinal)
                       || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
            }

            if (dataType == null || dataType == CellValues.Number) {
                if (!string.IsNullOrWhiteSpace(text)) {
                    return new DirectTypedCellValue(cell.DataType?.InnerText ?? "n", text);
                }

                return text;
            }

            if (dataType == CellValues.Error || dataType == CellValues.Date || dataType == CellValues.InlineString) {
                string dataTypeText = cell.DataType?.InnerText
                                      ?? (dataType == CellValues.Error
                                          ? "e"
                                          : dataType == CellValues.Date
                                              ? "d"
                                              : "inlineStr");
                return new DirectTypedCellValue(dataTypeText, text, cell.InlineString?.OuterXml);
            }

            return sheet.GetCellText(cell);
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
            } catch (OperationCanceledException) {
                throw;
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

            internal DirectDataSetWorkbookModel WithWorksheetMetadata(IReadOnlyList<DirectWorksheetMetadata?> metadata) {
                if (metadata == null) throw new ArgumentNullException(nameof(metadata));
                if (metadata.Count != Sheets.Count) {
                    throw new ArgumentException("Metadata count must match the sheet count.", nameof(metadata));
                }

                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    sheets[i] = Sheets[i].WithMetadata(metadata[i]);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy);
            }

            internal DirectDataSetWorkbookModel WithAutoFitColumns(
                string sheetName,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    double[] columnWidths = sheet.Table.CalculateColumnWidths(sheet.IncludeHeaders, dateTimeOffsetWriteStrategy, ct);
                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        sheet.IncludeAutoFilter,
                        sheet.HasTable,
                        autoFitColumns: true,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy);
            }

            internal DirectDataSetWorkbookModel WithTableAutoFilter(string sheetName, bool includeAutoFilter) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        includeAutoFilter,
                        sheet.HasTable,
                        sheet.AutoFitColumns,
                        sheet.OmitBlankCells,
                        sheet.ColumnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy);
            }

            internal DirectDataSetWorkbookModel WithAutoFitColumns(
                string sheetName,
                IReadOnlyList<int> columnIndexes,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        continue;
                    }

                    double[] columnWidths = sheet.Table.CalculateColumnWidths(sheet.IncludeHeaders, dateTimeOffsetWriteStrategy, ct, columnIndexes);
                    if (sheet.ColumnWidths != null && sheet.ColumnWidths.Length == columnWidths.Length) {
                        var mergedWidths = new double[columnWidths.Length];
                        Array.Copy(sheet.ColumnWidths, mergedWidths, mergedWidths.Length);
                        for (int columnIndex = 0; columnIndex < columnIndexes.Count; columnIndex++) {
                            int widthIndex = columnIndexes[columnIndex] - 1;
                            if (widthIndex >= 0 && widthIndex < mergedWidths.Length) {
                                mergedWidths[widthIndex] = columnWidths[widthIndex];
                            }
                        }

                        columnWidths = mergedWidths;
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        sheet.TableName,
                        sheet.Range,
                        sheet.Table,
                        sheet.TableStyle,
                        sheet.IncludeHeaders,
                        sheet.IncludeAutoFilter,
                        sheet.HasTable,
                        autoFitColumns: false,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats);
                }

                return new DirectDataSetWorkbookModel(sheets, Results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy);
            }

            internal DirectDataSetWorkbookModel WithTable(
                string sheetName,
                string tableName,
                bool includeHeaders,
                TableStyle tableStyle,
                bool includeAutoFilter,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                var results = new ExcelDataSetImportResult[Sheets.Count];
                bool canCancel = ct.CanBeCanceled;
                for (int i = 0; i < Sheets.Count; i++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    var sheet = Sheets[i];
                    if (!string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)) {
                        sheets[i] = sheet;
                        results[i] = new ExcelDataSetImportResult(sheet.SheetName, sheet.TableName, sheet.Range, sheet.Table.RowCount, sheet.Table.ColumnCount);
                        continue;
                    }

                    var table = includeHeaders ? sheet.Table : sheet.Table.WithGeneratedColumnNames();
                    double[]? columnWidths = sheet.ColumnWidths;
                    if (sheet.AutoFitColumns && !ReferenceEquals(table, sheet.Table)) {
                        columnWidths = table.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct);
                    }

                    sheets[i] = new DirectDataSetSheetModel(
                        sheet.Index,
                        sheet.SheetName,
                        tableName,
                        sheet.Range,
                        table,
                        tableStyle,
                        includeHeaders,
                        includeAutoFilter,
                        hasTable: true,
                        sheet.AutoFitColumns,
                        sheet.OmitBlankCells,
                        columnWidths,
                        sheet.UseCellValueNumberFormats,
                        sheet.Metadata,
                        sheet.ColumnNumberFormats);
                    results[i] = new ExcelDataSetImportResult(sheet.SheetName, tableName, sheet.Range, table.RowCount, table.ColumnCount);
                }

                return new DirectDataSetWorkbookModel(sheets, results, dateTimeOffsetWriteStrategy ?? DateTimeOffsetWriteStrategy);
            }

            internal DirectDataSetWorkbookModel WithColumnNumberFormat(string sheetName, int columnIndex, string numberFormat) {
                var sheets = new DirectDataSetSheetModel[Sheets.Count];
                for (int i = 0; i < Sheets.Count; i++) {
                    var sheet = Sheets[i];
                    sheets[i] = string.Equals(sheet.SheetName, sheetName, StringComparison.Ordinal)
                        ? sheet.WithColumnNumberFormat(columnIndex, numberFormat)
                        : sheet;
                }

                return new DirectDataSetWorkbookModel(sheets, Results, DateTimeOffsetWriteStrategy);
            }


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
                bool snapshotTables = false,
                bool omitBlankCells = false) {
                var sheets = new List<DirectDataSetSheetModel>();
                var results = new List<ExcelDataSetImportResult>();
                var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var usedTableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int index = 1;
                bool canCancel = ct.CanBeCanceled;
                foreach (DataTable table in dataSet.Tables) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

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
                    var sheet = new DirectDataSetSheetModel(index, sheetName, hasTable ? tableName : null, range, tableModel, tableStyle, includeHeaders, includeAutoFilter, hasTable, autoFit, omitBlankCells, columnWidths);
                    sheets.Add(sheet);
                    results.Add(new ExcelDataSetImportResult(sheetName, hasTable ? tableName : null, range, tableModel.RowCount, tableModel.ColumnCount));
                    index++;
                }

                return new DirectDataSetWorkbookModel(sheets, results, dateTimeOffsetWriteStrategy ?? DefaultDateTimeOffsetWriteStrategy);
            }

            internal static DirectDataSetWorkbookModel CreateSingle(
                string sheetName,
                string requestedName,
                string? tableName,
                string range,
                DirectDataSetTableModel tableModel,
                bool createTable,
                TableStyle tableStyle,
                bool includeHeaders,
                bool includeAutoFilter,
                bool autoFit,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                bool useCellValueNumberFormats = false) {
                int rowCount = tableModel.RowCount + (includeHeaders ? 1 : 0);
                ValidateWorksheetBounds(tableModel, rowCount, requestedName);
                bool hasTable = createTable && range.Length > 0;
                string? resolvedTableName = hasTable
                    ? SanitizeTableName(string.IsNullOrWhiteSpace(tableName) ? requestedName : tableName!)
                    : null;
                double[]? columnWidths = autoFit && tableModel.ColumnCount > 0
                    ? tableModel.CalculateColumnWidths(includeHeaders, dateTimeOffsetWriteStrategy, ct)
                    : null;
                var sheet = new DirectDataSetSheetModel(1, sheetName, resolvedTableName, range, tableModel, tableStyle, includeHeaders, includeAutoFilter, hasTable, autoFit, omitBlankCells: false, columnWidths: columnWidths, useCellValueNumberFormats: useCellValueNumberFormats);
                var result = new ExcelDataSetImportResult(sheetName, resolvedTableName, range, tableModel.RowCount, tableModel.ColumnCount);
                return new DirectDataSetWorkbookModel([sheet], [result], dateTimeOffsetWriteStrategy ?? DefaultDateTimeOffsetWriteStrategy);
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
                bool omitBlankCells,
                double[]? columnWidths,
                bool useCellValueNumberFormats = false,
                DirectWorksheetMetadata? metadata = null,
                IReadOnlyList<string?>? columnNumberFormats = null) {
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
                OmitBlankCells = omitBlankCells;
                ColumnWidths = columnWidths;
                UseCellValueNumberFormats = useCellValueNumberFormats;
                Metadata = metadata;
                ColumnNumberFormats = columnNumberFormats;
            }

            internal DirectDataSetSheetModel WithMetadata(DirectWorksheetMetadata? metadata) {
                if (ReferenceEquals(Metadata, metadata)) {
                    return this;
                }

                return new DirectDataSetSheetModel(
                    Index,
                    SheetName,
                    TableName,
                    Range,
                    Table,
                    TableStyle,
                    IncludeHeaders,
                    IncludeAutoFilter,
                    HasTable,
                    AutoFitColumns,
                    OmitBlankCells,
                    ColumnWidths,
                    UseCellValueNumberFormats,
                    metadata,
                    ColumnNumberFormats);
            }

            internal DirectDataSetSheetModel WithColumnNumberFormat(int columnIndex, string numberFormat) {
                if (columnIndex <= 0 || columnIndex > Table.ColumnCount) {
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
                }

                string?[] formats;
                if (ColumnNumberFormats == null || ColumnNumberFormats.Count != Table.ColumnCount) {
                    formats = new string?[Table.ColumnCount];
                } else {
                    formats = new string?[ColumnNumberFormats.Count];
                    for (int i = 0; i < formats.Length; i++) {
                        formats[i] = ColumnNumberFormats[i];
                    }
                }

                formats[columnIndex - 1] = numberFormat;
                return new DirectDataSetSheetModel(
                    Index,
                    SheetName,
                    TableName,
                    Range,
                    Table,
                    TableStyle,
                    IncludeHeaders,
                    IncludeAutoFilter,
                    HasTable,
                    AutoFitColumns,
                    OmitBlankCells,
                    ColumnWidths,
                    UseCellValueNumberFormats,
                    Metadata,
                    formats);
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

            internal bool OmitBlankCells { get; }

            internal double[]? ColumnWidths { get; }

            internal bool UseCellValueNumberFormats { get; }

            internal DirectWorksheetMetadata? Metadata { get; }

            internal IReadOnlyList<string?>? ColumnNumberFormats { get; }
        }

        private sealed class DirectWorksheetMetadata {
            internal static readonly DirectWorksheetMetadata Empty = new(
                null,
                null,
                null,
                null,
                Array.Empty<string>(),
                null,
                null,
                Array.Empty<string>(),
                Array.Empty<DirectOverlayCell>());

            internal DirectWorksheetMetadata(
                string? sheetPropertiesXml,
                string? sheetViewsXml,
                string? sheetFormatPropertiesXml,
                string? autoFilterXml,
                IReadOnlyList<string> conditionalFormattingXml,
                string? dataValidationsXml,
                string? drawingXml,
                IReadOnlyList<string> postDataValidationXml,
                IReadOnlyList<DirectOverlayCell> overlayCells) {
                SheetPropertiesXml = sheetPropertiesXml;
                SheetViewsXml = sheetViewsXml;
                SheetFormatPropertiesXml = sheetFormatPropertiesXml;
                AutoFilterXml = autoFilterXml;
                ConditionalFormattingXml = conditionalFormattingXml ?? Array.Empty<string>();
                DataValidationsXml = dataValidationsXml;
                DrawingXml = drawingXml;
                PostDataValidationXml = postDataValidationXml ?? Array.Empty<string>();
                OverlayCells = overlayCells ?? Array.Empty<DirectOverlayCell>();
            }

            internal DirectWorksheetMetadata WithSheetViewsXml(string? sheetViewsXml) {
                if (string.Equals(SheetViewsXml, sheetViewsXml, StringComparison.Ordinal)) {
                    return this;
                }

                return new DirectWorksheetMetadata(
                    SheetPropertiesXml,
                    sheetViewsXml,
                    SheetFormatPropertiesXml,
                    AutoFilterXml,
                    ConditionalFormattingXml,
                    DataValidationsXml,
                    DrawingXml,
                    PostDataValidationXml,
                    OverlayCells);
            }

            internal DirectWorksheetMetadata WithAutoFilterXml(string? autoFilterXml) {
                if (string.Equals(AutoFilterXml, autoFilterXml, StringComparison.Ordinal)) {
                    return this;
                }

                return new DirectWorksheetMetadata(
                    SheetPropertiesXml,
                    SheetViewsXml,
                    SheetFormatPropertiesXml,
                    autoFilterXml,
                    ConditionalFormattingXml,
                    DataValidationsXml,
                    DrawingXml,
                    PostDataValidationXml,
                    OverlayCells);
            }

            internal string? SheetPropertiesXml { get; }

            internal string? SheetViewsXml { get; }

            internal string? SheetFormatPropertiesXml { get; }

            internal string? AutoFilterXml { get; }

            internal IReadOnlyList<string> ConditionalFormattingXml { get; }

            internal string? DataValidationsXml { get; }

            internal string? DrawingXml { get; }

            internal IReadOnlyList<string> PostDataValidationXml { get; }

            internal IReadOnlyList<DirectOverlayCell> OverlayCells { get; }

            internal bool IsEmpty
                => SheetPropertiesXml == null
                   && SheetViewsXml == null
                   && SheetFormatPropertiesXml == null
                && AutoFilterXml == null
                && ConditionalFormattingXml.Count == 0
                && DataValidationsXml == null
                && DrawingXml == null
                && PostDataValidationXml.Count == 0
                && OverlayCells.Count == 0;
        }

        private readonly struct DirectOverlayCell {
            internal DirectOverlayCell(int row, int column, object? value, uint? styleIndex, string? numberFormat, bool isDeleted = false) {
                Row = row;
                Column = column;
                Value = value;
                StyleIndex = styleIndex;
                NumberFormat = numberFormat;
                IsDeleted = isDeleted;
            }

            internal int Row { get; }

            internal int Column { get; }

            internal object? Value { get; }

            internal uint? StyleIndex { get; }

            internal string? NumberFormat { get; }

            internal bool IsDeleted { get; }
        }

        private readonly struct DirectBufferedRows {
            private readonly object?[][]? _arrayRows;
            private readonly List<object?[]>? _listRows;

            internal DirectBufferedRows(object?[][] rows) {
                _arrayRows = rows;
                _listRows = null;
            }

            internal DirectBufferedRows(List<object?[]> rows) {
                _arrayRows = null;
                _listRows = rows;
            }

            internal int Count => _arrayRows?.Length ?? _listRows!.Count;

            internal object?[] this[int index] => _arrayRows != null
                ? _arrayRows[index]
                : _listRows![index];
        }

        private readonly struct DirectCellValueRows {
            internal DirectCellValueRows(object?[] values, int columnCount, int rowCount) {
                Values = values;
                ColumnCount = columnCount;
                Count = rowCount;
            }

            internal object?[] Values { get; }

            internal int ColumnCount { get; }

            internal int Count { get; }

            internal int GetRowOffset(int rowIndex) => rowIndex * ColumnCount;

            internal object? GetValue(int rowIndex, int columnIndex) {
                int index = GetRowOffset(rowIndex) + columnIndex;
                return Values[index];
            }
        }

        private sealed class DirectDataSetTableModel : IExcelSheetTabularRowSource {
            private const int MaxAutoFitStringWidthCacheEntriesPerColumn = 1024;
            private const long BufferedDictionaryCellLimit = 500_000;

            private enum AutoFitWidthKind {
                Object,
                String,
                Boolean,
                DateTime,
                DateTimeOffset,
                TimeSpan,
                Double,
                Float,
                Decimal,
                SByte,
                Byte,
                Int16,
                UInt16,
                Int32,
                UInt32,
                Int64,
                UInt64,
#if NET6_0_OR_GREATER
                DateOnly,
                TimeOnly,
#endif
            }

            private readonly DataTable? _sourceTable;
            private readonly DirectDataSetColumnModel[]? _columns;
            private readonly int[]? _stringCandidateColumnIndexes;
            private readonly object?[][]? _arrayRows;
            private readonly List<object?[]>? _listRows;
            private readonly DirectCellValueRows _cellValueRows;
            private readonly bool _hasCellValueRows;
            private readonly IReadOnlyList<Dictionary<string, object?>>? _exactDictionaryRows;
            private readonly IReadOnlyList<IReadOnlyDictionary<string, object?>>? _dictionaryRows;
            private readonly IReadOnlyList<System.Collections.IDictionary>? _legacyDictionaryRows;
            private readonly bool _legacyDictionaryExactKeyLookup;
            private string[]? _columnNameArray;

            private DirectDataSetTableModel(DataTable sourceTable) {
                _sourceTable = sourceTable;
                _columns = CreateColumns(sourceTable);
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(_columns);
            }

            private DirectDataSetTableModel(DataTable sourceTable, DirectDataSetColumnModel[] columns) {
                _sourceTable = sourceTable;
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<object?[]> rows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                if (rows is object?[][] arrayRows) {
                    _arrayRows = arrayRows;
                } else if (rows is List<object?[]> listRows) {
                    _listRows = listRows;
                } else {
                    var copiedRows = new object?[rows.Count][];
                    for (int i = 0; i < copiedRows.Length; i++) {
                        copiedRows[i] = rows[i];
                    }

                    _arrayRows = copiedRows;
                }
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, DirectCellValueRows cellValueRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _cellValueRows = cellValueRows;
                _hasCellValueRows = true;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<IReadOnlyDictionary<string, object?>> dictionaryRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _dictionaryRows = dictionaryRows;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<Dictionary<string, object?>> exactDictionaryRows) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _exactDictionaryRows = exactDictionaryRows;
            }

            private DirectDataSetTableModel(DirectDataSetColumnModel[] columns, IReadOnlyList<System.Collections.IDictionary> legacyDictionaryRows, bool exactKeyLookup = false) {
                _columns = columns;
                _stringCandidateColumnIndexes = CreateStringCandidateColumnIndexes(columns);
                _legacyDictionaryRows = legacyDictionaryRows;
                _legacyDictionaryExactKeyLookup = exactKeyLookup;
            }

            internal static DirectDataSetTableModel Reference(DataTable table) => new DirectDataSetTableModel(table);

            internal static DirectDataSetTableModel FromRows(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<object?[]> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            internal static DirectDataSetTableModel FromCellValues(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<Type> columnTypes,
                object?[] values,
                int columnCount,
                int rowCount) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                if (columnNames.Count != columnCount) {
                    throw new ArgumentException("Column count must match the column metadata.", nameof(columnCount));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, new DirectCellValueRows(values, columnCount, rowCount));
            }

            internal static DirectDataSetTableModel FromLegacyDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<System.Collections.IDictionary> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                return new DirectDataSetTableModel(columns, rows, HasCaseInsensitiveDuplicateColumnNames(columnNames));
            }

            internal static DirectDataSetTableModel FromExactDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<Dictionary<string, object?>> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                if (ShouldBufferDictionaryRows(rows.Count, columns.Length)) {
                    return new DirectDataSetTableModel(columns, SnapshotExactDictionaryRows(columnNames, rows));
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            internal static DirectDataSetTableModel FromDictionaries(IReadOnlyList<string> columnNames, IReadOnlyList<Type> columnTypes, IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                if (columnNames.Count != columnTypes.Count) {
                    throw new ArgumentException("Column name and type counts must match.", nameof(columnTypes));
                }

                var columns = new DirectDataSetColumnModel[columnNames.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(columnNames[i], columnTypes[i]);
                }

                if (ShouldBufferDictionaryRows(rows.Count, columns.Length)) {
                    return new DirectDataSetTableModel(columns, SnapshotDictionaryRows(columnNames, rows));
                }

                return new DirectDataSetTableModel(columns, rows);
            }

            private static bool ShouldBufferDictionaryRows(int rowCount, int columnCount)
                => rowCount > 0
                    && columnCount > 0
                    && (long)rowCount * columnCount <= BufferedDictionaryCellLimit;

            private static object?[][] SnapshotExactDictionaryRows(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<Dictionary<string, object?>> rows) {
                var bufferedRows = new object?[rows.Count][];
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    Dictionary<string, object?> source = rows[rowIndex];
                    var values = new object?[columnNames.Count];
                    for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++) {
                        values[columnIndex] = source.TryGetValue(columnNames[columnIndex], out object? value)
                            ? value
                            : null;
                    }

                    bufferedRows[rowIndex] = values;
                }

                return bufferedRows;
            }

            private static object?[][] SnapshotDictionaryRows(
                IReadOnlyList<string> columnNames,
                IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                var bufferedRows = new object?[rows.Count][];
                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                    IReadOnlyDictionary<string, object?> source = rows[rowIndex];
                    var values = new object?[columnNames.Count];
                    for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++) {
                        values[columnIndex] = source.TryGetValue(columnNames[columnIndex], out object? value)
                            ? value
                            : null;
                    }

                    bufferedRows[rowIndex] = values;
                }

                return bufferedRows;
            }

            internal DirectDataSetTableModel WithGeneratedColumnNames() {
                var columns = new DirectDataSetColumnModel[ColumnCount];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel("Column" + (i + 1).ToString(CultureInfo.InvariantCulture), GetColumnType(i));
                }

                if (_sourceTable != null) {
                    return new DirectDataSetTableModel(_sourceTable, columns);
                }

                if (_exactDictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _exactDictionaryRows);
                }

                if (_dictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _dictionaryRows);
                }

                if (_legacyDictionaryRows != null) {
                    return new DirectDataSetTableModel(columns, _legacyDictionaryRows, _legacyDictionaryExactKeyLookup);
                }

                if (_hasCellValueRows) {
                    return new DirectDataSetTableModel(columns, _cellValueRows);
                }

                return new DirectDataSetTableModel(columns, GetBufferedRowsForReuse());
            }

            internal static DirectDataSetTableModel Snapshot(DataTable table, CancellationToken ct) {
                var columns = CreateColumns(table);
                if (columns.Length == 8) {
                    return new DirectDataSetTableModel(columns, SnapshotEightColumnRows(table, ct));
                }

                var rows = new object?[table.Rows.Count][];
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

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

            private static object?[][] SnapshotEightColumnRows(DataTable table, CancellationToken ct) {
                var rows = new object?[table.Rows.Count][];
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    DataRow row = table.Rows[rowIndex];
                    object? value0 = row[0];
                    object? value1 = row[1];
                    object? value2 = row[2];
                    object? value3 = row[3];
                    object? value4 = row[4];
                    object? value5 = row[5];
                    object? value6 = row[6];
                    object? value7 = row[7];
                    rows[rowIndex] = new object?[] {
                        value0 == DBNull.Value ? null : value0,
                        value1 == DBNull.Value ? null : value1,
                        value2 == DBNull.Value ? null : value2,
                        value3 == DBNull.Value ? null : value3,
                        value4 == DBNull.Value ? null : value4,
                        value5 == DBNull.Value ? null : value5,
                        value6 == DBNull.Value ? null : value6,
                        value7 == DBNull.Value ? null : value7
                    };
                }

                return rows;
            }

            private static DirectDataSetColumnModel[] CreateColumns(DataTable table) {
                var columns = new DirectDataSetColumnModel[table.Columns.Count];
                for (int i = 0; i < columns.Length; i++) {
                    columns[i] = new DirectDataSetColumnModel(table.Columns[i].ColumnName, table.Columns[i].DataType);
                }

                return columns;
            }

            internal int ColumnCount => _columns!.Length;

            internal int RowCount => _sourceTable?.Rows.Count ?? _arrayRows?.Length ?? _listRows?.Count ?? (_hasCellValueRows ? _cellValueRows.Count : (int?)null) ?? _exactDictionaryRows?.Count ?? _dictionaryRows?.Count ?? _legacyDictionaryRows!.Count;

            internal string GetColumnName(int index) => _columns![index].Name;

            internal Type GetColumnType(int index) => _columns![index].DataType;

            int IExcelSheetTabularRowSource.ColumnCount => ColumnCount;

            int IExcelSheetTabularRowSource.RowCount => RowCount;

            string IExcelSheetTabularRowSource.GetColumnName(int index) => GetColumnName(index);

            Type IExcelSheetTabularRowSource.GetColumnType(int index) => GetColumnType(index);

            object? IExcelSheetTabularRowSource.GetValue(int rowIndex, int columnIndex) => GetValue(rowIndex, columnIndex);

            bool IExcelSheetTabularRowSource.TryGetBufferedRow(int rowIndex, out object?[]? values) {
                values = GetBufferedRow(rowIndex);
                return values != null;
            }

            bool IExcelSheetTabularRowSource.TryGetFlatValues(out object?[] values, out int columnCount) {
                if (_hasCellValueRows) {
                    values = _cellValueRows.Values;
                    columnCount = _cellValueRows.ColumnCount;
                    return true;
                }

                values = Array.Empty<object?>();
                columnCount = 0;
                return false;
            }

            internal string[] CreateColumnNameArray() {
                if (_columnNameArray != null) {
                    return _columnNameArray;
                }

                var columnNames = new string[_columns!.Length];
                for (int i = 0; i < columnNames.Length; i++) {
                    columnNames[i] = _columns[i].Name;
                }

                _columnNameArray = columnNames;
                return columnNames;
            }

            private static bool HasCaseInsensitiveDuplicateColumnNames(IReadOnlyList<string> columnNames) {
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < columnNames.Count; i++) {
                    if (!seen.Add(columnNames[i] ?? string.Empty)) {
                        return true;
                    }
                }

                return false;
            }

            internal int[]? GetStringCandidateColumnIndexes() => _stringCandidateColumnIndexes;

            private static int[]? CreateStringCandidateColumnIndexes(DirectDataSetColumnModel[] columns) {
                int[]? indexes = null;
                int count = 0;
                for (int i = 0; i < columns.Length; i++) {
                    Type dataType = columns[i].DataType;
                    if (dataType != typeof(string) && dataType != typeof(object)) {
                        continue;
                    }

                    indexes ??= new int[columns.Length];
                    indexes[count++] = i;
                }

                if (indexes == null) {
                    return null;
                }

                if (count == indexes.Length) {
                    return indexes;
                }

                Array.Resize(ref indexes, count);
                return indexes;
            }

            internal DataRow? GetSourceRow(int rowIndex) => _sourceTable?.Rows[rowIndex];

            internal bool HasSourceRows => _sourceTable != null;

            internal bool TryGetExactDictionaryRows(out IReadOnlyList<Dictionary<string, object?>> rows) {
                if (_exactDictionaryRows != null) {
                    rows = _exactDictionaryRows;
                    return true;
                }

                rows = Array.Empty<Dictionary<string, object?>>();
                return false;
            }

            internal bool TryGetDictionaryRows(out IReadOnlyList<IReadOnlyDictionary<string, object?>> rows) {
                if (_dictionaryRows != null) {
                    rows = _dictionaryRows;
                    return true;
                }

                rows = Array.Empty<IReadOnlyDictionary<string, object?>>();
                return false;
            }

            internal bool TryGetLegacyDictionaryRows(out IReadOnlyList<System.Collections.IDictionary> rows) {
                if (_legacyDictionaryRows != null) {
                    rows = _legacyDictionaryRows;
                    return true;
                }

                rows = Array.Empty<System.Collections.IDictionary>();
                return false;
            }

            internal bool TryGetBufferedRows(out DirectBufferedRows rows) {
                if (_arrayRows != null) {
                    rows = new DirectBufferedRows(_arrayRows);
                    return true;
                }

                if (_listRows != null) {
                    rows = new DirectBufferedRows(_listRows);
                    return true;
                }

                rows = default;
                return false;
            }

            internal bool TryGetCellValueRows(out DirectCellValueRows rows) {
                if (_hasCellValueRows) {
                    rows = _cellValueRows;
                    return true;
                }

                rows = default;
                return false;
            }

            internal object?[]? GetBufferedRow(int rowIndex) {
                if (_arrayRows != null) {
                    return _arrayRows[rowIndex];
                }

                return _listRows?[rowIndex];
            }

            internal object? GetValue(int rowIndex, int columnIndex) {
                object? value;
                if (_sourceTable != null) {
                    value = _sourceTable.Rows[rowIndex][columnIndex];
                } else if (_exactDictionaryRows != null) {
                    value = _exactDictionaryRows[rowIndex].TryGetValue(GetColumnName(columnIndex), out object? dictionaryValue)
                        ? dictionaryValue
                        : null;
                } else if (_dictionaryRows != null) {
                    value = _dictionaryRows[rowIndex].TryGetValue(GetColumnName(columnIndex), out object? dictionaryValue)
                        ? dictionaryValue
                        : null;
                } else if (_legacyDictionaryRows != null) {
                    value = GetLegacyDictionaryValue(_legacyDictionaryRows[rowIndex], GetColumnName(columnIndex), _legacyDictionaryExactKeyLookup);
                } else if (_hasCellValueRows) {
                    value = _cellValueRows.GetValue(rowIndex, columnIndex);
                } else {
                    value = GetBufferedRow(rowIndex)![columnIndex];
                }

                return value == DBNull.Value ? null : value;
            }

            internal static object? GetLegacyDictionaryValue(System.Collections.IDictionary dictionary, string column, bool exactKeyLookup = false) {
                if (dictionary.Contains(column)) {
                    return dictionary[column];
                }

                if (exactKeyLookup) {
                    return null;
                }

                foreach (System.Collections.DictionaryEntry entry in dictionary) {
                    string? key = entry.Key?.ToString();
                    if (string.Equals(key, column, StringComparison.OrdinalIgnoreCase)) {
                        return entry.Value;
                    }
                }

                return null;
            }

            internal double[] CalculateColumnWidths(
                bool includeHeaders,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                IReadOnlyList<int>? columnIndexes = null) {
                int columnCount = ColumnCount;
                var widths = new double[columnCount];
                if (columnCount == 0) {
                    return widths;
                }

                int[]? selectedColumnIndexes = CreateAutoFitSelectedColumnIndexes(columnCount, columnIndexes);
                if (selectedColumnIndexes != null && selectedColumnIndexes.Length == 0) {
                    return widths;
                }

                if (includeHeaders) {
                    if (selectedColumnIndexes == null) {
                        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                            widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(GetColumnName(columnIndex)));
                        }
                    } else {
                        for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                            int columnIndex = selectedColumnIndexes[i];
                            widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(GetColumnName(columnIndex)));
                        }
                    }
                }

                AutoFitWidthKind[] widthKinds = CreateAutoFitWidthKinds();
                Dictionary<string, double>?[]? stringWidthCaches = null;
                int rowCount = RowCount;
                DataRowCollection? sourceRows = _sourceTable?.Rows;
                bool canCancel = ct.CanBeCanceled;
                if (sourceRows != null) {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        DataRow sourceRow = sourceRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = sourceRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = sourceRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (TryGetBufferedRows(out DirectBufferedRows bufferedRows)) {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        object?[] bufferedRow = bufferedRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = bufferedRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = bufferedRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_exactDictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        Dictionary<string, object?> row = _exactDictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_dictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        IReadOnlyDictionary<string, object?> row = _dictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_legacyDictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        System.Collections.IDictionary row = _legacyDictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                }

                return widths;
            }

            private static int[]? CreateAutoFitSelectedColumnIndexes(int columnCount, IReadOnlyList<int>? columnIndexes) {
                if (columnIndexes == null) {
                    return null;
                }

                var selected = new int[columnIndexes.Count];
                int selectedCount = 0;
                bool allColumnsInOrder = columnIndexes.Count == columnCount;
                for (int i = 0; i < columnIndexes.Count; i++) {
                    int columnIndex = columnIndexes[i];
                    if (columnIndex <= 0 || columnIndex > columnCount) {
                        continue;
                    }

                    allColumnsInOrder &= columnIndex == i + 1;
                    selected[selectedCount++] = columnIndex - 1;
                }

                if (selectedCount == 0) {
                    return Array.Empty<int>();
                }

                if (allColumnsInOrder) {
                    return null;
                }

                if (selectedCount != selected.Length) {
                    Array.Resize(ref selected, selectedCount);
                }

                return selected;
            }

            private IReadOnlyList<object?[]> GetBufferedRowsForReuse() {
                if (_arrayRows != null) {
                    return _arrayRows;
                }

                return _listRows!;
            }

            private AutoFitWidthKind[] CreateAutoFitWidthKinds() {
                var kinds = new AutoFitWidthKind[_columns!.Length];
                for (int i = 0; i < kinds.Length; i++) {
                    kinds[i] = GetAutoFitWidthKind(_columns[i].DataType);
                }

                return kinds;
            }

            private static AutoFitWidthKind GetAutoFitWidthKind(Type dataType) {
                if (dataType == typeof(string)) return AutoFitWidthKind.String;
                if (dataType == typeof(bool)) return AutoFitWidthKind.Boolean;
                if (dataType == typeof(DateTime)) return AutoFitWidthKind.DateTime;
                if (dataType == typeof(DateTimeOffset)) return AutoFitWidthKind.DateTimeOffset;
                if (dataType == typeof(TimeSpan)) return AutoFitWidthKind.TimeSpan;
                if (dataType == typeof(double)) return AutoFitWidthKind.Double;
                if (dataType == typeof(float)) return AutoFitWidthKind.Float;
                if (dataType == typeof(decimal)) return AutoFitWidthKind.Decimal;
                if (dataType == typeof(sbyte)) return AutoFitWidthKind.SByte;
                if (dataType == typeof(byte)) return AutoFitWidthKind.Byte;
                if (dataType == typeof(short)) return AutoFitWidthKind.Int16;
                if (dataType == typeof(ushort)) return AutoFitWidthKind.UInt16;
                if (dataType == typeof(int)) return AutoFitWidthKind.Int32;
                if (dataType == typeof(uint)) return AutoFitWidthKind.UInt32;
                if (dataType == typeof(long)) return AutoFitWidthKind.Int64;
                if (dataType == typeof(ulong)) return AutoFitWidthKind.UInt64;
#if NET6_0_OR_GREATER
                if (dataType == typeof(DateOnly)) return AutoFitWidthKind.DateOnly;
                if (dataType == typeof(TimeOnly)) return AutoFitWidthKind.TimeOnly;
#endif
                return AutoFitWidthKind.Object;
            }

            private static double EstimateAutoFitWidth(string text) {
                if (string.IsNullOrEmpty(text)) {
                    return 0D;
                }

                if (text.IndexOf('\r') < 0 && text.IndexOf('\n') < 0) {
                    return EstimateAutoFitWidthFromLength(text.Length);
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

            private static double EstimateAutoFitWidth(object? value, AutoFitWidthKind widthKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                if (value == null || value == DBNull.Value) {
                    return 0D;
                }

                switch (widthKind) {
                    case AutoFitWidthKind.String:
                        return value is string stringValue ? EstimateAutoFitWidth(stringValue) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Boolean:
                        return value is bool boolValue ? EstimateAutoFitWidthFromLength(boolValue ? 4 : 5) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.DateTime:
                        return value is DateTime ? EstimateAutoFitWidthFromLength(16) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.DateTimeOffset:
                        return value is DateTimeOffset dateTimeOffsetValue ? EstimateDateTimeOffsetAutoFitWidth(dateTimeOffsetValue, dateTimeOffsetWriteStrategy) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.TimeSpan:
                        return value is TimeSpan timeSpanValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(timeSpanValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Double:
                        return value is double doubleValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(doubleValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Float:
                        return value is float floatValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(floatValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Decimal:
                        return value is decimal decimalValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(decimalValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.SByte:
                        return value is sbyte sbyteValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(sbyteValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Byte:
                        return value is byte byteValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(byteValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int16:
                        return value is short shortValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(shortValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt16:
                        return value is ushort ushortValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ushortValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int32:
                        return value is int intValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(intValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt32:
                        return value is uint uintValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(uintValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int64:
                        return value is long longValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(longValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt64:
                        return value is ulong ulongValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ulongValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
#if NET6_0_OR_GREATER
                    case AutoFitWidthKind.DateOnly:
                        return value is DateOnly ? EstimateAutoFitWidthFromLength(10) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.TimeOnly:
                        return value is TimeOnly ? EstimateAutoFitWidthFromLength(8) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
#endif
                    default:
                        return EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                }
            }

            private static double EstimateAutoFitWidth(
                object? value,
                AutoFitWidthKind widthKind,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ref Dictionary<string, double>?[]? stringWidthCaches,
                int columnIndex,
                int columnCount) {
                if (value is string stringValue && stringValue.Length > 0) {
                    stringWidthCaches ??= new Dictionary<string, double>?[columnCount];
                    Dictionary<string, double>? cache = stringWidthCaches[columnIndex];
                    if (cache == null) {
                        cache = new Dictionary<string, double>(StringComparer.Ordinal);
                        stringWidthCaches[columnIndex] = cache;
                    }

                    if (cache.TryGetValue(stringValue, out double cachedWidth)) {
                        return cachedWidth;
                    }

                    double width = EstimateAutoFitWidth(value, widthKind, dateTimeOffsetWriteStrategy);
                    if (cache.Count < MaxAutoFitStringWidthCacheEntriesPerColumn) {
                        cache[stringValue] = width;
                    }

                    return width;
                }

                return EstimateAutoFitWidth(value, widthKind, dateTimeOffsetWriteStrategy);
            }

            private static double EstimateDateTimeOffsetAutoFitWidth(DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                try {
                    _ = dateTimeOffsetWriteStrategy(value);
                    return EstimateAutoFitWidthFromLength(16);
                } catch (ArgumentException) {
                    return EstimateAutoFitWidth(value.ToString("o", CultureInfo.InvariantCulture));
                } catch (OverflowException) {
                    return EstimateAutoFitWidth(value.ToString("o", CultureInfo.InvariantCulture));
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
                    int rowCount = RowCount;
                    int columnCount = ColumnCount;
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        var values = new object?[columnCount];
                        for (int i = 0; i < values.Length; i++) {
                            values[i] = GetValue(rowIndex, i) ?? DBNull.Value;
                        }

                        table.Rows.Add(values);
                    }
                } finally {
                    table.EndLoadData();
                }

                return table;
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
            private readonly bool _subscribed;
            private bool _disposed;

            internal DirectDataSetSaveCandidate(DataSet dataSet, DirectDataSetWorkbookModel model, Action invalidate, bool isDeferred, bool subscribeToSourceChanges) {
                _dataSet = dataSet;
                Model = model;
                _invalidate = invalidate;
                IsDeferred = isDeferred;
                if (subscribeToSourceChanges) {
                    _subscribed = true;
                    Subscribe(dataSet);
                }
            }

            internal DirectDataSetWorkbookModel Model { get; }

            internal DataSet Owner => _dataSet;

            internal Action InvalidateCallback => _invalidate;

            internal bool IsDeferred { get; }

            internal bool SubscribesToSourceChanges => _subscribed;

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
                if (_subscribed) {
                    try {
                        Unsubscribe(_dataSet);
                    } catch {
                    }
                }
            }
        }
    }
}
