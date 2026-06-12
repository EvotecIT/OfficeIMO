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

                baseMetadata = MergeDirectWorksheetMetadata(baseMetadata, capturedMetadata, replaceOverlayCells: true);
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
                    if (_materializedDirectDataSetFastSaveModel != null) {
                        _materializedDirectDataSetFastSaveModel = _materializedDirectDataSetFastSaveModel.WithColumnNumberFormat(sheet.Name, columnIndex, numberFormat);
                        _preserveMaterializedDirectDataSetFastSaveModelForNextDirtyMark = true;
                    }
                } else {
                    _materializedDirectDataSetFastSaveModel = model;
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

        /// <summary>
        /// Returns the pending used range for a clean direct-tabular save candidate before the package is materialized.
        /// </summary>
        /// <param name="sheet">Worksheet whose pending direct-tabular range should be reported.</param>
        /// <param name="range">The candidate worksheet range in A1 notation when available.</param>
        /// <returns><c>true</c> when the workbook has a valid single-sheet direct-tabular candidate for the worksheet.</returns>
        internal bool TryGetDirectTabularSaveCandidateRange(ExcelSheet sheet, out string range) {
            range = string.Empty;
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
                || string.IsNullOrWhiteSpace(sheetModel.Range)) {
                return false;
            }

            range = sheetModel.Range;
            return true;
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
    }
}
