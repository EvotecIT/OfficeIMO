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
    }
}
