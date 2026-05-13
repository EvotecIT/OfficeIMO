using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (startRow < 1) throw new ArgumentOutOfRangeException(nameof(startRow));
            if (startColumn < 1) throw new ArgumentOutOfRangeException(nameof(startColumn));

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
                    string? fmt = null;
                    var t = col.DataType;
                    if (t == typeof(DateTime) || t == typeof(DateTimeOffset)) {
                        // General purpose date-time format; users can restyle later
                        fmt = "yyyy-mm-dd hh:mm";
                    } else if (t == typeof(TimeSpan)) {
                        fmt = "[h]:mm:ss";
                    }

                    if (fmt is null && value is TimeSpan)
                        fmt = "[h]:mm:ss";
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
            InsertDataTable(table, startRow, startColumn, includeHeaders, mode, ct);

            int rowsCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            int colsCount = Math.Max(1, table.Columns.Count);
            string startRef = A1.CellReference(startRow, startColumn);
            string endRef = A1.CellReference(startRow + rowsCount - 1, startColumn + colsCount - 1);
            string range = startRef + ":" + endRef;

            // Create the Table with optional AutoFilter and style
            AddTable(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter);
            return range;
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

            if (table.TotalsRowShown?.Value == true) {
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
            bool hasSyntheticHeaders = HasSyntheticTableColumnNames(tableColumnNames);
            DataTable appendTable = BuildAppendDataTable(dataTable, tableColumnNames, matchColumnsByHeader && hasHeaderRow && !hasSyntheticHeaders);
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

        private static bool HasSyntheticTableColumnNames(IReadOnlyList<string> tableColumnNames) {
            if (tableColumnNames.Count == 0) {
                return false;
            }

            for (int i = 0; i < tableColumnNames.Count; i++) {
                if (!string.Equals(tableColumnNames[i], $"Column{i + 1}", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }
            }

            return true;
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
