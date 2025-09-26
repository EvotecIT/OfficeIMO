using DocumentFormat.OpenXml;
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

                    for (int i = 0; i < prepared.Length; i++) {
                        var p = prepared[i];
                        var cell = GetCell(p.Row, p.Col);
                        cell.CellValue = p.Val;
                        cell.DataType = p.Type;
                        if (wrapFlags[i])
                            ApplyWrapText(cell);

                        var fmt = cells[i].NumFmt;
                        if (!string.IsNullOrWhiteSpace(fmt) && stylePlanner.TryGetCellFormatIndex(fmt, out uint idx)) {
                            cell.StyleIndex = idx;
                        }
                    }
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
            string startRef = GetColumnName(startColumn) + startRow.ToString(CultureInfo.InvariantCulture);
            string endRef = GetColumnName(startColumn + colsCount - 1) + (startRow + rowsCount - 1).ToString(CultureInfo.InvariantCulture);
            string range = startRef + ":" + endRef;

            // Create the Table with optional AutoFilter and style
            AddTable(range, includeHeaders, tableName ?? string.Empty, style, includeAutoFilter);
            return range;
        }
    }
}
