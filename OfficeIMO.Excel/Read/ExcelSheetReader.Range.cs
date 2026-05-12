using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        /// <summary>
        /// Returns the used range of the worksheet as an A1 string (e.g., "A1:C10").
        /// If the sheet is empty, returns "A1:A1".
        /// </summary>
        public string GetUsedRangeA1() {
            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            return reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
        }
        /// <summary>
        /// Reads a rectangular A1 range (e.g., "A1:C10") into a dense 2D array of typed values.
        /// </summary>
        public object?[,] ReadRange(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            var height = r2 - r1 + 1;
            var width = c2 - c1 + 1;
            var result = new object?[height, width];

            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = height * width;
            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                FillRangeSequential(result, r1, c1, r2, c2);
                return result;
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRange", mode, ct, workload);

            foreach (var cell in raw) {
                var rr = cell.Row - r1;
                var cc = cell.Col - c1;
                if ((uint)rr < (uint)height && (uint)cc < (uint)width)
                    result[rr, cc] = cell.TypedValue;
            }

            return result;
        }

        /// <summary>
        /// Reads a rectangular range to a DataTable. If headersInFirstRow = true, first row becomes column names.
        /// </summary>
        public DataTable ReadRangeAsDataTable(string a1Range, bool headersInFirstRow = true, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            var dt = new DataTable(_sheetName);
            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRangeAsDataTable", mode, ct, rows * cols);

            if (headersInFirstRow && rows > 0) {
                var headerValues = new object?[cols];
                foreach (var cell in raw) {
                    if (cell.Row != r1) continue;
                    int cc = cell.Col - c1;
                    if ((uint)cc < (uint)cols) {
                        headerValues[cc] = cell.TypedValue;
                    }
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) dt.Columns.Add($"Column{c + 1}", typeof(object));
            }

            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            var dataRows = new DataRow[dataRowCount];
            for (int r = 0; r < dataRowCount; r++) {
                dataRows[r] = dt.NewRow();
                dt.Rows.Add(dataRows[r]);
            }

            foreach (var cell in raw) {
                int rr = cell.Row - r1 - startRow;
                int cc = cell.Col - c1;
                if ((uint)rr < (uint)dataRowCount && (uint)cc < (uint)cols) {
                    dataRows[rr][cc] = cell.TypedValue ?? DBNull.Value;
                }
            }

            return dt;
        }

        /// <summary>
        /// Reads a rectangular range into a sequence of dictionaries using the first row as headers.
        /// </summary>
        public IEnumerable<Dictionary<string, object?>> ReadObjects(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (rows <= 1 || cols == 0) {
                return Array.Empty<Dictionary<string, object?>>();
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadObjects", mode, ct, rows * cols);

            var headerValues = new object?[cols];
            foreach (var cell in raw) {
                if (cell.Row != r1) continue;
                int cc = cell.Col - c1;
                if ((uint)cc < (uint)cols) {
                    headerValues[cc] = cell.TypedValue;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            int dataRowCount = rows - 1;
            var rowValues = new object?[dataRowCount][];

            foreach (var cell in raw) {
                if (cell.Row <= r1) continue;
                int rr = cell.Row - r1 - 1;
                int cc = cell.Col - c1;
                if ((uint)rr < (uint)dataRowCount && (uint)cc < (uint)cols) {
                    var values = rowValues[rr];
                    if (values == null) {
                        values = new object?[cols];
                        rowValues[rr] = values;
                    }

                    values[cc] = cell.TypedValue;
                }
            }

            var result = new List<Dictionary<string, object?>>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                var values = rowValues[r];
                var dict = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++) {
                    dict.Add(headers[c], values?[c]);
                }

                result.Add(dict);
            }

            return result;
        }

        private List<CellRaw> SnapshotAndConvertRangeCells(
            int r1,
            int c1,
            int r2,
            int c2,
            string operationName,
            OfficeIMO.Excel.ExecutionMode? mode,
            CancellationToken ct,
            int workload) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            var raw = new List<CellRaw>(capacity: Math.Max(1024, workload / 4));
            SnapshotCellsInto(raw, r1, c1, r2, c2);
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide(operationName, raw.Count);

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel && raw.Count > 0) {
                var po = new ParallelOptions {
                    CancellationToken = ct,
                    MaxDegreeOfParallelism = policy.MaxDegreeOfParallelism ?? -1
                };
                Parallel.For(0, raw.Count, po, i => raw[i] = ConvertRaw(raw[i]));
            } else {
                for (int i = 0; i < raw.Count; i++) {
                    ct.ThrowIfCancellationRequested();
                    raw[i] = ConvertRaw(raw[i]);
                }
            }

            return raw;
        }

        private void SnapshotCellsInto(List<CellRaw> buffer, int r1, int c1, int r2, int c2) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            foreach (var row in sheetData.Elements<Row>()) {
                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1 || rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    int cIndex = A1.ParseColumnIndexFromCellReference(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    var raw = new CellRaw {
                        Row = rIndex,
                        Col = cIndex,
                        TypeHint = cell.DataType?.Value,
                        StyleIndex = cell.StyleIndex?.Value,
                        HasFormula = cell.CellFormula is not null,
                        RawText = ExtractRawText(cell),
                        InlineText = ExtractInlineString(cell)
                    };

                    if (raw.RawText != null || raw.InlineText != null || CellHasExplicitBlank(cell) || _opt.FillBlanksInRanges)
                        buffer.Add(raw);
                }
            }
        }

        private void FillRangeSequential(object?[,] result, int r1, int c1, int r2, int c2) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            int height = result.GetLength(0);
            int width = result.GetLength(1);

            foreach (var row in sheetData.Elements<Row>()) {
                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                int rr = rIndex - r1;
                if ((uint)rr >= (uint)height) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    int cIndex = A1.ParseColumnIndexFromCellReference(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    int cc = cIndex - c1;
                    if ((uint)cc >= (uint)width) continue;

                    var raw = new CellRaw {
                        TypeHint = cell.DataType?.Value,
                        StyleIndex = cell.StyleIndex?.Value,
                        HasFormula = cell.CellFormula is not null,
                        RawText = ExtractRawText(cell),
                        InlineText = ExtractInlineString(cell)
                    };

                    if (raw.RawText != null || raw.InlineText != null || CellHasExplicitBlank(cell) || _opt.FillBlanksInRanges)
                        result[rr, cc] = ConvertRaw(raw).TypedValue;
                }
            }
        }
    }
}
