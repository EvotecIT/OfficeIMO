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
            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null)
                return "A1:A1";

            int minRow = int.MaxValue, maxRow = 0;
            int minCol = int.MaxValue, maxCol = 0;

            foreach (var row in sheetData.Elements<Row>()) {
                if (!row.HasChildren) continue;
                int rIndex = checked((int)row.RowIndex!.Value);
                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is null) continue;
                    var (r, c) = A1.ParseCellRef(cell.CellReference.Value);
                    if (r <= 0 || c <= 0) continue;
                    if (r < minRow) minRow = r;
                    if (r > maxRow) maxRow = r;
                    if (c < minCol) minCol = c;
                    if (c > maxCol) maxCol = c;
                }
            }

            if (maxRow == 0 || maxCol == 0)
                return "A1:A1";

            string start = A1.ColumnIndexToLetters(minCol) + minRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string end = A1.ColumnIndexToLetters(maxCol) + maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return start + ":" + end;
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

            var raw = new List<CellRaw>(capacity: Math.Max(1024, height * width / 4));
            SnapshotCellsInto(raw, r1, c1, r2, c2);

            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide("ReadRange", raw.Count);

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel && raw.Count > 0) {
                var po = new ParallelOptions {
                    CancellationToken = ct,
                    MaxDegreeOfParallelism = policy.MaxDegreeOfParallelism ?? -1
                };
                Parallel.For(0, raw.Count, po, i => raw[i] = ConvertRaw(raw[i]));
            } else {
                for (int i = 0; i < raw.Count; i++) raw[i] = ConvertRaw(raw[i]);
            }

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
            var values = ReadRange(a1Range, mode, ct);
            var dt = new DataTable(_sheetName);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            // Columns
            if (headersInFirstRow && rows > 0) {
                for (int c = 0; c < cols; c++) {
                    var hdr = values[0, c]?.ToString() ?? $"Column{c + 1}";
                    if (_opt.NormalizeHeaders) hdr = RegexNormalize(hdr);
                    dt.Columns.Add(hdr, typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) dt.Columns.Add($"Column{c + 1}", typeof(object));
            }

            // Rows
            int startRow = headersInFirstRow ? 1 : 0;
            for (int r = startRow; r < rows; r++) {
                var row = dt.NewRow();
                for (int c = 0; c < cols; c++) row[c] = values[r, c] ?? DBNull.Value;
                dt.Rows.Add(row);
            }
            return dt;
        }

        /// <summary>
        /// Reads a rectangular range into a sequence of dictionaries using the first row as headers.
        /// </summary>
        public IEnumerable<Dictionary<string, object?>> ReadObjects(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var values = ReadRange(a1Range, mode, ct);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows == 0 || cols == 0) yield break;

            var headers = new string[cols];
            for (int c = 0; c < cols; c++) {
                var hdr = values[0, c]?.ToString() ?? $"Column{c + 1}";
                headers[c] = _opt.NormalizeHeaders ? RegexNormalize(hdr) : hdr;
            }

            for (int r = 1; r < rows; r++) {
                var dict = new Dictionary<string, object?>(cols, System.StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++) dict[headers[c]] = values[r, c];
                yield return dict;
            }
        }

        private static string RegexNormalize(string text) {
            return System.Text.RegularExpressions.Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim();
        }

        private void SnapshotCellsInto(List<CellRaw> buffer, int r1, int c1, int r2, int c2) {
            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            foreach (var row in sheetData.Elements<Row>()) {
                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1 || rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    var (_, cIndex) = A1.ParseCellRef(cell.CellReference?.Value ?? string.Empty);
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
    }
}
