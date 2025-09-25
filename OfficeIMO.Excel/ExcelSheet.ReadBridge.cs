namespace OfficeIMO.Excel {
    /// <summary>
    /// Read convenience methods exposed directly on ExcelSheet to avoid separate reader usage.
    /// </summary>
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the used range A1 address for this sheet.
        /// Alias property for API ergonomics.
        /// </summary>
        public string UsedRangeA1 => GetUsedRangeA1();

        /// <summary>
        /// Reads the sheet's used range as a sequence of dictionaries using the first row as headers.
        /// </summary>
        /// <param name="options">Optional read options/presets.</param>
        public IEnumerable<Dictionary<string, object?>> Rows(ExcelReadOptions? options = null) {
            using var rdr = _excelDocument.CreateReader(options);
            var sh = rdr.GetSheet(this.Name);
            var a1 = sh.GetUsedRangeA1();
            // Reader's ReadObjects materializes a 2D array first, then yields rows; safe after disposing rdr.
            return sh.ReadObjects(a1);
        }

        /// <summary>
        /// Reads the specified A1 range as a sequence of dictionaries using the first row of the range as headers.
        /// </summary>
        /// <param name="a1Range">Inclusive A1 range (e.g., "A1:C100").</param>
        /// <param name="options">Optional read options/presets.</param>
        public IEnumerable<Dictionary<string, object?>> Rows(string a1Range, ExcelReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            using var rdr = _excelDocument.CreateReader(options);
            var sh = rdr.GetSheet(this.Name);
            return sh.ReadObjects(a1Range);
        }

        /// <summary>
        /// Maps the specified A1 range to a sequence of T using header-to-property mapping.
        /// </summary>
        public IEnumerable<T> RowsAs<T>(string a1Range, ExcelReadOptions? options = null) where T : new() {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            using var rdr = _excelDocument.CreateReader(options);
            var sh = rdr.GetSheet(this.Name);
            return sh.ReadObjects<T>(a1Range);
        }

        /// <summary>
        /// Reads the sheet's used range as editable rows. First row is treated as headers.
        /// </summary>
        public IEnumerable<RowEdit> RowsObjects(ExcelReadOptions? options = null) {
            using var rdr = _excelDocument.CreateReader(options);
            var sh = rdr.GetSheet(this.Name);
            var a1 = sh.GetUsedRangeA1();
            return BuildRowEditsFromRange(sh, a1, options ?? new ExcelReadOptions());
        }

        /// <summary>
        /// Reads the specified A1 range as editable rows. First row is treated as headers.
        /// </summary>
        public IEnumerable<RowEdit> RowsObjects(string a1Range, ExcelReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            using var rdr = _excelDocument.CreateReader(options);
            var sh = rdr.GetSheet(this.Name);
            return BuildRowEditsFromRange(sh, a1Range, options ?? new ExcelReadOptions());
        }

        private IEnumerable<RowEdit> BuildRowEditsFromRange(ExcelSheetReader sh, string a1Range, ExcelReadOptions opt) {
            var values = sh.ReadRange(a1Range);
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);
            if (rows == 0 || cols == 0) yield break;

            var headers = new string[cols];
            for (int c = 0; c < cols; c++) {
                var hdr = values[0, c]?.ToString() ?? $"Column{c + 1}";
                headers[c] = opt.NormalizeHeaders ? RegexNormalize(hdr) : hdr;
            }

            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < cols; c++) map[headers[c]] = c;

            for (int r = 1; r < rows; r++) {
                int sheetRowIndex = r1 + r; // offset + 1 for header
                // Build cell views for this row
                var cellEdits = new CellEdit[cols];
                for (int c = 0; c < cols; c++) {
                    int sheetColIndex = c1 + c;
                    cellEdits[c] = new CellEdit(this, sheetRowIndex, sheetColIndex, values[r, c]);
                }
                yield return new RowEdit(this, sheetRowIndex, headers, map, cellEdits);
            }
        }

        private static string RegexNormalize(string text) {
            return System.Text.RegularExpressions.Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim();
        }
    }
}
