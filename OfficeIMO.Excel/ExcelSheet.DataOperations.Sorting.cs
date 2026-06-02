namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // -------- Sorting (values-only, rewrites range) --------

        /// <summary>
        /// Sorts the sheet's UsedRange rows in-place (excluding header) by the column resolved via header.
        /// Values-only: rewrites cell values; formulas and styles are not preserved.
        /// When the header cannot be resolved the sort is skipped.
        /// </summary>
        public void SortUsedRangeByHeader(string header, bool ascending = true) {
            if (string.IsNullOrWhiteSpace(header)) throw new ArgumentNullException(nameof(header));

            var a1 = GetUsedRangeA1();
            var (r1, c1, r2, c2) = A1.ParseRange(a1);
            if (r2 - r1 < 1) return; // nothing to sort

            if (!TryGetColumnIndexByHeader(header, out var targetCol)) return;
            if (targetCol < c1 || targetCol > c2) return;
            int targetIndex = targetCol - c1; // zero-based in matrix

            // Read values
            using var rdr = _excelDocument.CreateReader();
            var sh = rdr.GetSheet(this.Name);
            var values = sh.ReadRange(a1);
            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            // Build list (rowIndexInSheet, rowValues)
            var list = new List<(int SheetRow, object?[] Row)>();
            for (int r = 1; r < rows; r++) // skip header at 0
            {
                var arr = new object?[cols];
                for (int c = 0; c < cols; c++) arr[c] = values[r, c];
                list.Add((r1 + r, arr));
            }

            int Comparer((int SheetRow, object?[] Row) a, (int SheetRow, object?[] Row) b) {
                var va = a.Row[targetIndex];
                var vb = b.Row[targetIndex];
                int cmp = CompareCell(va, vb);
                return ascending ? cmp : -cmp;
            }

            list.Sort(Comparer);

            // Rewrite values back
            WriteLock(() => {
                for (int i = 0; i < list.Count; i++) {
                    var row = list[i];
                    int sheetRow = r1 + 1 + i; // first data row
                    for (int c = 0; c < cols; c++) {
                        CellValueCore(sheetRow, c1 + c, row.Row[c] ?? string.Empty);
                    }
                }
            });

            static int CompareCell(object? a, object? b) {
                if (a == null && b == null) return 0;
                if (a == null) return -1;
                if (b == null) return 1;
                if (a is IComparable ca && b is IComparable cb && a.GetType() == b.GetType())
                    return ca.CompareTo(cb);
                // Fallback string compare
                return string.Compare(Convert.ToString(a), Convert.ToString(b), StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Sorts the sheet's UsedRange rows in-place by multiple headers in the given order (excluding header row).
        /// Values-only: rewrites cell values; formulas and styles are not preserved. Missing headers are ignored and when no
        /// valid headers remain the sort is skipped.
        /// </summary>
        public void SortUsedRangeByHeaders(params (string Header, bool Ascending)[] keys) {
            if (keys == null || keys.Length == 0) throw new ArgumentException("At least one key is required.", nameof(keys));

            var a1 = GetUsedRangeA1();
            var (r1, c1, r2, c2) = A1.ParseRange(a1);
            if (r2 - r1 < 1) return;

            // Resolve column indices and validate
            var cols = new List<(int Index, bool Asc)>();
            var headerMap = GetHeaderMapCached(DefaultHeaderReadOptions);
            foreach (var k in keys) {
                if (string.IsNullOrWhiteSpace(k.Header)) continue;
                if (!headerMap.TryGetValue(k.Header, out var col)) continue;
                if (col < c1 || col > c2) continue;
                cols.Add((col - c1, k.Ascending));
            }
            if (cols.Count == 0) return;

            using var rdr = _excelDocument.CreateReader();
            var sh = rdr.GetSheet(this.Name);
            var values = sh.ReadRange(a1);
            int rows = values.GetLength(0);
            int width = values.GetLength(1);

            var list = new List<(int SheetRow, object?[] Row)>();
            for (int r = 1; r < rows; r++) {
                var arr = new object?[width];
                for (int c = 0; c < width; c++) arr[c] = values[r, c];
                list.Add((r1 + r, arr));
            }

            list.Sort((a, b) => {
                foreach (var (idx, asc) in cols) {
                    int cmp = CompareCell(a.Row[idx], b.Row[idx]);
                    if (cmp != 0) return asc ? cmp : -cmp;
                }
                return 0;
            });

            WriteLock(() => {
                for (int i = 0; i < list.Count; i++) {
                    var row = list[i];
                    int sheetRow = r1 + 1 + i;
                    for (int c = 0; c < width; c++)
                        CellValueCore(sheetRow, c1 + c, row.Row[c] ?? string.Empty);
                }
            });

            static int CompareCell(object? a, object? b) {
                if (a == null && b == null) return 0;
                if (a == null) return -1;
                if (b == null) return 1;
                if (a is IComparable ca && b is IComparable cb && a.GetType() == b.GetType())
                    return ca.CompareTo(cb);
                return string.Compare(Convert.ToString(a), Convert.ToString(b), StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
