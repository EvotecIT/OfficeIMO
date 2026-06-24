namespace OfficeIMO.Excel {
    /// <summary>
    /// Header-based style helpers.
    /// </summary>
    public partial class ExcelSheet {


        /// <summary>
        /// Returns a builder for styling a column resolved by header with discoverable methods.
        /// When the header cannot be resolved an inert builder is returned so calls become no-ops.
        /// </summary>
        /// <param name="header">Header text used to resolve the target column after applying any configured normalization.</param>
        /// <param name="includeHeader">True to include the header row when styling; false to begin styling from the first data row.</param>
        /// <param name="options">Read options that control header normalization and other resolution behavior.</param>
        public ColumnStyleByHeaderBuilder ColumnStyleByHeader(string header, bool includeHeader = false, ExcelReadOptions? options = null) {
            TryGetColumnStyleByHeader(header, includeHeader, out var builder, out _, options);
            return builder;
        }

        /// <summary>
        /// Attempts to create a column style builder and returns the resolved 1-based column index.
        /// </summary>
        /// <param name="header">Header text used to resolve the target column after applying any configured normalization.</param>
        /// <param name="includeHeader">True to include the header row when styling; false to begin styling from the first data row.</param>
        /// <param name="builder">When successful, receives a builder for the resolved column.</param>
        /// <param name="columnIndex">When successful, receives the resolved 1-based column index.</param>
        /// <param name="options">Read options that control header normalization and other resolution behavior.</param>
        /// <param name="preferDirectTabularMetadata">True to resolve against pending direct tabular metadata when available; false to force worksheet materialization before resolving.</param>
        public bool TryGetColumnStyleByHeader(
            string header,
            bool includeHeader,
            out ColumnStyleByHeaderBuilder builder,
            out int columnIndex,
            ExcelReadOptions? options = null,
            bool preferDirectTabularMetadata = true) {
            if (preferDirectTabularMetadata &&
                _excelDocument.TryGetDirectTabularSaveCandidateColumnByHeader(this, header, includeHeader, options, out int directColumnIndex, out int directStartRow, out int directEndRow)) {
                columnIndex = directColumnIndex;
                builder = new ColumnStyleByHeaderBuilder(this, directColumnIndex, directStartRow, directEndRow);
                return true;
            }

            var a1 = GetUsedRangeA1();
            var (r1, _, r2, _) = A1.ParseRange(a1);
            int startRow = includeHeader ? r1 : r1 + 1;
            if (!TryGetColumnIndexByHeader(header, out var colIndex, options)) {
                columnIndex = 0;
                builder = new ColumnStyleByHeaderBuilder(this, 0, startRow, startRow - 1);
                return false;
            }

            columnIndex = colIndex;
            builder = new ColumnStyleByHeaderBuilder(this, colIndex, startRow, r2);
            return true;
        }
    }
}
