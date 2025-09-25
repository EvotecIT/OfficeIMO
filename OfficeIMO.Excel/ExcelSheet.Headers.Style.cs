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
            var a1 = GetUsedRangeA1();
            var (r1, _, r2, _) = A1.ParseRange(a1);
            int startRow = includeHeader ? r1 : r1 + 1;
            if (!TryGetColumnIndexByHeader(header, out var colIndex, options)) {
                // Use endRow one less than startRow so the builder's range is empty and therefore inert.
                return new ColumnStyleByHeaderBuilder(this, 0, startRow, startRow - 1);
            }
            return new ColumnStyleByHeaderBuilder(this, colIndex, startRow, r2);
        }
    }
}
