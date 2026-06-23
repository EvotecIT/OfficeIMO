namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a merged range parsed from a legacy XLS worksheet.
    /// </summary>
    public sealed class LegacyXlsMergedRange {
        /// <summary>
        /// Creates legacy merged range metadata.
        /// </summary>
        public LegacyXlsMergedRange(int startRow, int startColumn, int endRow, int endColumn) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }

        /// <summary>
        /// Gets the one-based first row.
        /// </summary>
        public int StartRow { get; }

        /// <summary>
        /// Gets the one-based first column.
        /// </summary>
        public int StartColumn { get; }

        /// <summary>
        /// Gets the one-based last row.
        /// </summary>
        public int EndRow { get; }

        /// <summary>
        /// Gets the one-based last column.
        /// </summary>
        public int EndColumn { get; }
    }
}
