namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents frozen pane metadata parsed from a legacy XLS worksheet view.
    /// </summary>
    public sealed class LegacyXlsFreezePane {
        /// <summary>
        /// Creates legacy frozen pane metadata.
        /// </summary>
        /// <param name="topRows">Number of top rows to freeze.</param>
        /// <param name="leftColumns">Number of left columns to freeze.</param>
        public LegacyXlsFreezePane(int topRows, int leftColumns) {
            TopRows = topRows;
            LeftColumns = leftColumns;
        }

        /// <summary>
        /// Gets the number of top rows to freeze.
        /// </summary>
        public int TopRows { get; }

        /// <summary>
        /// Gets the number of left columns to freeze.
        /// </summary>
        public int LeftColumns { get; }
    }
}
