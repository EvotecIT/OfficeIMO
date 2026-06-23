namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a selected worksheet cell range decoded from a BIFF Selection record.
    /// </summary>
    public sealed class LegacyXlsSelectedRange {
        /// <summary>
        /// Creates a selected range.
        /// </summary>
        public LegacyXlsSelectedRange(int startRow, int startColumn, int endRow, int endColumn) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }

        /// <summary>Gets the one-based first selected row.</summary>
        public int StartRow { get; }

        /// <summary>Gets the one-based first selected column.</summary>
        public int StartColumn { get; }

        /// <summary>Gets the one-based last selected row.</summary>
        public int EndRow { get; }

        /// <summary>Gets the one-based last selected column.</summary>
        public int EndColumn { get; }

        /// <summary>Gets the A1 reference for the selected range.</summary>
        public string Reference {
            get {
                string start = A1.CellReference(StartRow, StartColumn);
                string end = A1.CellReference(EndRow, EndColumn);
                return start == end ? start : start + ":" + end;
            }
        }
    }
}
