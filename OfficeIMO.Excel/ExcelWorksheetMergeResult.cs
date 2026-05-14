namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a completed worksheet merge operation.
    /// </summary>
    public sealed class ExcelWorksheetMergeResult {
        /// <summary>
        /// Creates a merge result.
        /// </summary>
        public ExcelWorksheetMergeResult(
            string sourceSheetName,
            string targetSheetName,
            string sourceRange,
            string targetRange,
            int rowsCopied,
            int columnsCopied,
            bool headerSkipped) {
            SourceSheetName = sourceSheetName;
            TargetSheetName = targetSheetName;
            SourceRange = sourceRange;
            TargetRange = targetRange;
            RowsCopied = rowsCopied;
            ColumnsCopied = columnsCopied;
            HeaderSkipped = headerSkipped;
        }

        /// <summary>Source worksheet name.</summary>
        public string SourceSheetName { get; }

        /// <summary>Target worksheet name.</summary>
        public string TargetSheetName { get; }

        /// <summary>Source A1 range read by the merge operation.</summary>
        public string SourceRange { get; }

        /// <summary>Target A1 range occupied by copied cells.</summary>
        public string TargetRange { get; }

        /// <summary>Number of source rows copied.</summary>
        public int RowsCopied { get; }

        /// <summary>Number of columns copied.</summary>
        public int ColumnsCopied { get; }

        /// <summary>Whether the first source row was skipped as a header.</summary>
        public bool HeaderSkipped { get; }
    }
}
