namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for merging worksheet data into another worksheet.
    /// </summary>
    public sealed class ExcelWorksheetMergeOptions {
        /// <summary>
        /// Source range to copy. When omitted, the source worksheet used range is copied.
        /// </summary>
        public string? SourceRange { get; set; }

        /// <summary>
        /// 1-based target row. When omitted, rows are appended after the target worksheet used range.
        /// </summary>
        public int? TargetStartRow { get; set; }

        /// <summary>
        /// 1-based target column. When omitted, the source range's starting column is used.
        /// </summary>
        public int? TargetStartColumn { get; set; }

        /// <summary>
        /// Indicates that the first source row is a header row.
        /// </summary>
        public bool SourceHasHeader { get; set; } = true;

        /// <summary>
        /// Copies the source header row. By default, headers are skipped when appending source rows.
        /// </summary>
        public bool IncludeSourceHeader { get; set; }

        /// <summary>
        /// Matches source columns to target columns by header text before copying values. Existing positional
        /// behavior is used when this is false.
        /// </summary>
        public bool MatchColumnsByHeader { get; set; }

        /// <summary>
        /// 1-based row containing target headers when <see cref="MatchColumnsByHeader"/> is true.
        /// When omitted, append operations use the first row of the target used range, and explicit
        /// target positions use the row immediately above <see cref="TargetStartRow"/>.
        /// </summary>
        public int? TargetHeaderRow { get; set; }

        /// <summary>
        /// Blank rows to leave before appended data when <see cref="TargetStartRow"/> is omitted.
        /// </summary>
        public int BlankRowsBefore { get; set; }

        /// <summary>
        /// Allows merge/join operations to overwrite existing target cell contents. When false, merge/join
        /// operations throw if a non-empty source cell would replace an existing target cell.
        /// </summary>
        public bool OverwriteExistingCells { get; set; }
    }
}
