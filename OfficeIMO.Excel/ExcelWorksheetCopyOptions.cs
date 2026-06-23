namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for copying worksheets between workbooks.
    /// </summary>
    public sealed class ExcelWorksheetCopyOptions {
        /// <summary>
        /// Gets or sets the copy strategy. Package copy preserves worksheet XML and avoids object materialization.
        /// </summary>
        public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Package;
    }
}
