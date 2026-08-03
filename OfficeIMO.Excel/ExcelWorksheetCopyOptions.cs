namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for copying worksheets between workbooks.
    /// </summary>
    public sealed class ExcelWorksheetCopyOptions {
        /// <summary>
        /// Gets or sets the copy strategy. Package copy preserves worksheet XML and avoids object materialization.
        /// </summary>
        public ExcelWorksheetCopyMode CopyMode { get; set; } = ExcelWorksheetCopyMode.Values;

        /// <summary>
        /// Gets or sets whether package-mode copies may preserve external-workbook references.
        /// External links are rejected unless callers opt in explicitly.
        /// </summary>
        public bool CopyExternalWorkbookReferences { get; set; }

        /// <summary>
        /// Gets or sets the maximum number of referenced defined names copied in package mode.
        /// </summary>
        public int MaxDefinedNames { get; set; } = 4096;

        /// <summary>
        /// Gets or sets the maximum aggregate formula characters across copied defined names.
        /// </summary>
        public int MaxDefinedNameCharacters { get; set; } = 1_000_000;

        /// <summary>
        /// Gets or sets the maximum number of in-cell image references copied in package mode.
        /// </summary>
        public int MaxInCellImages { get; set; } = 1024;

        /// <summary>
        /// Gets or sets the maximum payload size of one in-cell image copied in package mode.
        /// </summary>
        public long MaxInCellImageBytes { get; set; } = 32_000_000;

        /// <summary>
        /// Gets or sets the maximum aggregate resolved in-cell image bytes copied in package mode.
        /// Shared image assets are charged once per referencing cell.
        /// </summary>
        public long MaxTotalInCellImageBytes { get; set; } = 64_000_000;
    }
}
