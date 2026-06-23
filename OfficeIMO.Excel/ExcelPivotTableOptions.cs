namespace OfficeIMO.Excel {
    /// <summary>
    /// Workbook-interaction options for authored pivot tables.
    /// </summary>
    public sealed class ExcelPivotTableOptions {
        /// <summary>
        /// Gets or sets whether Excel should refresh the pivot cache when the workbook opens.
        /// </summary>
        public bool? RefreshOnOpen { get; set; }

        /// <summary>
        /// Gets or sets whether source cache records should be saved in the workbook package.
        /// </summary>
        public bool? SaveSourceData { get; set; }

        /// <summary>
        /// Gets or sets whether Excel should preserve pivot table formatting during refreshes.
        /// </summary>
        public bool? PreserveFormatting { get; set; }

        /// <summary>
        /// Gets or sets whether drill-down interaction is enabled for pivot details.
        /// </summary>
        public bool? EnableDrill { get; set; }
    }
}
