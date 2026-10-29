namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how worksheets are copied between workbooks.
    /// </summary>
    public enum ExcelWorksheetCopyMode {
        /// <summary>
        /// Copy cell values through the workbook reader and writer surface.
        /// </summary>
        Values = 0,

        /// <summary>
        /// Copy worksheet XML directly, rewriting workbook-scoped references such as shared strings and styles.
        /// Drawings, charts, legacy and threaded comments, header/footer VML and external hyperlinks are cloned with the sheet.
        /// </summary>
        Package = 1
    }
}
