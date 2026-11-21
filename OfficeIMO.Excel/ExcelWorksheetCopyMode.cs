namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how worksheets are copied between workbooks.
    /// </summary>
    public enum ExcelWorksheetCopyMode {
        /// <summary>
        /// Copy cell values through the workbook reader and writer surface, with tables and the sheet layout: column
        /// widths, row heights, views, merged cells, conditional formats, data validations and print settings.
        /// </summary>
        Values = 0,

        /// <summary>
        /// Copy worksheet XML directly, rewriting workbook-scoped references such as shared strings and styles.
        /// Drawings, charts, legacy and threaded comments, header/footer VML and external hyperlinks are cloned with the sheet.
        /// </summary>
        Package = 1
    }
}
