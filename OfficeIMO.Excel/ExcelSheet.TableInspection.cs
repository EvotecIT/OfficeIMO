namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the named Excel table definitions owned by this worksheet.
        /// Ordinary cell grids are not named tables. Metadata is read through the
        /// same locked snapshot owner as <see cref="ExcelDocument.GetTables()"/>.
        /// </summary>
        /// <remarks>Saving the workbook retains live worksheet handles, including when saving rebuilds the package.</remarks>
        /// <exception cref="System.InvalidOperationException">The worksheet was removed or no longer belongs to the current workbook package.</exception>
        public IReadOnlyList<ExcelTableInfo> GetTables() => _excelDocument.GetTables(_worksheetPart);
    }
}
