namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns the named Excel table definitions owned by this worksheet.
        /// Ordinary cell grids are not named tables. Metadata is read through the
        /// same locked snapshot owner as <see cref="ExcelDocument.GetTables()"/>.
        /// </summary>
        public IReadOnlyList<ExcelTableInfo> GetTables() => _excelDocument.GetTables(_worksheetPart);
    }
}
