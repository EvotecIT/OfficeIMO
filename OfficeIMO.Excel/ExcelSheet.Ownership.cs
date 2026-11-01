namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Gets the workbook that owns this worksheet.</summary>
        /// <remarks>The worksheet does not own its workbook's lifetime. Save and dispose the document explicitly.</remarks>
        public ExcelDocument Document => _excelDocument;
    }
}
