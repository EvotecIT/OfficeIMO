namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Whether the workbook has a cell metadata part (dynamic arrays and rich values).</summary>
        internal bool HasCellMetadataPart => WorkbookPartRoot.CellMetadataPart != null;
    }
}
