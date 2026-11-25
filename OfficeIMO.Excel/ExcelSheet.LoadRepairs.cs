namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool _cellRepairsRecorded;

        /// <summary>
        /// Excel repairs an empty cell style attribute (<c>s=""</c>) to the default style when it opens the file, keeping
        /// the value. Readers do the same, and the load report records it once per worksheet.
        /// </summary>
        private void RecordEmptyCellStyleRepair(int cells, string? firstCell) {
            if (_cellRepairsRecorded) return;
            _cellRepairsRecorded = true;
            string where = string.IsNullOrEmpty(firstCell) ? string.Empty : $" (first at {firstCell})";
            string message = cells == 1
                ? $"Worksheet '{Name}' has a cell with an empty style attribute{where}. It is read with the default style, as Excel repairs it."
                : $"Worksheet '{Name}' has {cells} cells with an empty style attribute{where}. They are read with the default style, as Excel repairs them.";
            _excelDocument.AddSourceLoadDiagnostic(new OfficeDiagnostic(
                OfficeDiagnosticCodes.ExcelCellStyleRepaired,
                OfficeDiagnosticSeverity.Warning,
                message,
                string.IsNullOrEmpty(firstCell) ? Name : Name + "!" + firstCell,
                "Excel"));
        }
    }
}
