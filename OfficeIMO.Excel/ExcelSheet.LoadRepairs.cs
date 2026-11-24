using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool _cellRepairsRecorded;

        /// <summary>
        /// Loads the worksheet and records cells that Excel repairs when it opens the file. The grid loader counts them
        /// while it reads; a worksheet the SDK parsed instead is scanned here.
        /// </summary>
        internal void RecordCellRepairs() {
            Worksheet worksheet = WorksheetRoot;
            if (Persistence.UsesCellStore || _cellRepairsRecorded) return;
            int emptyStyleCells = 0;
            string? firstEmptyStyleCell = null;
            foreach (SheetData sheetData in worksheet.Elements<SheetData>()) {
                foreach (Row row in sheetData.Elements<Row>()) {
                    foreach (Cell cell in row.Elements<Cell>()) {
                        if (cell.StyleIndex is { HasValue: false } style && string.IsNullOrEmpty(style.InnerText)
                            && emptyStyleCells++ == 0) {
                            firstEmptyStyleCell = cell.CellReference?.Value;
                        }
                    }
                }
            }
            if (emptyStyleCells > 0) RecordEmptyCellStyleRepair(emptyStyleCells, firstEmptyStyleCell);
        }

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
