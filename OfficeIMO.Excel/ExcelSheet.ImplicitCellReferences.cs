using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private IEnumerable<(Cell Cell, int Row, int Column)> EnumerateCellsWithEffectiveCoordinates() {
            uint previousRow = 0U;
            foreach (Row row in WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>()) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                int previousColumn = 0;
                foreach (Cell cell in row.Elements<Cell>()) {
                    int cellRow;
                    int cellColumn;
                    if (A1.TryParseCellReferenceFast(cell.CellReference?.Value, out int explicitRow, out int explicitColumn)) {
                        cellRow = explicitRow;
                        cellColumn = explicitColumn;
                    } else {
                        cellRow = checked((int)effectiveRow);
                        cellColumn = checked(previousColumn + 1);
                    }
                    previousColumn = cellColumn;
                    yield return (cell, cellRow, cellColumn);
                }
                previousRow = effectiveRow;
            }
        }

        private void NormalizeImplicitCellReferences() {
            uint previousRow = 0U;
            foreach (Row row in WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>()) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                row.RowIndex = effectiveRow;
                int previousColumn = 0;
                foreach (Cell cell in row.Elements<Cell>()) {
                    if (A1.TryParseCellReferenceFast(cell.CellReference?.Value, out _, out int explicitColumn)) {
                        previousColumn = explicitColumn;
                        continue;
                    }
                    int column = checked(previousColumn + 1);
                    if (column > A1.MaxColumns) {
                        throw new InvalidOperationException(
                            "Implicit worksheet cell order exceeds Excel's column limit.");
                    }
                    cell.CellReference = A1.CellReference(checked((int)effectiveRow), column);
                    previousColumn = column;
                }
                previousRow = effectiveRow;
            }
        }
    }
}
