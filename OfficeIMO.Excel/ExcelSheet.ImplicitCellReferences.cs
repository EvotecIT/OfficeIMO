using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
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
                    if (ExcelA1.TryParseCellReferenceFast(cell.CellReference?.Value, out int explicitRow, out int explicitColumn)) {
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

        /// <summary>
        /// Gets the coordinates of every cell, stored or DOM, in row order, for walks that need no cell element.
        /// Stored cells always carry their row's explicit index.
        /// </summary>
        private IEnumerable<(int Row, int Column)> EnumerateCellCoordinates() {
            uint previousRow = 0U;
            foreach (Row row in WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>()) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                if (row is ExcelStoredRow stored && stored.IsStoreAvailable) {
                    foreach ((int Row, int Column) item in EnumerateStoredRowCoordinates(stored, checked((int)effectiveRow))) yield return item;
                } else {
                    int previousColumn = 0;
                    foreach (Cell cell in row.Elements<Cell>()) {
                        int cellRow;
                        int cellColumn;
                        if (ExcelA1.TryParseCellReferenceFast(cell.CellReference?.Value, out int explicitRow, out int explicitColumn)) {
                            cellRow = explicitRow;
                            cellColumn = explicitColumn;
                        } else {
                            cellRow = checked((int)effectiveRow);
                            cellColumn = checked(previousColumn + 1);
                        }
                        previousColumn = cellColumn;
                        yield return (cellRow, cellColumn);
                    }
                }
                previousRow = effectiveRow;
            }
        }

        /// <summary>Merges a stored row's stored and DOM cells by column, as the row's materialized children would be ordered.</summary>
        private static List<(int Row, int Column)> EnumerateStoredRowCoordinates(ExcelStoredRow row, int rowIndex) {
            var coordinates = new List<(int Row, int Column)>(row.StoredCount + 4);
            ExcelStoredCell[] cells = row.Cells!;
            int count = row.StoredCount, index = 0;
            for (OpenXmlElement? child = ExcelWorksheetScan.FirstDomChild(row); child != null; child = ExcelWorksheetScan.NextDomSibling(child)) {
                if (child is not Cell cell) continue;
                ExcelA1.TryParseCellReferenceFast(cell.CellReference?.Value, out int cellRow, out int column);
                while (index < count && cells[index].Column < column) coordinates.Add((rowIndex, cells[index++].Column));
                coordinates.Add((cellRow, column));
            }
            while (index < count) coordinates.Add((rowIndex, cells[index++].Column));
            return coordinates;
        }

        /// <summary>Gets whether a row has any cell, stored or DOM, without moving stored cells into the DOM.</summary>
        private static bool RowHasCells(Row row) {
            if (row is ExcelStoredRow { StoredCount: > 0 }) return true;
            for (OpenXmlElement? child = ExcelWorksheetScan.FirstDomChild(row); child != null; child = ExcelWorksheetScan.NextDomSibling(child)) {
                if (child is Cell) return true;
            }
            return false;
        }

        /// <summary>Gets a row's DOM cells; a stored row's stored cells stay where they are.</summary>
        private static List<Cell> GetDomCells(Row row) {
            var cells = new List<Cell>();
            for (OpenXmlElement? child = ExcelWorksheetScan.FirstDomChild(row); child != null; child = ExcelWorksheetScan.NextDomSibling(child)) {
                if (child is Cell cell) cells.Add(cell);
            }
            return cells;
        }

        private void NormalizeImplicitCellReferences(bool preserveImplicitRowIndexes = false) {
            uint previousRow = 0U;
            foreach (Row row in WorksheetRoot.GetFirstChild<SheetData>()?.Elements<Row>()
                ?? Enumerable.Empty<Row>()) {
                uint effectiveRow = GetEffectiveRowIndex(row, previousRow);
                if (!preserveImplicitRowIndexes) row.RowIndex = effectiveRow;
                int previousColumn = 0;
                // A stored row's DOM cells all have explicit references, and stored cells take theirs from the row.
                foreach (Cell cell in row is ExcelStoredRow ? GetDomCells(row) : row.Elements<Cell>()) {
                    if (ExcelA1.TryParseCellReferenceFast(cell.CellReference?.Value, out _, out int explicitColumn)) {
                        previousColumn = explicitColumn;
                        continue;
                    }
                    int column = checked(previousColumn + 1);
                    if (column > ExcelA1.MaxColumns) {
                        throw new InvalidOperationException(
                            "Implicit worksheet cell order exceeds Excel's column limit.");
                    }
                    cell.CellReference = ExcelA1.CellReference(checked((int)effectiveRow), column);
                    previousColumn = column;
                }
                previousRow = effectiveRow;
            }
        }
    }
}
