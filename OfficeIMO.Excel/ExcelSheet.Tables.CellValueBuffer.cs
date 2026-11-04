using DocumentFormat.OpenXml.Spreadsheet;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // A table over buffered cells stays buffered when the header row needs no rewriting and
        // the range is complete; other tables take the DOM route, which fills and renames cells.
        private bool TryGetPendingTableHeaders(string range, out IReadOnlyList<string>? headers) {
            headers = null;
            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            if (buffer == null
                || !CanAcceptPendingDirectCell()
                || !ExcelA1.TryParseStrictRange(range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                || firstRow >= lastRow
                || firstColumn > lastColumn) {
                return false;
            }

            var names = new string[lastColumn - firstColumn + 1];
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = firstColumn; column <= lastColumn; column++) {
                int cellIndex = buffer.FindCell(firstRow, column);
                if (cellIndex < 0
                    || buffer.GetCell(cellIndex).Reference is not string name
                    || string.IsNullOrWhiteSpace(name)
                    || name.IndexOf('\n') >= 0
                    || name.IndexOf('\r') >= 0
                    || !used.Add(name)) {
                    return false;
                }

                names[column - firstColumn] = name;
            }

            for (int row = firstRow + 1; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    if (!buffer.Contains(row, column)) {
                        return false;
                    }
                }
            }

            headers = names;
            return true;
        }

        private Cell? InspectTableCell(int row, int column) {
            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            int cellIndex = buffer?.FindCell(row, column) ?? -1;
            if (cellIndex < 0) {
                return buffer == null ? TryGetExistingCell(row, column) : null;
            }

            ExcelBufferedCell buffered = buffer!.GetCell(cellIndex);
            var cell = new Cell { StyleIndex = GetEffectiveBufferedStyle(buffer, cellIndex) };
            switch (buffered.Kind) {
                case ExcelBufferedCellKind.StyleOnly:
                    break;
                case ExcelBufferedCellKind.Text:
                case ExcelBufferedCellKind.Blank:
                    cell.DataType = XCellValues.String;
                    cell.CellValue = new CellValue(buffered.Reference as string ?? string.Empty);
                    break;
                case ExcelBufferedCellKind.Boolean:
                    cell.DataType = XCellValues.Boolean;
                    cell.CellValue = new CellValue("0");
                    break;
                default:
                    if (buffered.Reference is DirectFormulaCellValue formula) {
                        cell.CellFormula = new CellFormula(formula.Formula);
                    } else {
                        cell.DataType = XCellValues.Number;
                        cell.CellValue = new CellValue("0");
                    }

                    break;
            }

            return cell;
        }

        private bool HasHorizontalAlignment(int row, int column) {
            ExcelCellValueBuffer? buffer = _pendingCellValueDirectSaveBuffer;
            int cellIndex = buffer?.FindCell(row, column) ?? -1;
            if (cellIndex < 0) {
                return !string.IsNullOrWhiteSpace(GetCellStyle(row, column).HorizontalAlignment);
            }

            uint style = buffer!.GetStyle(cellIndex);
            return style != 0U
                && _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>()
                    .ElementAtOrDefault((int)style)?.Alignment?.Horizontal != null;
        }
    }
}
