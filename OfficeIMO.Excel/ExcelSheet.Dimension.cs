using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Updates the SheetDimension element to reflect the current used range.
        /// Helps avoid rare "dimension" repair messages in Excel when generating sheets programmatically.
        /// </summary>
        internal void UpdateSheetDimension() {
            var ws = _worksheetPart.Worksheet;
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData == null) {
                // Ensure a minimal dimension
                var dim = ws.GetFirstChild<SheetDimension>();
                if (dim == null) {
                    ws.InsertAt(new SheetDimension { Reference = "A1" }, 0);
                } else {
                    dim.Reference = "A1";
                }
                return;
            }

            int minRow = int.MaxValue, maxRow = 0;
            int minCol = int.MaxValue, maxCol = 0;

            foreach (var row in sheetData.Elements<Row>()) {
                if (!row.HasChildren) continue;
                foreach (var cell in row.Elements<Cell>()) {
                    var cref = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cref)) continue;
                    var (r, c) = A1.ParseCellRef(cref!);
                    if (r <= 0 || c <= 0) continue;
                    if (r < minRow) minRow = r;
                    if (r > maxRow) maxRow = r;
                    if (c < minCol) minCol = c;
                    if (c > maxCol) maxCol = c;
                }
            }

            string reference;
            if (maxRow == 0 || maxCol == 0) {
                reference = "A1";
            } else {
                string start = A1.ColumnIndexToLetters(minCol) + minRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
                string end = A1.ColumnIndexToLetters(maxCol) + maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
                reference = start + ":" + end;
            }

            var dimEl = ws.GetFirstChild<SheetDimension>();
            if (dimEl == null) {
                ws.InsertAt(new SheetDimension { Reference = reference }, 0);
            } else {
                dimEl.Reference = reference;
            }
        }
    }
}
