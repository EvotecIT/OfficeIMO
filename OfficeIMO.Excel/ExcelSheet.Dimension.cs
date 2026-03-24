using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal static string ComputeSheetDimensionReference(Worksheet worksheet) {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return "A1";
            }

            int minRow = int.MaxValue;
            int maxRow = 0;
            int minCol = int.MaxValue;
            int maxCol = 0;
            int rowOrdinal = 0;

            foreach (var row in sheetData.Elements<Row>()) {
                rowOrdinal++;
                if (!row.HasChildren) continue;

                int rowIndex = checked((int)(row.RowIndex?.Value ?? (uint)rowOrdinal));
                if (rowIndex <= 0) {
                    rowIndex = rowOrdinal;
                }

                int cellOrdinal = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    cellOrdinal++;
                    int resolvedRow = rowIndex;
                    int resolvedCol = 0;

                    var cref = cell.CellReference?.Value;
                    if (!string.IsNullOrEmpty(cref)) {
                        var (parsedRow, parsedCol) = A1.ParseCellRef(cref!);
                        if (parsedRow > 0) {
                            resolvedRow = parsedRow;
                        }
                        if (parsedCol > 0) {
                            resolvedCol = parsedCol;
                        }
                    }

                    if (resolvedCol <= 0) {
                        resolvedCol = cellOrdinal;
                    }

                    if (resolvedRow <= 0 || resolvedCol <= 0) continue;

                    if (resolvedRow < minRow) minRow = resolvedRow;
                    if (resolvedRow > maxRow) maxRow = resolvedRow;
                    if (resolvedCol < minCol) minCol = resolvedCol;
                    if (resolvedCol > maxCol) maxCol = resolvedCol;
                }
            }

            if (maxRow == 0 || maxCol == 0) {
                return "A1";
            }

            string start = A1.ColumnIndexToLetters(minCol) + minRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string end = A1.ColumnIndexToLetters(maxCol) + maxRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return start == end ? start : start + ":" + end;
        }

        /// <summary>
        /// Updates the SheetDimension element to reflect the current used range.
        /// Helps avoid rare "dimension" repair messages in Excel when generating sheets programmatically.
        /// </summary>
        internal void UpdateSheetDimension() {
            var ws = WorksheetRoot;
            var dimensions = ws.Elements<SheetDimension>().ToList();
            SheetDimension? dimEl = dimensions.FirstOrDefault();
            foreach (var extraDimension in dimensions.Skip(1).ToList()) {
                extraDimension.Remove();
            }

            string reference = ComputeSheetDimensionReference(ws);

            if (dimEl == null) {
                ws.InsertAt(new SheetDimension { Reference = reference }, 0);
            } else {
                dimEl.Reference = reference;
            }
        }
    }
}
