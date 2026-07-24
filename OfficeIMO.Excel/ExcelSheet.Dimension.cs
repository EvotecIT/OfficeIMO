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

                uint? declaredRowIndex = row.RowIndex?.Value;
                int rowIndex = declaredRowIndex.HasValue
                    ? declaredRowIndex.Value >= 1U && declaredRowIndex.Value <= A1.MaxRows
                        ? (int)declaredRowIndex.Value
                        : 0
                    : rowOrdinal <= A1.MaxRows ? rowOrdinal : 0;

                int cellOrdinal = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    cellOrdinal++;
                    int resolvedRow = rowIndex;
                    int resolvedCol = 0;

                    var cref = cell.CellReference?.Value;
                    if (A1.TryParseCellReferenceFast(cref, out int parsedRow, out int parsedCol)) {
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

                    if (resolvedRow <= 0 || resolvedRow > A1.MaxRows || resolvedCol <= 0 || resolvedCol > A1.MaxColumns) continue;

                    if (resolvedRow < minRow) minRow = resolvedRow;
                    if (resolvedRow > maxRow) maxRow = resolvedRow;
                    if (resolvedCol < minCol) minCol = resolvedCol;
                    if (resolvedCol > maxCol) maxCol = resolvedCol;
                }
            }

            if (maxRow == 0 || maxCol == 0) {
                return "A1";
            }

            string start = A1.CellReference(minRow, minCol);
            string end = A1.CellReference(maxRow, maxCol);
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
                InsertSheetDimensionInSchemaOrder(ws, new SheetDimension { Reference = reference });
            } else {
                dimEl.Reference = reference;
            }
        }

        private static void InsertSheetDimensionInSchemaOrder(Worksheet worksheet, SheetDimension dimension) {
            var sheetProperties = worksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties != null) {
                worksheet.InsertAfter(dimension, sheetProperties);
                return;
            }

            worksheet.PrependChild(dimension);
        }
    }
}
