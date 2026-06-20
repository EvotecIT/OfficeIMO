namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static bool IsDefaultPdfExportedFormulaCell(
            ExcelFormulaCellInfo formula,
            IReadOnlyDictionary<string, ExcelSheet> exportedSheetsByName) {
            if (!exportedSheetsByName.TryGetValue(formula.SheetName, out ExcelSheet? sheet)) {
                return false;
            }

            if (!A1.TryParseCellReferenceFast(formula.CellReference, out int row, out int column)) {
                return true;
            }

            if (IsPdfHiddenRow(sheet, row) || IsPdfHiddenColumn(sheet, column)) {
                return false;
            }

            if (!TryGetDefaultPdfBodyRange(sheet, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return true;
            }

            ExcelPrintTitles titles = sheet.GetPrintTitles();
            if (titles.HasRows
                && row >= titles.FirstRow!.Value
                && row <= titles.LastRow!.Value
                && column >= firstColumn
                && column <= lastColumn) {
                return true;
            }

            return row >= firstRow
                   && row <= lastRow
                   && column >= firstColumn
                   && column <= lastColumn;
        }

        private static bool TryGetDefaultPdfBodyRange(ExcelSheet sheet, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            firstRow = 0;
            firstColumn = 0;
            lastRow = 0;
            lastColumn = 0;

            string? printArea = sheet.GetPrintArea();
            if (string.IsNullOrWhiteSpace(printArea)) {
                return false;
            }

            if (ContainsMultiplePrintAreas(printArea!)) {
                return false;
            }

            string range = StripSheetPrefix(printArea!).Replace("$", string.Empty);
            if (A1.TryParseRange(range, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                return true;
            }

            if (!A1.TryParseCellReferenceFast(range, out firstRow, out firstColumn)) {
                return false;
            }

            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static string StripSheetPrefix(string reference) {
            int separator = reference.LastIndexOf('!');
            return separator >= 0 && separator + 1 < reference.Length
                ? reference.Substring(separator + 1)
                : reference;
        }

        private static bool IsPdfHiddenRow(ExcelSheet sheet, int rowIndex) {
            IReadOnlyList<ExcelRowSnapshot> definitions = sheet.GetRowDefinitions();
            for (int i = definitions.Count - 1; i >= 0; i--) {
                if (definitions[i].Index == rowIndex) {
                    return definitions[i].Hidden;
                }
            }

            return false;
        }

        private static bool IsPdfHiddenColumn(ExcelSheet sheet, int columnIndex) {
            IReadOnlyList<ExcelColumnSnapshot> definitions = sheet.GetColumnDefinitions();
            for (int i = definitions.Count - 1; i >= 0; i--) {
                ExcelColumnSnapshot definition = definitions[i];
                if (columnIndex >= definition.StartIndex && columnIndex <= definition.EndIndex) {
                    return definition.Hidden;
                }
            }

            return false;
        }
    }
}
