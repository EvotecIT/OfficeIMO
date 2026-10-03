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

            IReadOnlyList<string> areas = sheet.GetPrintAreas();
            if (areas.Count == 0) return true;
            ExcelPrintTitles titles = sheet.GetPrintTitles();
            foreach (string area in areas) {
                // Invalid metadata is reported separately; remain conservative about cached formula results.
                if (!sheet.TryParsePrintAreaReference(area, out ExcelReference? reference) || reference == null) return true;
                if (reference.Contains(row, column)) return true;
                int firstColumn = reference.Kind == ExcelReferenceKind.WholeRow ? 1 : Math.Min(reference.Start.Column, reference.End.Column);
                int lastColumn = reference.Kind == ExcelReferenceKind.WholeRow ? A1.MaxColumns : Math.Max(reference.Start.Column, reference.End.Column);
                if (titles.HasRows && row >= titles.FirstRow!.Value && row <= titles.LastRow!.Value
                    && column >= firstColumn && column <= lastColumn) return true;
            }
            return false;
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
