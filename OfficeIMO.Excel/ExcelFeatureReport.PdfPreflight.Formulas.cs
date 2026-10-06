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

            return IsDefaultPdfSelectedCell(sheet, row, column);
        }

        private static bool IsDefaultPdfSelectedCell(ExcelSheet sheet, int row, int column) {
            if (IsPdfHiddenRow(sheet, row) || IsPdfHiddenColumn(sheet, column)) {
                return false;
            }

            IReadOnlyList<string> areas = sheet.GetPrintAreas();
            if (areas.Count == 0) return true;
            ExcelPrintTitles titles = sheet.GetPrintTitles();
            foreach (string area in areas) {
                // Invalid metadata is reported separately; remain conservative about cached formula results.
                if (!sheet.TryParsePrintAreaReference(area, out ExcelReference? reference) || reference == null) return true;
                int firstRow = reference.Kind == ExcelReferenceKind.WholeColumn ? 1 : Math.Min(reference.Start.Row, reference.End.Row);
                int lastRow = reference.Kind == ExcelReferenceKind.WholeColumn ? A1.MaxRows : Math.Max(reference.Start.Row, reference.End.Row);
                int firstColumn = reference.Kind == ExcelReferenceKind.WholeRow ? 1 : Math.Min(reference.Start.Column, reference.End.Column);
                int lastColumn = reference.Kind == ExcelReferenceKind.WholeRow ? A1.MaxColumns : Math.Max(reference.Start.Column, reference.End.Column);
                bool selectedRow = row >= firstRow && row <= lastRow
                    || titles.HasRows && row >= titles.FirstRow!.Value && row <= Math.Min(titles.LastRow!.Value, lastRow);
                bool selectedColumn = column >= firstColumn && column <= lastColumn
                    || titles.HasColumns && column >= titles.FirstColumn!.Value && column <= Math.Min(titles.LastColumn!.Value, lastColumn);
                if (selectedRow && selectedColumn) return true;
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
