using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects the supported XLSB workbook model into OfficeIMO's normal editable workbook surface.</summary>
    internal static class XlsbWorkbookProjector {
        internal static ExcelDocument ToExcelDocument(XlsbWorkbook workbook, System.Threading.CancellationToken cancellationToken = default,
            int maxCells = 4_000_000) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            cancellationToken.ThrowIfCancellationRequested();
            EnsureArrayProjectionBudget(workbook, maxCells, cancellationToken);

            ExcelDocument document = ExcelDocument.Create();
            try {
                if (workbook.Uses1904DateSystem) {
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                }
                if (workbook.Stylesheet != null) {
                    XlsbStylesheetProjector.Install(document, workbook.Stylesheet);
                }
                if (workbook.WorkbookProtection != null) {
                    XlsbWorkbookProtectionProjector.Apply(document, workbook.WorkbookProtection);
                }

                var relatedParts = new XlsbRelatedPartProjector(document, workbook, cancellationToken);
                foreach (XlsbSheet sourceSheet in workbook.Sheets) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (sourceSheet.IsChartSheet) {
                        relatedParts.ProjectChartSheet(sourceSheet);
                        continue;
                    }
                    ExcelSheet targetSheet = document.AddSheet(sourceSheet.Name);
                    targetSheet.Batch(sheet => {
                        foreach (XlsbCell cell in sourceSheet.Cells) {
                            cancellationToken.ThrowIfCancellationRequested();
                            switch (cell.Kind) {
                                case XlsbCellValueKind.Blank:
                                    break;
                                case XlsbCellValueKind.Error:
                                    sheet.SetLegacyErrorCellValue(cell.Row, cell.Column, cell.Value as string ?? "#VALUE!");
                                    break;
                                default:
                                    sheet.CellValue(cell.Row, cell.Column, cell.Value);
                                    break;
                            }

                            if (!string.IsNullOrWhiteSpace(cell.FormulaText)) {
                                sheet.CellFormula(cell.Row, cell.Column, cell.FormulaText!);
                            }

                            sheet.SetCellStyleIndex(cell.Row, cell.Column, cell.StyleIndex, save: false);
                        }
                    });

                    var arrays = sourceSheet.FormulaGroups.Where(group => group.Type == 426).ToArray();
                    var sourceCells = arrays.Length == 0 ? new Dictionary<(int Row, int Column), XlsbCell>()
                        : sourceSheet.Cells.Where(cell => !string.IsNullOrWhiteSpace(cell.FormulaText)).ToDictionary(cell => (cell.Row, cell.Column));
                    foreach (var group in arrays) {
                        sourceCells.TryGetValue((group.AnchorRow, group.AnchorColumn), out var anchor);
                        if (string.IsNullOrWhiteSpace(anchor?.FormulaText)) continue;
                        targetSheet.SetLegacyArrayFormula(group.Range.ToA1Reference(), anchor!.FormulaText!, cancellationToken);
                    }
                    XlsbWorksheetPropertiesProjector.Apply(targetSheet, sourceSheet.Properties);
                    XlsbWorksheetGeometryProjector.Apply(targetSheet, sourceSheet);
                    XlsbWorksheetProtectionProjector.Apply(targetSheet, sourceSheet.Protection);
                    XlsbWorksheetAutoFilterProjector.Apply(targetSheet, sourceSheet.AutoFilter);
                    XlsbWorksheetPrintSettingsProjector.Apply(
                        targetSheet,
                        sourceSheet.PrintOptions,
                        sourceSheet.PageMargins,
                        sourceSheet.PageSetup,
                        sourceSheet.HeaderFooter);
                    XlsbWorksheetHyperlinkProjector.Apply(targetSheet, sourceSheet);
                    targetSheet.EnsureWorksheetElementOrder();
                    targetSheet.WorksheetPart.Worksheet?.Save();

                    if (sourceSheet.State == 1) {
                        targetSheet.SetHidden(true);
                    } else if (sourceSheet.State == 2) {
                        targetSheet.SetVeryHidden(true);
                    }
                }

                relatedParts.ProjectTheme();
                relatedParts.ProjectVba();
                XlsbDefinedNameProjector.Apply(document, workbook.DefinedNames);

                if (workbook.CalculationProperties != null) {
                    XlsbCalculationPropertiesProjector.Apply(document, workbook.CalculationProperties);
                }

                return document;
            } catch {
                document.Dispose();
                throw;
            }
        }

        private static void EnsureArrayProjectionBudget(XlsbWorkbook workbook, int maxCells,
            System.Threading.CancellationToken cancellationToken) {
            long projectedCells = workbook.Sheets.Sum(sheet => (long)sheet.Cells.Count);
            if (projectedCells > maxCells) throw new InvalidDataException($"The XLSB workbook exceeds the configured limit of {maxCells} cells.");
            foreach (XlsbSheet sheet in workbook.Sheets) {
                cancellationToken.ThrowIfCancellationRequested();
                if (sheet.IsChartSheet || !sheet.FormulaGroups.Any(group => group.Type == 426)) continue;
                var formulaAnchors = new HashSet<(int Row, int Column)>();
                foreach (XlsbCell cell in sheet.Cells) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!string.IsNullOrWhiteSpace(cell.FormulaText)) formulaAnchors.Add((cell.Row, cell.Column));
                }
                foreach (var group in sheet.FormulaGroups) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (group.Type != 426 || !formulaAnchors.Contains((group.AnchorRow, group.AnchorColumn))) continue;
                    long area = (long)(group.Range.LastRow - group.Range.FirstRow + 1) *
                        (group.Range.LastColumn - group.Range.FirstColumn + 1);
                    // The anchor is already counted as an input cell. Counting other
                    // explicit members again is conservative and keeps the preflight cheap.
                    if (area - 1 > maxCells - projectedCells)
                        throw new InvalidDataException($"The XLSB array formulas exceed the configured limit of {maxCells} projected cells.");
                    projectedCells += area - 1;
                }
            }
        }
    }
}
