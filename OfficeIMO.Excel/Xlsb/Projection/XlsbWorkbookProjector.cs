using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects the supported XLSB workbook model into OfficeIMO's normal editable workbook surface.</summary>
    internal static class XlsbWorkbookProjector {
        internal static ExcelDocument ToExcelDocument(XlsbWorkbook workbook, System.Threading.CancellationToken cancellationToken = default) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            cancellationToken.ThrowIfCancellationRequested();

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
                    ExcelSheet targetSheet = document.AddWorksheet(sourceSheet.Name);
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

                    foreach (var group in sourceSheet.FormulaGroups.Where(group => group.Type == 426)) {
                        var anchor = sourceSheet.Cells.FirstOrDefault(cell => cell.Row == group.AnchorRow && cell.Column == group.AnchorColumn);
                        if (string.IsNullOrWhiteSpace(anchor?.FormulaText)) continue;
                        targetSheet.SetLegacyArrayFormula(group.Range.ToA1Reference(), anchor!.FormulaText!);
                        foreach (var cell in targetSheet.WorksheetPart.Worksheet!.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()) {
                            if (cell.CellReference?.Value == null) continue;
                            var position = A1.ParseCellRef(cell.CellReference.Value);
                            if (group.Contains(position.Row, position.Col) && (position.Row != group.AnchorRow || position.Col != group.AnchorColumn)) cell.CellFormula = null;
                        }
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
    }
}
