using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects the supported XLSB workbook model into OfficeIMO's normal editable workbook surface.</summary>
    internal static class XlsbWorkbookProjector {
        internal static ExcelDocument ToExcelDocument(XlsbWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            ExcelDocument document = ExcelDocument.Create();
            if (workbook.Uses1904DateSystem) {
                document.DateSystem = ExcelDateSystem.NineteenFour;
            }
            if (workbook.Stylesheet != null) {
                XlsbStylesheetProjector.Install(document, workbook.Stylesheet);
            }
            if (workbook.WorkbookProtection != null) {
                XlsbWorkbookProtectionProjector.Apply(document, workbook.WorkbookProtection);
            }

            foreach (XlsbWorksheet sourceSheet in workbook.Worksheets) {
                ExcelSheet targetSheet = document.AddWorksheet(sourceSheet.Name);
                targetSheet.Batch(sheet => {
                    foreach (XlsbCell cell in sourceSheet.Cells) {
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

                XlsbWorksheetGeometryProjector.Apply(targetSheet, sourceSheet);
                XlsbWorksheetHyperlinkProjector.Apply(targetSheet, sourceSheet);

                if (sourceSheet.State == 1) {
                    targetSheet.SetHidden(true);
                } else if (sourceSheet.State == 2) {
                    targetSheet.SetVeryHidden(true);
                }
            }

            XlsbDefinedNameProjector.Apply(document, workbook.DefinedNames);

            if (workbook.CalculationProperties != null) {
                XlsbCalculationPropertiesProjector.Apply(document, workbook.CalculationProperties);
            }

            return document;
        }
    }
}
