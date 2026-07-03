using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaReferenceOperators() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("ReferenceOps");
                    sheet.CellValue(1, 2, 10d);
                    sheet.CellValue(1, 3, 20d);
                    sheet.CellValue(1, 4, 30d);
                    sheet.CellValue(1, 1, 30d);
                    sheet.CellFormula(1, 1, "SUM((B1:C1))");
                    sheet.CellValue(2, 1, 30d);
                    sheet.CellFormula(2, 1, "SUM((B1,C1))");
                    sheet.CellValue(3, 1, 20d);
                    sheet.CellFormula(3, 1, "SUM((B1:D1 C1:C1))");
                    sheet.CellValue(4, 1, 30d);
                    sheet.CellFormula(4, 1, "SUM(B:C)");
                    sheet.CellValue(5, 1, 4d);
                    sheet.CellFormula(5, 1, "COUNT(1:1)");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                AssertNumericFormula(worksheet, 1, 30d, "SUM((B1:C1))");
                AssertNumericFormula(worksheet, 2, 30d, "SUM((B1,C1))");
                AssertNumericFormula(worksheet, 3, 20d, "SUM((B1:D1 C1:C1))");
                AssertNumericFormula(worksheet, 4, 30d, "SUM(B:C)");
                AssertNumericFormula(worksheet, 5, 4d, "COUNT(1:1)");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token => token.TokenName == "PtgRange");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token => token.TokenName == "PtgUnion");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token => token.TokenName == "PtgIsect");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgArea"
                    && token.OperandText == "B:C");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgArea"
                    && token.OperandText == "1:1");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesWorkbookInternal3dReferenceUnions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet first = document.AddWorkSheet("Region 1");
                    ExcelSheet second = document.AddWorkSheet("Region 2");
                    ExcelSheet third = document.AddWorkSheet("Region 3");
                    ExcelSheet fourth = document.AddWorkSheet("Region 4");
                    ExcelSheet calc = document.AddWorkSheet("Calc");

                    first.CellValue(1, 1, 1d);
                    second.CellValue(1, 1, 2d);
                    third.CellValue(1, 1, 3d);
                    fourth.CellValue(1, 1, 4d);
                    calc.CellValue(1, 1, 7d);
                    calc.CellFormula(1, 1, "SUM(('Region 1:Region 2'!A1,'Region 4'!A1))");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Calc");

                AssertNumericFormula(worksheet, 1, 7d, "SUM(('Region 1:Region 2'!A1,'Region 4'!A1))");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token => token.TokenName == "PtgUnion");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgRef3d"
                    && token.OperandText == "ExternSheet:3;Reference:A1");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgRef3d"
                    && token.OperandText == "ExternSheet:5;Reference:A1");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
