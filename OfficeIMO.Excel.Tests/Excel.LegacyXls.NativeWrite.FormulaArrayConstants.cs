using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaArrayConstants() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("ArrayConstants");
                    sheet.CellValue(1, 1, 6d);
                    sheet.CellFormula(1, 1, "SUM({1,2,3})");
                    sheet.CellValue(2, 1, 2d);
                    sheet.CellFormula(2, 1, "SUM({TRUE,FALSE,1})");
                    sheet.CellValue(3, 1, 14d);
                    sheet.CellFormula(3, 1, "SUM(B1:B3*{1;2;3})");
                    sheet.CellValue(4, 1, "Blue");
                    sheet.CellFormula(4, 1, "INDEX({\"Red\",\"Blue\"},2)");
                    sheet.CellValue(5, 1, "#N/A");
                    sheet.CellFormula(5, 1, "INDEX({#N/A,5},1)");
                    SetOpenXmlErrorCell(sheet, "A5", "#N/A");
                    sheet.CellValue(6, 1, 3d);
                    sheet.CellFormula(6, 1, "IF(B1>0,SUM({1,2}),0)");
                    sheet.CellValue(7, 1, 0d);
                    sheet.CellFormula(7, 1, "IF(B1<0,SUM({1,2}))");
                    sheet.CellValue(8, 1, 3d);
                    sheet.CellFormula(8, 1, "CHOOSE(B1,SUM({1,2}),0)");
                    sheet.CellValue(9, 1, 7d);
                    sheet.CellFormula(9, 1, "CHOOSE(2,SUM({1,2}),SUM({3,4}))");
                    sheet.CellValue(1, 2, 1d);
                    sheet.CellValue(2, 2, 2d);
                    sheet.CellValue(3, 2, 3d);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));
                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);

                AssertNumericFormula(worksheet, 1, 6d, "SUM({1,2,3})");
                AssertNumericFormula(worksheet, 2, 2d, "SUM({TRUE,FALSE,1})");
                AssertNumericFormula(worksheet, 3, 14d, "SUM(B1:B3*{1;2;3})");
                AssertTextFormula(worksheet, 4, "Blue", "INDEX({\"Red\",\"Blue\"},2)");
                AssertNumericFormula(worksheet, 6, 3d, "IF(B1>0,SUM({1,2}),0)");
                AssertNumericFormula(worksheet, 7, 0d, "IF(B1<0,SUM({1,2}))");
                AssertNumericFormula(worksheet, 8, 3d, "CHOOSE(B1,SUM({1,2}),0)");
                AssertNumericFormula(worksheet, 9, 7d, "CHOOSE(2,SUM({1,2}),SUM({3,4}))");

                LegacyXlsCell errorFormula = Assert.Single(worksheet.Cells, cell => cell.Row == 5 && cell.Column == 1);
                Assert.True(errorFormula.IsFormula);
                Assert.Equal(LegacyXlsCellValueKind.Error, errorFormula.Kind);
                Assert.Equal("#N/A", Assert.IsType<string>(errorFormula.Value));
                Assert.Equal("INDEX({#N/A,5},1)", errorFormula.FormulaText);

                Assert.Equal(10, result.Workbook.FormulaTokenRecords.Count(token => token.OperandKind == "ArrayLiteral"));
                Assert.Contains(result.Workbook.FormulaTokenRecords, record =>
                    record.TokenName == "PtgFuncVar"
                    && string.Equals(record.FunctionName, "IF", StringComparison.Ordinal)
                    && record.FunctionId == 0x0001
                    && record.FunctionParameterCount == 2);
                Assert.Contains(result.Workbook.FormulaTokenRecords, record =>
                    record.TokenName == "PtgFuncVar"
                    && string.Equals(record.FunctionName, "IF", StringComparison.Ordinal)
                    && record.FunctionId == 0x0001
                    && record.FunctionParameterCount == 3);
                Assert.Contains(result.Workbook.FormulaTokenRecords, record =>
                    record.TokenName == "PtgFuncVar"
                    && string.Equals(record.FunctionName, "CHOOSE", StringComparison.Ordinal)
                    && record.FunctionId == 0x0064
                    && record.FunctionParameterCount == 3);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
