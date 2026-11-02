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
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddSheet("ArrayConstants");
                    sheet.Cell(1, 1).SetValue(6d);
                    sheet.Cell(1, 1).SetFormula("SUM({1,2,3})");
                    sheet.Cell(2, 1).SetValue(2d);
                    sheet.Cell(2, 1).SetFormula("SUM({TRUE,FALSE,1})");
                    sheet.Cell(3, 1).SetValue(14d);
                    sheet.Cell(3, 1).SetFormula("SUM(B1:B3*{1;2;3})");
                    sheet.Cell(4, 1).SetValue("Blue");
                    sheet.Cell(4, 1).SetFormula("INDEX({\"Red\",\"Blue\"},2)");
                    sheet.Cell(5, 1).SetValue("#N/A");
                    sheet.Cell(5, 1).SetFormula("INDEX({#N/A,5},1)");
                    SetOpenXmlErrorCell(sheet, "A5", "#N/A");
                    sheet.Cell(6, 1).SetValue(3d);
                    sheet.Cell(6, 1).SetFormula("IF(B1>0,SUM({1,2}),0)");
                    sheet.Cell(7, 1).SetValue(0d);
                    sheet.Cell(7, 1).SetFormula("IF(B1<0,SUM({1,2}))");
                    sheet.Cell(8, 1).SetValue(3d);
                    sheet.Cell(8, 1).SetFormula("CHOOSE(B1,SUM({1,2}),0)");
                    sheet.Cell(9, 1).SetValue(7d);
                    sheet.Cell(9, 1).SetFormula("CHOOSE(2,SUM({1,2}),SUM({3,4}))");
                    sheet.Cell(1, 2).SetValue(1d);
                    sheet.Cell(2, 2).SetValue(2d);
                    sheet.Cell(3, 2).SetValue(3d);

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
