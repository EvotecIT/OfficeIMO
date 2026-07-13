using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesSubtotalWithMoreThanThirtyArguments() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            string formula = "SUBTOTAL(9," + string.Join(",", Enumerable.Range(1, 31).Select(row => $"A{row}")) + ")";

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SubtotalArgs");
                    for (int row = 1; row <= 31; row++) {
                        sheet.CellValue(row, 1, row);
                    }

                    sheet.CellValue(1, 2, 496d);
                    sheet.CellFormula(1, 2, formula);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCell formulaCell = Assert.Single(worksheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
                Assert.True(formulaCell.IsFormula);
                Assert.Equal(496d, Assert.IsType<double>(formulaCell.Value));
                Assert.Equal(formula, formulaCell.FormulaText);
                Assert.Contains(result.Workbook.FormulaTokenRecords, record =>
                    record.TokenName == "PtgFuncVar"
                    && string.Equals(record.FunctionName, "SUBTOTAL", StringComparison.Ordinal)
                    && record.FunctionId == 0x0158
                    && record.FunctionParameterCount == 32);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
