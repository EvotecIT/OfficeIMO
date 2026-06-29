using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesXlmChartObjectAndEvaluationInformationFunctions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("XlmInfo");

                    sheet.CellValue(1, 1, "ChartItem");
                    sheet.CellFormula(1, 1, "GET.CHART.ITEM(1)");
                    sheet.CellValue(2, 1, "ObjectInfo");
                    sheet.CellFormula(2, 1, "GET.OBJECT(1)");
                    sheet.CellValue(3, 1, "TextBox");
                    sheet.CellFormula(3, 1, "TEXT.BOX(\"Caption\")");
                    sheet.CellValue(4, 1, 2d);
                    sheet.CellFormula(4, 1, "EVALUATE(\"1+1\")");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertTextFormula(worksheet, 1, "ChartItem", "GET.CHART.ITEM(1)");
                AssertTextFormula(worksheet, 2, "ObjectInfo", "GET.OBJECT(1)");
                AssertTextFormula(worksheet, 3, "TextBox", "TEXT.BOX(\"Caption\")");
                AssertNumericFormula(worksheet, 4, 2d, "EVALUATE(\"1+1\")");
                AssertVariableFunctionToken(result.Workbook, "GET.CHART.ITEM", 0x00a0, 1);
                AssertVariableFunctionToken(result.Workbook, "GET.OBJECT", 0x00f6, 1);
                AssertVariableFunctionToken(result.Workbook, "TEXT.BOX", 0x00f3, 1);
                AssertVariableFunctionToken(result.Workbook, "EVALUATE", 0x0101, 1);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
