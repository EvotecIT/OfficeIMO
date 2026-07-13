using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;
using OpenXmlCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesArrayFormulaTextCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ArrayText");
                    sheet.CellValue(1, 2, "AB");
                    sheet.SetArrayFormula("B1:B2", "CONCATENATE(\"A\",\"B\")");
                    sheet.CellValue(2, 2, "AB");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertTextArrayFormulaCell(worksheet, 1, 2, "AB", "CONCATENATE(\"A\",\"B\")");
                AssertTextArrayFormulaCell(worksheet, 2, 2, "AB", "CONCATENATE(\"A\",\"B\")");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("B1:B2", arrayFormula.Range);
                Assert.Equal(2, arrayFormula.DeclaredCellCount);
                Assert.Equal(2, arrayFormula.MatchedFormulaCellCount);
                Assert.True(arrayFormula.FormulaTextProjected);

                using ExcelDocument normalLoaded = ExcelDocument.Load(xlsOutputPath);
                ExcelSheet loadedSheet = normalLoaded.Sheets.Single();
                Assert.True(loadedSheet.TryGetCellText(1, 2, out string firstText));
                Assert.True(loadedSheet.TryGetCellText(2, 2, out string secondText));
                Assert.Equal("AB", firstText);
                Assert.Equal("AB", secondText);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesArrayFormulaDateCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ArrayDates");
                    sheet.CellValue(1, 2, "2026-01-02T00:00:00");
                    sheet.SetArrayFormula("B1:B2", "DATE(2026,1,2)");
                    sheet.CellValue(2, 2, "2026-01-02T00:00:00");

                    foreach (string reference in new[] { "B1", "B2" }) {
                        OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                            .Single(item => string.Equals(item.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
                        cell.DataType = OpenXmlCellValues.Date;
                        cell.CellValue = new OpenXmlCellValue("2026-01-02T00:00:00");
                    }

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                double expectedSerial = new DateTime(2026, 1, 2).ToOADate();
                AssertArrayFormulaCell(worksheet, 1, 2, expectedSerial, "DATE(2026,1,2)");
                AssertArrayFormulaCell(worksheet, 2, 2, expectedSerial, "DATE(2026,1,2)");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("B1:B2", arrayFormula.Range);
                Assert.Equal(2, arrayFormula.DeclaredCellCount);
                Assert.Equal(2, arrayFormula.MatchedFormulaCellCount);
                Assert.True(arrayFormula.FormulaTextProjected);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesArrayFormulaBooleanCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ArrayBoolean");
                    sheet.CellValue(1, 2, true);
                    sheet.SetArrayFormula("B1:B2", "A1=1");
                    sheet.CellValue(2, 2, false);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertBooleanArrayFormulaCell(worksheet, 1, 2, expectedValue: true, "A1=1");
                AssertBooleanArrayFormulaCell(worksheet, 2, 2, expectedValue: false, "A1=1");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("B1:B2", arrayFormula.Range);
                Assert.Equal(2, arrayFormula.DeclaredCellCount);
                Assert.Equal(2, arrayFormula.MatchedFormulaCellCount);
                Assert.True(arrayFormula.FormulaTextProjected);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesArrayFormulaErrorCachedResults() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ArrayErrors");
                    sheet.CellValue(1, 2, "#N/A");
                    sheet.SetArrayFormula("B1:B2", "NA()");
                    sheet.CellValue(2, 2, "#N/A");

                    foreach (string reference in new[] { "B1", "B2" }) {
                        OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                            .Single(item => string.Equals(item.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
                        cell.DataType = OpenXmlCellValues.Error;
                        cell.CellValue = new OpenXmlCellValue("#N/A");
                    }

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertErrorArrayFormulaCell(worksheet, 1, 2, "#N/A", "NA()");
                AssertErrorArrayFormulaCell(worksheet, 2, 2, "#N/A", "NA()");

                LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(worksheet.ArrayFormulaRecords);
                Assert.Equal("B1:B2", arrayFormula.Range);
                Assert.Equal(2, arrayFormula.DeclaredCellCount);
                Assert.Equal(2, arrayFormula.MatchedFormulaCellCount);
                Assert.True(arrayFormula.FormulaTextProjected);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        private static void AssertTextArrayFormulaCell(LegacyXlsWorksheet worksheet, int row, int column, string expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == column);
            Assert.True(cell.IsFormula);
            Assert.Equal(LegacyXlsCellValueKind.Text, cell.Kind);
            Assert.Equal(expectedValue, Assert.IsType<string>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertBooleanArrayFormulaCell(LegacyXlsWorksheet worksheet, int row, int column, bool expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == column);
            Assert.True(cell.IsFormula);
            Assert.Equal(LegacyXlsCellValueKind.Boolean, cell.Kind);
            Assert.Equal(expectedValue, Assert.IsType<bool>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static void AssertErrorArrayFormulaCell(LegacyXlsWorksheet worksheet, int row, int column, string expectedValue, string expectedFormulaText) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == column);
            Assert.True(cell.IsFormula);
            Assert.Equal(LegacyXlsCellValueKind.Error, cell.Kind);
            Assert.Equal(expectedValue, Assert.IsType<string>(cell.Value));
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }
    }
}
