using OpenXmlCell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using OpenXmlCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;
using OpenXmlCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidOpenXmlDateCellsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("date value 'not-a-date'", (document, sheet) => {
                sheet.CellValue(1, 1, "not-a-date");
                SetOpenXmlDateCell(sheet, "A1", "not-a-date");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksInvalidOpenXmlFormulaCachedDatesBeforeWriting() {
            AssertNativeXlsSaveNotSupported("cached date result 'not-a-date'", (document, sheet) => {
                sheet.CellValue(1, 1, 1d);
                sheet.CellFormula(2, 1, "A1+1");

                OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                    .Single(item => string.Equals(item.CellReference?.Value, "A2", StringComparison.OrdinalIgnoreCase));
                cell.DataType = OpenXmlCellValues.Date;
                cell.CellValue = new OpenXmlCellValue("not-a-date");
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesOpenXmlFormulaCachedDates() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("FormulaDates");
                    sheet.CellValue(1, 1, "2026-01-02T00:00:00");
                    sheet.CellFormula(1, 1, "DATE(2026,1,2)");

                    OpenXmlCell cell = sheet.WorksheetPart.Worksheet.Descendants<OpenXmlCell>()
                        .Single(item => string.Equals(item.CellReference?.Value, "A1", StringComparison.OrdinalIgnoreCase));
                    cell.DataType = OpenXmlCellValues.Date;
                    cell.CellValue = new OpenXmlCellValue("2026-01-02T00:00:00");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertNumericFormula(worksheet, 1, new DateTime(2026, 1, 2).ToOADate(), "DATE(2026,1,2)");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
