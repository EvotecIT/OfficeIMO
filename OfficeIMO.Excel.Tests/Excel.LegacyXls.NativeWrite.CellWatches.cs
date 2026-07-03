using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesCellWatches() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("CellWatches");
                    sheet.CellValue(3, 3, "Watched");
                    sheet.WorksheetPart.Worksheet.Append(new CellWatches(
                        new CellWatch { CellReference = "C3" }));
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsCellWatch cellWatch = Assert.Single(worksheet.CellWatches);
                Assert.Equal("C3", cellWatch.CellReference);
                Assert.Equal(3, cellWatch.Row);
                Assert.Equal(3, cellWatch.Column);

                CellWatch projectedCellWatch = result.Document.Sheets.Single()
                    .WorksheetPart.Worksheet
                    .Elements<CellWatches>()
                    .Single()
                    .Elements<CellWatch>()
                    .Single();
                Assert.Equal("C3", projectedCellWatch.CellReference!.Value);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }
    }
}
