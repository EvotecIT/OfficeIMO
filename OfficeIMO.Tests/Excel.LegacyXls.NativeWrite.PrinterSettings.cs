using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesWorksheetPrinterSettings() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath, autoSave: false)) {
                    ExcelSheet sheet = document.AddWorkSheet("PrinterSettings");
                    sheet.CellValue(1, 1, "Printer settings");

                    SpreadsheetPrinterSettingsPart printerSettingsPart = sheet.WorksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>();
                    using var stream = new MemoryStream(new byte[] { 0x01, 0x02, 0x03, 0x04 });
                    printerSettingsPart.FeedData(stream);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Contains(worksheet.MetadataRecords, record => record.Kind == LegacyXlsWorksheetMetadataKind.PrinterSettings);
                Assert.Equal(1, result.ImportReport.WorksheetMetadataRecordsByKind[LegacyXlsWorksheetMetadataKind.PrinterSettings]);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedWorksheetPrinterSettingsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("printer settings payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized printer settings");

                SpreadsheetPrinterSettingsPart printerSettingsPart = sheet.WorksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>();
                using var stream = new MemoryStream(new byte[65534]);
                printerSettingsPart.FeedData(stream);
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksMultipleWorksheetPrinterSettingsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("multiple worksheet printer settings parts", (document, sheet) => {
                sheet.CellValue(1, 1, "Multiple printer settings");

                SpreadsheetPrinterSettingsPart firstPrinterSettingsPart = sheet.WorksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>();
                using var firstStream = new MemoryStream(new byte[] { 0x01, 0x02 });
                firstPrinterSettingsPart.FeedData(firstStream);

                SpreadsheetPrinterSettingsPart secondPrinterSettingsPart = sheet.WorksheetPart.AddNewPart<SpreadsheetPrinterSettingsPart>();
                using var secondStream = new MemoryStream(new byte[] { 0x03, 0x04 });
                secondPrinterSettingsPart.FeedData(secondStream);
            });
        }
    }
}
