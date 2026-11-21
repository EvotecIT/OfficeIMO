using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // Binary numeric fields hold finite numbers without a negative zero, so filter text that reads as an
        // infinity stays text and "-0" is stored as zero.
        [Fact]
        public void LegacyXls_NativeSave_KeepsOverflowingFilterTextAsTextAndStoresNegativeZeroAsZero() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");
            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddSheet("Filter");
                    sheet.Cell(1, 1).SetValue("Value");
                    sheet.Cell(2, 1).SetValue(1d);
                    sheet.Cell(3, 1).SetValue(2d);
                    sheet.Range("A1:A3").AutoFilter();
                    sheet.ApplyAutoFilterCustomCriteria(
                        "A1:A3",
                        0,
                        matchAll: false,
                        new[] { (FilterOperatorValues.Equal, "1e400"), (FilterOperatorValues.Equal, "-0") });
                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                Assert.False(result.HasImportErrors);
                LegacyXlsAutoFilterCriteria criteria = Assert.Single(Assert.Single(result.Workbook.Worksheets).AutoFilterCriteria);
                Assert.Contains(criteria.Conditions, condition =>
                    condition.ValueKind == LegacyXlsAutoFilterValueKind.Text && condition.Value == "1e400");
                Assert.Contains(criteria.Conditions, condition =>
                    condition.ValueKind == LegacyXlsAutoFilterValueKind.Number && condition.Value == "0");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void Xlsb_NativeSave_StoresNegativeZeroFilterValueAsZero() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Filter");
                sheet.Cell(1, 1).SetValue("Value");
                sheet.Cell(2, 1).SetValue(1d);
                sheet.Range("A1:A2").AutoFilter();
                sheet.ApplyAutoFilterCustomCriteria(
                    "A1:A2",
                    0,
                    matchAll: true,
                    new[] { (FilterOperatorValues.GreaterThan, "-0") });
                source = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
            }

            using var archive = new ZipArchive(new MemoryStream(source, writable: false), ZipArchiveMode.Read);
            using Stream worksheetStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/worksheets/sheet1.bin")).Open();
            XlsbRecord custom = Assert.Single(XlsbRecordReader.ReadAll(worksheetStream),
                record => record.Type == XlsbAutoFilterCriteria.BrtCustomFilter);
            Assert.Equal(XlsbAutoFilterCriteria.NumberValue, custom.Data[0]);
            Assert.Equal(new byte[8], custom.Data.Skip(2).Take(8).ToArray());
        }

        [Fact]
        public void Xlsb_NativeSave_KeepsOverflowingFilterTextAsText() {
            byte[] source;
            using (ExcelDocument document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Filter");
                sheet.Cell(1, 1).SetValue("Value");
                sheet.Cell(2, 1).SetValue(1d);
                sheet.Range("A1:A2").AutoFilter();
                sheet.ApplyAutoFilterCustomCriteria(
                    "A1:A2",
                    0,
                    matchAll: true,
                    new[] { (FilterOperatorValues.Equal, "1e400") });
                source = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
            }

            using var archive = new ZipArchive(new MemoryStream(source, writable: false), ZipArchiveMode.Read);
            using Stream worksheetStream = Assert.IsType<ZipArchiveEntry>(archive.GetEntry("xl/worksheets/sheet1.bin")).Open();
            XlsbRecord custom = Assert.Single(XlsbRecordReader.ReadAll(worksheetStream),
                record => record.Type == XlsbAutoFilterCriteria.BrtCustomFilter);
            Assert.Equal(XlsbAutoFilterCriteria.StringValue, custom.Data[0]);
            Assert.Contains("1e400", System.Text.Encoding.Unicode.GetString(custom.Data));
        }
    }
}
