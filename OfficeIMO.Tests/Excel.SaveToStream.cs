using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_Save_ToMemoryStream_PackageIsReadable() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveToStream.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                var sheet = document.AddWorkSheet("StreamData");
                sheet.CellValue(1, 1, "Hello Stream");

                using var memory = new MemoryStream();
                document.Save(memory, new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.True(memory.Length > 0);

                // Document should remain usable after stream save
                document.AddWorkSheet("PostSave");
                Assert.Equal(2, document.Sheets.Count);

                memory.Position = 0;
                using var spreadsheet = SpreadsheetDocument.Open(memory, false);
                Assert.NotNull(spreadsheet.WorkbookPart);
                Assert.NotNull(spreadsheet.WorkbookPart!.Workbook);
                var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
                var sheetInfo = Assert.Single(sheets);
                Assert.Equal("StreamData", sheetInfo.Name?.Value);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_SaveAsync_ToMemoryStream_PackageIsReadable() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveToStreamAsync.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                var sheet = document.AddWorkSheet("AsyncStream");
                sheet.CellValue(2, 2, 42);

                using var memory = new MemoryStream();
                await document.SaveAsync(memory, new ExcelSaveOptions { ValidateOpenXml = true });
                Assert.True(memory.Length > 0);

                document.AddWorkSheet("PostAsyncSave");
                Assert.Equal(2, document.Sheets.Count);

                memory.Position = 0;
                using var spreadsheet = SpreadsheetDocument.Open(memory, false);
                Assert.NotNull(spreadsheet.WorkbookPart);
                var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
                var sheetInfo = Assert.Single(sheets);
                Assert.Equal("AsyncStream", sheetInfo.Name?.Value);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
