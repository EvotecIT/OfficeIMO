using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Tests.TestStreams;
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

        [Fact]
        public void Test_Save_ToReusedMemoryStream_ReplacesExistingContent() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveToReusedStream.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                var sheet = document.AddWorkSheet("StreamData");
                sheet.CellValue(1, 1, "Fresh");

                using var memory = new MemoryStream();
                memory.Write(new byte[256], 0, 256);
                memory.Position = memory.Length;

                document.Save(memory, new ExcelSaveOptions { ValidateOpenXml = true });

                memory.Position = 0;
                using var spreadsheet = SpreadsheetDocument.Open(memory, false);
                var savedSheet = Assert.Single(spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>());
                Assert.Equal("StreamData", savedSheet.Name?.Value);
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_Save_ToFailingStream_KeepsDocumentUsable() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveToFailingStream.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                document.AddWorkSheet("Original");

                using var failing = new ThrowAfterBytesWriteStream(64);
                Assert.Throws<IOException>(() => document.Save(failing));

                document.AddWorkSheet("Recovered");

                using var memory = new MemoryStream();
                document.Save(memory, new ExcelSaveOptions { ValidateOpenXml = true });

                memory.Position = 0;
                using var spreadsheet = SpreadsheetDocument.Open(memory, false);
                var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
                Assert.Equal(2, sheets.Count);
                Assert.Contains(sheets, s => s.Name?.Value == "Original");
                Assert.Contains(sheets, s => s.Name?.Value == "Recovered");
            }
            finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_Save_ToReadOnlyPath_PreservesExistingFile_AndKeepsDocumentUsable() {
            string sourcePath = Path.Combine(_directoryWithFiles, "SaveToLockedPath.Source.xlsx");
            string destinationPath = Path.Combine(_directoryWithFiles, "SaveToLockedPath.Destination.xlsx");

            try {
                using (var existing = ExcelDocument.Create(destinationPath)) {
                    existing.AddWorkSheet("Original");
                    existing.Save(destinationPath, openExcel: false);
                }

                var destinationFile = new FileInfo(destinationPath);
                destinationFile.IsReadOnly = true;

                using var document = ExcelDocument.Create(sourcePath);
                document.AddWorkSheet("Updated");

                var exception = Assert.Throws<IOException>(() => document.Save(destinationPath, openExcel: false));
                Assert.Contains("read-only", exception.Message, StringComparison.OrdinalIgnoreCase);

                using (var spreadsheet = SpreadsheetDocument.Open(destinationPath, false)) {
                    var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
                    var savedSheet = Assert.Single(sheets);
                    Assert.Equal("Original", savedSheet.Name?.Value);
                }

                document.AddWorkSheet("Recovered");

                using var memory = new MemoryStream();
                document.Save(memory, new ExcelSaveOptions { ValidateOpenXml = true });

                memory.Position = 0;
                using var recoveredSpreadsheet = SpreadsheetDocument.Open(memory, false);
                var recoveredSheets = recoveredSpreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
                Assert.Equal(2, recoveredSheets.Count);
                Assert.Contains(recoveredSheets, s => s.Name?.Value == "Updated");
                Assert.Contains(recoveredSheets, s => s.Name?.Value == "Recovered");
            }
            finally {
                if (File.Exists(sourcePath)) {
                    File.Delete(sourcePath);
                }

                if (File.Exists(destinationPath)) {
                    new FileInfo(destinationPath).IsReadOnly = false;
                    File.Delete(destinationPath);
                }
            }
        }
    }
}
