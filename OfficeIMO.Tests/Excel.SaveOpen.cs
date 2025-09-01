using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact()]
        public void Test_Save_OpensWithoutSharingViolation() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveOpen.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Test");

                // Attempt to save the document, do not open, as it may fail on systems without associated application, and not really great for testing
                try {
                    document.Save(false);
                } catch {
                    // Opening the file may fail on systems without associated application
                }

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                }

                document.AddWorkSheet("Second");
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);
                Assert.NotNull(spreadsheet.WorkbookPart);
                Assert.NotNull(spreadsheet.WorkbookPart!.Workbook);
                Assert.NotNull(spreadsheet.WorkbookPart!.Workbook!.Sheets);
                Assert.Equal(2, spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().Count());
            }
        }
    }
}

