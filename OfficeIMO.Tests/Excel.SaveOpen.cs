using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_Save_OpensWithoutSharingViolation() {
            string filePath = Path.Combine(_directoryWithFiles, "SaveOpen.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Test");

                try {
                    document.Save(true);
                } catch {
                    // Opening the file may fail on systems without associated application
                }

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                }

                document.AddWorkSheet("Second");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Assert.Equal(2, spreadsheet.WorkbookPart.Workbook.Sheets.OfType<Sheet>().Count());
            }
        }
    }
}

