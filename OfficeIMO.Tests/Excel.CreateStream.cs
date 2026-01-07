using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Create_ToMemoryStream_AutoSaveWritesPackage() {
            using var memory = new MemoryStream();
            using (var document = ExcelDocument.Create(memory)) {
                var sheet = document.AddWorkSheet("StreamData");
                sheet.CellValue(1, 1, "Hello Stream");
            }

            Assert.True(memory.Length > 0);
            memory.Position = 0;

            using var spreadsheet = SpreadsheetDocument.Open(memory, false);
            Assert.NotNull(spreadsheet.WorkbookPart);
            var sheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().ToList();
            var sheetInfo = Assert.Single(sheets);
            Assert.Equal("StreamData", sheetInfo.Name?.Value);
        }
    }
}
