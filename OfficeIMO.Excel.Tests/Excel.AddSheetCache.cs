using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void SheetsAddedOneAtATimeStayInWorkbookOrderWithDistinctIds(bool cacheSheets) {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.SheetCachingEnabled = cacheSheets;
                for (int i = 0; i < 4; i++) {
                    ExcelSheet added = document.AddSheet($"S{i}");
                    // Reading the collection between adds must see the sheet just added.
                    Assert.Equal(added.Name, document.Sheets[document.Sheets.Count - 1].Name);
                    added.Cell(1, 1).SetValue(i);
                }

                document.RemoveWorksheet("S1");
                document.AddSheet("S4").Cell(1, 1).SetValue(4);
                Assert.Equal(new[] { "S0", "S2", "S3", "S4" }, document.Sheets.Select(sheet => sheet.Name).ToArray());
                document.Save();
            }

            stream.Position = 0;
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false)) {
                Sheet[] sheets = package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().ToArray();
                Assert.Equal(new[] { "S0", "S2", "S3", "S4" }, sheets.Select(sheet => sheet.Name!.Value).ToArray());
                Assert.Equal(sheets.Length, sheets.Select(sheet => sheet.SheetId!.Value).Distinct().Count());
            }

            stream.Position = 0;
            using var loaded = ExcelDocument.Load(stream);
            Assert.Equal(new[] { "0", "2", "3", "4" },
                loaded.Sheets.Select(sheet => sheet.TryGetCellText(1, 1, out string text) ? text : null).ToArray());
        }

        [Fact]
        public void AddSheetSeesSheetsReorderedThroughTheOpenXmlDocument() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddSheet("A");
            document.AddSheet("B");
            _ = document.Sheets;

            Sheets sheets = document.OpenXmlDocument.WorkbookPart!.Workbook.Sheets!;
            Sheet b = sheets.Elements<Sheet>().Single(sheet => sheet.Name == "B");
            b.Remove();
            sheets.InsertAt(b, 0);
            document.AddSheet("C");

            Assert.Equal(new[] { "B", "A", "C" }, document.Sheets.Select(sheet => sheet.Name).ToArray());
        }

        [Fact]
        public void AddSheetSeesSheetIdsChangedThroughTheOpenXmlDocument() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                document.AddSheet("A");
                document.AddSheet("B");
                _ = document.Sheets;

                Sheet b = document.OpenXmlDocument.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "B");
                b.SheetId = 3U;
                document.AddSheet("C");
                document.Save();
            }

            stream.Position = 0;
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            uint[] ids = package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Select(sheet => sheet.SheetId!.Value).ToArray();
            Assert.Equal(3, ids.Distinct().Count());
        }

        [Fact]
        public void AddSheetSeesSheetPartsSwappedThroughTheOpenXmlDocument() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddSheet("A").Cell(1, 1).SetValue("in A");
            document.AddSheet("B").Cell(1, 1).SetValue("in B");
            _ = document.Sheets;

            Sheet[] elements = document.OpenXmlDocument.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().ToArray();
            (elements[0].Id, elements[1].Id) = (elements[1].Id!.Value, elements[0].Id!.Value);
            document.AddSheet("C");

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string text));
            Assert.Equal("in B", text);
        }
    }
}
