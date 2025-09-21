using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        [Trait("Category", "ExcelLinks")]
        public void Excel_SetHyperlink_ReplacesExistingEntry() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelHyperlinkReplace.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Links");
                sheet.SetHyperlink(1, 1, "https://initial.example/", display: "First");
                sheet.SetHyperlink(1, 1, "https://final.example/", display: "Second");
                document.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = package.WorkbookPart!;
                var sheetRef = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name!.Value == "Links");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetRef.Id!);
                var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                Assert.Single(items);
                var hyperlink = items[0];
                Assert.Equal("A1", hyperlink.Reference!.Value);
                var relationships = worksheetPart.HyperlinkRelationships.ToList();
                Assert.Single(relationships);
                Assert.Equal("https://final.example/", relationships[0].Uri.ToString());
            }
        }

        [Fact]
        [Trait("Category", "ExcelLinks")]
        public void Excel_SetInternalLink_ReplacesExistingEntry() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelInternalLinkReplace.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Links");
                var target1 = document.AddWorkSheet("Target1");
                var target2 = document.AddWorkSheet("Target2");
                sheet.SetInternalLink(2, 1, target1, "A1", display: "First");
                sheet.SetInternalLink(2, 1, target2, "B5", display: "Second");
                document.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = package.WorkbookPart!;
                var sheetRef = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name!.Value == "Links");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetRef.Id!);
                var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                Assert.Single(items);
                var hyperlink = items[0];
                Assert.Equal("A2", hyperlink.Reference!.Value);
                Assert.Equal("'Target2'!B5", hyperlink.Location!.Value);
                Assert.Null(hyperlink.Id);
                Assert.Empty(worksheetPart.HyperlinkRelationships);
            }
        }
    }
}
