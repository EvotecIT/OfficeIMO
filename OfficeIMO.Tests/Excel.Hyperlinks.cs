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

        [Fact]
        [Trait("Category", "ExcelLinks")]
        public void Excel_SetHyperlink_PreservesSharedRelationshipForOtherCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelHyperlinkSharedRelationship.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Links");
                sheet.SetHyperlink(1, 1, "https://shared.example/", display: "Primary");
                sheet.SetHyperlink(1, 2, "https://shared.example/", display: "Secondary");
                document.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, true)) {
                var workbookPart = package.WorkbookPart!;
                var sheetRef = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name!.Value == "Links");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetRef.Id!);
                var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().First();
                var items = hyperlinks.Elements<Hyperlink>().ToList();
                Assert.Equal(2, items.Count);
                var sharedId = items[0].Id!.Value;
                var redundantId = items[1].Id!.Value;
                if (!string.Equals(sharedId, redundantId, StringComparison.OrdinalIgnoreCase)) {
                    var redundantRelationship = worksheetPart.HyperlinkRelationships.First(r => r.Id == redundantId);
                    worksheetPart.DeleteReferenceRelationship(redundantRelationship);
                    items[1].Id = sharedId;
                    worksheetPart.Worksheet.Save();
                }
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First(s => s.Name == "Links");
                sheet.SetHyperlink(1, 1, "https://updated.example/", display: "Updated");
                document.Save(false);
            }

            using (var package = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = package.WorkbookPart!;
                var sheetRef = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name!.Value == "Links");
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetRef.Id!);
                var hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(hyperlinks);
                var items = hyperlinks!.Elements<Hyperlink>().ToList();
                Assert.Equal(2, items.Count);

                var a1 = items.First(h => h.Reference!.Value == "A1");
                Assert.Equal("https://updated.example/", worksheetPart.HyperlinkRelationships.First(r => r.Id == a1.Id!.Value).Uri.ToString());

                var b1 = items.First(h => h.Reference!.Value == "B1");
                Assert.False(string.IsNullOrEmpty(b1.Id?.Value));
                var remainingRel = worksheetPart.HyperlinkRelationships.FirstOrDefault(r => r.Id == b1.Id!.Value);
                Assert.NotNull(remainingRel);
                Assert.Equal("https://shared.example/", remainingRel!.Uri.ToString());
            }
        }
    }
}
