using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ExternalDocumentBookmarkSurvivesInspectionAndReapplication() {
            string path = Path.Combine(_directoryWithFiles, "ExternalBookmark.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Links");
                sheet.SetHyperlink("A1", "doc1.docx", "Document", style: false, tooltip: "Bookmark");
                sheet.SetHyperlink("A2", "doc1.docx", "Whole document", style: false);
                document.Save();
            }
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
                worksheet.Descendants<Hyperlink>().Single(link => link.Reference!.Value == "A1").Location = "b3";
                worksheet.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                var sheet = document.Sheets[0];
                var link = sheet.GetHyperlinks()["A1"];
                Assert.True(link.IsExternal);
                Assert.Equal("doc1.docx#b3", link.Target);
                Assert.Equal("Bookmark", link.Tooltip);
                Assert.Equal("doc1.docx", sheet.GetHyperlinks()["A2"].Target);
                var snapshot = document.Inspect().Worksheets.Single().Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
                Assert.Equal("doc1.docx#b3", snapshot.Hyperlink!.Target);
                sheet.SetHyperlink("B1", link.Target, "Same bookmark", style: false);
                document.Save();
            }

            using var saved = SpreadsheetDocument.Open(path, false);
            var part = saved.WorkbookPart!.WorksheetParts.Single();
            var original = part.Worksheet.Descendants<Hyperlink>().Single(link => link.Reference!.Value == "A1");
            Assert.Equal("b3", original.Location!.Value);
            Assert.Equal("doc1.docx", part.HyperlinkRelationships.Single(rel => rel.Id == original.Id!.Value).Uri.OriginalString);
            var reapplied = part.Worksheet.Descendants<Hyperlink>().Single(link => link.Reference!.Value == "B1");
            Assert.Equal("doc1.docx#b3", part.HyperlinkRelationships.Single(rel => rel.Id == reapplied.Id!.Value).Uri.OriginalString);
        }
    }
}
