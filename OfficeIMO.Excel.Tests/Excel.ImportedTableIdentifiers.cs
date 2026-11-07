using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ImportedTableRangeMatchesNameAndDistinctDisplayNameWithoutRenaming() {
            string path = Path.Combine(_directoryWithFiles, "ImportedTableIdentifiers.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(2, 1).SetValue("Alpha");
                sheet.AddTable("A1:A2", name: "Original");
                document.Save();
            }
            using (var package = SpreadsheetDocument.Open(path, true)) {
                var table = package.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
                table.Name = "1";
                table.DisplayName = "_1";
                table.Save();
            }

            using (var document = ExcelDocument.Load(path)) {
                var sheet = document.Sheets[0];
                Assert.Equal("A1:A2", sheet.GetTableRange("1"));
                Assert.Equal("A1:A2", sheet.GetTableRange("_1"));
                Assert.Null(sheet.GetTableRange("Missing"));
                Assert.Equal("_1", Assert.Single(sheet.GetTables()).DisplayName);
                sheet.Cell(2, 1).SetValue("Changed");
                document.Save();
            }
            using var saved = SpreadsheetDocument.Open(path, false);
            var retained = saved.WorkbookPart!.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
            Assert.Equal("1", retained.Name!.Value);
            Assert.Equal("_1", retained.DisplayName!.Value);
            Assert.Equal("A1:A2", retained.Reference!.Value);
        }
    }
}
