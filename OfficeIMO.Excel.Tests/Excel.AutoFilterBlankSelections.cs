using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void AutoFilterCombinesBlanksWithValuesAndKeepsLiteralWhitespace(bool tableOwner) {
            string path = Path.Combine(_directoryWithFiles, "BlankSelection." + tableOwner + ".xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Color");
                sheet.Cell(2, 1).SetValue("Orange");
                sheet.Cell(3, 1).SetValue("Red");
                sheet.Cell(5, 1).SetValue(" ");
                sheet.Cell(6, 1).SetValue("Blue");
                if (tableOwner) sheet.AddTable("A1:A6", name: "Colors");
                var filter = sheet.Range("A1:A6").AutoFilter();
                filter.Column("Color").Equals("Orange", "Red");
                filter.Column("Color").Equals("", "Orange", "Red", " ", "", "orange");
                Assert.True(Assert.Single(filter.Inspect()!.Columns).IncludeBlank);
                filter.Column("Color").Equals("Orange", "Red", " ");
                Assert.False(Assert.Single(filter.Inspect()!.Columns).IncludeBlank);
                // The nonblank values match exactly, so this also exercises the idempotence check.
                filter.Column("Color").Equals("Orange", "Red", " ", "");
                document.Save();
            }

            using var saved = SpreadsheetDocument.Open(path, false);
            var part = saved.WorkbookPart!.WorksheetParts.Single();
            var owner = tableOwner ? part.TableDefinitionParts.Single().Table!.AutoFilter
                : part.Worksheet.GetFirstChild<AutoFilter>();
            var values = owner!.Elements<FilterColumn>().Single().GetFirstChild<Filters>()!;
            Assert.True(values.Blank!.Value);
            Assert.Equal(new[] { "Orange", "Red", " " }, values.Elements<Filter>().Select(item => item.Val!.Value));
        }
    }
}
