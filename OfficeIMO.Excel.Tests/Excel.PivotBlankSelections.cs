using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("hide", "", 0, 14d)]
        [InlineData("only", "", 0, 1d)]
        [InlineData("page", "", 0, 1d)]
        [InlineData("hide", " ", 1, 13d)]
        [InlineData("only", " ", 1, 2d)]
        [InlineData("page", " ", 1, 2d)]
        [InlineData("only", " A ", 3, 8d)]
        public void PivotItemSelectionsDistinguishBlanksFromLiteralCaptions(string selection, string caption,
            int selectedIndex, double expectedTotal) {
            string path = Path.Combine(_directoryWithFiles, "PivotBlankSelection.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Category");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(2, 2).SetValue(1d);
                sheet.Cell(3, 1).SetValue(" ");
                sheet.Cell(3, 2).SetValue(2d);
                sheet.Cell(4, 1).SetValue("A");
                sheet.Cell(4, 2).SetValue(4d);
                sheet.Cell(5, 1).SetValue(" A ");
                sheet.Cell(5, 2).SetValue(8d);
                ExcelPivotFieldOptions field = selection == "hide"
                    ? ExcelPivotFieldOptions.HideItems("Category", caption)
                    : selection == "only" ? ExcelPivotFieldOptions.ShowOnlyItems("Category", caption)
                    : ExcelPivotFieldOptions.SelectPageItem("Category", caption);
                sheet.AddPivotTable("A1:B5", "D1", options: new ExcelPivotTableOptions {
                    Name = "Selection",
                    RowFields = selection == "page" ? Array.Empty<string>() : new[] { "Category" },
                    PageFields = selection == "page" ? new[] { "Category" } : Array.Empty<string>(),
                    DataFields = new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Sum") },
                    FieldOptions = new[] { field }
                });
                sheet.MaterializePivotTable("Selection");
                Assert.Equal(expectedTotal, sheet.GetPivotData("Selection", "Sum").Value);
                document.Save();
            }

            using (var package = SpreadsheetDocument.Open(path, false)) {
                var pivotPart = package.WorkbookPart!.WorksheetParts.Single().PivotTableParts.Single();
                var cacheItems = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!
                    .CacheFields!.Elements<CacheField>().First().SharedItems!.ChildElements;
                Assert.IsType<MissingItem>(cacheItems[0]);
                Assert.Equal(" ", Assert.IsType<StringItem>(cacheItems[1]).Val!.Value);
                Assert.Equal("A", Assert.IsType<StringItem>(cacheItems[2]).Val!.Value);
                Assert.Equal(" A ", Assert.IsType<StringItem>(cacheItems[3]).Val!.Value);
                var definition = pivotPart.PivotTableDefinition!;
                if (selection == "page") {
                    Assert.Equal((uint)selectedIndex, definition.PageFields!.Elements<PageField>().Single().Item!.Value);
                } else {
                    var hidden = definition.PivotFields!.Elements<PivotField>().First().Items!.Elements<Item>()
                        .Where(item => item.Hidden?.Value == true).Select(item => (int)item.Index!.Value).ToArray();
                    Assert.Equal(selection == "hide" ? new[] { selectedIndex }
                        : Enumerable.Range(0, 4).Where(index => index != selectedIndex).ToArray(), hidden);
                }
            }

            using var loaded = ExcelDocument.Load(path);
            Assert.Equal(expectedTotal, loaded.Sheets[0].GetPivotData("Selection", "Sum").Value);
        }
    }
}
