using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for adding and persisting auto filters in Excel sheets.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AddAutoFilterPersists() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 20d);
                Dictionary<uint, IEnumerable<string>> criteria = new Dictionary<uint, IEnumerable<string>> {
                    { 0, new[] { "A" } }
                };
                sheet.AddAutoFilter("A1:B3", criteria);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter? autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.NotNull(autoFilter);
                Assert.NotNull(autoFilter!.Reference);
                Assert.Equal("A1:B3", autoFilter.Reference!.Value);
                FilterColumn? filterColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault();
                Assert.NotNull(filterColumn);
                Filters? filters = filterColumn!.GetFirstChild<Filters>();
                Assert.NotNull(filters);
                Filter filter = filters!.Elements<Filter>().First();
                Assert.NotNull(filter.Val);
                Assert.Equal("A", filter.Val!.Value);
            }
        }

        [Fact]
        public async Task Test_AddAutoFilterConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.Concurrent.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 20d);

                var tasks = Enumerable.Range(0, 5)
                    .Select(_ => Task.Run(() => sheet.AddAutoFilter("A1:B3")))
                    .ToArray();
                await Task.WhenAll(tasks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter? autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.NotNull(autoFilter);
                Assert.NotNull(autoFilter!.Reference);
                Assert.Equal("A1:B3", autoFilter.Reference!.Value);
            }
        }

        [Fact]
        public void Test_CustomAutoFilterHelpersPersistExpectedOperators() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.CustomHelpers.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Department");
                sheet.CellValue(1, 3, "Score");
                sheet.CellValue(1, 4, "Notes");
                sheet.CellValue(1, 5, "Delta");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, "Finance");
                sheet.CellValue(2, 3, 10d);
                sheet.CellValue(2, 4, "keep");
                sheet.CellValue(2, 5, 5d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, "Operations");
                sheet.CellValue(3, 3, 20d);
                sheet.CellValue(3, 4, "review later");
                sheet.CellValue(3, 5, 10d);
                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "Support");
                sheet.CellValue(4, 3, 30d);
                sheet.CellValue(4, 4, "done");
                sheet.CellValue(4, 5, 15d);
                sheet.AutoFilterAdd("A1:E4");
                sheet.AutoFilterByHeaderStartsWith("Region", "No");
                sheet.AutoFilterByHeaderEndsWith("Department", "ions");
                sheet.AutoFilterByHeaderGreaterThanOrEqual("Score", 20d);
                sheet.AutoFilterByHeaderDoesNotContain("Notes", "view");
                sheet.AutoFilterByHeaderNotEqual("Delta", 10d);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter autoFilter = Assert.Single(wsPart.Worksheet.Elements<AutoFilter>());
                Assert.Equal("A1:E4", autoFilter.Reference!.Value);

                var filterColumns = autoFilter.Elements<FilterColumn>().OrderBy(fc => fc.ColumnId?.Value ?? 0U).ToList();
                Assert.Equal(5, filterColumns.Count);

                var startsWithFilter = filterColumns[0].GetFirstChild<CustomFilters>();
                Assert.NotNull(startsWithFilter);
                var startsWithCondition = Assert.Single(startsWithFilter!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.Equal, startsWithCondition.Operator?.Value);
                Assert.Equal("No*", startsWithCondition.Val!.Value);

                var endsWithFilter = filterColumns[1].GetFirstChild<CustomFilters>();
                Assert.NotNull(endsWithFilter);
                var endsWithCondition = Assert.Single(endsWithFilter!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.Equal, endsWithCondition.Operator?.Value);
                Assert.Equal("*ions", endsWithCondition.Val!.Value);

                var numericFilter = filterColumns[2].GetFirstChild<CustomFilters>();
                Assert.NotNull(numericFilter);
                var numericCondition = Assert.Single(numericFilter!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, numericCondition.Operator?.Value);
                Assert.Equal("20", numericCondition.Val!.Value);

                var notContainsFilter = filterColumns[3].GetFirstChild<CustomFilters>();
                Assert.NotNull(notContainsFilter);
                var notContainsCondition = Assert.Single(notContainsFilter!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.NotEqual, notContainsCondition.Operator?.Value);
                Assert.Equal("*view*", notContainsCondition.Val!.Value);

                var notEqualFilter = filterColumns[4].GetFirstChild<CustomFilters>();
                Assert.NotNull(notEqualFilter);
                var notEqualCondition = Assert.Single(notEqualFilter!.Elements<CustomFilter>());
                Assert.Equal(FilterOperatorValues.NotEqual, notEqualCondition.Operator?.Value);
                Assert.Equal("10", notEqualCondition.Val!.Value);
            }
        }

        [Fact]
        public void Test_CustomAutoFilterRangeHelpersPersistExpectedOperators() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.CustomRangeHelpers.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Budget");
                sheet.CellValue(1, 2, "Score");
                sheet.CellValue(2, 1, 10d);
                sheet.CellValue(2, 2, 5d);
                sheet.CellValue(3, 1, 20d);
                sheet.CellValue(3, 2, 15d);
                sheet.CellValue(4, 1, 30d);
                sheet.CellValue(4, 2, 25d);
                sheet.AutoFilterAdd("A1:B4");
                sheet.AutoFilterByHeaderBetween("Budget", 12d, 28d);
                sheet.AutoFilterByHeaderNotBetween("Score", 10d, 20d);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter autoFilter = Assert.Single(wsPart.Worksheet.Elements<AutoFilter>());
                var filterColumns = autoFilter.Elements<FilterColumn>().OrderBy(fc => fc.ColumnId?.Value ?? 0U).ToList();
                Assert.Equal(2, filterColumns.Count);

                var betweenFilter = filterColumns[0].GetFirstChild<CustomFilters>();
                Assert.NotNull(betweenFilter);
                Assert.True(betweenFilter!.And?.Value);
                var betweenConditions = betweenFilter.Elements<CustomFilter>().ToList();
                Assert.Equal(2, betweenConditions.Count);
                Assert.Equal(FilterOperatorValues.GreaterThanOrEqual, betweenConditions[0].Operator?.Value);
                Assert.Equal("12", betweenConditions[0].Val!.Value);
                Assert.Equal(FilterOperatorValues.LessThanOrEqual, betweenConditions[1].Operator?.Value);
                Assert.Equal("28", betweenConditions[1].Val!.Value);

                var notBetweenFilter = filterColumns[1].GetFirstChild<CustomFilters>();
                Assert.NotNull(notBetweenFilter);
                Assert.False(notBetweenFilter!.And?.Value ?? false);
                var notBetweenConditions = notBetweenFilter.Elements<CustomFilter>().ToList();
                Assert.Equal(2, notBetweenConditions.Count);
                Assert.Equal(FilterOperatorValues.LessThan, notBetweenConditions[0].Operator?.Value);
                Assert.Equal("10", notBetweenConditions[0].Val!.Value);
                Assert.Equal(FilterOperatorValues.GreaterThan, notBetweenConditions[1].Operator?.Value);
                Assert.Equal("20", notBetweenConditions[1].Val!.Value);
            }
        }
    }
}
