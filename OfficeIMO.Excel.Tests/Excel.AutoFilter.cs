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
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("A");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("B");
                sheet.Cell(3, 2).SetValue(20d);
                Dictionary<uint, IEnumerable<string>> criteria = new Dictionary<uint, IEnumerable<string>> {
                    { 0, new[] { "A" } }
                };
                sheet.Range("A1:B3").AutoFilter();
                foreach (var criterion in criteria) sheet.Range("A1:B3").AutoFilter().Column(criterion.Key).Equals(criterion.Value);
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public async Task Test_AddAutoFilterConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.Concurrent.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Value");
                sheet.Cell(2, 1).SetValue("A");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("B");
                sheet.Cell(3, 2).SetValue(20d);

                var tasks = Enumerable.Range(0, 5)
                    .Select(_ => Task.Run(() => sheet.Range("A1:B3").AutoFilter()))
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AutoFilterByHeadersEquals_BatchesAndReplacesExistingFilters() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.BatchHeaders.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Department");
                sheet.Cell(1, 3).SetValue("Score");
                sheet.Cell(1, 4).SetValue("Notes");
                sheet.Cell(2, 1).SetValue("North");
                sheet.Cell(2, 2).SetValue("Finance");
                sheet.Cell(2, 3).SetValue(20d);
                sheet.Cell(2, 4).SetValue("keep");
                sheet.Cell(3, 1).SetValue("South");
                sheet.Cell(3, 2).SetValue("Operations");
                sheet.Cell(3, 3).SetValue(30d);
                sheet.Cell(3, 4).SetValue("review");
                sheet.Range("A1:D3").AutoFilter();
                sheet.Range("A1:D3").AutoFilter().Column("Region").Equals(new[] { "Old" });
                sheet.Range("A1:D3").AutoFilter()
                    .Column("Region").Equals(new[] { "North", "north", "South" })
                    .Column("Score").Equals(new[] { "20", "20", "30" })
                    .Column("Region").Equals(new[] { "East" });
                sheet.Range("A1:D3").AutoFilter().Column("Region").Equals(new[] { "East" }).Column("Score").Equals(new[] { "20", "30" });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                AutoFilter autoFilter = Assert.Single(wsPart.Worksheet.Elements<AutoFilter>());
                Assert.Equal("A1:D3", autoFilter.Reference!.Value);

                var filterColumns = autoFilter.Elements<FilterColumn>().OrderBy(fc => fc.ColumnId?.Value ?? 0U).ToList();
                Assert.Equal(2, filterColumns.Count);

                Assert.Equal(0U, filterColumns[0].ColumnId!.Value);
                var regionFilters = Assert.Single(filterColumns[0].Elements<Filters>());
                Filter regionFilter = Assert.Single(regionFilters.Elements<Filter>());
                Assert.Equal("East", regionFilter.Val!.Value);

                Assert.Equal(2U, filterColumns[1].ColumnId!.Value);
                var scoreFilters = Assert.Single(filterColumns[1].Elements<Filters>());
                Assert.Equal(new[] { "20", "30" }, scoreFilters.Elements<Filter>().Select(filter => filter.Val!.Value).ToArray());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_CustomAutoFilterHelpersPersistExpectedOperators() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.CustomHelpers.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Region");
                sheet.Cell(1, 2).SetValue("Department");
                sheet.Cell(1, 3).SetValue("Score");
                sheet.Cell(1, 4).SetValue("Notes");
                sheet.Cell(1, 5).SetValue("Delta");
                sheet.Cell(2, 1).SetValue("North");
                sheet.Cell(2, 2).SetValue("Finance");
                sheet.Cell(2, 3).SetValue(10d);
                sheet.Cell(2, 4).SetValue("keep");
                sheet.Cell(2, 5).SetValue(5d);
                sheet.Cell(3, 1).SetValue("South");
                sheet.Cell(3, 2).SetValue("Operations");
                sheet.Cell(3, 3).SetValue(20d);
                sheet.Cell(3, 4).SetValue("review later");
                sheet.Cell(3, 5).SetValue(10d);
                sheet.Cell(4, 1).SetValue("East");
                sheet.Cell(4, 2).SetValue("Support");
                sheet.Cell(4, 3).SetValue(30d);
                sheet.Cell(4, 4).SetValue("done");
                sheet.Cell(4, 5).SetValue(15d);
                sheet.Range("A1:E4").AutoFilter();
                sheet.Range("A1:E4").AutoFilter().Column("Region").BeginsWith("No");
                sheet.Range("A1:E4").AutoFilter().Column("Department").EndsWith("ions");
                sheet.Range("A1:E4").AutoFilter().Column("Score").GreaterThanOrEqual(20d);
                sheet.Range("A1:E4").AutoFilter().Column("Notes").DoesNotContain("view");
                sheet.Range("A1:E4").AutoFilter().Column("Delta").NotEqual(10d);
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_CustomAutoFilterRangeHelpersPersistExpectedOperators() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.CustomRangeHelpers.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Budget");
                sheet.Cell(1, 2).SetValue("Score");
                sheet.Cell(2, 1).SetValue(10d);
                sheet.Cell(2, 2).SetValue(5d);
                sheet.Cell(3, 1).SetValue(20d);
                sheet.Cell(3, 2).SetValue(15d);
                sheet.Cell(4, 1).SetValue(30d);
                sheet.Cell(4, 2).SetValue(25d);
                sheet.Range("A1:B4").AutoFilter();
                sheet.Range("A1:B4").AutoFilter().Column("Budget").Between(12d, 28d);
                sheet.Range("A1:B4").AutoFilter().Column("Score").NotBetween(10d, 20d);
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

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
