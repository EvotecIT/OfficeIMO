using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("date-quarter-row-conformance.xlsx", "A4:D20", 17, 4)]
        [InlineData("date-quarter-column-conformance.xlsx", "A4:Q8", 5, 17)]
        public void Test_PivotQuarterMaterialization_ImportedViewAndLookupMatchExcel(
            string file, string range, int rows, int columns) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(range);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B5");
            object?[] qualifiedLookups = { 235d, 85d, 30d, 60d, "#REF!" };
            for (int row = 0; row < qualifiedLookups.Length; row++)
                AssertPivotLookupOracleValue(qualifiedLookups[row], expectedLookups[row, 0]);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(235d, grouped.GetPivotData("PivotQuarterGrouped", "Metric").Value);
                var result = grouped.MaterializePivotTable("PivotQuarterGrouped");
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(5, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(range);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < columns; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B5");
            for (int row = 0; row < 5; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Fact]
        public void Test_PivotQuarterMaterialization_TemplateFreePublicApi() {
            string output = Path.Combine(_directoryWithFiles, "QuarterHierarchy.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "OrderDate");
                source.CellValue(1, 2, "Sales");
                var entries = new[] {
                    (new DateTime(2025, 1, 15), 10d), (new DateTime(2025, 3, 20), 20d),
                    (new DateTime(2025, 7, 1), 30d), (new DateTime(2025, 11, 5), 25d),
                    (new DateTime(2026, 1, 10), 40d), (new DateTime(2026, 4, 5), 50d),
                    (new DateTime(2026, 9, 9), 60d)
                };
                for (int index = 0; index < entries.Length; index++) {
                    source.CellValue(index + 2, 1, entries[index].Item1);
                    source.CellValue(index + 2, 2, entries[index].Item2);
                }
                source.Pivot("A1:B8").Rows("OrderDate").Sum("Sales", "Metric")
                    .DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Quarters,
                        ExcelPivotGroupBy.Months)
                    .Layout(ExcelPivotLayout.Tabular).At("F4", "QuarterHierarchy");
                var result = source.MaterializePivotTable("QuarterHierarchy");
                Assert.Equal("F4:I20", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(235d, source.GetPivotData("QuarterHierarchy", "Metric").Value);
                Assert.Equal(30d, source.GetPivotData("QuarterHierarchy", "Metric",
                    new Dictionary<string, object?> { ["OrderDate Years"] = 2025d,
                        ["OrderDate Quarters"] = "Q1" }).Value);
                Assert.Equal(60d, source.GetPivotData("QuarterHierarchy", "Metric",
                    new Dictionary<string, object?> { ["OrderDate Years"] = 2026d,
                        ["OrderDate Quarters"] = "Q3" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(235d, reopened.GetSheet("Source").GetPivotData("QuarterHierarchy", "Metric").Value);
            var part = reopened.GetSheet("Source").WorksheetPart.PivotTableParts.Single();
            var cacheFields = part.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!
                .Elements<CacheField>().ToArray();
            Assert.Equal(2U, cacheFields[0].FieldGroup!.ParentId!.Value);
            Assert.Null(cacheFields[0].FieldGroup!.GetFirstChild<RangeProperties>());
            Assert.Equal(new[] { 4, 6, 14 }, cacheFields.Skip(2).Take(3)
                .Select(field => field.FieldGroup!.GetFirstChild<GroupItems>()!.ChildElements.Count));
            Assert.All(cacheFields.Skip(2).Take(3), field => {
                var range = field.FieldGroup!.GetFirstChild<RangeProperties>()!;
                Assert.Equal(new DateTime(2025, 1, 15), range.StartDate!.Value);
                Assert.Equal(new DateTime(2026, 9, 10), range.EndDate!.Value);
            });
            var pivotFields = part.PivotTableDefinition!.PivotFields!.Elements<PivotField>().ToArray();
            Assert.All(pivotFields.Skip(2).Take(3), field => Assert.False(field.ShowAll!.Value));
        }
    }
}
