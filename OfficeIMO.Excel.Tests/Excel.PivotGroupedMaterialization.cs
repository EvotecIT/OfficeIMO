using OfficeIMO.Excel;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotNumericGroupOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "numeric-group-conformance.xlsx");
        private static string PivotNumericGroupRefreshOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "numeric-group-refresh-conformance.xlsx");

        [Fact]
        public void Test_PivotNumericGroupMaterialization_MatchesExcelOutputAndLookup() {
            string output = Path.Combine(_directoryWithFiles, "NumericGroup.Materialized.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotNumericGroupOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:B10");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            var selections = new IReadOnlyDictionary<string, object?>[] {
                new Dictionary<string, object?>(),
                new Dictionary<string, object?> { ["Quantity"] = "<0" },
                new Dictionary<string, object?> { ["Quantity"] = "0-9" },
                new Dictionary<string, object?> { ["Quantity"] = "10-19" },
                new Dictionary<string, object?> { ["Quantity"] = "20-30" },
                new Dictionary<string, object?> { ["Quantity"] = ">30" },
                new Dictionary<string, object?> { ["Quantity"] = 4d }
            };
            using (var document = ExcelDocument.Load(PivotNumericGroupOraclePath)) {
                var sheet = document.GetSheet("Grouped");
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotGrouped", "Metric", selections[index]).Value);
                var result = sheet.MaterializePivotTable("PivotGrouped");
                Assert.Equal("A4:B10", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotGrouped", "Metric", selections[index]).Value);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange("A4:B10");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int index = 0; index < 7; index++)
                AssertPivotLookupOracleValue(expectedLookups[index, 0], actualLookups[index, 0]);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_PivotNumericGroupMaterialization_TemplateFreePublicApiMatchesExcel(bool numbersStoredAsText) {
            string output = Path.Combine(_directoryWithFiles, numbersStoredAsText
                ? "NumericGroup.TextNumbers.xlsx" : "NumericGroup.TemplateFree.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotNumericGroupOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:B9");
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:B10");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                for (int row = 0; row < 9; row++)
                    for (int column = 0; column < 2; column++)
                        sheet.CellValue(row + 1, column + 1, numbersStoredAsText && row > 0 && column == 0
                            ? Convert.ToString(sourceValues[row, column], CultureInfo.InvariantCulture)
                            : sourceValues[row, column]);
                sheet.Pivot("A1:B9").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 10, 0, 30).Layout(ExcelPivotLayout.Tabular)
                    .At("D4", "PivotGrouped");
                var result = sheet.MaterializePivotTable("PivotGrouped");
                Assert.Equal("D4:E10", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("D4:E10");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotNumericGroupMaterialization_RefreshRegroupsAndClearsOldTail() {
            string output = Path.Combine(_directoryWithFiles, "NumericGroup.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotNumericGroupRefreshOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:B10");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            using (var document = ExcelDocument.Load(PivotNumericGroupOraclePath)) {
                var source = document.GetSheet("Source");
                source.CellValue(2, 1, 15d);
                source.CellValue(3, 1, 31d);
                var sheet = document.GetSheet("Grouped");
                Assert.Equal("A4:B9", sheet.MaterializePivotTable("PivotGrouped").OutputRange);
                var selections = new IReadOnlyDictionary<string, object?>[] {
                    new Dictionary<string, object?>(),
                    new Dictionary<string, object?> { ["Quantity"] = "<0" },
                    new Dictionary<string, object?> { ["Quantity"] = "0-9" },
                    new Dictionary<string, object?> { ["Quantity"] = "10-19" },
                    new Dictionary<string, object?> { ["Quantity"] = "20-30" },
                    new Dictionary<string, object?> { ["Quantity"] = ">30" },
                    new Dictionary<string, object?> { ["Quantity"] = 4d }
                };
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotGrouped", "Metric", selections[index]).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange("A4:B10");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
        }

        [Fact]
        public void Test_PivotNumericGroupAuthoring_FieldOptionsUseGroupLabels() {
            string output = Path.Combine(_directoryWithFiles, "NumericGroup.HiddenBucket.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotNumericGroupOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:B9");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                for (int row = 0; row < 9; row++)
                    for (int column = 0; column < 2; column++)
                        sheet.CellValue(row + 1, column + 1, sourceValues[row, column]);
                sheet.Pivot("A1:B9").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 10, 0, 30).HideItems("Quantity", "0-9")
                    .At("D4", "PivotGrouped");
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            var pivot = Assert.Single(reopened.GetPivotTables());
            var field = Assert.Single(pivot.Fields, field => field.FieldName == "Quantity");
            Assert.Equal(new[] { "0-9" }, field.HiddenItems);
            Assert.Equal(new[] { "<0", "10-19", "20-30", ">30" }, field.VisibleItems);
            Assert.Empty(reopened.ValidateOpenXml());
        }
    }
}
