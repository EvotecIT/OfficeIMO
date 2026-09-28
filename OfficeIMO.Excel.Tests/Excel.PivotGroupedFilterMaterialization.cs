using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("manual-group-filter-conformance.xlsx", "PivotManualGrouped", "A4:C9", "A4:C12", 9, 3, 6)]
        [InlineData("numeric-group-filter-conformance.xlsx", "PivotGrouped", "A4:B9", "A4:B10", 7, 2, 7)]
        [InlineData("date-group-filter-conformance.xlsx", "PivotDateGrouped", "A4:C9", "A4:C12", 9, 3, 6)]
        public void Test_PivotGroupedItemFilter_MaterializationMatchesExcel(string file, string pivotName,
            string outputRange, string comparisonRange, int rows, int columns, int lookupRows) {
            string oraclePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Filtered." + file);
            using var oracle = ExcelDocumentReader.Open(oraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(comparisonRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange($"B1:B{lookupRows}");
            using (var document = ExcelDocument.Load(oraclePath)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable(pivotName);
                Assert.Equal(outputRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(lookupRows, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(comparisonRange);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < columns; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange($"B1:B{lookupRows}");
            for (int row = 0; row < lookupRows; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Fact]
        public void Test_PivotGroupedItemFilter_ManualAuthoringMatchesExcel() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "manual-group-filter-conformance.xlsx");
            using var oracle = ExcelDocumentReader.Open(path);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:C9");
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Filtered.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var sheet = document.AddWorksheet("Source");
                object?[,] source = oracle.GetSheet("Source").ReadRange("A1:B6");
                for (int row = 0; row < 6; row++)
                    for (int column = 0; column < 2; column++)
                        sheet.CellValue(row + 1, column + 1, source[row, column]);
                sheet.Pivot("A1:B6").Rows("Product").Sum("Sales", "Metric")
                    .HideItems("Product", "Apple")
                    .Layout(ExcelPivotLayout.Tabular).At("E4", "PivotManualGrouped");
                sheet.AddPivotManualGrouping("PivotManualGrouped", "Product", "Product2",
                    new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } },
                    hiddenGroupItems: new[] { "Carrot" });
                Assert.True(sheet.MaterializePivotTable("PivotManualGrouped").Mutation.PackageIsValid);
                Assert.Equal(60d, sheet.GetPivotData("PivotManualGrouped", "Metric").Value);
                Assert.Equal(20d, sheet.GetPivotData("PivotManualGrouped", "Metric",
                    new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("E4:G9");
            for (int row = 0; row < 6; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotGroupedItemFilter_NumericAndDateAuthoring() {
            using var document = ExcelDocument.Create();
            var numeric = document.AddWorksheet("Numbers");
            using (var oracle = ExcelDocumentReader.Open(PivotNumericGroupOraclePath)) {
                object?[,] source = oracle.GetSheet("Source").ReadRange("A1:B9");
                for (int row = 0; row < 9; row++)
                    for (int column = 0; column < 2; column++)
                        numeric.CellValue(row + 1, column + 1, source[row, column]);
            }
            numeric.Pivot("A1:B9").Rows("Quantity").Sum("Sales", "Metric")
                .NumberGroup("Quantity", 10, 0, 30).HideItems("Quantity", "0-9")
                .At("D4", "FilteredBuckets");
            Assert.True(numeric.MaterializePivotTable("FilteredBuckets").Mutation.PackageIsValid);
            Assert.Equal(813d, numeric.GetPivotData("FilteredBuckets", "Metric").Value);
            Assert.Equal("#REF!", numeric.GetPivotData("FilteredBuckets", "Metric",
                new Dictionary<string, object?> { ["Quantity"] = "0-9" }).Value);

            var dates = document.AddWorksheet("Dates");
            dates.CellValue(1, 1, "OrderDate"); dates.CellValue(1, 2, "Sales");
            var values = new[] { new DateTime(2025, 1, 15), new DateTime(2025, 3, 20),
                new DateTime(2025, 7, 1), new DateTime(2026, 1, 10), new DateTime(2026, 4, 5) };
            for (int index = 0; index < values.Length; index++) {
                dates.CellValue(index + 2, 1, values[index]);
                dates.CellValue(index + 2, 2, (index + 1) * 10d);
            }
            dates.Pivot("A1:B6").Rows("OrderDate").Sum("Sales", "Metric")
                .DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Months)
                .HideItems("OrderDate Years", "2026")
                .At("D4", "FilteredYears");
            Assert.True(dates.MaterializePivotTable("FilteredYears").Mutation.PackageIsValid);
            Assert.Equal(60d, dates.GetPivotData("FilteredYears", "Metric").Value);
            Assert.Equal("#REF!", dates.GetPivotData("FilteredYears", "Metric",
                new Dictionary<string, object?> { ["OrderDate Years"] = "2026" }).Value);
        }

        [Fact]
        public void Test_PivotGroupedItemFilter_RefreshPreservesHiddenItemMapping() {
            string corpus = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus");
            string sourcePath = Path.Combine(corpus, "manual-group-filter-conformance.xlsx");
            string oraclePath = Path.Combine(corpus, "manual-group-filter-refresh-conformance.xlsx");
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Filtered.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(oraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B6");
            using (var document = ExcelDocument.Load(sourcePath)) {
                document.GetSheet("Source").CellValue(2, 1, "Broccoli");
                var grouped = document.GetSheet("Grouped");
                Assert.Equal("A4:C9", grouped.MaterializePivotTable("PivotManualGrouped").OutputRange);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(6, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:C12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B6");
            for (int row = 0; row < 6; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }
    }
}
