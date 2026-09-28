using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("label", "LabelPivot", "A4:B8", 5, 115d, "#REF!", 40d, 60d, "#REF!", 15d)]
        [InlineData("value", "ValuePivot", "A4:B8", 5, 170d, "#REF!", 40d, 60d, 70d, "#REF!")]
        [InlineData("combined", "CombinedPivot", "A4:B7", 4, 100d, "#REF!", 40d, 60d, "#REF!", "#REF!")]
        [InlineData("caption", "CaptionPivot", "A4:B9", 6, 145d, 30d, 40d, 60d, "#REF!", 15d)]
        public void Test_PivotLabelValueFilters_ImportedViewAndLookupMatchExcel(
            string kind, string pivotName, string range, int rows, double total,
            object north, object northeast, object east, object west, object southeast) {
            string file = $"pivot-filter-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(range);
            const int lookupRows = 7;
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange($"B1:B{lookupRows}");
            object?[] qualifiedLookups = kind == "caption"
                ? new object?[] { total, north, northeast, east, west, southeast, 30d }
                : new object?[] { total, north, northeast, east, west, southeast, "#REF!" };
            for (int row = 0; row < qualifiedLookups.Length; row++)
                AssertPivotLookupOracleValue(qualifiedLookups[row], expectedLookups[row, 0]);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData(pivotName, "Metric").Value);
                if (kind == "caption") {
                    Assert.Equal(30d, grouped.GetPivotData(pivotName, "Metric",
                        new Dictionary<string, object?> { ["Region"] = "North" }).Value);
                    Assert.Equal(30d, grouped.GetPivotData(pivotName, "Metric",
                        new Dictionary<string, object?> { ["Region"] = "Easterly" }).Value);
                }
                var result = grouped.MaterializePivotTable(pivotName);
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(lookupRows, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(range);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange($"B1:B{lookupRows}");
            for (int row = 0; row < lookupRows; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("label", 115d)]
        [InlineData("value", 170d)]
        [InlineData("combined", 100d)]
        public void Test_PivotLabelValueFilters_TemplateFreePublicApi(string kind, double total) {
            string output = Path.Combine(_directoryWithFiles, $"Filter.{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Sales");
                var entries = new[] {
                    ("East", 25d), ("East", 35d), ("North", 10d), ("North", 20d),
                    ("Northeast", 40d), ("Southeast", 15d), ("West", 70d)
                };
                for (int index = 0; index < entries.Length; index++) {
                    source.CellValue(index + 2, 1, entries[index].Item1);
                    source.CellValue(index + 2, 2, entries[index].Item2);
                }
                var builder = source.Pivot("A1:B8").Rows("Region").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular);
                if (kind == "label" || kind == "combined")
                    builder.Filter(ExcelPivotFilter.LabelContains("Region", "east"));
                if (kind == "value" || kind == "combined")
                    builder.Filter(ExcelPivotFilter.ValueGreaterThan("Region", "Metric", 30));
                builder.At("D4", "FilteredPivot");
                var result = source.MaterializePivotTable("FilteredPivot");
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(total, source.GetPivotData("FilteredPivot", "Metric").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(total, reopened.GetSheet("Source").GetPivotData("FilteredPivot", "Metric").Value);
        }

        [Fact]
        public void Test_PivotLabelValueFilters_RefreshRequalifiesValueItem() {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus");
            string input = Path.Combine(directory, "pivot-filter-combined-conformance.xlsx");
            string oraclePath = Path.Combine(directory, "pivot-filter-combined-refresh-conformance.xlsx");
            string output = Path.Combine(_directoryWithFiles, "Filter.Combined.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(oraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:B8");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            AssertPivotLookupOracleValue(155d, expectedLookups[0, 0]);
            AssertPivotLookupOracleValue(55d, expectedLookups[5, 0]);
            using (var document = ExcelDocument.Load(input)) {
                document.GetSheet("Source").CellValue(7, 2, 55d);
                var grouped = document.GetSheet("Grouped");
                Assert.Equal("A4:B8", grouped.MaterializePivotTable("CombinedPivot").OutputRange);
                Assert.Equal(155d, grouped.GetPivotData("CombinedPivot", "Metric").Value);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange("A4:B8");
            for (int row = 0; row < 5; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int row = 0; row < 7; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }
    }
}
