using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("between", "A4:B8", 5, 90d)]
        [InlineData("not-between", "A4:B8", 5, 110d)]
        [InlineData("top-count", "A4:B7", 4, 100d)]
        [InlineData("bottom-count", "A4:B6", 3, 10d)]
        [InlineData("top-percent", "A4:B7", 4, 100d)]
        [InlineData("bottom-percent", "A4:B9", 6, 100d)]
        [InlineData("top-sum", "A4:B7", 4, 100d)]
        [InlineData("bottom-sum", "A4:B8", 5, 60d)]
        public void Test_PivotValueRangeRanking_ImportedViewAndLookupMatchExcel(
            string kind, string range, int rows, double total) {
            string file = $"pivot-value-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(range);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(range);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int row = 0; row < 7; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("between", 90d)]
        [InlineData("not-between", 110d)]
        [InlineData("top-count", 100d)]
        [InlineData("bottom-count", 10d)]
        [InlineData("top-percent", 100d)]
        [InlineData("bottom-percent", 100d)]
        [InlineData("top-sum", 100d)]
        [InlineData("bottom-sum", 60d)]
        public void Test_PivotValueRangeRanking_TemplateFreePublicApi(string kind, double total) {
            string output = Path.Combine(_directoryWithFiles, $"Filter.value-{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Sales");
                var entries = new[] {
                    ("Alpha", 10d), ("Bravo", 20d), ("Charlie", 30d),
                    ("Delta", 40d), ("Echo", 50d), ("Foxtrot", 50d)
                };
                for (int index = 0; index < entries.Length; index++) {
                    source.CellValue(index + 2, 1, entries[index].Item1);
                    source.CellValue(index + 2, 2, entries[index].Item2);
                }
                ExcelPivotFilter filter = kind switch {
                    "between" => ExcelPivotFilter.ValueBetween("Region", "Metric", 20, 40),
                    "not-between" => ExcelPivotFilter.ValueNotBetween("Region", "Metric", 20, 40),
                    "top-count" => ExcelPivotFilter.TopCount("Region", "Metric", 1),
                    "bottom-count" => ExcelPivotFilter.BottomCount("Region", "Metric", 1),
                    "top-percent" => ExcelPivotFilter.TopPercent("Region", "Metric", 40),
                    "bottom-percent" => ExcelPivotFilter.BottomPercent("Region", "Metric", 40),
                    "top-sum" => ExcelPivotFilter.TopSum("Region", "Metric", 90),
                    "bottom-sum" => ExcelPivotFilter.BottomSum("Region", "Metric", 35),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:B7").Rows("Region").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("D4", "FilteredPivot");
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
        public void Test_PivotTopPercent_LargeFiniteTotalDoesNotOverflowThreshold() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Sales");
            source.CellValue(2, 1, "Alpha");
            source.CellValue(2, 2, 6e306);
            source.CellValue(3, 1, "Bravo");
            source.CellValue(3, 2, 4e306);
            source.Pivot("A1:B3").Rows("Region").Sum("Sales", "Metric")
                .Layout(ExcelPivotLayout.Tabular)
                .Filter(ExcelPivotFilter.TopPercent("Region", "Metric", 40))
                .At("D4", "FilteredPivot");
            var result = source.MaterializePivotTable("FilteredPivot");
            Assert.True(result.Mutation.PackageIsValid);
            Assert.Equal(6e306, source.GetPivotData("FilteredPivot", "Metric").Value);
        }
    }
}
