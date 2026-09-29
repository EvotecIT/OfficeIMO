using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("mixed-row-greater50", "A4:F11", "F4:K11", 9, 50d, 220d)]
        [InlineData("mixed-row-between45and65", "A4:F11", "F4:K11", 13, 45d, 170d)]
        [InlineData("mixed-row-top1", "A4:F10", "F4:K10", 1, 1d, 160d)]
        [InlineData("mixed-row-bottom1", "A4:F10", "F4:K10", 2, 1d, 105d)]
        [InlineData("mixed-column-bottom1", "A4:D12", "F4:I12", 2, 1d, 85d)]
        [InlineData("mixed-column-greater85", "A4:E12", "F4:J12", 9, 85d, 180d)]
        [InlineData("mixed-column-top1", "A4:E12", "F4:J12", 1, 1d, 180d)]
        public void Test_PivotThreeLevelMixedValue_MatchesExcel(
            string kind, string oracleRange, string authoredRange, int filterType, double threshold, double total) {
            string file = $"pivot-value-three-level-{kind}-conformance.xlsx";
            string path = ThreeLevelPivotOraclePath(file);
            VerifyThreeLevelPivotOracle(path, oracleRange, filterType, threshold);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(oracleRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B13");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);

            string importedOutput = Path.Combine(_directoryWithFiles, "Imported." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                AssertPivotLookupOracleValue(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(oracleRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(13, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var imported = ExcelDocumentReader.Open(importedOutput)) {
                AssertThreeLevelMixedPivotView(expectedView, imported.GetSheet("Grouped").ReadRange(oracleRange));
                var actualLookups = imported.GetSheet("Lookups").ReadRange("B1:B13");
                for (int row = 0; row < 13; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "Authored." + file);
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
                ExcelPivotFilter filter = kind switch {
                    "mixed-row-greater50" => ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 50d),
                    "mixed-row-between45and65" => ExcelPivotFilter.ValueBetween("Product", "Metric", 45d, 65d),
                    "mixed-row-top1" => ExcelPivotFilter.TopCount("Product", "Metric", 1),
                    "mixed-row-bottom1" => ExcelPivotFilter.BottomCount("Product", "Metric", 1),
                    "mixed-column-bottom1" => ExcelPivotFilter.BottomCount("Channel", "Metric", 1),
                    "mixed-column-greater85" => ExcelPivotFilter.ValueGreaterThan("Channel", "Metric", 85d),
                    "mixed-column-top1" => ExcelPivotFilter.TopCount("Channel", "Metric", 1),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:D13").Rows("Region", "Product").Columns("Channel")
                    .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                    .Filter(filter).At("F4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal(authoredRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(authoredOutput);
            }
            using (var reopened = ExcelDocument.Load(authoredOutput)) {
                Assert.Empty(reopened.ValidateOpenXml());
                var source = reopened.GetSheet("Source");
                AssertPivotLookupOracleValue(expectedLookups[0, 0], source.GetPivotData("ValuePivot", "Metric").Value);
                for (int index = 0; index < ThreeLevelTopTwoRows.Length; index++) {
                    var row = ThreeLevelTopTwoRows[index];
                    var actual = source.GetPivotData("ValuePivot", "Metric", new Dictionary<string, object?> {
                        ["Region"] = row.Region, ["Product"] = row.Product, ["Channel"] = row.Channel
                    });
                    AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], actual.Value);
                }
            }
            using var authored = ExcelDocumentReader.Open(authoredOutput);
            AssertThreeLevelMixedPivotView(expectedView, authored.GetSheet("Source").ReadRange(authoredRange));
        }

        private static void AssertThreeLevelMixedPivotView(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            for (int column = 0; column < expected.GetLength(1); column++)
                AssertPivotLookupOracleValue(expected[0, column], actual[0, column]);
            AssertPivotLookupOracleValue(expected[1, 0], actual[1, 0]);
            AssertPivotLookupOracleValue(expected[1, 1], actual[1, 1]);

            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? region = null;
                int width = view.GetLength(1);
                for (int row = 2; row < view.GetLength(0); row++) {
                    string label = view[row, 0]?.ToString() ?? "";
                    if (label is "East" or "West") region = label;
                    string rowKey = label == "Grand Total" || label.EndsWith(" Total", StringComparison.Ordinal)
                        ? label
                        : $"{region}/{view[row, 1]}";
                    Assert.False(rowKey.EndsWith("/", StringComparison.Ordinal), $"Missing pivot row label at row {row}.");
                    for (int column = 2; column < width; column++) {
                        string columnKey = view[1, column]?.ToString() ?? "";
                        Assert.NotEmpty(columnKey);
                        string key = $"{rowKey}/{columnKey}";
                        Assert.False(cells.ContainsKey(key), $"Duplicate pivot cell {key}.");
                        cells.Add(key, JsonSerializer.Serialize(view[row, column]));
                    }
                }
                return cells;
            }
            var expectedCells = Cells(expected);
            var actualCells = Cells(actual);
            Assert.Equal(expectedCells.OrderBy(pair => pair.Key, StringComparer.Ordinal),
                actualCells.OrderBy(pair => pair.Key, StringComparer.Ordinal));
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Test_PivotThreeLevelMixedValue_UnqualifiedRuleFailsClosed(bool outerRow) {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
            ExcelPivotFilter filter = outerRow
                ? ExcelPivotFilter.ValueGreaterThan("Region", "Metric", 50d)
                : ExcelPivotFilter.TopCount("Channel", "Metric", 2);
            source.Pivot("A1:D13").Rows("Region", "Product").Columns("Channel")
                .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                .Filter(filter).At("F4", "ValuePivot");
            var error = Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ValuePivot"));
            Assert.Contains("qualified axis rule", error.Message, StringComparison.Ordinal);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
