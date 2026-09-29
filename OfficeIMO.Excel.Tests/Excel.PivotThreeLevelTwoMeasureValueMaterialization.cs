using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("metric", "Metric", 160d, 33d)]
        [InlineData("units", "UnitsMetric", 110d, 60d)]
        public void Test_PivotThreeLevelMixedValue_TwoMeasuresSelectsNamedMeasure(
            string kind, string selectedMeasure, double metricTotal, double unitsTotal) {
            string file = $"pivot-value-three-level-mixed-{kind}-top1-two-measures-conformance.xlsx";
            string path = ThreeLevelPivotOraclePath(file);
            VerifyThreeLevelPivotOracle(path, "A4:J11", 1, 1d);
            using (var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")))) {
                var root = provenance.RootElement;
                Assert.Equal("Source!A1:E13", root.GetProperty("sourceRange").GetString());
                Assert.Equal(selectedMeasure, root.GetProperty("selectedMeasure").GetString());
                Assert.Equal(metricTotal, root.GetProperty("grandTotal").GetDouble());
                Assert.Equal(unitsTotal, root.GetProperty("unitsGrandTotal").GetDouble());
            }
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:J11");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:C13");
            AssertPivotLookupOracleValue(metricTotal, expectedLookups[0, 0]);
            AssertPivotLookupOracleValue(unitsTotal, expectedLookups[0, 1]);

            string importedOutput = Path.Combine(_directoryWithFiles, "Imported." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                AssertPivotLookupOracleValue(metricTotal, grouped.GetPivotData("ValuePivot", "Metric").Value);
                AssertPivotLookupOracleValue(unitsTotal, grouped.GetPivotData("ValuePivot", "UnitsMetric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal("A4:J11", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(26, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var imported = ExcelDocumentReader.Open(importedOutput)) {
                AssertTwoMeasureMixedPivotView(expectedView, imported.GetSheet("Grouped").ReadRange("A4:J11"));
                AssertTwoMeasurePivotLookups(expectedLookups, imported.GetSheet("Lookups").ReadRange("B1:C13"));
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "Authored." + file);
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
                source.CellValue(1, 5, "Units");
                for (int index = 0; index < ThreeLevelTopTwoRows.Length; index++)
                    source.CellValue(index + 2, 5, index is >= 3 and <= 8 ? 10d : 1d);
                source.Pivot("A1:E13").Rows("Region", "Product").Columns("Channel")
                    .Sum("Sales", "Metric").Sum("Units", "UnitsMetric")
                    .Layout(ExcelPivotLayout.Tabular)
                    .Filter(ExcelPivotFilter.TopCount("Product", selectedMeasure, 1))
                    .At("G4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal("G4:P11", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(authoredOutput);
            }
            using (var reopened = ExcelDocument.Load(authoredOutput)) {
                Assert.Empty(reopened.ValidateOpenXml());
                var source = reopened.GetSheet("Source");
                for (int measure = 0; measure < 2; measure++) {
                    string caption = measure == 0 ? "Metric" : "UnitsMetric";
                    AssertPivotLookupOracleValue(expectedLookups[0, measure], source.GetPivotData("ValuePivot", caption).Value);
                    for (int index = 0; index < ThreeLevelTopTwoRows.Length; index++) {
                        var row = ThreeLevelTopTwoRows[index];
                        var actual = source.GetPivotData("ValuePivot", caption, new Dictionary<string, object?> {
                            ["Region"] = row.Region, ["Product"] = row.Product, ["Channel"] = row.Channel
                        });
                        AssertPivotLookupOracleValue(expectedLookups[index + 1, measure], actual.Value);
                    }
                }
            }
            using var authored = ExcelDocumentReader.Open(authoredOutput);
            AssertTwoMeasureMixedPivotView(expectedView, authored.GetSheet("Source").ReadRange("G4:P11"));
        }

        private static void AssertTwoMeasurePivotLookups(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(2, actual.GetLength(1));
            for (int row = 0; row < expected.GetLength(0); row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        private static void AssertTwoMeasureMixedPivotView(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? region = null;
                for (int row = 3; row < view.GetLength(0); row++) {
                    string label = view[row, 0]?.ToString() ?? "";
                    if (label is "East" or "West") region = label;
                    string rowKey = label == "Grand Total" || label.EndsWith(" Total", StringComparison.Ordinal)
                        ? label : $"{region}/{view[row, 1]}";
                    Assert.False(rowKey.EndsWith("/", StringComparison.Ordinal));
                    string? channel = null;
                    for (int column = 2; column < view.GetLength(1); column++) {
                        string heading = view[1, column]?.ToString() ?? "";
                        string measure = view[2, column]?.ToString() ?? "";
                        if (measure.Length == 0 && heading.StartsWith("Total ", StringComparison.Ordinal))
                            measure = heading.Substring("Total ".Length);
                        Assert.True(measure is "Metric" or "UnitsMetric");
                        if (heading is "Retail" or "Online" or "Partner") channel = heading;
                        string columnKey = heading is "Grand Total" or "Total Metric" or "Total UnitsMetric"
                            ? $"Grand Total/{measure}" : $"{channel}/{measure}";
                        Assert.False(columnKey.StartsWith("/", StringComparison.Ordinal));
                        string key = $"{rowKey}/{columnKey}";
                        Assert.False(cells.ContainsKey(key), $"Duplicate pivot cell {key}.");
                        cells.Add(key, JsonSerializer.Serialize(view[row, column]));
                    }
                }
                return cells;
            }
            Assert.Equal(Cells(expected).OrderBy(pair => pair.Key, StringComparer.Ordinal),
                Cells(actual).OrderBy(pair => pair.Key, StringComparer.Ordinal));
        }
    }
}
