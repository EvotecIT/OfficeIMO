using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("metric", "Metric", 160d, 33d, false, false, false, "A4:J11", "G4:P11")]
        [InlineData("units", "UnitsMetric", 110d, 60d, false, false, false, "A4:J11", "G4:P11")]
        [InlineData("metric-row-values", "Metric", 160d, 33d, true, false, false, "A4:G15", "G4:M15")]
        [InlineData("units-row-values", "UnitsMetric", 110d, 60d, true, false, false, "A4:G15", "G4:M15")]
        [InlineData("metric-column-values-first", "Metric", 160d, 33d, false, true, false, "A4:J11", "G4:P11")]
        [InlineData("units-column-values-first", "UnitsMetric", 110d, 60d, false, true, false, "A4:J11", "G4:P11")]
        [InlineData("metric-row-values-first", "Metric", 160d, 66d, true, true, false, "A4:G17", "G4:M17")]
        [InlineData("units-row-values-first", "UnitsMetric", 265d, 60d, true, true, false, "A4:G17", "G4:M17")]
        [InlineData("metric-row-values-middle", "Metric", 160d, 66d, true, false, true, "A4:G17", "G4:M17")]
        [InlineData("units-row-values-middle", "UnitsMetric", 265d, 60d, true, false, true, "A4:G17", "G4:M17")]
        public void Test_PivotThreeLevelMixedValue_TwoMeasuresSelectsNamedMeasure(
            string kind, string selectedMeasure, double metricTotal, double unitsTotal,
            bool valuesOnRows, bool valuesFirst, bool valuesMiddle, string oracleRange, string authoredRange) {
            string file = $"pivot-value-three-level-mixed-{kind}-top1-two-measures-conformance.xlsx";
            string path = ThreeLevelPivotOraclePath(file);
            VerifyThreeLevelPivotOracle(path, oracleRange, 1, 1d);
            using (var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")))) {
                var root = provenance.RootElement;
                Assert.Equal("Source!A1:E13", root.GetProperty("sourceRange").GetString());
                Assert.Equal(selectedMeasure, root.GetProperty("selectedMeasure").GetString());
                Assert.Equal(valuesOnRows, root.TryGetProperty("valuesOnRows", out var savedAxis)
                    && savedAxis.GetBoolean());
                if (valuesFirst || valuesMiddle) Assert.Equal(valuesFirst ? 1 : 2, root.GetProperty("valuesPosition").GetInt32());
                Assert.Equal(metricTotal, root.GetProperty("grandTotal").GetDouble());
                Assert.Equal(unitsTotal, root.GetProperty("unitsGrandTotal").GetDouble());
            }
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(oracleRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:C13");
            AssertPivotLookupOracleValue(metricTotal, expectedLookups[0, 0]);
            AssertPivotLookupOracleValue(unitsTotal, expectedLookups[0, 1]);

            string importedOutput = Path.Combine(_directoryWithFiles, $"Imported.TwoMeasure.{kind}.xlsx");
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(valuesFirst ? 0 : valuesMiddle ? 1 : valuesOnRows ? 2 : 1,
                    Assert.Single(grouped.GetPivotTables()).ValuesAxisPosition);
                AssertPivotLookupOracleValue(metricTotal, grouped.GetPivotData("ValuePivot", "Metric").Value);
                AssertPivotLookupOracleValue(unitsTotal, grouped.GetPivotData("ValuePivot", "UnitsMetric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(oracleRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(26, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var imported = ExcelDocumentReader.Open(importedOutput)) {
                AssertTwoMeasureMixedPivotView(expectedView, imported.GetSheet("Grouped").ReadRange(oracleRange), valuesOnRows, valuesFirst, valuesMiddle);
                AssertTwoMeasurePivotLookups(expectedLookups, imported.GetSheet("Lookups").ReadRange("B1:C13"));
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, $"Authored.TwoMeasure.{kind}.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
                source.CellValue(1, 5, "Units");
                for (int index = 0; index < ThreeLevelTopTwoRows.Length; index++)
                    source.CellValue(index + 2, 5, index is >= 3 and <= 8 ? 10d : 1d);
                var pivot = source.Pivot("A1:E13").Rows("Region", "Product").Columns("Channel")
                    .Sum("Sales", "Metric").Sum("Units", "UnitsMetric")
                    .Layout(ExcelPivotLayout.Tabular)
                    .Display(dataOnRows: valuesOnRows)
                    .Filter(ExcelPivotFilter.TopCount("Product", selectedMeasure, 1));
                if (valuesFirst || valuesMiddle) pivot.ValuesPosition(valuesFirst ? 0 : 1);
                pivot.At("G4", "ValuePivot");
                Assert.Equal(valuesFirst ? 0 : valuesMiddle ? 1 : valuesOnRows ? 2 : 1,
                    Assert.Single(source.GetPivotTables()).ValuesAxisPosition);
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal(authoredRange, result.OutputRange);
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
            AssertTwoMeasureMixedPivotView(expectedView, authored.GetSheet("Source").ReadRange(authoredRange), valuesOnRows, valuesFirst, valuesMiddle);
        }

        private static void AssertTwoMeasurePivotLookups(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(2, actual.GetLength(1));
            for (int row = 0; row < expected.GetLength(0); row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Theory]
        [InlineData("average-measure")]
        [InlineData("second-filter")]
        public void Test_PivotThreeLevelMixedValue_TwoMeasureUnqualifiedShapesFailClosed(string shape) {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
            source.CellValue(1, 5, "Units");
            for (int index = 0; index < ThreeLevelTopTwoRows.Length; index++)
                source.CellValue(index + 2, 5, index is >= 3 and <= 8 ? 10d : 1d);
            var pivot = source.Pivot("A1:E13").Rows("Region", "Product").Columns("Channel")
                .Sum("Sales", "Metric");
            if (shape == "average-measure") pivot.Average("Units", "UnitsMetric");
            else pivot.Sum("Units", "UnitsMetric");
            pivot.Layout(ExcelPivotLayout.Tabular)
                .Filter(ExcelPivotFilter.TopCount("Product", "Metric", 1));
            if (shape == "second-filter")
                pivot.Filter(ExcelPivotFilter.TopCount("Channel", "Metric", 1));
            pivot.At("G4", "ValuePivot");
            var error = Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ValuePivot"));
            Assert.Contains("qualified axis rule", error.Message, StringComparison.Ordinal);
            Assert.Empty(document.ValidateOpenXml());
        }

        private static void AssertTwoMeasureMixedPivotView(object?[,] expected, object?[,] actual, bool valuesOnRows, bool valuesFirst, bool valuesMiddle) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            if (valuesOnRows) {
                for (int row = 0; row < expected.GetLength(0); row++)
                    for (int column = 0; column < 3; column++)
                        Assert.True(string.Equals(expected[row, column]?.ToString(), actual[row, column]?.ToString(), StringComparison.Ordinal),
                            $"Pivot row label at [{row}, {column}]: expected '{expected[row, column]}', actual '{actual[row, column]}'.");
                if (valuesFirst) AssertTwoMeasureFirstRowValuesView(expected, actual);
                else if (valuesMiddle) AssertTwoMeasureMiddleRowValuesView(expected, actual);
                else AssertTwoMeasureRowValuesView(expected, actual);
                return;
            }
            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? region = null;
                for (int row = 3; row < view.GetLength(0); row++) {
                    string label = view[row, 0]?.ToString() ?? "";
                    if (label is "East" or "West") region = label;
                    string rowKey = label == "Grand Total" || label.EndsWith(" Total", StringComparison.Ordinal)
                        ? label : $"{region}/{view[row, 1]}";
                    Assert.False(rowKey.EndsWith("/", StringComparison.Ordinal));
                    string? channel = null, activeMeasure = null;
                    for (int column = 2; column < view.GetLength(1); column++) {
                        string heading = view[valuesFirst ? 2 : 1, column]?.ToString() ?? "";
                        string measure = view[valuesFirst ? 1 : 2, column]?.ToString() ?? "";
                        if (measure.Length == 0 && heading.StartsWith("Total ", StringComparison.Ordinal))
                            measure = heading.Substring("Total ".Length);
                        if (valuesFirst && heading.Length == 0 && measure.StartsWith("Total ", StringComparison.Ordinal)) {
                            measure = measure.Substring("Total ".Length);
                            heading = "Grand Total";
                        }
                        if (valuesFirst && measure.Length == 0) measure = activeMeasure ?? "";
                        if (measure is "Metric" or "UnitsMetric") activeMeasure = measure;
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

        private static void AssertTwoMeasureFirstRowValuesView(object?[,] expected, object?[,] actual) {
            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? measure = null, region = null;
                for (int row = 2; row < view.GetLength(0); row++) {
                    string first = view[row, 0]?.ToString() ?? "";
                    string second = view[row, 1]?.ToString() ?? "";
                    string product = view[row, 2]?.ToString() ?? "";
                    if (first is "Metric" or "UnitsMetric") measure = first;
                    if (second is "East" or "West") region = second;
                    string rowKey = first.StartsWith("Total ", StringComparison.Ordinal)
                        ? "Grand Total/" + first.Substring("Total ".Length)
                        : second.EndsWith(" Total", StringComparison.Ordinal)
                            ? $"{measure}/{second}"
                            : $"{measure}/{region}/{product}";
                    Assert.False(rowKey.EndsWith("/", StringComparison.Ordinal));
                    for (int column = 3; column < view.GetLength(1); column++) {
                        string channel = view[1, column]?.ToString() ?? "";
                        Assert.NotEmpty(channel);
                        string key = $"{rowKey}/{channel}";
                        Assert.False(cells.ContainsKey(key), $"Duplicate pivot cell {key}.");
                        cells.Add(key, JsonSerializer.Serialize(view[row, column]));
                    }
                }
                return cells;
            }
            Assert.Equal(Cells(expected).OrderBy(pair => pair.Key, StringComparer.Ordinal),
                Cells(actual).OrderBy(pair => pair.Key, StringComparer.Ordinal));
        }

        private static void AssertTwoMeasureMiddleRowValuesView(object?[,] expected, object?[,] actual) {
            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? region = null, measure = null;
                for (int row = 2; row < view.GetLength(0); row++) {
                    string first = view[row, 0]?.ToString() ?? "";
                    string second = view[row, 1]?.ToString() ?? "";
                    string product = view[row, 2]?.ToString() ?? "";
                    if (first is "East" or "West") region = first;
                    if (second is "Metric" or "UnitsMetric") measure = second;
                    string rowKey;
                    if (first is "Total Metric" or "Total UnitsMetric")
                        rowKey = "Grand Total/" + first.Substring("Total ".Length);
                    else if (first is "East Metric" or "East UnitsMetric" or "West Metric" or "West UnitsMetric") {
                        int separator = first.IndexOf(' ');
                        rowKey = first.Substring(0, separator) + " Total/" + first.Substring(separator + 1);
                    } else {
                        Assert.NotNull(region);
                        Assert.NotNull(measure);
                        Assert.NotEmpty(product);
                        rowKey = $"{region}/{measure}/{product}";
                    }
                    for (int column = 3; column < view.GetLength(1); column++) {
                        string channel = view[1, column]?.ToString() ?? "";
                        Assert.NotEmpty(channel);
                        string key = $"{rowKey}/{channel}";
                        Assert.False(cells.ContainsKey(key), $"Duplicate pivot cell {key}.");
                        cells.Add(key, JsonSerializer.Serialize(view[row, column]));
                    }
                }
                return cells;
            }
            Assert.Equal(Cells(expected).OrderBy(pair => pair.Key, StringComparer.Ordinal),
                Cells(actual).OrderBy(pair => pair.Key, StringComparer.Ordinal));
        }

        private static void AssertTwoMeasureRowValuesView(object?[,] expected, object?[,] actual) {
            Dictionary<string, string> Cells(object?[,] view) {
                var cells = new Dictionary<string, string>(StringComparer.Ordinal);
                string? region = null, product = null;
                for (int row = 2; row < view.GetLength(0); row++) {
                    string first = view[row, 0]?.ToString() ?? "";
                    string second = view[row, 1]?.ToString() ?? "";
                    string third = view[row, 2]?.ToString() ?? "";
                    if (first is "East" or "West") region = first;
                    if (second is "A" or "B") product = second;
                    string rowKey;
                    if (first is "Total Metric" or "Total UnitsMetric")
                        rowKey = "Grand Total/" + first.Substring("Total ".Length);
                    else if ((first == "Grand Total" || first.EndsWith(" Total", StringComparison.Ordinal))
                        && (third is "Metric" or "UnitsMetric"))
                        rowKey = first + "/" + third;
                    else if (first.EndsWith(" Metric", StringComparison.Ordinal)
                        || first.EndsWith(" UnitsMetric", StringComparison.Ordinal)) {
                        int separator = first.IndexOf(' ');
                        rowKey = first.Substring(0, separator) + " Total/" + first.Substring(separator + 1);
                    }
                    else
                        rowKey = $"{region}/{product}/{third}";
                    for (int column = 3; column < view.GetLength(1); column++) {
                        string channel = view[1, column]?.ToString() ?? "";
                        Assert.NotEmpty(channel);
                        string key = $"{rowKey}/{channel}";
                        Assert.False(cells.ContainsKey(key),
                            $"Duplicate pivot cell {key} at row {row}: '{first}' / '{second}' / '{third}'.");
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
