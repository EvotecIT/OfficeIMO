using OfficeIMO.Excel;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private const string ThreeLevelPivotFile = "pivot-value-three-level-top1-conformance.xlsx";
        private static readonly (string Region, string Product, string Channel, double Sales)[] ThreeLevelPivotRows = {
            ("East", "A", "Retail", 10d), ("East", "A", "Online", 30d),
            ("East", "B", "Retail", 40d), ("East", "B", "Online", 20d),
            ("West", "A", "Retail", 5d), ("West", "A", "Online", 25d),
            ("West", "B", "Retail", 35d), ("West", "B", "Online", 15d)
        };

        [Fact]
        public void Test_PivotThreeLevelValue_ImportedTopCountMatchesExcel() {
            string path = ThreeLevelPivotOraclePath();
            VerifyThreeLevelPivotOracle(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:D15");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B9");
            string output = Path.Combine(_directoryWithFiles, "Imported." + ThreeLevelPivotFile);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(130d, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal("A4:D15", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(9, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            AssertThreeLevelPivotView(expectedView, reopened.GetSheet("Grouped").ReadRange("A4:D15"));
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B9");
            for (int row = 0; row < 9; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Fact]
        public void Test_PivotThreeLevelValue_TemplateFreeTopCountMatchesExcel() {
            string path = ThreeLevelPivotOraclePath();
            VerifyThreeLevelPivotOracle(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:D15");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B9");
            string output = Path.Combine(_directoryWithFiles, "Authored." + ThreeLevelPivotFile);
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateThreeLevelPivotSource(source);
                source.Pivot("A1:D9").Rows("Region", "Product", "Channel")
                    .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                    .Filter(ExcelPivotFilter.TopCount("Channel", "Metric", 1))
                    .At("F4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal("F4:I15", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(output);
            }
            using (var reopened = ExcelDocument.Load(output)) {
                Assert.Empty(reopened.ValidateOpenXml());
                var source = reopened.GetSheet("Source");
                AssertPivotLookupOracleValue(expectedLookups[0, 0], source.GetPivotData("ValuePivot", "Metric").Value);
                for (int index = 0; index < ThreeLevelPivotRows.Length; index++) {
                    var row = ThreeLevelPivotRows[index];
                    var actual = source.GetPivotData("ValuePivot", "Metric", new Dictionary<string, object?> {
                        ["Region"] = row.Region, ["Product"] = row.Product, ["Channel"] = row.Channel
                    });
                    AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], actual.Value);
                }
            }
            using var authored = ExcelDocumentReader.Open(output);
            AssertThreeLevelPivotView(expectedView, authored.GetSheet("Source").ReadRange("F4:I15"));
        }

        [Fact]
        public void Test_PivotThreeLevelValue_UnqualifiedTopCountFailsClosed() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            PopulateThreeLevelPivotSource(source);
            source.Pivot("A1:D9").Rows("Region", "Product", "Channel")
                .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                .Filter(ExcelPivotFilter.TopCount("Channel", "Metric", 2))
                .At("F4", "ValuePivot");
            var error = Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ValuePivot"));
            Assert.Contains("qualified axis rule", error.Message, StringComparison.Ordinal);
            Assert.Empty(document.ValidateOpenXml());
        }

        private static void PopulateThreeLevelPivotSource(ExcelSheet source) {
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Product");
            source.CellValue(1, 3, "Channel");
            source.CellValue(1, 4, "Sales");
            for (int index = 0; index < ThreeLevelPivotRows.Length; index++) {
                var row = ThreeLevelPivotRows[index];
                source.CellValue(index + 2, 1, row.Region);
                source.CellValue(index + 2, 2, row.Product);
                source.CellValue(index + 2, 3, row.Channel);
                source.CellValue(index + 2, 4, row.Sales);
            }
        }

        private static string ThreeLevelPivotOraclePath() => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", ThreeLevelPivotFile);

        private static void VerifyThreeLevelPivotOracle(string path) {
            using var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")));
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(path);
            string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
            Assert.Equal("Microsoft Excel", provenance.RootElement.GetProperty("producer").GetString());
            Assert.Equal(provenance.RootElement.GetProperty("sha256").GetString(), hash);
            Assert.Equal("A4:D15", provenance.RootElement.GetProperty("outputRange").GetString());
        }

        private static void AssertThreeLevelPivotView(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            string[] Rows(object?[,] view) {
                var rows = new List<string>();
                object? region = null;
                for (int row = 0; row < view.GetLength(0); row++) {
                    var values = Enumerable.Range(0, view.GetLength(1))
                        .Select(column => view[row, column]).ToArray();
                    // Excel leaves a repeated outer label blank; OfficeIMO writes it on every row.
                    if (row > 0 && values[0] is string label && label is "East" or "West")
                        region = label;
                    if (values[0] == null && values[1] != null) values[0] = region;
                    rows.Add(JsonSerializer.Serialize(values));
                }
                return rows.OrderBy(row => row, StringComparer.Ordinal).ToArray();
            }
            Assert.Equal(Rows(expected), Rows(actual));
        }
    }
}
