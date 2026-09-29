using OfficeIMO.Excel;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly (string Region, string Product, string Channel, double Sales)[] ThreeLevelTopOneRows = {
            ("East", "A", "Retail", 10d), ("East", "A", "Online", 30d),
            ("East", "B", "Retail", 40d), ("East", "B", "Online", 20d),
            ("West", "A", "Retail", 5d), ("West", "A", "Online", 25d),
            ("West", "B", "Retail", 35d), ("West", "B", "Online", 15d)
        };
        private static readonly (string Region, string Product, string Channel, double Sales)[] ThreeLevelTopTwoRows = {
            ("East", "A", "Retail", 10d), ("East", "A", "Online", 30d), ("East", "A", "Partner", 20d),
            ("East", "B", "Retail", 40d), ("East", "B", "Online", 20d), ("East", "B", "Partner", 5d),
            ("West", "A", "Retail", 5d), ("West", "A", "Online", 25d), ("West", "A", "Partner", 15d),
            ("West", "B", "Retail", 35d), ("West", "B", "Online", 15d), ("West", "B", "Partner", 45d)
        };

        [Theory]
        [InlineData(1, "A4:D15", 130d)]
        [InlineData(2, "A4:D19", 230d)]
        public void Test_PivotThreeLevelValue_ImportedTopCountMatchesExcel(int topCount, string viewRange, double total) {
            var rows = ThreeLevelSourceRows(topCount);
            string file = ThreeLevelPivotFile(topCount);
            string path = ThreeLevelPivotOraclePath(file);
            VerifyThreeLevelPivotOracle(path, viewRange, topCount);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(viewRange);
            string lookupRange = $"B1:B{rows.Length + 1}";
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange(lookupRange);
            string output = Path.Combine(_directoryWithFiles, "Imported." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(viewRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(rows.Length + 1, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            AssertThreeLevelPivotView(expectedView, reopened.GetSheet("Grouped").ReadRange(viewRange));
            var actualLookups = reopened.GetSheet("Lookups").ReadRange(lookupRange);
            for (int row = 0; row < rows.Length + 1; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData(1, "A4:D15", "F4:I15")]
        [InlineData(2, "A4:D19", "F4:I19")]
        public void Test_PivotThreeLevelValue_TemplateFreeTopCountMatchesExcel(
            int topCount, string oracleRange, string authoredRange) {
            var rows = ThreeLevelSourceRows(topCount);
            string file = ThreeLevelPivotFile(topCount);
            string path = ThreeLevelPivotOraclePath(file);
            VerifyThreeLevelPivotOracle(path, oracleRange, topCount);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(oracleRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange($"B1:B{rows.Length + 1}");
            string output = Path.Combine(_directoryWithFiles, "Authored." + file);
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateThreeLevelPivotSource(source, rows);
                source.Pivot($"A1:D{rows.Length + 1}").Rows("Region", "Product", "Channel")
                    .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                    .Filter(ExcelPivotFilter.TopCount("Channel", "Metric", topCount))
                    .At("F4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal(authoredRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(output);
            }
            using (var reopened = ExcelDocument.Load(output)) {
                Assert.Empty(reopened.ValidateOpenXml());
                var source = reopened.GetSheet("Source");
                AssertPivotLookupOracleValue(expectedLookups[0, 0], source.GetPivotData("ValuePivot", "Metric").Value);
                for (int index = 0; index < rows.Length; index++) {
                    var row = rows[index];
                    var actual = source.GetPivotData("ValuePivot", "Metric", new Dictionary<string, object?> {
                        ["Region"] = row.Region, ["Product"] = row.Product, ["Channel"] = row.Channel
                    });
                    AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], actual.Value);
                }
            }
            using var authored = ExcelDocumentReader.Open(output);
            AssertThreeLevelPivotView(expectedView, authored.GetSheet("Source").ReadRange(authoredRange));
        }

        [Fact]
        public void Test_PivotThreeLevelValue_UnqualifiedTopCountFailsClosed() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            PopulateThreeLevelPivotSource(source, ThreeLevelTopTwoRows);
            source.Pivot("A1:D13").Rows("Region", "Product", "Channel")
                .Sum("Sales", "Metric").Layout(ExcelPivotLayout.Tabular)
                .Filter(ExcelPivotFilter.TopCount("Channel", "Metric", 3))
                .At("F4", "ValuePivot");
            var error = Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ValuePivot"));
            Assert.Contains("qualified axis rule", error.Message, StringComparison.Ordinal);
            Assert.Empty(document.ValidateOpenXml());
        }

        private static void PopulateThreeLevelPivotSource(ExcelSheet source,
            (string Region, string Product, string Channel, double Sales)[] rows) {
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Product");
            source.CellValue(1, 3, "Channel");
            source.CellValue(1, 4, "Sales");
            for (int index = 0; index < rows.Length; index++) {
                var row = rows[index];
                source.CellValue(index + 2, 1, row.Region);
                source.CellValue(index + 2, 2, row.Product);
                source.CellValue(index + 2, 3, row.Channel);
                source.CellValue(index + 2, 4, row.Sales);
            }
        }

        private static (string Region, string Product, string Channel, double Sales)[] ThreeLevelSourceRows(int topCount)
            => topCount switch {
                1 => ThreeLevelTopOneRows,
                2 => ThreeLevelTopTwoRows,
                _ => throw new ArgumentOutOfRangeException(nameof(topCount))
            };

        private static string ThreeLevelPivotFile(int topCount)
            => $"pivot-value-three-level-top{topCount}-conformance.xlsx";

        private static string ThreeLevelPivotOraclePath(string file) => Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);

        private static void VerifyThreeLevelPivotOracle(string path, string expectedRange, int topCount) {
            using var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")));
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(path);
            string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
            Assert.Equal("Microsoft Excel", provenance.RootElement.GetProperty("producer").GetString());
            Assert.Equal(provenance.RootElement.GetProperty("sha256").GetString(), hash);
            Assert.Equal(topCount, provenance.RootElement.GetProperty("threshold").GetInt32());
            Assert.Equal(expectedRange, provenance.RootElement.GetProperty("outputRange").GetString());
        }

        private static void AssertThreeLevelPivotView(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            string[] Rows(object?[,] view) {
                var rows = new List<string>();
                object? region = null;
                object? product = null;
                for (int row = 0; row < view.GetLength(0); row++) {
                    var values = Enumerable.Range(0, view.GetLength(1))
                        .Select(column => view[row, column]).ToArray();
                    // Excel leaves repeated row labels blank; OfficeIMO may write an outer label.
                    if (row > 0 && values[0] is string label && label is "East" or "West")
                        region = label;
                    if (values[1] is string item && item is "A" or "B")
                        product = item;
                    if (values[0] == null && (values[1] != null || values[2] != null)) values[0] = region;
                    if (values[1] == null && values[2] != null) values[1] = product;
                    rows.Add(JsonSerializer.Serialize(values));
                }
                return rows.OrderBy(row => row, StringComparer.Ordinal).ToArray();
            }
            Assert.Equal(Rows(expected), Rows(actual));
        }
    }
}
