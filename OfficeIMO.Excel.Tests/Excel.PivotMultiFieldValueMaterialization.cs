using OfficeIMO.Excel;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("outer", "A4:C8", 5, 65d)]
        [InlineData("inner", "A4:C11", 8, 150d)]
        [InlineData("inner-top1", "A4:C11", 8, 150d)]
        public void Test_PivotMultiFieldValue_ImportedViewAndLookupMatchExcel(
            string kind, string expectedRange, int viewRows, double total) {
            string file = $"pivot-value-multifield-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(expectedRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(expectedRange);
            // Excel sorts parent captions, while OfficeIMO preserves source order.
            // Compare the full saved row multiset, including parent subtotal rows.
            string[] Rows(object?[,] view) => Enumerable.Range(0, viewRows)
                .Select(row => JsonSerializer.Serialize(Enumerable.Range(0, 3)
                    .Select(column => view[row, column]).ToArray()))
                .OrderBy(row => row, StringComparer.Ordinal).ToArray();
            Assert.Equal(Rows(expectedView), Rows(actualView));
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int row = 0; row < 7; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("outer", 65d)]
        [InlineData("inner", 150d)]
        [InlineData("inner-top1", 150d)]
        public void Test_PivotMultiFieldValue_TemplateFreePublicApi(string kind, double total) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                $"pivot-value-multifield-{kind}-conformance.xlsx");
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            string output = Path.Combine(_directoryWithFiles, $"Value.multifield-{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Product");
                source.CellValue(1, 3, "Sales");
                var rows = new[] {
                    ("East", "A", 10d), ("East", "B", 50d),
                    ("West", "A", 40d), ("West", "B", 20d),
                    ("South", "A", 5d), ("South", "B", 60d)
                };
                for (int index = 0; index < rows.Length; index++) {
                    source.CellValue(index + 2, 1, rows[index].Item1);
                    source.CellValue(index + 2, 2, rows[index].Item2);
                    source.CellValue(index + 2, 3, rows[index].Item3);
                }
                ExcelPivotFilter filter = kind == "outer"
                    ? ExcelPivotFilter.ValueGreaterThan("Region", "Metric", 62d)
                    : kind == "inner"
                        ? ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 30d)
                        : ExcelPivotFilter.TopCount("Product", "Metric", 1);
                source.Pivot("A1:C7").Rows("Region", "Product").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("E4", "ValuePivot");
                Assert.True(source.MaterializePivotTable("ValuePivot").Mutation.PackageIsValid);
                Assert.Equal(total, source.GetPivotData("ValuePivot", "Metric").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            var sheet = reopened.GetSheet("Source");
            Assert.Equal(total, sheet.GetPivotData("ValuePivot", "Metric").Value);
            string[] regions = { "East", "East", "West", "West", "South", "South" };
            string[] products = { "A", "B", "A", "B", "A", "B" };
            for (int index = 0; index < regions.Length; index++) {
                var result = sheet.GetPivotData("ValuePivot", "Metric",
                    new Dictionary<string, object?> { ["Region"] = regions[index], ["Product"] = products[index] });
                AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], result.Value);
            }
        }

        [Fact]
        public void Test_PivotMultiFieldValue_ColumnHierarchyMatchesExcelLookups() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                "pivot-value-multifield-column-inner-conformance.xlsx");
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:H7");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            string importedOutput = Path.Combine(_directoryWithFiles, "Column.Imported.xlsx");
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(150d, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal("A4:H7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var reopened = ExcelDocumentReader.Open(importedOutput)) {
                var actualView = reopened.GetSheet("Grouped").ReadRange("A4:H7");
                AssertPivotColumnGridMatchesExcel(expectedView, actualView);
                var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
                for (int row = 0; row < 7; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "Column.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Product");
                source.CellValue(1, 3, "Sales");
                var rows = new[] {
                    ("East", "A", 10d), ("East", "B", 50d),
                    ("West", "A", 40d), ("West", "B", 20d),
                    ("South", "A", 5d), ("South", "B", 60d)
                };
                for (int index = 0; index < rows.Length; index++) {
                    source.CellValue(index + 2, 1, rows[index].Item1);
                    source.CellValue(index + 2, 2, rows[index].Item2);
                    source.CellValue(index + 2, 3, rows[index].Item3);
                }
                source.Pivot("A1:C7").Columns("Region", "Product").Sum("Sales", "Metric")
                    .Filter(ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 30d))
                    .At("E4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal("E4:L7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(authoredOutput);
            }
            using var authored = ExcelDocument.Load(authoredOutput);
            Assert.Empty(authored.ValidateOpenXml());
            var sheet = authored.GetSheet("Source");
            Assert.Equal(150d, sheet.GetPivotData("ValuePivot", "Metric").Value);
            string[] regions = { "East", "East", "West", "West", "South", "South" };
            string[] products = { "A", "B", "A", "B", "A", "B" };
            for (int index = 0; index < regions.Length; index++) {
                var result = sheet.GetPivotData("ValuePivot", "Metric",
                    new Dictionary<string, object?> { ["Region"] = regions[index], ["Product"] = products[index] });
                AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], result.Value);
            }
            using var authoredView = ExcelDocumentReader.Open(authoredOutput);
            AssertPivotColumnGridMatchesExcel(expectedView, authoredView.GetSheet("Source").ReadRange("E4:L7"));
        }

        private static void AssertPivotColumnGridMatchesExcel(object?[,] expected, object?[,] actual) {
            // Excel uses "Column Labels" and caption order; OfficeIMO uses the field name
            // and source order. Compare every displayed child, subtotal, and grand-total cell.
            Assert.Equal("Column Labels", expected[0, 1]);
            Assert.Equal("Region", actual[0, 1]);
            for (int row = 0; row < 4; row++) {
                AssertPivotLookupOracleValue(expected[row, 0], actual[row, 0]);
                if (row > 0) AssertPivotLookupOracleValue(expected[row, 7], actual[row, 7]);
            }
            for (int column = 2; column < 8; column++)
                AssertPivotLookupOracleValue(expected[0, column], actual[0, column]);
            string[] Groups(object?[,] view) => Enumerable.Range(0, 3).Select(index => {
                int column = 1 + 2 * index;
                return JsonSerializer.Serialize(Enumerable.Range(1, 3).Select(row =>
                    new[] { view[row, column], view[row, column + 1] }).ToArray());
            }).OrderBy(group => group, StringComparer.Ordinal).ToArray();
            Assert.Equal(Groups(expected), Groups(actual));
        }

        private static void AssertPivotMultiFieldFixtureHash(string path) {
            using var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")));
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(path);
            string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
            Assert.Equal(provenance.RootElement.GetProperty("sha256").GetString(), hash);
        }
    }
}
