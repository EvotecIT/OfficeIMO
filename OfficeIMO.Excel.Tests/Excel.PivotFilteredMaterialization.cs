using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotFilteredOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "filtered-conformance.xlsx");
        private static string PivotFilteredRefreshOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "filtered-refresh-conformance.xlsx");
        private static string PivotFilteredBlankOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "filtered-blank-conformance.xlsx");

        [Fact]
        public void Test_PivotFilteredMaterialization_MatchesExcelViewAndLookupsAfterReopen() {
            string output = Path.Combine(_directoryWithFiles, "Filtered.Materialized.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotFilteredOraclePath);
            var expectedView = oracle.GetSheet("Filtered").ReadRange("A4:B7");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            var selections = new IReadOnlyDictionary<string, object?>[] {
                new Dictionary<string, object?>(),
                new Dictionary<string, object?> { ["Region"] = "North" },
                new Dictionary<string, object?> { ["Region"] = "South" },
                new Dictionary<string, object?> { ["Region"] = "West" },
                new Dictionary<string, object?> { ["Product"] = "B" },
                new Dictionary<string, object?> { ["Product"] = "A" },
                new Dictionary<string, object?> { ["Region"] = "North", ["Product"] = "B" }
            };
            using (var document = ExcelDocument.Load(PivotFilteredOraclePath)) {
                var sheet = document.GetSheet("Filtered");
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotFiltered", "Metric", selections[index]).Value);

                var result = sheet.MaterializePivotTable("PivotFiltered");
                Assert.Equal("A4:B7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotFiltered", "Metric", selections[index]).Value);

                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(selections.Length, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Filtered").ReadRange("A4:B7");
            for (int row = 0; row < 4; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int index = 0; index < selections.Length; index++)
                AssertPivotLookupOracleValue(expectedLookups[index, 0], actualLookups[index, 0]);
        }

        [Fact]
        public void Test_PivotFilteredMaterialization_RefreshMatchesExcelWhenNewFilteredKeyAppears() {
            string output = Path.Combine(_directoryWithFiles, "Filtered.RefreshedSource.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotFilteredRefreshOraclePath);
            var expectedView = oracle.GetSheet("Filtered").ReadRange("A4:B7");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B8");
            var selections = new IReadOnlyDictionary<string, object?>[] {
                new Dictionary<string, object?>(),
                new Dictionary<string, object?> { ["Region"] = "North" },
                new Dictionary<string, object?> { ["Region"] = "South" },
                new Dictionary<string, object?> { ["Region"] = "West" },
                new Dictionary<string, object?> { ["Product"] = "B" },
                new Dictionary<string, object?> { ["Product"] = "A" },
                new Dictionary<string, object?> { ["Region"] = "North", ["Product"] = "B" },
                new Dictionary<string, object?> { ["Region"] = "East" }
            };
            using (var document = ExcelDocument.Load(PivotFilteredOraclePath)) {
                document.GetSheet("Source").CellValue(2, 1, "East");
                document.GetSheet("Source").CellValue(2, 2, "B");
                var sheet = document.GetSheet("Filtered");
                Assert.Equal("A4:B7", sheet.MaterializePivotTable("PivotFiltered").OutputRange);
                for (int index = 0; index < selections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotFiltered", "Metric", selections[index]).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Filtered").ReadRange("A4:B7");
            for (int row = 0; row < 4; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
        }

        [Fact]
        public void Test_PivotFilteredMaterialization_TemplateFreePublicApiMatchesExcel() {
            string output = Path.Combine(_directoryWithFiles, "Filtered.TemplateFree.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotFilteredOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:C7");
            var expected = oracle.GetSheet("Filtered").ReadRange("A4:B7");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                for (int row = 0; row < 7; row++)
                    for (int column = 0; column < 3; column++)
                        sheet.CellValue(row + 1, column + 1, sourceValues[row, column]);
                sheet.AddPivotTable("A1:C7", "E4", "PivotFiltered", rowFields: new[] { "Region" },
                    pageFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Metric") },
                    fieldOptions: new[] {
                        new ExcelPivotFieldOptions("Region", hiddenItems: new[] { "West" }),
                        new ExcelPivotFieldOptions("Product", selectedItem: "B")
                    }, layout: ExcelPivotLayout.Tabular);
                var result = sheet.MaterializePivotTable("PivotFiltered");
                Assert.Equal("E4:F7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(38d, sheet.GetPivotData("PivotFiltered", "Metric").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("E4:F7");
            for (int row = 0; row < 4; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotFilteredMaterialization_SelectedBlankPageMatchesExcelLookups() {
            string output = Path.Combine(_directoryWithFiles, "Filtered.BlankPage.xlsx");
            var blank = new Dictionary<string, object?> { ["Product"] = "(blank)" };
            using var oracle = ExcelDocumentReader.Open(PivotFilteredBlankOraclePath);
            var expectedView = oracle.GetSheet("Filtered").ReadRange("A4:B6");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B9");
            using (var document = ExcelDocument.Load(PivotFilteredBlankOraclePath)) {
                var sheet = document.GetSheet("Filtered");
                AssertPivotLookupOracleValue(expectedLookups[8, 0], sheet.GetPivotData("PivotFiltered", "Metric", blank).Value);
                Assert.Equal("A4:B6", sheet.MaterializePivotTable("PivotFiltered").OutputRange);
                AssertPivotLookupOracleValue(expectedLookups[8, 0], sheet.GetPivotData("PivotFiltered", "Metric", blank).Value);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(9, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Filtered").ReadRange("A4:B6");
            for (int row = 0; row < 3; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B9");
            for (int index = 0; index < 9; index++)
                AssertPivotLookupOracleValue(expectedLookups[index, 0], actualLookups[index, 0]);
        }
    }
}
