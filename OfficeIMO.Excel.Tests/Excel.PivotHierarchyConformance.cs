using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotHierarchyOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", "hierarchy-conformance.xlsx");

        [Fact]
        public void Test_PivotHierarchyLookup_MatchesNativeSubtotalAndAmbiguityRules() {
            using var oracle = ExcelDocumentReader.Open(PivotHierarchyOraclePath);
            var expected = oracle.GetSheet("Lookups").ReadRange("B1:B924");
            using var document = ExcelDocument.Load(PivotHierarchyOraclePath);
            var lookups = document.GetSheet("Lookups");
            lookups.ClearCachedFormulaResults();
            Assert.Equal(924, lookups.RecalculateSupportedFormulas());
            for (int row = 1; row <= 924; row++) {
                var actual = lookups.CellAt(row, 2).GetValue().Value;
                Assert.True(expected[row - 1, 0]?.GetType() == actual?.GetType(), $"B{row}: Excel={expected[row - 1, 0]}; OfficeIMO={actual}");
                AssertPivotLookupOracleValue(expected[row - 1, 0], actual);
            }
            Assert.Equal(0d, document.GetSheet("Row2").GetPivotData("PivotRow2", "Revenue",
                new Dictionary<string, object?> { ["City"] = "Unique" }).Value);
            Assert.Equal("#REF!", document.GetSheet("Row2").GetPivotData("PivotRow2", "Revenue",
                new Dictionary<string, object?> { ["City"] = "East" }).Value);

            string manifestPath = Path.Combine(Path.GetDirectoryName(PivotHierarchyOraclePath)!, "hierarchy-conformance.provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(manifestPath));
            var selections = manifest.RootElement.GetProperty("selections").EnumerateArray().ToArray();
            var captions = manifest.RootElement.GetProperty("measures").EnumerateArray().Select(m => m.GetProperty("Caption").GetString()!).ToArray();
            int checkedLookups = 0;
            foreach (var profile in manifest.RootElement.GetProperty("profiles").EnumerateArray()) {
                var sheet = document.GetSheet(profile.GetProperty("sheet").GetString()!);
                int firstLookup = profile.GetProperty("firstLookup").GetInt32();
                for (int measure = 0; measure < profile.GetProperty("measures").GetInt32(); measure++) {
                    for (int selection = 0; selection < selections.Length; selection++) {
                        var keys = selections[selection].EnumerateArray().Select(key => key.GetString()!).ToArray();
                        var criteria = new Dictionary<string, object?>();
                        for (int index = 0; index < keys.Length; index += 2) criteria.Add(keys[index], keys[index + 1]);
                        var actual = sheet.GetPivotData(profile.GetProperty("name").GetString()!, captions[measure], criteria);
                        var native = expected[firstLookup + measure * selections.Length + selection - 1, 0];
                        Assert.Equal(native is double ? ExcelCellDataKind.Number : ExcelCellDataKind.Error, actual.Kind);
                        AssertPivotLookupOracleValue(native, actual.Value);
                        checkedLookups++;
                    }
                }
            }
            Assert.Equal(924, checkedLookups);
            string output = Path.Combine(_directoryWithFiles, "HierarchyLookup.xlsx");
            document.Save(output);
            using var reopened = ExcelDocumentReader.Open(output);
            var saved = reopened.GetSheet("Lookups").ReadRange("B1:B924");
            for (int row = 0; row < 924; row++) AssertPivotLookupOracleValue(expected[row, 0], saved[row, 0]);
        }

        [Fact]
        public void Test_PivotHierarchyLookup_RejectsPrefixLongerThanPreviousSubtotal() {
            using var document = ExcelDocument.Load(PivotHierarchyOraclePath);
            var sheet = document.GetSheet("Row3");
            var items = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.RowItems!.Elements<RowItem>().ToArray();
            int subtotal = Array.FindIndex(items, item => item.ItemType?.Value == ItemValues.Default);
            items[subtotal + 1].RepeatedItemCount = 3;
            Assert.Throws<NotSupportedException>(() => sheet.GetPivotData("PivotRow3", "Revenue"));
        }
    }
}
