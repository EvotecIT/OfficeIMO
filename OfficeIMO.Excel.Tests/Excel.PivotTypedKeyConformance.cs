using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(1900, false)]
        [InlineData(1904, false)]
        [InlineData(1900, true)]
        [InlineData(1904, true)]
        public void Test_PivotTypedKeys_MatchIndependentExcelLookupsAndReopen(int system, bool materialize) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", $"typed-keys-{system}.xlsx");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, ".provenance.json")));
            using var oracle = ExcelDocumentReader.Open(path);
            int count = manifest.RootElement.GetProperty("lookupFormulaCells").GetInt32();
            var expected = oracle.GetSheet("Lookups").ReadRange($"B1:B{count}");
            string output = Path.Combine(_directoryWithFiles, $"TypedKeys{system}.{(materialize ? "Materialized" : "Recalculated")}.xlsx");
            using (var document = ExcelDocument.Load(path)) {
                foreach (var profile in manifest.RootElement.GetProperty("profiles").EnumerateArray()) {
                    string name = profile.GetProperty("sheet").GetString()!;
                    string pivot = profile.GetProperty("name").GetString()!;
                    var sheet = document.GetSheet(name);
                    if (materialize) {
                        var result = sheet.MaterializePivotTable(pivot);
                        Assert.Equal(profile.GetProperty("outputRange").GetString(), result.OutputRange);
                        Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                        if (name.StartsWith("DatesOuter", StringComparison.Ordinal) || name.StartsWith("DatesDeep", StringComparison.Ordinal)) {
                            Assert.True(A1.TryParseRange(result.OutputRange, out int top, out int left, out int bottom, out int right));
                            var grid = new object?[bottom - top + 1, right - left + 1];
                            for (int r = 0; r < grid.GetLength(0); r++) for (int c = 0; c < grid.GetLength(1); c++)
                                grid[r, c] = sheet.CellAt(top + r, left + c).GetValue().Value;
                            var labels = grid.Cast<object?>().OfType<string>().ToArray();
                            Assert.Contains((system == 1900 ? "1900-01-01" : "1904-01-01") + " 00:00:00 Total", labels);
                            if (system == 1900) {
                                Assert.Contains("1900-02-29 00:00:00 Total", labels);
                                Assert.Contains("1900-02-29 12:00:00 Total", labels);
                            }
                            if (name.StartsWith("DatesDeep", StringComparison.Ordinal)) {
                                bool columns = name.EndsWith("Cols", StringComparison.Ordinal);
                                int ancestors = 0;
                                for (int r = 0; r < grid.GetLength(0); r++) for (int c = 0; c < grid.GetLength(1); c++) {
                                    if (grid[r, c] is not string caption || caption != "North Total") continue;
                                    int dateRow = columns ? r - 1 : r;
                                    int dateColumn = columns ? c : c - 1;
                                    Assert.IsType<double>(grid[dateRow, dateColumn]);
                                    var cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == A1.CellReference(top + dateRow, left + dateColumn));
                                    Assert.True(cell.StyleIndex?.Value > 0);
                                    ancestors++;
                                }
                                Assert.True(ancestors > 0);
                            }
                        }
                        foreach (var field in sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>()) {
                            var items = field.SharedItems!;
                            Assert.False(items.MinDate != null && items.MinValue != null);
                            Assert.False(items.MaxDate != null && items.MaxValue != null);
                            if (items.Elements<DateTimeItem>().Any()) {
                                Assert.False(items.ContainsNumber!.Value);
                                Assert.Null(items.ContainsInteger);
                                if (!items.Elements<NumberItem>().Any() && !items.Elements<StringItem>().Any()
                                    && !items.Elements<ErrorItem>().Any() && !items.Elements<BooleanItem>().Any()) Assert.False(items.ContainsNonDate!.Value);
                            }
                        }
                    }
                    int row = profile.GetProperty("firstLookup").GetInt32();
                    foreach (var selection in profile.GetProperty("selections").EnumerateArray()) {
                        var pairs = selection.GetProperty("Pairs").EnumerateArray().ToArray();
                        var criteria = new Dictionary<string, object?>();
                        for (int index = 0; index < pairs.Length; index += 2) criteria.Add(pairs[index].GetString()!, TypedPivotOracleValue(pairs[index + 1]));
                        var actual = sheet.GetPivotData(pivot, "Revenue", criteria);
                        Assert.True(expected[row - 1, 0]?.GetType() == actual.Value?.GetType(), $"{system}/{name}/{selection.GetProperty("Kind").GetString()}: Excel={expected[row - 1, 0]}; OfficeIMO={actual.Value}");
                        AssertPivotLookupOracleValue(expected[row - 1, 0], actual.Value);
                        if (selection.GetProperty("Kind").GetString() == "serial" && !(system == 1900 && criteria.Values.OfType<double>().Any(serial => Math.Floor(serial) == 60))) {
                            foreach (string key in criteria.Keys.ToArray()) if (criteria[key] is double serial)
                                criteria[key] = ExcelDateSystemConverter.FromSerial(serial, document.DateSystem);
                            AssertPivotLookupOracleValue(expected[row - 1, 0], sheet.GetPivotData(pivot, "Revenue", criteria).Value);
                        }
                        row++;
                    }
                    if (profile.GetProperty("errorLookupValue").ValueKind == JsonValueKind.Number) {
                        var errorCriteria = new Dictionary<string, object?> { ["MixedKey"] = new ExcelCellData(ExcelCellDataKind.Error, "#DIV/0!") };
                        AssertPivotLookupOracleValue(profile.GetProperty("errorLookupValue").GetDouble(), sheet.GetPivotData(pivot, "Revenue", errorCriteria).Value);
                    }
                }
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(count, lookups.RecalculateSupportedFormulas());
                for (int row = 0; row < count; row++) AssertPivotLookupOracleValue(expected[row, 0], lookups.CellAt(row + 1, 2).GetValue().Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var saved = reopened.GetSheet("Lookups").ReadRange($"B1:B{count}");
            for (int row = 0; row < count; row++) AssertPivotLookupOracleValue(expected[row, 0], saved[row, 0]);
        }

        private static object? TypedPivotOracleValue(JsonElement value) => value.ValueKind switch {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            _ => null
        };

        [Theory]
        [InlineData(ExcelDateSystem.NineteenHundred)]
        [InlineData(ExcelDateSystem.NineteenFour)]
        public void Test_PivotTypedKeys_TemplateFreePreservesDateFormatsAndDistinctErrors(ExcelDateSystem system) {
            using var document = ExcelDocument.Create();
            document.DateSystem = system;
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Group"); sheet.CellValue(1, 2, "Key"); sheet.CellValue(1, 3, "Amount");
            var dates = system == ExcelDateSystem.NineteenHundred
                ? new[] { new DateTime(1900, 1, 1), new DateTime(1900, 2, 28), new DateTime(1900, 3, 1), new DateTime(2024, 1, 1, 12, 0, 0) }
                : new[] { new DateTime(1904, 1, 1), new DateTime(1904, 2, 28), new DateTime(1904, 3, 1), new DateTime(2024, 1, 1, 12, 0, 0) };
            for (int index = 0; index < dates.Length; index++) sheet.CellValue(index + 2, 2, dates[index]);
            sheet.CellError(6, 2, "#DIV/0!"); sheet.CellValue(7, 2, "#DIV/0!");
            for (int row = 2; row <= 7; row++) { sheet.CellValue(row, 1, "All"); sheet.CellValue(row, 3, (double)row); }
            sheet.AddPivotTable("A1:C7", "E1", "TypedPivot", rowFields: new[] { "Group", "Key" },
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue") });
            Assert.Equal("E1:G9", sheet.MaterializePivotTable("TypedPivot").OutputRange);
            for (int index = 0; index < dates.Length; index++) {
                Assert.Equal((double)index + 2, sheet.GetPivotData("TypedPivot", "Revenue", new Dictionary<string, object?> { ["Key"] = dates[index] }).Value);
                var cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "F" + (index + 2));
                Assert.True(cell.StyleIndex?.Value > 0);
            }
            Assert.Equal(6d, sheet.GetPivotData("TypedPivot", "Revenue", new Dictionary<string, object?> { ["Key"] = new ExcelCellData(ExcelCellDataKind.Error, "#DIV/0!") }).Value);
            Assert.Equal(7d, sheet.GetPivotData("TypedPivot", "Revenue", new Dictionary<string, object?> { ["Key"] = "#DIV/0!" }).Value);
            string output = Path.Combine(_directoryWithFiles, $"TypedKeys{(system == ExcelDateSystem.NineteenHundred ? 1900 : 1904)}.TemplateFree.xlsx");
            document.Save(output);
            using var reopened = ExcelDocument.Load(output);
            Assert.Equal(6d, reopened.GetSheet("Source").GetPivotData("TypedPivot", "Revenue", new Dictionary<string, object?> { ["Key"] = new ExcelCellData(ExcelCellDataKind.Error, "#DIV/0!") }).Value);
        }

        [Fact]
        public void Test_PivotTypedKeys_RematerializationUsesCurrentSourceDateStyles() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Key"); sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, 1d); sheet.CellValue(2, 2, 7d);
            sheet.AddPivotTable("A1:B2", "D1", "Pivot", rowFields: new[] { "Key" },
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue") });
            sheet.MaterializePivotTable("Pivot");
            var date = new DateTime(1900, 1, 1);
            sheet.CellValue(2, 1, date);
            sheet.MaterializePivotTable("Pivot");
            Assert.Equal(7d, sheet.GetPivotData("Pivot", "Revenue", new Dictionary<string, object?> { ["Key"] = date }).Value);
            var field = sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().First();
            Assert.Equal(new DateTime(1899, 12, 31), Assert.Single(field.SharedItems!.Elements<DateTimeItem>()).Val!.Value);
        }
    }
}
