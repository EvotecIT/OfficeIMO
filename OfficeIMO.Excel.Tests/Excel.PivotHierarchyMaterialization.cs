using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("Row2", "Row2")]
        [InlineData("Col2", "Col2")]
        [InlineData("Both2", "Both2")]
        [InlineData("Row3", "Row3")]
        [InlineData("Col3", "Col3")]
        [InlineData("ManyCol", "ManyCol")]
        [InlineData("ManyRow", "ManyRow")]
        [InlineData("MiddleCol", "MiddleCol")]
        [InlineData("MiddleRow", "MiddleRow")]
        [InlineData("OuterCol", "OuterCol")]
        [InlineData("OuterRow", "OuterRow")]
        [InlineData("NoSubtotals", "NoSubtotals")]
        [InlineData("CustomSum", "Both2")]
        [InlineData("CustomTwo", "Both2")]
        [InlineData("NoGrand", "NoGrand")]
        [InlineData("NoSubtotal3", "NoSubtotal3")]
        [InlineData("CompactManyRow", "MiddleRow")]
        [InlineData("OutlineManyCol", "OuterCol")]
        [InlineData("TopSubtotalRow3", "Row3")]
        [InlineData("CollapsedRow", "Both2")]
        [InlineData("CompactOuterRow", "OuterRow")]
        [InlineData("CollapsedManyRow", "ManyRow")]
        public void Test_PivotHierarchyMaterialization_MatchesNativeTabularViewAndReopens(string name, string expectedProfile) {
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(Path.GetDirectoryName(PivotHierarchyOraclePath)!, "hierarchy-conformance.provenance.json")));
            var profiles = manifest.RootElement.GetProperty("profiles").EnumerateArray().ToArray();
            var original = profiles.Single(profile => profile.GetProperty("sheet").GetString() == name);
            var target = profiles.Single(profile => profile.GetProperty("sheet").GetString() == expectedProfile);
            int first = original.GetProperty("firstLookup").GetInt32();
            int expectedFirst = target.GetProperty("firstLookup").GetInt32();
            int count = original.GetProperty("lookupCount").GetInt32();
            using var oracle = ExcelDocumentReader.Open(PivotHierarchyOraclePath);
            var expected = oracle.GetSheet("Lookups").ReadRange($"B{expectedFirst}:B{expectedFirst + count - 1}");
            string output = Path.Combine(_directoryWithFiles, name + ".HierarchyMaterialized.xlsx");
            using (var document = ExcelDocument.Load(PivotHierarchyOraclePath)) {
                var sheet = document.GetSheet(name);
                var result = sheet.MaterializePivotTable("Pivot" + name);
                Assert.Equal(target.GetProperty("outputRange").GetString(), result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(924, lookups.RecalculateSupportedFormulas());
                for (int index = 0; index < count; index++) {
                    var actual = lookups.CellAt(first + index, 2).GetValue().Value;
                    Assert.True(expected[index, 0]?.GetType() == actual?.GetType(), $"{name} B{first + index}: Excel={expected[index, 0]}; OfficeIMO={actual}");
                    AssertPivotLookupOracleValue(expected[index, 0], actual);
                }
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var saved = reopened.GetSheet("Lookups").ReadRange($"B{first}:B{first + count - 1}");
            for (int index = 0; index < count; index++) AssertPivotLookupOracleValue(expected[index, 0], saved[index, 0]);
        }

        [Fact]
        public void Test_PivotHierarchyMaterialization_BudgetsSubtotalWorkBeforeMutation() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            var headers = new[] { "Region", "City", "Product", "Channel", "Amount", "Units" };
            for (int column = 0; column < headers.Length; column++) sheet.CellValue(1, column + 1, headers[column]);
            for (int row = 2; row <= 101; row++) {
                sheet.CellValue(row, 1, "North"); sheet.CellValue(row, 2, "East");
                sheet.CellValue(row, 3, "A"); sheet.CellValue(row, 4, "Retail");
                sheet.CellValue(row, 5, 1d); sheet.CellValue(row, 6, 1d);
            }
            sheet.AddPivotTable("A1:F101", "I1", "Pivot", rowFields: new[] { "Region", "City" }, columnFields: new[] { "Product", "Channel" },
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue"),
                    new ExcelPivotDataField("Units", ExcelPivotDataFunction.Average, "AverageUnits"),
                    new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Count, "Entries") });
            var part = sheet.WorksheetPart.PivotTableParts.Single();
            string definition = part.PivotTableDefinition!.OuterXml;
            string cache = part.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml;
            var error = Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("Pivot", new ExcelMutationPlanOptions { MaximumAffectedCells = 700 }));
            Assert.Contains("measure input visits", error.Message);
            Assert.Equal(definition, part.PivotTableDefinition.OuterXml);
            Assert.Equal(cache, part.PivotTableCacheDefinitionPart.PivotCacheDefinition.OuterXml);
            var result = sheet.MaterializePivotTable("Pivot", new ExcelMutationPlanOptions { MaximumAffectedCells = 1200 });
            Assert.Equal("I1:S7", result.OutputRange);
            Assert.Equal(100d, sheet.GetPivotData("Pivot", "Revenue").Value);
            string output = Path.Combine(_directoryWithFiles, "HierarchyTemplateFree.xlsx");
            document.Save(output);
            using var reopened = ExcelDocument.Load(output);
            Assert.Equal(100d, reopened.GetSheet("Source").GetPivotData("Pivot", "Revenue").Value);
            Assert.Equal(1d, reopened.GetSheet("Source").GetPivotData("Pivot", "AverageUnits", new Dictionary<string, object?> { ["Region"] = "North" }).Value);
        }

        [Fact]
        public void Test_PivotHierarchyMaterialization_BudgetsAllCriteriaBeforeMutation() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            var headers = Enumerable.Range(0, 6).Select(field => "Key" + field).Concat(new[] { "Amount" }).ToArray();
            for (int column = 0; column < headers.Length; column++) sheet.CellValue(1, column + 1, headers[column]);
            for (int row = 2; row <= 101; row++) {
                for (int column = 1; column <= 6; column++) sheet.CellValue(row, column, "Item" + row);
                sheet.CellValue(row, 7, 1d);
            }
            sheet.AddPivotTable("A1:G101", "I1", "Pivot", rowFields: headers.Take(6).ToArray(),
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue") });
            var part = sheet.WorksheetPart.PivotTableParts.Single();
            var definition = part.PivotTableDefinition!;
            foreach (var field in definition.PivotFields!.Elements<PivotField>()) field.DefaultSubtotal = false;
            definition.ColumnGrandTotals = false;
            string before = definition.OuterXml;
            string cache = part.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml;
            var error = Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("Pivot", new ExcelMutationPlanOptions { MaximumAffectedCells = 1000 }));
            Assert.Contains("criteria index", error.Message);
            Assert.Equal(before, part.PivotTableDefinition.OuterXml);
            Assert.Equal(cache, part.PivotTableCacheDefinitionPart.PivotCacheDefinition.OuterXml);
            Assert.Equal("I1:O101", sheet.MaterializePivotTable("Pivot", new ExcelMutationPlanOptions { MaximumAffectedCells = 1200 }).OutputRange);
            var criteria = headers.Take(6).ToDictionary(field => field, _ => (object?)"Item101");
            Assert.Equal(1d, sheet.GetPivotData("Pivot", "Revenue", criteria).Value);
        }

        [Fact]
        public void Test_PivotHierarchyMaterialization_RejectsDuplicatedRealAxisField() {
            using var document = ExcelDocument.Load(PivotHierarchyOraclePath);
            var sheet = document.GetSheet("Both2");
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            definition.ColumnFields!.Elements<Field>().Last().Index = 1;
            string before = definition.OuterXml;
            Assert.Throws<NotSupportedException>(() => sheet.MaterializePivotTable("PivotBoth2"));
            Assert.Equal(before, definition.OuterXml);
        }
    }
}
