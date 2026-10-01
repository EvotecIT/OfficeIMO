using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(1)]
        [InlineData(2)]
        [InlineData(3)]
        [InlineData(4)]
        public void Test_PivotMaterialization_RejectsUncachedFormulasInEverySourceField(int column) {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            document.GetSheet("Source").CellFormula(2, column, "IF(TRUE,1,2)");
            document.GetSheet("Source").ClearCachedFormulaResults();
            string before = sheet.WorksheetPart.Worksheet.OuterXml;
            var cache = sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!;
            string cacheBefore = cache.PivotCacheDefinition!.OuterXml;
            Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotSum"));
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(cacheBefore, cache.PivotCacheDefinition!.OuterXml);
        }

        [Fact]
        public void Test_PivotMaterialization_CreatesWithoutTemplateOrOfficeRefresh() {
            string path = Path.Combine(_directoryWithFiles, "HeadlessPivot.TemplateFree.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotLookupOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:D9");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Sales");
                for (int row = 1; row <= 9; row++) for (int column = 1; column <= 4; column++) sheet.CellValue(row, column, sourceValues[row - 1, column - 1]);
                sheet.AddPivotTable("A1:D9", "F1", "SalesPivot", rowFields: new[] { "Region" }, columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Metric") }, layout: ExcelPivotLayout.Tabular);
                var result = sheet.MaterializePivotTable("SalesPivot");
                Assert.Equal("F1:I6", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(48d, sheet.GetPivotData("SalesPivot", "Metric").Value);
                var cache = sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!;
                Assert.Equal(cache.GetIdOfPart(cache.PivotTableCacheRecordsPart!), cache.PivotCacheDefinition!.Id!.Value);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Equal(48d, reopened.GetSheet("Sales").GetPivotData("SalesPivot", "Metric").Value);
            Assert.Equal(8U, reopened.GetSheet("Sales").WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.RecordCount!.Value);
        }

        [Theory]
        [InlineData(ExcelPivotDataFunction.Sum)]
        [InlineData(ExcelPivotDataFunction.Count)]
        [InlineData(ExcelPivotDataFunction.CountNumbers)]
        [InlineData(ExcelPivotDataFunction.Average)]
        [InlineData(ExcelPivotDataFunction.Minimum)]
        [InlineData(ExcelPivotDataFunction.Maximum)]
        [InlineData(ExcelPivotDataFunction.Product)]
        [InlineData(ExcelPivotDataFunction.StandardDeviation)]
        [InlineData(ExcelPivotDataFunction.StandardDeviationP)]
        [InlineData(ExcelPivotDataFunction.Variance)]
        [InlineData(ExcelPivotDataFunction.VarianceP)]
        public void Test_PivotMaterialization_MatchesIndependentExcelAndReopens(ExcelPivotDataFunction function) {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                using var oracle = ExcelDocumentReader.Open(PivotLookupOraclePath);
                var functions = new[] { ExcelPivotDataFunction.Sum, ExcelPivotDataFunction.Count, ExcelPivotDataFunction.CountNumbers,
                    ExcelPivotDataFunction.Average, ExcelPivotDataFunction.Minimum, ExcelPivotDataFunction.Maximum,
                    ExcelPivotDataFunction.Product, ExcelPivotDataFunction.StandardDeviation, ExcelPivotDataFunction.StandardDeviationP,
                    ExcelPivotDataFunction.Variance, ExcelPivotDataFunction.VarianceP };
                int firstRow = Array.IndexOf(functions, function) * 9 + 1;
                var expected = oracle.GetSheet("Lookups").ReadRange($"B{firstRow}:B{firstRow + 8}");
                using (var document = ExcelDocument.Load(PivotLookupOraclePath)) {
                    var sheet = document.GetSheet(function.ToString());
                    var result = sheet.MaterializePivotTable("Pivot" + function);
                    Assert.Equal("A1:D6", result.OutputRange);
                    Assert.Equal(8, result.SourceRecordCount);
                    Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                    var pivot = Assert.Single(sheet.GetPivotTables());
                    Assert.False(pivot.RefreshOnOpen);
                    Assert.True(pivot.SaveSourceData);
                    var selections = new (string? Region, string? Product)[] { (null, null), ("North", null), ("South", null), ("West", null),
                        (null, "A"), (null, "B"), ("North", "A"), ("South", "B"), ("missing", null) };
                    for (int index = 0; index < selections.Length; index++) {
                        var criteria = new Dictionary<string, object?>();
                        if (selections[index].Region != null) criteria["Region"] = selections[index].Region;
                        if (selections[index].Product != null) criteria["Product"] = selections[index].Product;
                        AssertPivotOracleValue(expected[index, 0], sheet.GetPivotData("Pivot" + function, "Metric", criteria));
                    }
                    document.GetSheet("Lookups").ClearCachedFormulaResults();
                    Assert.Equal(165, document.GetSheet("Lookups").RecalculateSupportedFormulas());
                    document.Save(output);
                    File.Copy(output, Path.Combine(_directoryWithFiles, function + ".Materialized.xlsx"), true);
                }
                using var reopened = ExcelDocumentReader.Open(output);
                var actual = reopened.GetSheet("Lookups").ReadRange($"B{firstRow}:B{firstRow + 8}");
                for (int index = 0; index < 9; index++) AssertPivotLookupOracleValue(expected[index, 0], actual[index, 0]);
                using var model = ExcelDocument.Load(output);
                Assert.Equal(8U, model.GetSheet(function.ToString()).WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
            } finally { if (File.Exists(output)) File.Delete(output); }
        }

        [Fact]
        public void Test_PivotMaterialization_RefreshesSourceAndClearsStaleOwnedTail() {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var source = document.GetSheet("Source");
            var sheet = document.GetSheet("Sum");
            source.CellValue(2, 3, 999d);
            sheet.MaterializePivotTable("PivotSum");
            Assert.Equal(1037d, sheet.GetPivotData("PivotSum", "Metric").Value);
            source.CellValue(7, 1, "North");
            var result = sheet.MaterializePivotTable("PivotSum");
            Assert.Equal("A1:D5", result.OutputRange);
            for (int column = 1; column <= 4; column++) Assert.Null(sheet.CellAt(6, column).GetValue().Value);
            Assert.Equal(1037d, sheet.GetPivotData("PivotSum", "Metric").Value);
        }

        [Theory]
        [InlineData("RowsOnly", "A1:B5", 23, 1)]
        [InlineData("ColumnsOnly", "A1:D3", 24, 1)]
        [InlineData("Scalar", "A1:A2", 25, 1)]
        [InlineData("Keys", "A1:B9", 26, 8)]
        [InlineData("ErrorCount", "A1:B8", 34, 1)]
        [InlineData("ErrorNumbers", "A1:B8", 35, 1)]
        public void Test_PivotMaterialization_IndependentLayoutsAndTypedKeys(string name, string range, int firstRow, int count) {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            using var oracle = ExcelDocumentReader.Open(PivotLookupOraclePath);
            var sheet = document.GetSheet(name);
            if (name == "RowsOnly") {
                var importedField = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.PivotFields!.Elements<PivotField>().First();
                Assert.True(importedField.SumSubtotal!.Value);
                Assert.True(importedField.CountASubtotal!.Value);
            }
            var result = sheet.MaterializePivotTable("Pivot" + name);
            Assert.Equal(range, result.OutputRange);
            Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
            if (name == "RowsOnly") {
                var normalizedField = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.PivotFields!.Elements<PivotField>().First();
                Assert.False(normalizedField.SumSubtotal!.Value);
                Assert.False(normalizedField.CountASubtotal!.Value);
                Assert.True(normalizedField.DefaultSubtotal!.Value);
            }
            var lookup = document.GetSheet("LookupEdges");
            lookup.ClearCachedFormulaResults();
            Assert.Equal(35, lookup.RecalculateSupportedFormulas());
            var expected = oracle.GetSheet("LookupEdges").ReadRange($"A{firstRow}:A{firstRow + count - 1}");
            for (int index = 0; index < count; index++) AssertPivotLookupOracleValue(expected[index, 0], lookup.CellAt(firstRow + index, 1).GetValue().Value);
            document.Save(Path.Combine(_directoryWithFiles, name + ".Materialized.xlsx"));
            if (name == "Keys") {
                var shared = sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().First().SharedItems!;
                Assert.Equal(7U, shared.Count!.Value);
                Assert.Single(shared.Elements<NumberItem>());
                Assert.Single(shared.Elements<BooleanItem>());
                Assert.Single(shared.Elements<MissingItem>());
                Assert.Equal(4, shared.Elements<StringItem>().Count());
            }
        }

        [Fact]
        public void Test_PivotMaterialization_BudgetsAndCollisionsPreserveSourceAndView() {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            string before = sheet.WorksheetPart.Worksheet.OuterXml;
            string cacheBefore = sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml;
            Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotSum", new ExcelMutationPlanOptions { MaximumAffectedCells = 1 }));
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(cacheBefore, sheet.WorksheetPart.PivotTableParts.Single().PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml);
            Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotSum", new ExcelMutationPlanOptions { MaximumSnapshotCharacters = 1 }));
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
            document.GetSheet("Source").CellValue(7, 2, "C");
            sheet.CellValue(3, 5, "keep");
            string collisionBefore = sheet.WorksheetPart.Worksheet.OuterXml;
            Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotSum"));
            Assert.Equal(collisionBefore, sheet.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal("keep", sheet.CellAt(3, 5).GetValue().Value);
        }
    }
}
