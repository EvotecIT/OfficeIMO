using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotManualGroupOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "manual-group-conformance.xlsx");
        private static string PivotManualGroupRefreshOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "manual-group-refresh-conformance.xlsx");

        private static IReadOnlyDictionary<string, object?>[] ManualGroupSelections => new IReadOnlyDictionary<string, object?>[] {
            new Dictionary<string, object?>(),
            new Dictionary<string, object?> { ["Product2"] = "Fruit" },
            new Dictionary<string, object?> { ["Product2"] = "Fruit", ["Product"] = "Apple" },
            new Dictionary<string, object?> { ["Product2"] = "Fruit", ["Product"] = "Pear" },
            new Dictionary<string, object?> { ["Product2"] = "Carrot" },
            new Dictionary<string, object?> { ["Product2"] = "Unknown" }
        };

        [Fact]
        public void Test_PivotManualGroup_ImportedMaterializationAndLookupMatchExcel() {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Materialized.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotManualGroupOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B6");
            using (var document = ExcelDocument.Load(PivotManualGroupOraclePath)) {
                var sheet = document.GetSheet("Grouped");
                for (int index = 0; index < ManualGroupSelections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotManualGrouped", "Metric", ManualGroupSelections[index]).Value);
                var result = sheet.MaterializePivotTable("PivotManualGrouped");
                Assert.Equal("A4:C12", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(6, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:C12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actual[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B6");
            for (int index = 0; index < 6; index++)
                AssertPivotLookupOracleValue(expectedLookups[index, 0], actualLookups[index, 0]);
        }

        [Fact]
        public void Test_PivotManualGroup_RefreshReindexesDiscreteMapping() {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotManualGroupRefreshOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B6");
            using (var document = ExcelDocument.Load(PivotManualGroupOraclePath)) {
                document.GetSheet("Source").CellValue(2, 1, "Carrot");
                var sheet = document.GetSheet("Grouped");
                Assert.Equal("A4:C12", sheet.MaterializePivotTable("PivotManualGrouped").OutputRange);
                for (int index = 0; index < ManualGroupSelections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotManualGrouped", "Metric", ManualGroupSelections[index]).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:C12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotManualGroup_PublicAuthoringMaterializesExcelContract() {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Authored.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotManualGroupOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            using (var document = ExcelDocument.Create()) {
                var sheet = document.AddWorksheet("Source");
                sheet.CellValue(1, 1, "Product"); sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, "Apple"); sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "Pear"); sheet.CellValue(3, 2, 20d);
                sheet.CellValue(4, 1, "Carrot"); sheet.CellValue(4, 2, 30d);
                sheet.CellValue(5, 1, "Broccoli"); sheet.CellValue(5, 2, 40d);
                sheet.CellValue(6, 1, "Apple"); sheet.CellValue(6, 2, 5d);
                sheet.Pivot("A1:B6").Rows("Product").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular)
                    .At("E4", "PivotManualGrouped");
                sheet.AddPivotManualGrouping("PivotManualGrouped", "Product", "Product2",
                    new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
                Assert.True(sheet.MaterializePivotTable("PivotManualGrouped").Mutation.PackageIsValid);
                Assert.Equal(35d, sheet.GetPivotData("PivotManualGrouped", "Metric",
                    new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("E4:G12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotManualGroup_NewSourceItemRemainsVisibleAfterRefresh() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Product"); sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "Apple"); sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "Pear"); sheet.CellValue(3, 2, 20d);
            sheet.CellValue(4, 1, "Carrot"); sheet.CellValue(4, 2, 30d);
            sheet.Pivot("A1:B4").Rows("Product").Sum("Sales", "Metric")
                .At("E4", "Grouped");
            sheet.AddPivotManualGrouping("Grouped", "Product", "Product2",
                new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
            sheet.MaterializePivotTable("Grouped");
            sheet.CellValue(2, 1, "Kiwi");
            Assert.True(sheet.MaterializePivotTable("Grouped").Mutation.PackageIsValid);
            Assert.Equal(20d, sheet.GetPivotData("Grouped", "Metric",
                new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
            Assert.Equal(10d, sheet.GetPivotData("Grouped", "Metric",
                new Dictionary<string, object?> { ["Product2"] = "Kiwi" }).Value);
        }

        [Fact]
        public void Test_PivotManualGroup_MultipleGroupsOnColumnAxis() {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Columns.xlsx");
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Product"); sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "Apple"); sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "Pear"); sheet.CellValue(3, 2, 20d);
            sheet.CellValue(4, 1, "Carrot"); sheet.CellValue(4, 2, 30d);
            sheet.CellValue(5, 1, "Broccoli"); sheet.CellValue(5, 2, 40d);
            sheet.CellValue(6, 1, "Banana"); sheet.CellValue(6, 2, 5d);
            sheet.Pivot("A1:B6").Columns("Product").Sum("Sales", "Metric")
                .At("D4", "GroupedColumns");
            sheet.AddPivotManualGrouping("GroupedColumns", "Product", "Product2",
                new Dictionary<string, string[]> {
                    ["Fruit"] = new[] { "Apple", "Pear", "Banana" },
                    ["Vegetable"] = new[] { "Carrot", "Broccoli" }
                });
            Assert.True(sheet.MaterializePivotTable("GroupedColumns").Mutation.PackageIsValid);
            Assert.Equal(35d, sheet.GetPivotData("GroupedColumns", "Metric",
                new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
            Assert.Equal(70d, sheet.GetPivotData("GroupedColumns", "Metric",
                new Dictionary<string, object?> { ["Product2"] = "Vegetable" }).Value);
            Assert.Equal(30d, sheet.GetPivotData("GroupedColumns", "Metric",
                new Dictionary<string, object?> { ["Product2"] = "Vegetable", ["Product"] = "Carrot" }).Value);
            document.Save(output);
        }

        [Fact]
        public void Test_PivotManualGroup_ValidatesBeforeMutatingNamedPivot() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Product"); sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "Apple"); sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "Pear"); sheet.CellValue(3, 2, 20d);
            sheet.CellValue(4, 1, "Carrot"); sheet.CellValue(4, 2, 30d);
            sheet.AddPivotTable("A1:B4", "E4", "Grouped", rowFields: new[] { "Product" });
            sheet.AddPivotTable("A1:B4", "J4", "Grouped", rowFields: new[] { "Product" });
            string renamed = document.GetPivotTables().Single(pivot => pivot.Name != "Grouped").Name;
            var target = sheet.WorksheetPart.PivotTableParts.Single(pivot => pivot.PivotTableDefinition?.Name?.Value == renamed);
            string beforeDefinition = target.PivotTableDefinition!.OuterXml;
            string beforeCache = target.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml;
            Assert.Throws<ArgumentException>(() => sheet.AddPivotManualGrouping(renamed, "Product", "Product2",
                new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Unknown" } }));
            Assert.Equal(beforeDefinition, target.PivotTableDefinition.OuterXml);
            Assert.Equal(beforeCache, target.PivotTableCacheDefinitionPart.PivotCacheDefinition.OuterXml);
            sheet.AddPivotManualGrouping(renamed, "Product", "Product2",
                new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
            Assert.Equal(2, sheet.WorksheetPart.PivotTableParts.Single(pivot => pivot.PivotTableDefinition?.Name?.Value == "Grouped")
                .PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.ChildElements.Count);
            Assert.Equal(3, target.PivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields!.ChildElements.Count);
        }
    }
}
