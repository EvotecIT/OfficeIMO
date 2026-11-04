using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
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
                document.GetSheet("Source").Cell(2, 1).SetValue("Carrot");
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
        public void Test_PivotManualGroup_PublicAuthoringMaterializesExcelContract()
        {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Authored.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotManualGroupOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            using (var document = ExcelDocument.Create())
            {
                var sheet = document.AddSheet("Source");
                sheet.Cell(1, 1).SetValue("Product");
                sheet.Cell(1, 2).SetValue("Sales");
                sheet.Cell(2, 1).SetValue("Apple");
                sheet.Cell(2, 2).SetValue(10d);
                sheet.Cell(3, 1).SetValue("Pear");
                sheet.Cell(3, 2).SetValue(20d);
                sheet.Cell(4, 1).SetValue("Carrot");
                sheet.Cell(4, 2).SetValue(30d);
                sheet.Cell(5, 1).SetValue("Broccoli");
                sheet.Cell(5, 2).SetValue(40d);
                sheet.Cell(6, 1).SetValue("Apple");
                sheet.Cell(6, 2).SetValue(5d);
                {
                    var pivotOptions1 = new ExcelPivotTableOptions();
                    pivotOptions1.RowFields.Add("Product");
                    pivotOptions1.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Metric"));
                    pivotOptions1.Layout = ExcelPivotLayout.Tabular;
                    pivotOptions1.Name = "PivotManualGrouped";
                    sheet.AddPivotTable("A1:B6", "E4", pivotOptions1);
                }

                var grouping = sheet.AddPivotManualGrouping("PivotManualGrouped", "Product", "Product2", new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
                Assert.Equal("Product2", grouping.FieldName);
                Assert.Equal(new[] { "Fruit", "Broccoli", "Carrot" }, grouping.GroupItems);
                var projectedGroup = sheet.GetPivotTables().Single(pivot => pivot.Name == "PivotManualGrouped").Groupings.Single();
                Assert.Equal(grouping.GroupItems, projectedGroup.GroupItems);
                Assert.Equal(grouping.BaseFieldIndex, projectedGroup.BaseFieldIndex);
                Assert.True(sheet.MaterializePivotTable("PivotManualGrouped").Mutation.PackageIsValid);
                Assert.Equal(35d, sheet.GetPivotData("PivotManualGrouped", "Metric", new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
                document.Save(output);
            }

            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("E4:G12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotManualGroup_NewSourceItemRemainsVisibleAfterRefresh()
        {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Source");
            sheet.Cell(1, 1).SetValue("Product");
            sheet.Cell(1, 2).SetValue("Sales");
            sheet.Cell(2, 1).SetValue("Apple");
            sheet.Cell(2, 2).SetValue(10d);
            sheet.Cell(3, 1).SetValue("Pear");
            sheet.Cell(3, 2).SetValue(20d);
            sheet.Cell(4, 1).SetValue("Carrot");
            sheet.Cell(4, 2).SetValue(30d);
            {
                var pivotOptions2 = new ExcelPivotTableOptions();
                pivotOptions2.RowFields.Add("Product");
                pivotOptions2.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Metric"));
                pivotOptions2.Name = "Grouped";
                sheet.AddPivotTable("A1:B4", "E4", pivotOptions2);
            }

            sheet.AddPivotManualGrouping("Grouped", "Product", "Product2", new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
            sheet.MaterializePivotTable("Grouped");
            sheet.Cell(2, 1).SetValue("Kiwi");
            Assert.True(sheet.MaterializePivotTable("Grouped").Mutation.PackageIsValid);
            Assert.Equal(20d, sheet.GetPivotData("Grouped", "Metric", new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
            Assert.Equal(10d, sheet.GetPivotData("Grouped", "Metric", new Dictionary<string, object?> { ["Product2"] = "Kiwi" }).Value);
        }

        [Fact]
        public void Test_PivotManualGroup_MultipleGroupsOnColumnAxis()
        {
            string output = Path.Combine(_directoryWithFiles, "ManualGroup.Columns.xlsx");
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Source");
            sheet.Cell(1, 1).SetValue("Product");
            sheet.Cell(1, 2).SetValue("Sales");
            sheet.Cell(2, 1).SetValue("Apple");
            sheet.Cell(2, 2).SetValue(10d);
            sheet.Cell(3, 1).SetValue("Pear");
            sheet.Cell(3, 2).SetValue(20d);
            sheet.Cell(4, 1).SetValue("Carrot");
            sheet.Cell(4, 2).SetValue(30d);
            sheet.Cell(5, 1).SetValue("Broccoli");
            sheet.Cell(5, 2).SetValue(40d);
            sheet.Cell(6, 1).SetValue("Banana");
            sheet.Cell(6, 2).SetValue(5d);
            {
                var pivotOptions3 = new ExcelPivotTableOptions();
                pivotOptions3.ColumnFields.Add("Product");
                pivotOptions3.DataFields.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, displayName: "Metric"));
                pivotOptions3.Name = "GroupedColumns";
                sheet.AddPivotTable("A1:B6", "D4", pivotOptions3);
            }

            sheet.AddPivotManualGrouping("GroupedColumns", "Product", "Product2", new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear", "Banana" }, ["Vegetable"] = new[] { "Carrot", "Broccoli" } });
            Assert.True(sheet.MaterializePivotTable("GroupedColumns").Mutation.PackageIsValid);
            Assert.Equal(35d, sheet.GetPivotData("GroupedColumns", "Metric", new Dictionary<string, object?> { ["Product2"] = "Fruit" }).Value);
            Assert.Equal(70d, sheet.GetPivotData("GroupedColumns", "Metric", new Dictionary<string, object?> { ["Product2"] = "Vegetable" }).Value);
            Assert.Equal(30d, sheet.GetPivotData("GroupedColumns", "Metric", new Dictionary<string, object?> { ["Product2"] = "Vegetable", ["Product"] = "Carrot" }).Value);
            document.Save(output);
        }

        [Fact]
        public void Test_PivotManualGroup_ValidatesBeforeMutatingNamedPivot()
        {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Source");
            sheet.Cell(1, 1).SetValue("Product");
            sheet.Cell(1, 2).SetValue("Sales");
            sheet.Cell(2, 1).SetValue("Apple");
            sheet.Cell(2, 2).SetValue(10d);
            sheet.Cell(3, 1).SetValue("Pear");
            sheet.Cell(3, 2).SetValue(20d);
            sheet.Cell(4, 1).SetValue("Carrot");
            sheet.Cell(4, 2).SetValue(30d);
            sheet.AddPivotTable("A1:B4", "E4", options: new ExcelPivotTableOptions { Name = "Grouped", RowFields = new[] { "Product" } });
            sheet.AddPivotTable("A1:B4", "J4", options: new ExcelPivotTableOptions { Name = "Grouped", RowFields = new[] { "Product" } });
            string renamed = document.GetPivotTables().Single(pivot => pivot.Name != "Grouped").Name;
            var target = sheet.WorksheetPart.PivotTableParts.Single(pivot => pivot.PivotTableDefinition?.Name?.Value == renamed);
            string beforeDefinition = target.PivotTableDefinition!.OuterXml;
            string beforeCache = target.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.OuterXml;
            Assert.Throws<ArgumentException>(() => sheet.AddPivotManualGrouping(renamed, "Product", "Product2", new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Unknown" } }));
            Assert.Equal(beforeDefinition, target.PivotTableDefinition.OuterXml);
            Assert.Equal(beforeCache, target.PivotTableCacheDefinitionPart.PivotCacheDefinition.OuterXml);
            sheet.AddPivotManualGrouping(renamed, "Product", "Product2", new Dictionary<string, string[]> { ["Fruit"] = new[] { "Apple", "Pear" } });
            Assert.Equal(2, sheet.WorksheetPart.PivotTableParts.Single(pivot => pivot.PivotTableDefinition?.Name?.Value == "Grouped").PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.ChildElements.Count);
            Assert.Equal(3, target.PivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields!.ChildElements.Count);
        }
    }
}