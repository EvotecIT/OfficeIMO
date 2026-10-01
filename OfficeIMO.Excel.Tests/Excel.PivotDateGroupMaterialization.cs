using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotDateGroupOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-conformance.xlsx");
        private static string PivotDateGroupRefreshOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-refresh-conformance.xlsx");
        private static string PivotDateGroupValuesFirstOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-values-first-conformance.xlsx");
        private static string PivotDateGroupValuesMiddleOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-values-middle-conformance.xlsx");
        private static string PivotDateGroupValuesLastOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-values-last-conformance.xlsx");
        private static string PivotDateGroupColumnsOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "date-group-columns-conformance.xlsx");

        [Fact]
        public void Test_PivotDateGroupMaterialization_MatchesExcelViewAndLookup() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.Materialized.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:C12");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B6");
            using (var document = ExcelDocument.Load(PivotDateGroupOraclePath)) {
                var sheet = document.GetSheet("Grouped");
                Assert.Equal(150d, sheet.GetPivotData("PivotDateGrouped", "Metric").Value);
                Assert.Equal(60d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2025" }).Value);
                Assert.Equal(60d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = 2025d }).Value);
                Assert.Equal(10d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2025", ["Months (OrderDate)"] = "sty" }).Value);
                var result = sheet.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("A4:C12", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(40d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2026", ["Months (OrderDate)"] = "sty" }).Value);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(6, lookups.RecalculateSupportedFormulas());
                for (int row = 0; row < 6; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], lookups.CellAt(row + 1, 2).GetValue().Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:C12");
            for (int row = 0; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_TemplateFreePublicApi() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.TemplateFree.xlsx");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                sheet.CellValue(1, 1, "OrderDate");
                sheet.CellValue(1, 2, "Sales");
                var dates = new[] { new DateTime(2025, 1, 15), new DateTime(2025, 3, 20),
                    new DateTime(2025, 7, 1), new DateTime(2026, 1, 10), new DateTime(2026, 4, 5) };
                for (int index = 0; index < dates.Length; index++) {
                    sheet.CellValue(index + 2, 1, dates[index]);
                    sheet.CellValue(index + 2, 2, (index + 1) * 10d);
                }
                sheet.Pivot("A1:B6").Rows("OrderDate").Sum("Sales", "Metric")
                    .DateHierarchy("OrderDate", ExcelPivotGroupBy.Years, ExcelPivotGroupBy.Months)
                    .Layout(ExcelPivotLayout.Tabular).At("D4", "PivotDateGrouped");
                var result = sheet.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("D4:F12", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(150d, sheet.GetPivotData("PivotDateGrouped", "Metric").Value);
                Assert.Equal(60d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["OrderDate Years"] = "2025" }).Value);
                Assert.Equal(40d, sheet.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["OrderDate Years"] = "2026", ["OrderDate Months"] = "January" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(90d, reopened.GetSheet("Source").GetPivotData("PivotDateGrouped", "Metric",
                new Dictionary<string, object?> { ["OrderDate Years"] = "2026" }).Value);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_RefreshRegroupsAndClearsOldTail() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupRefreshOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:C10");
            using (var document = ExcelDocument.Load(PivotDateGroupOraclePath)) {
                var source = document.GetSheet("Source");
                source.CellValue(3, 1, new DateTime(2025, 1, 15));
                source.CellValue(4, 1, new DateTime(2025, 1, 15));
                var grouped = document.GetSheet("Grouped");
                Assert.Equal("A4:C10", grouped.MaterializePivotTable("PivotDateGrouped").OutputRange);
                var pivotPart = grouped.WorksheetPart.PivotTableParts.Single();
                int savedSourceKeys = pivotPart.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!
                    .Elements<CacheField>().First().SharedItems!.ChildElements.Count;
                Assert.Equal(3, savedSourceKeys);
                Assert.All(pivotPart.PivotTableDefinition!.PivotFields!.Elements<PivotField>().First().Items!
                    .Elements<Item>().Where(item => item.Index != null), item => Assert.True(item.Index!.Value < savedSourceKeys));
                Assert.Equal(60d, grouped.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2025", ["Months (OrderDate)"] = "sty" }).Value);
                Assert.Equal("#REF!", grouped.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2025", ["Months (OrderDate)"] = "mar" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:C12");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
            for (int row = 7; row < 9; row++)
                for (int column = 0; column < 3; column++)
                    Assert.Null(actual[row, column]);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_ValuesFirstMultipleMeasuresMatchesExcelView() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.ValuesFirst.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupValuesFirstOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:D20");
            using (var document = ExcelDocument.Load(PivotDateGroupValuesFirstOraclePath)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("A4:D20", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(3d, grouped.GetPivotData("PivotDateGrouped", "Count",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = 2025d }).Value);
                Assert.Equal(40d, grouped.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = "2026", ["Months (OrderDate)"] = "sty" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:D20");
            for (int row = 0; row < 17; row++)
                for (int column = 0; column < 4; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_ValuesBetweenDateLevelsMatchesExcelView() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.ValuesMiddle.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupValuesMiddleOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:D20");
            using (var document = ExcelDocument.Load(PivotDateGroupValuesMiddleOraclePath)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("A4:D20", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(2d, grouped.GetPivotData("PivotDateGrouped", "Count",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = 2026d }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:D20");
            for (int row = 0; row < 17; row++)
                for (int column = 0; column < 4; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_ValuesLastMultipleMeasuresMatchesExcelView() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.ValuesLast.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupValuesLastOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:D20");
            using (var document = ExcelDocument.Load(PivotDateGroupValuesLastOraclePath)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("A4:D20", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(5d, grouped.GetPivotData("PivotDateGrouped", "Count").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:D20");
            for (int row = 0; row < 17; row++)
                for (int column = 0; column < 4; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDateGroupMaterialization_ColumnDateHierarchyMatchesExcelView() {
            string output = Path.Combine(_directoryWithFiles, "DateGroup.Columns.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDateGroupColumnsOraclePath);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:I7");
            using (var document = ExcelDocument.Load(PivotDateGroupColumnsOraclePath)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("PivotDateGrouped");
                Assert.Equal("A4:I7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(50d, grouped.GetPivotData("PivotDateGrouped", "Metric",
                    new Dictionary<string, object?> { ["Years (OrderDate)"] = 2026d,
                        ["Months (OrderDate)"] = "kwi" }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:I7");
            for (int row = 0; row < 4; row++)
                for (int column = 0; column < 9; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }
    }
}
