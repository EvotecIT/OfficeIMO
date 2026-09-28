using OfficeIMO.Excel;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotDecimalGroupOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "decimal-group-conformance.xlsx");
        private static string PivotDecimalGroupRefreshOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
            "Documents", "ExcelPivotCorpus", "decimal-group-refresh-conformance.xlsx");

        private static IReadOnlyDictionary<string, object?>[] DecimalGroupSelections => new IReadOnlyDictionary<string, object?>[] {
            new Dictionary<string, object?>(),
            new Dictionary<string, object?> { ["Quantity"] = "<0.25" },
            new Dictionary<string, object?> { ["Quantity"] = "0.25-0.75" },
            new Dictionary<string, object?> { ["Quantity"] = "0.75-1.25" },
            new Dictionary<string, object?> { ["Quantity"] = "1.25-1.75" },
            new Dictionary<string, object?> { ["Quantity"] = "1.75-2.25" },
            new Dictionary<string, object?> { ["Quantity"] = ">2.25" },
            new Dictionary<string, object?> { ["Quantity"] = 0.25d }
        };

        [Fact]
        public void Test_PivotDecimalGroup_ImportedMaterializationAndLookupMatchExcel() {
            string output = Path.Combine(_directoryWithFiles, "DecimalGroup.Materialized.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDecimalGroupOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:B11");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B8");
            using (var document = ExcelDocument.Load(PivotDecimalGroupOraclePath)) {
                var sheet = document.GetSheet("Grouped");
                for (int index = 0; index < DecimalGroupSelections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotGrouped", "Metric", DecimalGroupSelections[index]).Value);
                var result = sheet.MaterializePivotTable("PivotGrouped");
                Assert.Equal("A4:B11", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(8, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange("A4:B11");
            for (int row = 0; row < 8; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B8");
            for (int index = 0; index < 8; index++)
                AssertPivotLookupOracleValue(expectedLookups[index, 0], actualLookups[index, 0]);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_PivotDecimalGroup_TemplateFreePublicApiMatchesExcel(bool numbersStoredAsText) {
            string output = Path.Combine(_directoryWithFiles, numbersStoredAsText
                ? "DecimalGroup.TextNumbers.xlsx" : "DecimalGroup.TemplateFree.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDecimalGroupOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:B12");
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:B11");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                for (int row = 0; row < 12; row++)
                    for (int column = 0; column < 2; column++)
                        sheet.CellValue(row + 1, column + 1, numbersStoredAsText && row > 0 && column == 0
                            ? Convert.ToString(sourceValues[row, column], CultureInfo.InvariantCulture)
                            : sourceValues[row, column]);
                sheet.Pivot("A1:B12").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 0.5, 0.25, 2.25).Layout(ExcelPivotLayout.Tabular)
                    .At("D4", "PivotGrouped");
                var result = sheet.MaterializePivotTable("PivotGrouped");
                Assert.Equal("D4:E11", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("D4:E11");
            for (int row = 0; row < 8; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDecimalGroup_RefreshRegroupsAndClearsOldTail() {
            string output = Path.Combine(_directoryWithFiles, "DecimalGroup.Refreshed.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDecimalGroupRefreshOraclePath);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:B11");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B8");
            using (var document = ExcelDocument.Load(PivotDecimalGroupOraclePath)) {
                var source = document.GetSheet("Source");
                source.CellValue(2, 1, 0.75d);
                source.CellValue(3, 1, 2.26d);
                var sheet = document.GetSheet("Grouped");
                Assert.Equal("A4:B10", sheet.MaterializePivotTable("PivotGrouped").OutputRange);
                for (int index = 0; index < DecimalGroupSelections.Length; index++)
                    AssertPivotLookupOracleValue(expectedLookups[index, 0],
                        sheet.GetPivotData("PivotGrouped", "Metric", DecimalGroupSelections[index]).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Grouped").ReadRange("A4:B11");
            for (int row = 0; row < 8; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotDecimalGroup_AuthoringFieldOptionsUseDisplayedLabels() {
            string output = Path.Combine(_directoryWithFiles, "DecimalGroup.HiddenBucket.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotDecimalGroupOraclePath);
            var sourceValues = oracle.GetSheet("Source").ReadRange("A1:B12");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                for (int row = 0; row < 12; row++)
                    for (int column = 0; column < 2; column++)
                        sheet.CellValue(row + 1, column + 1, sourceValues[row, column]);
                sheet.Pivot("A1:B12").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 0.5, 0.25, 2.25).HideItems("Quantity", "0.25-0.75")
                    .At("D4", "PivotGrouped");
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            var pivot = Assert.Single(reopened.GetPivotTables());
            var field = Assert.Single(pivot.Fields, field => field.FieldName == "Quantity");
            Assert.Equal(new[] { "0.25-0.75" }, field.HiddenItems);
            Assert.Equal(new[] { "<0.25", "0.75-1.25", "1.25-1.75", "1.75-2.25", ">2.25" },
                field.VisibleItems);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Fact]
        public void Test_PivotDecimalGroup_RejectsNonFiniteAndReversedAuthoringBounds() {
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelPivotGrouping.Number("Quantity", double.NaN, 0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelPivotGrouping.Number("Quantity", double.PositiveInfinity, 0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelPivotGrouping.Number("Quantity", 0.5, double.NaN, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelPivotGrouping.Number("Quantity", 0.5, 0, double.NegativeInfinity));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelPivotGrouping.Number("Quantity", 0.5, 1, 0));
        }

        [Fact]
        public void Test_PivotDecimalGroup_TenthsStayOnTheirBoundaries() {
            string output = Path.Combine(_directoryWithFiles, "DecimalGroup.Tenths.xlsx");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                sheet.CellValue(1, 1, "Quantity");
                sheet.CellValue(1, 2, "Sales");
                for (int index = 1; index <= 4; index++) {
                    sheet.CellValue(index + 1, 1, index / 10d);
                    sheet.CellValue(index + 1, 2, index);
                }
                sheet.Pivot("A1:B5").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 0.1, 0.1, 0.4).Layout(ExcelPivotLayout.Tabular)
                    .At("D4", "PivotTenths");
                Assert.True(sheet.MaterializePivotTable("PivotTenths").Mutation.PackageIsValid);
                Assert.Equal(7d, sheet.GetPivotData("PivotTenths", "Metric",
                    new Dictionary<string, object?> { ["Quantity"] = 0.3d }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var view = reopened.GetSheet("Source").ReadRange("D5:E8");
            Assert.Equal("0.1-0.2", view[0, 0]);
            Assert.Equal(1d, view[0, 1]);
            Assert.Equal("0.2-0.3", view[1, 0]);
            Assert.Equal(2d, view[1, 1]);
            Assert.Equal("0.3-0.4", view[2, 0]);
            Assert.Equal(7d, view[2, 1]);
            Assert.Equal(10d, view[3, 1]);
        }

        [Fact]
        public void Test_PivotDecimalGroup_PreservesAdjacentLargeIntegerKeys() {
            const double start = 8_999_999_999_999_998d;
            const double next = 8_999_999_999_999_999d;
            const double end = 9_000_000_000_000_000d;
            string output = Path.Combine(_directoryWithFiles, "NumericGroup.LargeIntegers.xlsx");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Source");
                sheet.CellValue(1, 1, "Quantity");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(2, 1, start);
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, next);
                sheet.CellValue(3, 2, 2d);
                sheet.CellValue(4, 1, end);
                sheet.CellValue(4, 2, 3d);
                sheet.Pivot("A1:B4").Rows("Quantity").Sum("Sales", "Metric")
                    .NumberGroup("Quantity", 1, start, end).Layout(ExcelPivotLayout.Tabular)
                    .At("D4", "PivotLargeIntegers");
                Assert.True(sheet.MaterializePivotTable("PivotLargeIntegers").Mutation.PackageIsValid);
                Assert.Equal(1d, sheet.GetPivotData("PivotLargeIntegers", "Metric",
                    new Dictionary<string, object?> { ["Quantity"] = start }).Value);
                Assert.Equal(5d, sheet.GetPivotData("PivotLargeIntegers", "Metric",
                    new Dictionary<string, object?> { ["Quantity"] = next }).Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var view = reopened.GetSheet("Source").ReadRange("D5:E7");
            Assert.Equal("8999999999999998-8999999999999998", view[0, 0]);
            Assert.Equal(1d, view[0, 1]);
            Assert.Equal("8999999999999999-9000000000000000", view[1, 0]);
            Assert.Equal(5d, view[1, 1]);
            Assert.Equal(6d, view[2, 1]);
        }
    }
}
