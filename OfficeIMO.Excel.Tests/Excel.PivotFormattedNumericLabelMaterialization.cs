using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("contains-comma", "A4:B7", 4, 4, 50d)]
        [InlineData("equals-grouped", "A4:B6", 3, 4, 20d)]
        [InlineData("not-equals-grouped", "A4:B7", 4, 4, 40d)]
        [InlineData("general-contains-one", "A4:B7", 4, 4, 30d)]
        [InlineData("midpoint-positive", "A4:B6", 3, 3, 10d)]
        [InlineData("midpoint-negative", "A4:B6", 3, 3, 10d)]
        [InlineData("decimal-two", "A4:B6", 3, 3, 10d)]
        [InlineData("percent-one", "A4:B6", 3, 3, 10d)]
        [InlineData("decimal-midpoint", "A4:B6", 3, 3, 10d)]
        public void Test_PivotFormattedNumericLabel_ImportedViewAndLookupMatchExcel(
            string kind, string range, int rows, int lookupRows, double total) {
            string file = $"pivot-label-number-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(range);
            string lookupRange = $"B1:B{lookupRows}";
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange(lookupRange);
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("LabelPivot");
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                if (NumericLabelFormat(kind) is string format)
                    Assert.Equal(format, grouped.GetCellStyle(5, 1).NumberFormatCode);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(lookupRows, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(range);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange(lookupRange);
            for (int row = 0; row < lookupRows; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("contains-comma", 50d)]
        [InlineData("equals-grouped", 20d)]
        [InlineData("not-equals-grouped", 40d)]
        [InlineData("general-contains-one", 30d)]
        [InlineData("midpoint-positive", 10d)]
        [InlineData("midpoint-negative", 10d)]
        [InlineData("decimal-two", 10d)]
        [InlineData("percent-one", 10d)]
        [InlineData("decimal-midpoint", 10d)]
        public void Test_PivotFormattedNumericLabel_TemplateFreePublicApi(string kind, double total) {
            string output = Path.Combine(_directoryWithFiles, $"Filter.label-number-{kind}.Authored.xlsx");
            ExcelPivotFilter filter = kind switch {
                "contains-comma" => ExcelPivotFilter.LabelContains("Item", ","),
                "equals-grouped" => ExcelPivotFilter.LabelEquals("Item", "1,000"),
                "general-contains-one" => ExcelPivotFilter.LabelContains("Item", "1"),
                "midpoint-positive" => ExcelPivotFilter.LabelEquals("Item", "3"),
                "midpoint-negative" => ExcelPivotFilter.LabelEquals("Item", "-3"),
                "decimal-two" => ExcelPivotFilter.LabelEquals("Item", "1.20"),
                "percent-one" => ExcelPivotFilter.LabelEquals("Item", "12.5%"),
                "decimal-midpoint" => ExcelPivotFilter.LabelEquals("Item", "1.3"),
                _ => ExcelPivotFilter.LabelNotEquals("Item", "1,000")
            };
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Item");
                source.CellValue(1, 2, "Sales");
                double[] keys = kind == "midpoint-positive" ? new[] { 2.5d, 4.5d }
                    : kind == "midpoint-negative" ? new[] { -2.5d, 4.5d }
                    : kind == "decimal-two" ? new[] { 1.2d, 2.3d }
                    : kind == "percent-one" ? new[] { 0.125d, 0.25d }
                    : kind == "decimal-midpoint" ? new[] { 1.25d, 2.25d }
                    : new[] { 10d, 1000d, 2000d };
                for (int index = 0; index < keys.Length; index++) {
                    source.CellValue(index + 2, 1, keys[index]);
                    source.CellValue(index + 2, 2, 10 * (index + 1));
                }
                var builder = source.Pivot($"A1:B{keys.Length + 1}").Rows("Item").Sum("Sales", "Metric");
                if (NumericLabelFormat(kind) is string fieldFormat) builder.FieldNumberFormat("Item", fieldFormat);
                builder.Layout(ExcelPivotLayout.Tabular).Filter(filter).At("D4", "FilteredPivot");
                var result = source.MaterializePivotTable("FilteredPivot");
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(total, source.GetPivotData("FilteredPivot", "Metric").Value);
                if (NumericLabelFormat(kind) is string outputFormat)
                    Assert.Equal(outputFormat, source.GetCellStyle(5, 4).NumberFormatCode);
                var pivot = source.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
                Assert.False(pivot.PivotFields!.Elements<PivotField>().First().ShowAll!.Value);
                if (kind is "equals-grouped" or "midpoint-positive" or "midpoint-negative"
                    or "decimal-two" or "percent-one" or "decimal-midpoint") {
                    var column = pivot.PivotFilters!.Elements<PivotFilter>().Single().AutoFilter!
                        .Elements<FilterColumn>().Single();
                    Assert.Null(column.GetFirstChild<CustomFilters>());
                    Assert.Equal(filter.Value1, Assert.Single(column.GetFirstChild<Filters>()!.Elements<Filter>()).Val!.Value);
                }
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(total, reopened.GetSheet("Source").GetPivotData("FilteredPivot", "Metric").Value);
            if (NumericLabelFormat(kind) is string savedFormat)
                Assert.Equal(savedFormat, reopened.GetSheet("Source").GetCellStyle(5, 4).NumberFormatCode);
        }

        private static string? NumericLabelFormat(string kind) => kind switch {
            "general-contains-one" => null,
            "decimal-two" => "0.00",
            "percent-one" => "0.0%",
            "decimal-midpoint" => "0.0",
            _ => "#,##0"
        };

        [Fact]
        public void Test_PivotFormattedNumericLabel_MidpointsUseExcelDisplayRounding() {
            Assert.Equal("3", ExcelNumberFormatDisplay.FormatNumericText(2.5d, 3U, null, "2.5"));
            Assert.Equal("-3", ExcelNumberFormatDisplay.FormatNumericText(-2.5d, 3U, null, "-2.5"));
        }

        [Theory]
        [InlineData("pivot-label-number-subtotal-conformance.xlsx")]
        [InlineData("pivot-label-number-subtotal-repeat-conformance.xlsx")]
        public void Test_PivotFormattedNumericSubtotal_ImportedViewAndLookupMatchExcel(string file) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:C10");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B4");
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                var result = grouped.MaterializePivotTable("LabelPivot");
                Assert.Equal("A4:C10", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal("#,##0", grouped.GetCellStyle(5, 1).NumberFormatCode);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(4, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange("A4:C10");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B4");
            for (int row = 0; row < 4; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Fact]
        public void Test_PivotFormattedNumericSubtotal_TemplateFreePublicApi() {
            string output = Path.Combine(_directoryWithFiles, "Filter.label-number-subtotal.Authored.xlsx");
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                "pivot-label-number-subtotal-conformance.xlsx");
            using var oracle = ExcelDocumentReader.Open(path);
            var expected = oracle.GetSheet("Grouped").ReadRange("A4:C10");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Quantity");
                source.CellValue(1, 2, "Product");
                source.CellValue(1, 3, "Sales");
                source.CellValue(2, 1, 1000d); source.CellValue(2, 2, "A"); source.CellValue(2, 3, 10d);
                source.CellValue(3, 1, 1000d); source.CellValue(3, 2, "B"); source.CellValue(3, 3, 20d);
                source.CellValue(4, 1, 2000d); source.CellValue(4, 2, "A"); source.CellValue(4, 3, 30d);
                source.Pivot("A1:C4").Rows("Quantity", "Product").Sum("Sales", "Metric")
                    .FieldNumberFormat("Quantity", "#,##0")
                    .Layout(ExcelPivotLayout.Tabular).At("E4", "FilteredPivot");
                var result = source.MaterializePivotTable("FilteredPivot");
                Assert.Equal("E4:G10", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(60d, source.GetPivotData("FilteredPivot", "Metric").Value);
                var pivot = source.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
                Assert.All(pivot.PivotFields!.Elements<PivotField>().Take(2),
                    field => Assert.False(field.ShowAll!.Value));
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Source").ReadRange("E4:G10");
            for (int row = 0; row < 7; row++)
                for (int column = 0; column < 3; column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void Test_PivotExplicitShowAll_RejectsHeadlessMaterializationWithoutChangingView() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            source.CellValue(1, 1, "Quantity");
            source.CellValue(1, 2, "Product");
            source.CellValue(1, 3, "Sales");
            source.CellValue(2, 1, 1000d); source.CellValue(2, 2, "A"); source.CellValue(2, 3, 10d);
            source.CellValue(3, 1, 1000d); source.CellValue(3, 2, "B"); source.CellValue(3, 3, 20d);
            source.CellValue(4, 1, 2000d); source.CellValue(4, 2, "A"); source.CellValue(4, 3, 30d);
            source.Pivot("A1:C4").Rows("Quantity", "Product").Sum("Sales", "Metric")
                .FieldDisplay("Quantity", showAll: true).Layout(ExcelPivotLayout.Tabular)
                .At("E4", "ShowAllPivot");
            var definition = source.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            string beforePivot = definition.OuterXml;
            string beforeView = source.WorksheetPart.Worksheet.OuterXml;
            Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ShowAllPivot"));
            Assert.Equal(beforePivot, definition.OuterXml);
            Assert.Equal(beforeView, source.WorksheetPart.Worksheet.OuterXml);
        }
    }
}
