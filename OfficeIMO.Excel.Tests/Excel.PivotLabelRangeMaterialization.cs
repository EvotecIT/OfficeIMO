using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("greater", "A4:B8", 5, 150d)]
        [InlineData("greater-equal", "A4:B9", 6, 180d)]
        [InlineData("less", "A4:B7", 4, 30d)]
        [InlineData("less-equal", "A4:B8", 5, 60d)]
        [InlineData("between", "A4:B9", 6, 140d)]
        [InlineData("not-between", "A4:B7", 4, 70d)]
        public void Test_PivotLabelRange_ImportedViewAndLookupMatchExcel(
            string kind, string range, int rows, double total) {
            string file = $"pivot-label-range-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(range);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("LabelPivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("LabelPivot");
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(diagnostic => diagnostic.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(range);
            for (int row = 0; row < rows; row++)
                for (int column = 0; column < 2; column++)
                    AssertPivotLookupOracleValue(expectedView[row, column], actualView[row, column]);
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int row = 0; row < 7; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("greater", 150d)]
        [InlineData("greater-equal", 180d)]
        [InlineData("less", 30d)]
        [InlineData("less-equal", 60d)]
        [InlineData("between", 140d)]
        [InlineData("not-between", 70d)]
        public void Test_PivotLabelRange_TemplateFreePublicApi(string kind, double total) {
            string output = Path.Combine(_directoryWithFiles, $"Filter.label-range-{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "Region");
                source.CellValue(1, 2, "Sales");
                var entries = new[] {
                    ("Alpha", 10d), ("Bravo", 20d), ("charlie", 30d),
                    ("Delta", 40d), ("Echo", 50d), ("Foxtrot", 60d)
                };
                for (int index = 0; index < entries.Length; index++) {
                    source.CellValue(index + 2, 1, entries[index].Item1);
                    source.CellValue(index + 2, 2, entries[index].Item2);
                }
                ExcelPivotFilter filter = kind switch {
                    "greater" => ExcelPivotFilter.LabelGreaterThan("Region", "Charlie"),
                    "greater-equal" => ExcelPivotFilter.LabelGreaterThanOrEqual("Region", "Charlie"),
                    "less" => ExcelPivotFilter.LabelLessThan("Region", "Charlie"),
                    "less-equal" => ExcelPivotFilter.LabelLessThanOrEqual("Region", "Charlie"),
                    "between" => ExcelPivotFilter.LabelBetween("Region", "Bravo", "Echo"),
                    "not-between" => ExcelPivotFilter.LabelNotBetween("Region", "Bravo", "Echo"),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:B7").Rows("Region").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("D4", "FilteredPivot");
                var result = source.MaterializePivotTable("FilteredPivot");
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(diagnostic => diagnostic.Message)));
                Assert.Equal(total, source.GetPivotData("FilteredPivot", "Metric").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(total, reopened.GetSheet("Source").GetPivotData("FilteredPivot", "Metric").Value);
        }

        [Fact]
        public void Test_PivotLabelRange_RejectsMismatchedSavedPredicateBeforeWriting() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                "pivot-label-range-between-conformance.xlsx");
            using var document = ExcelDocument.Load(path);
            var grouped = document.GetSheet("Grouped");
            var pivot = grouped.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            pivot.PivotFilters!.Elements<PivotFilter>().Single().AutoFilter!
                .Elements<FilterColumn>().Single().GetFirstChild<CustomFilters>()!.And = false;
            string beforeView = grouped.WorksheetPart.Worksheet.OuterXml;
            string beforePivot = pivot.OuterXml;
            Assert.Throws<NotSupportedException>(() => grouped.MaterializePivotTable("LabelPivot"));
            Assert.Equal(beforeView, grouped.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(beforePivot, pivot.OuterXml);
        }

        [Fact]
        public void Test_PivotLabelRange_RejectsUnqualifiedExcelCollation() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                "pivot-label-range-collation-conformance.xlsx");
            using var document = ExcelDocument.Load(path);
            var grouped = document.GetSheet("Grouped");
            var pivot = grouped.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            Assert.Equal(346d, grouped.GetPivotData("LabelPivot", "Metric").Value);
            string beforeView = grouped.WorksheetPart.Worksheet.OuterXml;
            string beforePivot = pivot.OuterXml;
            var error = Assert.Throws<NotSupportedException>(() => grouped.MaterializePivotTable("LabelPivot"));
            Assert.Contains("localized text ordering", error.Message);
            Assert.Equal(beforeView, grouped.WorksheetPart.Worksheet.OuterXml);
            Assert.Equal(beforePivot, pivot.OuterXml);
        }
    }
}
