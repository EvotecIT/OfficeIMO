using OfficeIMO.Excel;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("outer", "A4:C8", 5, 65d)]
        [InlineData("inner", "A4:C11", 8, 150d)]
        [InlineData("inner-top1", "A4:C11", 8, 150d)]
        [InlineData("inner-less", "A4:C11", 8, 35d)]
        [InlineData("inner-between", "A4:C8", 5, 60d)]
        [InlineData("inner-notbetween", "A4:C11", 8, 125d)]
        [InlineData("inner-bottom1", "A4:C11", 8, 35d)]
        [InlineData("inner-top50pct", "A4:C11", 8, 150d)]
        [InlineData("inner-topsum30", "A4:C11", 8, 150d)]
        [InlineData("outer-bottom1", "A4:C11", 8, 120d)]
        [InlineData("inner-equal", "A4:C7", 4, 40d)]
        [InlineData("inner-notequal", "A4:C13", 10, 145d)]
        [InlineData("inner-greater-equal", "A4:C11", 8, 150d)]
        [InlineData("inner-less-equal", "A4:C11", 8, 35d)]
        [InlineData("inner-bottom25pct", "A4:C13", 10, 145d)]
        [InlineData("inner-bottomsum15", "A4:C13", 10, 145d)]
        [InlineData("outer-top2", "A4:C14", 11, 185d)]
        [InlineData("outer-equal", "A4:C11", 8, 120d)]
        [InlineData("outer-notequal", "A4:C8", 5, 65d)]
        [InlineData("outer-greater-equal", "A4:C8", 5, 65d)]
        [InlineData("outer-less", "A4:C11", 8, 120d)]
        [InlineData("outer-less-equal", "A4:C11", 8, 120d)]
        [InlineData("outer-between", "A4:C8", 5, 65d)]
        [InlineData("outer-notbetween", "A4:C11", 8, 120d)]
        public void Test_PivotMultiFieldValue_ImportedViewAndLookupMatchExcel(
            string kind, string expectedRange, int viewRows, double total) {
            string file = $"pivot-value-multifield-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(expectedRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);
            string output = Path.Combine(_directoryWithFiles, "Materialized." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actualView = reopened.GetSheet("Grouped").ReadRange(expectedRange);
            // Excel sorts parent captions, while OfficeIMO preserves source order.
            // Compare the full saved row multiset, including parent subtotal rows.
            string[] Rows(object?[,] view) => Enumerable.Range(0, viewRows)
                .Select(row => JsonSerializer.Serialize(Enumerable.Range(0, 3)
                    .Select(column => view[row, column]).ToArray()))
                .OrderBy(row => row, StringComparer.Ordinal).ToArray();
            Assert.Equal(Rows(expectedView), Rows(actualView));
            var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
            for (int row = 0; row < 7; row++)
                AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
        }

        [Theory]
        [InlineData("outer", 65d)]
        [InlineData("inner", 150d)]
        [InlineData("inner-top1", 150d)]
        [InlineData("inner-less", 35d)]
        [InlineData("inner-between", 60d)]
        [InlineData("inner-notbetween", 125d)]
        [InlineData("inner-bottom1", 35d)]
        [InlineData("inner-top50pct", 150d)]
        [InlineData("inner-topsum30", 150d)]
        [InlineData("outer-bottom1", 120d)]
        [InlineData("inner-equal", 40d)]
        [InlineData("inner-notequal", 145d)]
        [InlineData("inner-greater-equal", 150d)]
        [InlineData("inner-less-equal", 35d)]
        [InlineData("inner-bottom25pct", 145d)]
        [InlineData("inner-bottomsum15", 145d)]
        [InlineData("outer-top2", 185d)]
        [InlineData("outer-equal", 120d)]
        [InlineData("outer-notequal", 65d)]
        [InlineData("outer-greater-equal", 65d)]
        [InlineData("outer-less", 120d)]
        [InlineData("outer-less-equal", 120d)]
        [InlineData("outer-between", 65d)]
        [InlineData("outer-notbetween", 120d)]
        public void Test_PivotMultiFieldValue_TemplateFreePublicApi(string kind, double total) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                $"pivot-value-multifield-{kind}-conformance.xlsx");
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            string output = Path.Combine(_directoryWithFiles, $"Value.multifield-{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateMultiFieldPivotSource(source);
                ExcelPivotFilter filter = kind switch {
                    "outer" => ExcelPivotFilter.ValueGreaterThan("Region", "Metric", 62d),
                    "inner" => ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 30d),
                    "inner-top1" => ExcelPivotFilter.TopCount("Product", "Metric", 1),
                    "inner-less" => ExcelPivotFilter.ValueLessThan("Product", "Metric", 30d),
                    "inner-between" => ExcelPivotFilter.ValueBetween("Product", "Metric", 15d, 45d),
                    "inner-notbetween" => ExcelPivotFilter.ValueNotBetween("Product", "Metric", 15d, 45d),
                    "inner-bottom1" => ExcelPivotFilter.BottomCount("Product", "Metric", 1),
                    "inner-top50pct" => ExcelPivotFilter.TopPercent("Product", "Metric", 50),
                    "inner-topsum30" => ExcelPivotFilter.TopSum("Product", "Metric", 30d),
                    "outer-bottom1" => ExcelPivotFilter.BottomCount("Region", "Metric", 1),
                    "inner-equal" => ExcelPivotFilter.ValueEquals("Product", "Metric", 40d),
                    "inner-notequal" => ExcelPivotFilter.ValueNotEquals("Product", "Metric", 40d),
                    "inner-greater-equal" => ExcelPivotFilter.ValueGreaterThanOrEqual("Product", "Metric", 40d),
                    "inner-less-equal" => ExcelPivotFilter.ValueLessThanOrEqual("Product", "Metric", 20d),
                    "inner-bottom25pct" => ExcelPivotFilter.BottomPercent("Product", "Metric", 25),
                    "inner-bottomsum15" => ExcelPivotFilter.BottomSum("Product", "Metric", 15d),
                    "outer-top2" => ExcelPivotFilter.TopCount("Region", "Metric", 2),
                    "outer-equal" => ExcelPivotFilter.ValueEquals("Region", "Metric", 60d),
                    "outer-notequal" => ExcelPivotFilter.ValueNotEquals("Region", "Metric", 60d),
                    "outer-greater-equal" => ExcelPivotFilter.ValueGreaterThanOrEqual("Region", "Metric", 65d),
                    "outer-less" => ExcelPivotFilter.ValueLessThan("Region", "Metric", 65d),
                    "outer-less-equal" => ExcelPivotFilter.ValueLessThanOrEqual("Region", "Metric", 60d),
                    "outer-between" => ExcelPivotFilter.ValueBetween("Region", "Metric", 61d, 65d),
                    "outer-notbetween" => ExcelPivotFilter.ValueNotBetween("Region", "Metric", 61d, 65d),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:C7").Rows("Region", "Product").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("E4", "ValuePivot");
                Assert.True(source.MaterializePivotTable("ValuePivot").Mutation.PackageIsValid);
                Assert.Equal(total, source.GetPivotData("ValuePivot", "Metric").Value);
                document.Save(output);
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            var sheet = reopened.GetSheet("Source");
            Assert.Equal(total, sheet.GetPivotData("ValuePivot", "Metric").Value);
            AssertMultiFieldPivotLookups(sheet, expectedLookups);
        }

        [Fact]
        public void Test_PivotMultiFieldValue_ColumnHierarchyMatchesExcelLookups() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                "pivot-value-multifield-column-inner-conformance.xlsx");
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange("A4:H7");
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            string importedOutput = Path.Combine(_directoryWithFiles, "Column.Imported.xlsx");
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(150d, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal("A4:H7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var reopened = ExcelDocumentReader.Open(importedOutput)) {
                var actualView = reopened.GetSheet("Grouped").ReadRange("A4:H7");
                AssertPivotColumnGridMatchesExcel(expectedView, actualView);
                var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
                for (int row = 0; row < 7; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "Column.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateMultiFieldPivotSource(source);
                source.Pivot("A1:C7").Columns("Region", "Product").Sum("Sales", "Metric")
                    .Filter(ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 30d))
                    .At("E4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal("E4:L7", result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(authoredOutput);
            }
            using var authored = ExcelDocument.Load(authoredOutput);
            Assert.Empty(authored.ValidateOpenXml());
            var sheet = authored.GetSheet("Source");
            Assert.Equal(150d, sheet.GetPivotData("ValuePivot", "Metric").Value);
            AssertMultiFieldPivotLookups(sheet, expectedLookups);
            using var authoredView = ExcelDocumentReader.Open(authoredOutput);
            AssertPivotColumnGridMatchesExcel(expectedView, authoredView.GetSheet("Source").ReadRange("E4:L7"));
        }

        [Theory]
        [InlineData("mixed-column", "A4:C9", "E4:G9", 6, 3, 130d)]
        [InlineData("mixed-row", "A4:D7", "E4:H7", 4, 4, 65d)]
        [InlineData("mixed-column-top1", "A4:C9", "E4:G9", 6, 3, 130d)]
        [InlineData("mixed-column-less", "A4:C9", "E4:G9", 6, 3, 55d)]
        [InlineData("mixed-column-equal", "A4:C9", "E4:G9", 6, 3, 55d)]
        [InlineData("mixed-column-notequal", "A4:C9", "E4:G9", 6, 3, 130d)]
        [InlineData("mixed-column-between", "A4:C9", "E4:G9", 6, 3, 55d)]
        [InlineData("mixed-column-bottom1", "A4:C9", "E4:G9", 6, 3, 55d)]
        [InlineData("mixed-column-top2", "A4:D9", "E4:H9", 6, 4, 185d)]
        [InlineData("mixed-column-top50pct", "A4:C9", "E4:G9", 6, 3, 130d)]
        [InlineData("mixed-column-bottomsum60", "A4:D9", "E4:H9", 6, 4, 185d)]
        [InlineData("mixed-row-less", "A4:D8", "E4:H8", 5, 4, 120d)]
        [InlineData("mixed-row-notequal", "A4:D7", "E4:H7", 4, 4, 65d)]
        [InlineData("mixed-row-between", "A4:D7", "E4:H7", 4, 4, 65d)]
        [InlineData("mixed-row-notbetween", "A4:D8", "E4:H8", 5, 4, 120d)]
        [InlineData("mixed-row-bottom1", "A4:D8", "E4:H8", 5, 4, 120d)]
        [InlineData("mixed-row-top2", "A4:D9", "E4:H9", 6, 4, 185d)]
        [InlineData("mixed-row-top25pct", "A4:D7", "E4:H7", 4, 4, 65d)]
        [InlineData("mixed-row-bottomsum70", "A4:D8", "E4:H8", 5, 4, 120d)]
        public void Test_PivotMultiFieldValue_MixedAxesMatchExcel(
            string kind, string expectedRange, string authoredRange, int height, int width, double total) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                $"pivot-value-multifield-{kind}-conformance.xlsx");
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(expectedRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B7");
            string importedOutput = Path.Combine(_directoryWithFiles, $"Mixed.{kind}.Imported.xlsx");
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                Assert.Equal(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(7, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var reopened = ExcelDocumentReader.Open(importedOutput)) {
                var actualView = reopened.GetSheet("Grouped").ReadRange(expectedRange);
                AssertPivotMixedGridMatchesExcel(expectedView, actualView, height, width);
                var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B7");
                for (int row = 0; row < 7; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, $"Mixed.{kind}.Authored.xlsx");
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateMultiFieldPivotSource(source);
                var filter = kind switch {
                    "mixed-column" => ExcelPivotFilter.ValueGreaterThan("Product", "Metric", 100d),
                    "mixed-row" => ExcelPivotFilter.ValueGreaterThan("Region", "Metric", 62d),
                    "mixed-column-top1" => ExcelPivotFilter.TopCount("Product", "Metric", 1),
                    "mixed-column-less" => ExcelPivotFilter.ValueLessThan("Product", "Metric", 100d),
                    "mixed-column-equal" => ExcelPivotFilter.ValueEquals("Product", "Metric", 55d),
                    "mixed-column-notequal" => ExcelPivotFilter.ValueNotEquals("Product", "Metric", 55d),
                    "mixed-column-between" => ExcelPivotFilter.ValueBetween("Product", "Metric", 50d, 60d),
                    "mixed-column-bottom1" => ExcelPivotFilter.BottomCount("Product", "Metric", 1),
                    "mixed-column-top2" => ExcelPivotFilter.TopCount("Product", "Metric", 2),
                    "mixed-column-top50pct" => ExcelPivotFilter.TopPercent("Product", "Metric", 50),
                    "mixed-column-bottomsum60" => ExcelPivotFilter.BottomSum("Product", "Metric", 60d),
                    "mixed-row-less" => ExcelPivotFilter.ValueLessThan("Region", "Metric", 62d),
                    "mixed-row-notequal" => ExcelPivotFilter.ValueNotEquals("Region", "Metric", 60d),
                    "mixed-row-between" => ExcelPivotFilter.ValueBetween("Region", "Metric", 61d, 65d),
                    "mixed-row-notbetween" => ExcelPivotFilter.ValueNotBetween("Region", "Metric", 61d, 65d),
                    "mixed-row-bottom1" => ExcelPivotFilter.BottomCount("Region", "Metric", 1),
                    "mixed-row-top2" => ExcelPivotFilter.TopCount("Region", "Metric", 2),
                    "mixed-row-top25pct" => ExcelPivotFilter.TopPercent("Region", "Metric", 25),
                    "mixed-row-bottomsum70" => ExcelPivotFilter.BottomSum("Region", "Metric", 70d),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:C7").Rows("Region").Columns("Product").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("E4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal(authoredRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                document.Save(authoredOutput);
            }
            using (var authored = ExcelDocument.Load(authoredOutput)) {
                Assert.Empty(authored.ValidateOpenXml());
                var sheet = authored.GetSheet("Source");
                Assert.Equal(total, sheet.GetPivotData("ValuePivot", "Metric").Value);
                AssertMultiFieldPivotLookups(sheet, expectedLookups);
            }
            using var authoredView = ExcelDocumentReader.Open(authoredOutput);
            AssertPivotMixedGridMatchesExcel(expectedView,
                authoredView.GetSheet("Source").ReadRange(authoredRange), height, width);
        }

        [Theory]
        [InlineData("mixed-row-bottomsum50", 60d)]
        [InlineData("outer-tie-bottomsum50", 60d)]
        [InlineData("outer-tie-top50pct", 125d)]
        [InlineData("outer-tie-bottom25pct", 60d)]
        [InlineData("outer-tie-bottomsum50-reversed", 60d)]
        [InlineData("outer-tie-bottomsum50-three", 60d)]
        public void Test_PivotMultiFieldValue_TiedPercentSumCutoffRejectsUnqualifiedItemOrder(
            string kind, double excelTotal) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus",
                $"pivot-value-multifield-{kind}-conformance.xlsx");
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            AssertPivotLookupOracleValue(excelTotal, oracle.GetSheet("Lookups").ReadRange("B1:B1")[0, 0]);
            using var document = ExcelDocument.Load(path);
            var grouped = document.GetSheet("Grouped");
            Assert.Equal(excelTotal, grouped.GetPivotData("ValuePivot", "Metric").Value);
            var error = Assert.Throws<NotSupportedException>(() => grouped.MaterializePivotTable("ValuePivot"));
            Assert.Contains("tied cutoff", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(excelTotal, grouped.GetPivotData("ValuePivot", "Metric").Value);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_PivotMultiFieldValue_TemplateFreeTiedSumFailsBeforeWritingView() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            PopulateMultiFieldPivotSource(source);
            source.Pivot("A1:C7").Rows("Region").Columns("Product").Sum("Sales", "Metric")
                .Layout(ExcelPivotLayout.Tabular)
                .Filter(ExcelPivotFilter.BottomSum("Region", "Metric", 50d))
                .At("E4", "ValuePivot");
            var error = Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("ValuePivot"));
            Assert.Contains("tied cutoff", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Theory]
        [InlineData("inner-wide-top2", "A4:C14", "E4:G14", 210d)]
        [InlineData("inner-wide-bottom2", "A4:C14", "E4:G14", 45d)]
        [InlineData("inner-wide-top50pct", "A4:C11", "E4:G11", 150d)]
        [InlineData("inner-wide-bottomsum15", "A4:C15", "E4:G15", 95d)]
        [InlineData("inner-wide-error-top2", "A4:C14", "E4:G14", 210d)]
        [InlineData("outer-wide-top50pct", "A4:C13", "E4:G13", 155d)]
        [InlineData("outer-wide-bottom50pct", "A4:C13", "E4:G13", 100d)]
        [InlineData("outer-wide-topsum50", "A4:C9", "E4:G9", 95d)]
        [InlineData("outer-wide-bottomsum50", "A4:C13", "E4:G13", 100d)]
        [InlineData("outer-negative-top40pct", "A4:C13", "E4:G13", -90d)]
        [InlineData("outer-negative-bottom40pct", "A4:C9", "E4:G9", -60d)]
        [InlineData("outer-negative-topsum30", "A4:C17", "E4:G17", -150d)]
        [InlineData("outer-negative-bottomsum30", "A4:C17", "E4:G17", -150d)]
        [InlineData("outer-zero-bottom1", "A4:C13", "E4:G13", 0d)]
        [InlineData("outer-zero-top50pct", "A4:C9", "E4:G9", 40d)]
        [InlineData("outer-zero-bottom50pct", "A4:C17", "E4:G17", 40d)]
        [InlineData("inner-errorparent-top1", "A4:C11", "E4:G11", "#DIV/0!")]
        [InlineData("inner-errorparent-bottom1", "A4:C11", "E4:G11", "#N/A")]
        public void Test_PivotMultiFieldValue_WideRankingMatchesExcel(
            string kind, string oracleRange, string authoredRange, object total) {
            string file = $"pivot-value-multifield-{kind}-conformance.xlsx";
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", file);
            AssertPivotMultiFieldFixtureHash(path);
            using var oracle = ExcelDocumentReader.Open(path);
            var expectedView = oracle.GetSheet("Grouped").ReadRange(oracleRange);
            var expectedLookups = oracle.GetSheet("Lookups").ReadRange("B1:B10");
            AssertPivotLookupOracleValue(total, expectedLookups[0, 0]);

            string importedOutput = Path.Combine(_directoryWithFiles, "Imported." + file);
            using (var document = ExcelDocument.Load(path)) {
                var grouped = document.GetSheet("Grouped");
                AssertPivotLookupOracleValue(total, grouped.GetPivotData("ValuePivot", "Metric").Value);
                var result = grouped.MaterializePivotTable("ValuePivot");
                Assert.Equal(oracleRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookups = document.GetSheet("Lookups");
                lookups.ClearCachedFormulaResults();
                Assert.Equal(10, lookups.RecalculateSupportedFormulas());
                document.Save(importedOutput);
            }
            using (var reopened = ExcelDocumentReader.Open(importedOutput)) {
                AssertWidePivotRowsMatchExcel(expectedView, reopened.GetSheet("Grouped").ReadRange(oracleRange));
                var actualLookups = reopened.GetSheet("Lookups").ReadRange("B1:B10");
                for (int row = 0; row < 10; row++)
                    AssertPivotLookupOracleValue(expectedLookups[row, 0], actualLookups[row, 0]);
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "Authored." + file);
            using (var document = ExcelDocument.Create()) {
                var source = document.AddWorksheet("Source");
                PopulateWidePivotSource(source, kind);
                var filter = kind switch {
                    "inner-wide-top2" => ExcelPivotFilter.TopCount("Product", "Metric", 2),
                    "inner-wide-bottom2" => ExcelPivotFilter.BottomCount("Product", "Metric", 2),
                    "inner-wide-top50pct" => ExcelPivotFilter.TopPercent("Product", "Metric", 50),
                    "inner-wide-bottomsum15" => ExcelPivotFilter.BottomSum("Product", "Metric", 15d),
                    "inner-wide-error-top2" => ExcelPivotFilter.TopCount("Product", "Metric", 2),
                    "outer-wide-top50pct" => ExcelPivotFilter.TopPercent("Region", "Metric", 50),
                    "outer-wide-bottom50pct" => ExcelPivotFilter.BottomPercent("Region", "Metric", 50),
                    "outer-wide-topsum50" => ExcelPivotFilter.TopSum("Region", "Metric", 50d),
                    "outer-wide-bottomsum50" => ExcelPivotFilter.BottomSum("Region", "Metric", 50d),
                    "outer-negative-top40pct" => ExcelPivotFilter.TopPercent("Region", "Metric", 40),
                    "outer-negative-bottom40pct" => ExcelPivotFilter.BottomPercent("Region", "Metric", 40),
                    "outer-negative-topsum30" => ExcelPivotFilter.TopSum("Region", "Metric", 30d),
                    "outer-negative-bottomsum30" => ExcelPivotFilter.BottomSum("Region", "Metric", 30d),
                    "outer-zero-bottom1" => ExcelPivotFilter.BottomCount("Region", "Metric", 1),
                    "outer-zero-top50pct" => ExcelPivotFilter.TopPercent("Region", "Metric", 50),
                    "outer-zero-bottom50pct" => ExcelPivotFilter.BottomPercent("Region", "Metric", 50),
                    "inner-errorparent-top1" => ExcelPivotFilter.TopCount("Product", "Metric", 1),
                    "inner-errorparent-bottom1" => ExcelPivotFilter.BottomCount("Product", "Metric", 1),
                    _ => throw new ArgumentOutOfRangeException(nameof(kind))
                };
                source.Pivot("A1:C10").Rows("Region", "Product").Sum("Sales", "Metric")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("E4", "ValuePivot");
                var result = source.MaterializePivotTable("ValuePivot");
                Assert.Equal(authoredRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                document.Save(authoredOutput);
            }
            using (var reopened = ExcelDocument.Load(authoredOutput)) {
                Assert.Empty(reopened.ValidateOpenXml());
                var sheet = reopened.GetSheet("Source");
                AssertPivotLookupOracleValue(total, sheet.GetPivotData("ValuePivot", "Metric").Value);
                AssertWideMultiFieldPivotLookups(sheet, expectedLookups);
            }
            using var authoredView = ExcelDocumentReader.Open(authoredOutput);
            AssertWidePivotRowsMatchExcel(expectedView,
                authoredView.GetSheet("Source").ReadRange(authoredRange));
        }

        private static void AssertWidePivotRowsMatchExcel(object?[,] expected, object?[,] actual) {
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            string[] Rows(object?[,] view) => Enumerable.Range(0, view.GetLength(0))
                .Select(row => JsonSerializer.Serialize(Enumerable.Range(0, view.GetLength(1))
                    .Select(column => view[row, column]).ToArray()))
                .OrderBy(row => row, StringComparer.Ordinal).ToArray();
            Assert.Equal(Rows(expected), Rows(actual));
        }

        private static void PopulateWidePivotSource(ExcelSheet source, string kind) {
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Product");
            source.CellValue(1, 3, "Sales");
            string[] regions = { "East", "East", "East", "West", "West", "West", "South", "South", "South" };
            string[] products = { "A", "B", "C", "A", "B", "C", "A", "B", "C" };
            double[] amounts = kind.StartsWith("outer-negative-", StringComparison.Ordinal)
                ? new[] { -30d, -30d, 0d, -30d, -20d, 0d, -20d, -20d, 0d }
                : kind.StartsWith("outer-zero-", StringComparison.Ordinal)
                    ? new[] { -10d, 10d, 0d, -5d, 5d, 0d, 20d, 20d, 0d }
                    : new[] { 10d, 50d, -20d, 40d, 20d, 0d, 5d, 60d, 30d };
            for (int index = 0; index < amounts.Length; index++) {
                source.CellValue(index + 2, 1, regions[index]);
                source.CellValue(index + 2, 2, products[index]);
                if (kind.StartsWith("inner-errorparent-", StringComparison.Ordinal) && index < 3) {
                    string[] errors = { "#N/A", "#DIV/0!", "#VALUE!" };
                    source.CellError(index + 2, 3, errors[index]);
                } else if (kind == "inner-wide-error-top2" && index == 2)
                    source.CellError(index + 2, 3, "#N/A");
                else source.CellValue(index + 2, 3, amounts[index]);
            }
        }

        private static void AssertWideMultiFieldPivotLookups(ExcelSheet sheet, object?[,] expectedLookups) {
            string[] regions = { "East", "East", "East", "West", "West", "West", "South", "South", "South" };
            string[] products = { "A", "B", "C", "A", "B", "C", "A", "B", "C" };
            for (int index = 0; index < regions.Length; index++) {
                var result = sheet.GetPivotData("ValuePivot", "Metric",
                    new Dictionary<string, object?> { ["Region"] = regions[index], ["Product"] = products[index] });
                AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], result.Value);
            }
        }

        private static void AssertPivotMixedGridMatchesExcel(object?[,] expected, object?[,] actual,
            int height, int width) {
            string Row(object?[,] view, int row) => JsonSerializer.Serialize(Enumerable.Range(0, width)
                .Select(column => view[row, column]).ToArray());
            // Excel sorts row captions while OfficeIMO retains source order.
            Assert.Equal(Row(expected, 0), Row(actual, 0));
            Assert.Equal(Row(expected, 1), Row(actual, 1));
            Assert.Equal(Row(expected, height - 1), Row(actual, height - 1));
            string[] Body(object?[,] view) => Enumerable.Range(2, height - 3)
                .Select(row => Row(view, row)).OrderBy(row => row, StringComparer.Ordinal).ToArray();
            Assert.Equal(Body(expected), Body(actual));
        }

        private static void AssertPivotColumnGridMatchesExcel(object?[,] expected, object?[,] actual) {
            // Excel uses "Column Labels" and caption order; OfficeIMO uses the field name
            // and source order. Compare every displayed child, subtotal, and grand-total cell.
            Assert.Equal("Column Labels", expected[0, 1]);
            Assert.Equal("Region", actual[0, 1]);
            for (int row = 0; row < 4; row++) {
                AssertPivotLookupOracleValue(expected[row, 0], actual[row, 0]);
                if (row > 0) AssertPivotLookupOracleValue(expected[row, 7], actual[row, 7]);
            }
            for (int column = 2; column < 8; column++)
                AssertPivotLookupOracleValue(expected[0, column], actual[0, column]);
            string[] Groups(object?[,] view) => Enumerable.Range(0, 3).Select(index => {
                int column = 1 + 2 * index;
                return JsonSerializer.Serialize(Enumerable.Range(1, 3).Select(row =>
                    new[] { view[row, column], view[row, column + 1] }).ToArray());
            }).OrderBy(group => group, StringComparer.Ordinal).ToArray();
            Assert.Equal(Groups(expected), Groups(actual));
        }

        private static void AssertPivotMultiFieldFixtureHash(string path) {
            using var provenance = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, "provenance.json")));
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(path);
            string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
            Assert.Equal(provenance.RootElement.GetProperty("sha256").GetString(), hash);
        }

        private static void PopulateMultiFieldPivotSource(ExcelSheet source) {
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Product");
            source.CellValue(1, 3, "Sales");
            var rows = new[] {
                ("East", "A", 10d), ("East", "B", 50d),
                ("West", "A", 40d), ("West", "B", 20d),
                ("South", "A", 5d), ("South", "B", 60d)
            };
            for (int index = 0; index < rows.Length; index++) {
                source.CellValue(index + 2, 1, rows[index].Item1);
                source.CellValue(index + 2, 2, rows[index].Item2);
                source.CellValue(index + 2, 3, rows[index].Item3);
            }
        }

        private static void AssertMultiFieldPivotLookups(ExcelSheet sheet, object?[,] expectedLookups) {
            string[] regions = { "East", "East", "West", "West", "South", "South" };
            string[] products = { "A", "B", "A", "B", "A", "B" };
            for (int index = 0; index < regions.Length; index++) {
                var result = sheet.GetPivotData("ValuePivot", "Metric",
                    new Dictionary<string, object?> { ["Region"] = regions[index], ["Product"] = products[index] });
                AssertPivotLookupOracleValue(expectedLookups[index + 1, 0], result.Value);
            }
        }
    }
}
