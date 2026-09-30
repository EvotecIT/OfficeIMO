using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_PivotLayoutMaterialization_RejectsAmbiguousMeasureCaptionsBeforeMutation() {
            using var document = ExcelDocument.Load(PivotLayoutOraclePath);
            var sheet = document.GetSheet("ScalarCol");
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            definition.DataFields!.Elements<DataField>().Last().Name = "revenue";
            string before = sheet.WorksheetPart.Worksheet.OuterXml;
            Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotScalarCol"));
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
        }

        [Fact]
        public void Test_AddPivotTable_RejectsAmbiguousMeasureCaptionsBeforeCreatingParts() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Key"); sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "A"); sheet.CellValue(2, 2, 1d);
            Assert.Throws<ArgumentException>(() => sheet.AddPivotTable("A1:B2", "D1", "Pivot", rowFields: new[] { "Key" },
                dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue"),
                    new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Count, "revenue") }));
            Assert.Empty(sheet.WorksheetPart.PivotTableParts);
        }

        [Theory]
        [InlineData(false, false, 2)]
        [InlineData(true, false, 2)]
        [InlineData(false, true, 0)]
        public void Test_AddPivotTable_RejectsInvalidValuesPositionBeforeCreatingParts(
            bool valuesOnRows, bool singleMeasure, int position) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Channel"); sheet.CellValue(1, 3, "Sales");
            sheet.CellValue(2, 1, "East"); sheet.CellValue(2, 2, "Retail"); sheet.CellValue(2, 3, 10d);
            var measures = new List<ExcelPivotDataField> {
                new("Sales", ExcelPivotDataFunction.Sum, "Metric")
            };
            if (!singleMeasure) measures.Add(new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Count, "Count"));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddPivotTable("A1:C2", "E1", "Pivot",
                rowFields: new[] { "Region" }, columnFields: new[] { "Channel" }, dataFields: measures,
                dataOnRows: valuesOnRows, options: new ExcelPivotTableOptions { ValuesAxisPosition = position }));
            Assert.Empty(sheet.WorksheetPart.PivotTableParts);
            Assert.Null(document.WorkbookPartRoot.Workbook.PivotCaches);
            Assert.Empty(sheet.GetPivotTables());
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_PivotLayoutLookup_RejectsRepeatedPrefixWithoutPreviousItem() {
            using var document = ExcelDocument.Load(PivotLayoutOraclePath);
            var sheet = document.GetSheet("BothRow");
            sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.RowItems!.Elements<RowItem>().First().RepeatedItemCount = 1;
            Assert.Throws<NotSupportedException>(() => sheet.GetPivotData("PivotBothRow", "Revenue"));
        }

        [Theory]
        [InlineData(false, "F1:O7")]
        [InlineData(true, "F1:J14")]
        public void Test_PivotLayoutMaterialization_CreatesMultipleMeasuresWithoutTemplate(bool valuesOnRows, string range) {
            string output = Path.Combine(_directoryWithFiles, valuesOnRows ? "MultiTemplateRows.xlsx" : "MultiTemplateColumns.xlsx");
            using var oracle = ExcelDocumentReader.Open(PivotLayoutOraclePath);
            var source = oracle.GetSheet("Source").ReadRange("A1:D9");
            using (var document = ExcelDocument.Create(output)) {
                var sheet = document.AddWorksheet("Sales");
                for (int row = 1; row <= 9; row++) for (int column = 1; column <= 4; column++)
                    if (source[row - 1, column - 1] != null) sheet.CellValue(row, column, source[row - 1, column - 1]);
                sheet.AddPivotTable("A1:D9", "F1", "SalesPivot", rowFields: new[] { "Region" }, columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Sum, "Revenue"),
                        new ExcelPivotDataField("Units", ExcelPivotDataFunction.Average, "AvgUnits"), new ExcelPivotDataField("Amount", ExcelPivotDataFunction.Count, "Entries") },
                    layout: ExcelPivotLayout.Tabular, dataOnRows: valuesOnRows);
                Assert.True(Assert.Single(sheet.GetPivotTables()).HasValuesAxisField);
                var result = sheet.MaterializePivotTable("SalesPivot");
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(48d, sheet.GetPivotData("SalesPivot", "Revenue").Value);
                Assert.Equal(2.375d, sheet.GetPivotData("SalesPivot", "AvgUnits").Value);
                Assert.Equal(7d, sheet.GetPivotData("SalesPivot", "Entries").Value);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(output);
            Assert.Equal(2.375d, reopened.GetSheet("Sales").GetPivotData("SalesPivot", "AvgUnits").Value);
        }

        [Fact]
        public void Test_PivotLayoutMaterialization_BudgetsMeasureVisitsBeforeMutation() {
            using var document = ExcelDocument.Load(PivotLayoutOraclePath);
            var sheet = document.GetSheet("ScalarCol");
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            var original = definition.DataFields!.Elements<DataField>().First();
            definition.DataFields.RemoveAllChildren();
            for (int index = 0; index < 256; index++) {
                var measure = (DataField)original.CloneNode(true);
                measure.Name = "Measure" + index;
                definition.DataFields.AppendChild(measure);
            }
            definition.DataFields.Count = 256;
            string before = sheet.WorksheetPart.Worksheet.OuterXml;
            var exception = Assert.Throws<InvalidOperationException>(() => sheet.MaterializePivotTable("PivotScalarCol", new ExcelMutationPlanOptions { MaximumAffectedCells = 512 }));
            Assert.Contains("measure input visits", exception.Message);
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
        }

        private static string PivotLayoutOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", "layout-conformance.xlsx");
        [Fact]
        public void Test_PivotLayoutLookup_MatchesNativeValuesAxesAndRepeatedPrefixes() {
            using var oracle = ExcelDocumentReader.Open(PivotLayoutOraclePath);
            var expected = oracle.GetSheet("Lookups").ReadRange("B1:B324");
            using var document = ExcelDocument.Load(PivotLayoutOraclePath);
            var lookup = document.GetSheet("Lookups");
            lookup.ClearCachedFormulaResults();
            Assert.Equal(324, lookup.RecalculateSupportedFormulas());
            for (int row = 1; row <= 324; row++) AssertPivotLookupOracleValue(expected[row - 1, 0], lookup.CellAt(row, 2).GetValue().Value);
        }

        [Theory]
        [InlineData("BothCol", "A1:J7", 0)]
        [InlineData("BothRow", "A1:E14", 1)]
        [InlineData("RowCol", "A1:D5", 2)]
        [InlineData("RowRow", "A1:C13", 3)]
        [InlineData("ColCol", "A1:I4", 4)]
        [InlineData("ColRow", "A1:D5", 5)]
        [InlineData("ScalarCol", "A1:C2", 6)]
        [InlineData("ScalarRow", "A1:B4", 7)]
        [InlineData("BothColOuter", "A1:J7", 8)]
        [InlineData("BothRowOuter", "A1:E14", 9)]
        [InlineData("RowRowOuter", "A1:C13", 10)]
        [InlineData("ColColOuter", "A1:I4", 11)]
        public void Test_PivotLayoutMaterialization_MatchesNativeMeasuresAndReopens(string name, string range, int profile) {
            using var oracle = ExcelDocumentReader.Open(PivotLayoutOraclePath);
            var expected = oracle.GetSheet("Lookups").ReadRange($"B{profile * 27 + 1}:B{profile * 27 + 27}");
            string output = Path.Combine(_directoryWithFiles, name + ".MultiMeasure.xlsx");
            using (var document = ExcelDocument.Load(PivotLayoutOraclePath)) {
                var result = document.GetSheet(name).MaterializePivotTable("Pivot" + name);
                Assert.Equal(range, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid, string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                var lookup = document.GetSheet("Lookups");
                lookup.ClearCachedFormulaResults();
                Assert.Equal(324, lookup.RecalculateSupportedFormulas());
                for (int index = 0; index < 27; index++)
                    AssertPivotLookupOracleValue(expected[index, 0], lookup.CellAt(profile * 27 + index + 1, 2).GetValue().Value);
                document.Save(output);
            }
            using var reopened = ExcelDocumentReader.Open(output);
            var actual = reopened.GetSheet("Lookups").ReadRange($"B{profile * 27 + 1}:B{profile * 27 + 27}");
            for (int index = 0; index < 27; index++) AssertPivotLookupOracleValue(expected[index, 0], actual[index, 0]);
        }
    }
}
