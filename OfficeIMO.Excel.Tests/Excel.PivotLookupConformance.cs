using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static string PivotLookupOraclePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", "aggregation-conformance.xlsx");

        [Fact]
        public void Test_PivotLookup_RecalculatesIndependentExcelLookupsAndPersistsTypedResults() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                using var oracle = ExcelDocumentReader.Open(PivotLookupOraclePath);
                using (var document = ExcelDocument.Load(PivotLookupOraclePath)) {
                    foreach (var name in new[] { "Lookups", "LookupEdges" }) {
                        var sheet = document.GetSheet(name);
                        int rows = name == "Lookups" ? 165 : 22;
                        int column = name == "Lookups" ? 2 : 1;
                        sheet.ClearCachedFormulaResults();
                        Assert.Equal(rows, sheet.RecalculateSupportedFormulas());
                        var expected = oracle.GetSheet(name).ReadRange(column == 2 ? $"B1:B{rows}" : $"A1:A{rows}");
                        for (int row = 1; row <= rows; row++) {
                            object? actual = sheet.CellAt(row, column).GetValue().Value;
                            AssertPivotLookupOracleValue(expected[row - 1, 0], actual);
                        }
                    }
                    Assert.Equal(191, document.Calculate());
                    document.Save(output);
                }
                using var reopened = ExcelDocumentReader.Open(output);
                foreach (var item in new[] { (Name: "Lookups", Range: "B1:B165"), (Name: "LookupEdges", Range: "A1:A22") }) {
                    var expected = oracle.GetSheet(item.Name).ReadRange(item.Range);
                    var actual = reopened.GetSheet(item.Name).ReadRange(item.Range);
                    for (int row = 0; row < expected.GetLength(0); row++) AssertPivotLookupOracleValue(expected[row, 0], actual[row, 0]);
                }
            } finally { if (File.Exists(output)) File.Delete(output); }
        }

        private static void AssertPivotLookupOracleValue(object? expected, object? actual) {
            Assert.Equal(expected?.GetType(), actual?.GetType());
            if (expected is double number) {
                double value = Assert.IsType<double>(actual);
                Assert.True(Math.Abs(number - value) <= Math.Max(1e-12, Math.Abs(number) * 1e-12), $"Excel={number}; OfficeIMO={value}");
            } else Assert.Equal(expected, actual);
        }

        [Fact]
        public void Test_PivotLookup_PublicApiUsesSavedViewAndTypedErrors() {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            Assert.Equal(48d, sheet.GetPivotData("PivotSum", "Metric").Value);
            Assert.Equal(48d, sheet.GetPivotData("PivotSum", "Amount").Value);
            Assert.Equal(28d, sheet.GetPivotData("PivotSum", "Metric", new Dictionary<string, object?> { ["region"] = "north" }).Value);
            Assert.Equal(0d, sheet.GetPivotData("PivotSum", "Metric", new Dictionary<string, object?> { ["Region"] = "West", ["Product"] = "A" }).Value);
            Assert.Equal("#REF!", sheet.GetPivotData("missing", "Metric").Value);
            Assert.Equal("#REF!", sheet.GetPivotData("PivotSum", "missing").Value);
            Assert.Equal("#REF!", sheet.GetPivotData("PivotSum", "Metric", new Dictionary<string, object?> { ["Product"] = 1 }).Value);
            document.GetSheet("Source").CellValue(2, 3, 999d);
            Assert.Equal(48d, sheet.GetPivotData("PivotSum", "Metric").Value);
            var error = document.GetSheet("ErrorGroups").GetPivotData("PivotErrorGroups", "Count");
            Assert.Equal(ExcelCellDataKind.Error, error.Kind);
            Assert.Equal("#N/A", error.Value);
        }

        [Fact]
        public void Test_PivotLookup_RejectsUnsupportedOrOutOfBoundsMetadataWithoutReadingOutsideView() {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            definition.Location!.FirstDataRow = uint.MaxValue;
            Assert.Equal("#REF!", sheet.GetPivotData("PivotSum", "Metric").Value);
            definition.Location.FirstDataRow = 2;
            definition.RowItems = null;
            Assert.Throws<NotSupportedException>(() => sheet.GetPivotData("PivotSum", "Metric"));
            document.GetSheet("LookupEdges").ClearCachedFormulaResults();
            Assert.True(document.GetSheet("LookupEdges").RecalculateSupportedFormulas() < 22);
            Assert.Null(document.GetSheet("LookupEdges").CellAt(2, 1).GetValue().Value);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Test_PivotLookup_MissingDataOffsetCannotSelectAHeaderAreaCell(bool rowOffset) {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            var location = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!.Location!;
            if (rowOffset) location.FirstDataRow = null;
            else location.FirstDataColumn = null;
            var actual = sheet.GetPivotData("PivotSum", "Metric");
            Assert.Equal(ExcelCellDataKind.Error, actual.Kind);
            Assert.Equal("#REF!", actual.Value);
            var lookup = document.GetSheet("LookupEdges");
            lookup.ClearCachedFormulaResults();
            lookup.RecalculateSupportedFormulas();
            Assert.Equal("#REF!", lookup.CellAt(2, 1).GetValue().Value);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_PivotLookup_MultipleMeasuresRequireOneUnambiguousValuesAxis(bool duplicateAxis) {
            using var document = ExcelDocument.Load(PivotLookupOraclePath);
            var sheet = document.GetSheet("Sum");
            var definition = sheet.WorksheetPart.PivotTableParts.Single().PivotTableDefinition!;
            var second = (DataField)definition.DataFields!.Elements<DataField>().Single().CloneNode(true);
            second.Name = "SecondMeasure";
            definition.DataFields.AppendChild(second);
            definition.DataFields.Count = 2;
            if (duplicateAxis) {
                definition.RowFields!.AppendChild(new Field { Index = -2 });
                definition.ColumnFields!.AppendChild(new Field { Index = -2 });
            }
            Assert.Throws<NotSupportedException>(() => sheet.GetPivotData("PivotSum", "SecondMeasure"));
            var lookup = document.GetSheet("LookupEdges");
            lookup.CellFormula(1, 1, "GETPIVOTDATA(\"SecondMeasure\",Sum!A1)");
            lookup.ClearCachedFormulaResults();
            lookup.RecalculateSupportedFormulas();
            Assert.Null(lookup.CellAt(1, 1).GetValue().Value);
        }
    }
}
