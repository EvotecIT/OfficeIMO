using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
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
        public void Test_PivotAggregate_MatchesIndependentExcelGroupsAndMergedTotals(ExcelPivotDataFunction function) {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelPivotCorpus", "aggregation-conformance.xlsx");
            using var document = ExcelDocument.Load(source);
            var sourceSheet = document.GetSheet("Source");
            var pivot = Assert.Single(document.GetSheet(function.ToString()).GetPivotTables());
            Assert.Equal(function, Assert.Single(pivot.DataFields).Function);
            Assert.Equal(new[] { "Region" }, pivot.RowFields);
            Assert.Equal(new[] { "Product" }, pivot.ColumnFields);
            using var reader = ExcelDocumentReader.Open(source);
            var oracleSheet = reader.GetSheet("Lookups");
            var functions = new[] { ExcelPivotDataFunction.Sum, ExcelPivotDataFunction.Count, ExcelPivotDataFunction.CountNumbers,
                ExcelPivotDataFunction.Average, ExcelPivotDataFunction.Minimum, ExcelPivotDataFunction.Maximum,
                ExcelPivotDataFunction.Product, ExcelPivotDataFunction.StandardDeviation, ExcelPivotDataFunction.StandardDeviationP,
                ExcelPivotDataFunction.Variance, ExcelPivotDataFunction.VarianceP };
            int firstRow = Array.IndexOf(functions, function) * 9 + 1;
            var expected = oracleSheet.ReadRange($"B{firstRow}:B{firstRow + 8}");
            var filters = new (string? Region, string? Product)[] { (null, null), ("North", null), ("South", null), ("West", null),
                (null, "A"), (null, "B"), ("North", "A"), ("South", "B") };
            for (int index = 0; index < filters.Length; index++) {
                var direct = new ExcelPivotAggregateAccumulator();
                var leaves = new Dictionary<string, ExcelPivotAggregateAccumulator>();
                for (int row = 2; row <= 9; row++) {
                    string region = (string)sourceSheet.CellAt(row, 1).GetValue().Value!;
                    string product = (string)sourceSheet.CellAt(row, 2).GetValue().Value!;
                    if ((filters[index].Region != null && filters[index].Region != region)
                        || (filters[index].Product != null && filters[index].Product != product)) continue;
                    var value = sourceSheet.CellAt(row, 3).GetValue();
                    direct.Add(value.Value, value.Kind == ExcelCellDataKind.Error);
                    string key = region + "/" + product;
                    if (!leaves.TryGetValue(key, out var leaf)) leaves.Add(key, leaf = new ExcelPivotAggregateAccumulator());
                    leaf.Add(value.Value, value.Kind == ExcelCellDataKind.Error);
                }
                var merged = new ExcelPivotAggregateAccumulator();
                foreach (var leaf in leaves.Values) merged.Merge(leaf);
                AssertPivotOracleValue(expected[index, 0], direct.GetValue(function));
                AssertPivotOracleValue(expected[index, 0], merged.GetValue(function));
            }
            Assert.Equal("#REF!", expected[8, 0]);
            var errorSource = document.GetSheet("ErrorSource");
            var errorValues = reader.GetSheet("ErrorSource").ReadRange("B2:B8");
            int errorFirstRow = 100 + Array.IndexOf(functions, function) * 6;
            var errorExpected = oracleSheet.ReadRange($"B{errorFirstRow}:B{errorFirstRow + 5}");
            var groups = new[] { "Error", "Literal", "Empty", "Single", "Boolean", "TextNumber" };
            for (int index = 0; index < groups.Length; index++) {
                var direct = new ExcelPivotAggregateAccumulator();
                var merged = new ExcelPivotAggregateAccumulator();
                for (int row = 2; row <= 8; row++) {
                    if ((string)errorSource.CellAt(row, 1).GetValue().Value! != groups[index]) continue;
                    errorSource.TryGetCellValueSnapshot(row, 2, out var snapshot);
                    bool isError = snapshot?.OpenXmlType == ExcelCellValueType.Error;
                    var value = errorSource.CellAt(row, 2).GetValue();
                    Assert.Equal(errorValues[row - 2, 0]?.GetType(), value.Value?.GetType());
                    Assert.Equal(errorValues[row - 2, 0], value.Value);
                    direct.Add(value.Value, isError);
                    var leaf = new ExcelPivotAggregateAccumulator();
                    leaf.Add(value.Value, isError);
                    merged.Merge(leaf);
                }
                AssertPivotOracleValue(errorExpected[index, 0], direct.GetValue(function));
                AssertPivotOracleValue(errorExpected[index, 0], merged.GetValue(function));
            }
        }

        private static void AssertPivotOracleValue(object? expected, ExcelCellData actual) {
            if (expected is double number) {
                Assert.Equal(ExcelCellDataKind.Number, actual.Kind);
                double value = Assert.IsType<double>(actual.Value);
                Assert.True(Math.Abs(number - value) <= Math.Max(1e-12, Math.Abs(number) * 1e-12), $"Excel={number}; OfficeIMO={value}");
            } else {
                Assert.Equal(ExcelCellDataKind.Error, actual.Kind);
                Assert.Equal(expected, actual.Value);
            }
        }
    }
}
