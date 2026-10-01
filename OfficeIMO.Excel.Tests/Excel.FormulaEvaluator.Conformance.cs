using System.Globalization;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("1900")]
        [InlineData("1904")]
        public void Test_FormulaEvaluator_MatchesIndependentFunctionAndDateOracle(string dateSystem) {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", $"function-conformance-{dateSystem}.xlsx");
            string path = Path.Combine(_directoryWithFiles, $"FunctionOracle{dateSystem}.xlsx");
            File.Copy(source, path);
            object?[,] expectedValues;
            using (var reader = ExcelDocumentReader.Open(source)) {
                expectedValues = reader.GetSheet("Expressions").ReadRange("B1:B166");
            }
            using (var document = ExcelDocument.Load(path)) {
                var expected = document.InspectFormulas().Formulas.ToDictionary(cell => cell.SheetName + "!" + cell.CellReference);
                Assert.Equal(168, expected.Count);
                int calculated = document.Calculate();
                var failures = new List<string>();
                if (calculated != expected.Count - 1) failures.Add($"Calculated {calculated} of {expected.Count - 1} supported producer formula cells.");
                foreach (var actual in document.InspectFormulas().Formulas) {
                    string key = actual.SheetName + "!" + actual.CellReference;
                    var original = expected[key];
                    if (key == "Expressions!C1") {
                        if (actual.IsSupportedByOfficeIMO || original.CachedValue != actual.CachedValue)
                            failures.Add($"{key}: Named expression must remain unsupported with its producer cache intact.");
                        continue;
                    }
                    if (!actual.IsSupportedByOfficeIMO || !EquivalentFormulaCache(original.CachedValue, actual.CachedValue)) {
                        failures.Add($"{key}: {original.Formula}: Excel={original.CachedValue}, OfficeIMO={actual.CachedValue}, supported={actual.IsSupportedByOfficeIMO}: {actual.UnsupportedReason}");
                    }
                }
                Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
                document.Save();
                Assert.Empty(document.ValidateOpenXml());
            }
            using var reopened = ExcelDocumentReader.Open(path);
            object?[,] actualValues = reopened.GetSheet("Expressions").ReadRange("B1:B166");
            for (int row = 0; row < expectedValues.GetLength(0); row++) {
                Assert.Equal(expectedValues[row, 0]?.GetType(), actualValues[row, 0]?.GetType());
                Assert.True(EquivalentFormulaCache(Convert.ToString(expectedValues[row, 0], CultureInfo.InvariantCulture),
                    Convert.ToString(actualValues[row, 0], CultureInfo.InvariantCulture)), $"Saved row {row + 1} differs from Excel.");
            }
        }

        [Theory]
        [InlineData(60d, "1900-02-29")]
        [InlineData(0d, "1900-01-00")]
        public void Test_FormulaEvaluator_InvariantDateTextRetainsEarlyCalendar(double serial, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Dates");
            sheet.CellFormula(1, 1, $"TEXT({serial.ToString(CultureInfo.InvariantCulture)},\"yyyy-mm-dd\")");
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? actual));
            Assert.Equal(expected, actual);
        }

        private static bool EquivalentFormulaCache(string? expected, string? actual) {
            if (double.TryParse(expected, NumberStyles.Float, CultureInfo.InvariantCulture, out double left)
                && double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out double right)) {
                return Math.Abs(left - right) <= Math.Max(1e-12, Math.Abs(left) * 1e-12);
            }
            return string.Equals(expected, actual, StringComparison.Ordinal);
        }
    }
}
