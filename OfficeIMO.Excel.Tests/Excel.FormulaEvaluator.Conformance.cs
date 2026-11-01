using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void FormulaEvaluator_LookupExpansionMatchesNativeExcelCachesAndSpills() {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus");
            string input = Path.Combine(directory, "lookup-conformance-input.xlsx");
            string path = Path.Combine(_directoryWithFiles, "LookupOracle.xlsx");
            using var catalog = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "lookup-conformance.cases.json")));
            using var native = SpreadsheetDocument.Open(Path.Combine(directory, "lookup-conformance-windows.xlsx"), false);
            var nativeCells = LookupOracleReportCells(native);
            var spills = nativeCells.Values.Where(cell => cell.CellFormula?.Reference?.Value?.Contains(":") == true)
                .ToDictionary(cell => cell.CellReference!.Value!, cell => cell.CellFormula!.Reference!.Value!);
            var references = new HashSet<string>(StringComparer.Ordinal);
            using (var source = SpreadsheetDocument.Open(input, false)) {
                Assert.All(source.WorkbookPart!.WorksheetParts.SelectMany(part => part.Worksheet.Descendants<Cell>())
                    .Where(cell => cell.CellFormula != null), cell => Assert.True(string.IsNullOrEmpty(cell.CellValue?.Text),
                        $"Input formula {cell.CellReference} already has a cache."));
            }
            using (var document = ExcelDocument.Load(input)) {
                var report = document.Sheets.Single(sheet => sheet.Name == "Report");
                foreach (var item in catalog.RootElement.EnumerateArray()) {
                    int row = item.GetProperty("Row").GetInt32();
                    string formula = item.GetProperty("Formula").GetString()!;
                    string reference = A1.CellReference(row, 1);
                    Assert.True(references.Add(reference), $"Duplicate catalog case {reference}.");
                    if (spills.ContainsKey(reference)) report.SetDynamicArrayFormula(reference, formula);
                    else report.CellFormula(row, 1, formula);
                }
                Assert.Equal(208, references.Count);
                // Includes the independently authored NA() source cell, with no imported native caches.
                Assert.Equal(references.Count + 1, document.Calculate());
                Assert.All(document.InspectFormulas().Formulas, formula => Assert.True(formula.IsSupportedByOfficeIMO,
                    $"{formula.SheetName}!{formula.CellReference}: {formula.UnsupportedReason}"));
                document.Save(path);
                Assert.Empty(document.ValidateOpenXml());
            }
            Assert.Equal(12, spills.Count);
            using var saved = SpreadsheetDocument.Open(path, false);
            var actualCells = LookupOracleReportCells(saved);
            foreach (var spill in spills) {
                Assert.True(actualCells.TryGetValue(spill.Key, out Cell? anchor), $"Missing spill anchor {spill.Key}.");
                Assert.Equal(spill.Value, anchor!.CellFormula?.Reference?.Value);
                Assert.True(A1.TryParseRange(spill.Value, out int r1, out int c1, out int r2, out int c2));
                for (int row = r1; row <= r2; row++) {
                    for (int column = c1; column <= c2; column++) references.Add(A1.CellReference(row, column));
                }
            }
            Assert.Equal(232, references.Count);
            var failures = new List<string>();
            foreach (string reference in references) {
                Assert.True(nativeCells.TryGetValue(reference, out Cell? expected), $"Missing native cell {reference}.");
                if (!actualCells.TryGetValue(reference, out Cell? actual)) {
                    failures.Add($"{reference}: missing calculated cell.");
                    continue;
                }
                var expectedCache = LookupOracleCache(native, expected!);
                var actualCache = LookupOracleCache(saved, actual);
                bool equivalent = expectedCache.Kind == CellValues.Number
                    ? EquivalentFormulaCache(expectedCache.Value, actualCache.Value)
                    : string.Equals(expectedCache.Value, actualCache.Value, StringComparison.Ordinal);
                if (expectedCache.Kind != actualCache.Kind || !equivalent) {
                    failures.Add($"{reference}: Excel={expectedCache.Kind}:{expectedCache.Value}, OfficeIMO={actualCache.Kind}:{actualCache.Value}");
                }
            }
            Assert.True(failures.Count == 0, string.Join(Environment.NewLine, failures));
        }

        private static Dictionary<string, Cell> LookupOracleReportCells(SpreadsheetDocument package) {
            var workbook = package.WorkbookPart!;
            var sheet = workbook.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Report");
            return ((WorksheetPart)workbook.GetPartById(sheet.Id!)).Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
        }

        private static (CellValues Kind, string? Value) LookupOracleCache(SpreadsheetDocument package, Cell cell) {
            var kind = cell.DataType?.Value ?? CellValues.Number;
            string? value = cell.CellValue?.Text;
            if (kind == CellValues.SharedString) {
                value = package.WorkbookPart!.SharedStringTablePart!.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(int.Parse(value!, CultureInfo.InvariantCulture)).InnerText;
                kind = CellValues.String;
            } else if (kind == CellValues.InlineString) {
                value = cell.InlineString?.InnerText;
                kind = CellValues.String;
            }
            return (kind, value);
        }

        [Theory]
        [InlineData("1900")]
        [InlineData("1904")]
        public void Test_FormulaEvaluator_MatchesIndependentFunctionAndDateOracle(string dateSystem) {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", $"function-conformance-{dateSystem}.xlsx");
            string path = Path.Combine(_directoryWithFiles, $"FunctionOracle{dateSystem}.xlsx");
            File.Copy(source, path);
            object?[,] expectedValues;
            object? expectedNamedValue;
            using (var reader = ExcelDocumentReader.Open(source)) {
                expectedValues = reader.GetSheet("Expressions").ReadRange("B1:B166");
                expectedNamedValue = reader.GetSheet("Expressions").ReadRange("C1:C1")[0, 0];
            }
            using (var document = ExcelDocument.Load(path)) {
                var expected = document.InspectFormulas().Formulas.ToDictionary(cell => cell.SheetName + "!" + cell.CellReference);
                Assert.Equal(168, expected.Count);
                int calculated = document.Calculate();
                var failures = new List<string>();
                if (calculated != expected.Count) failures.Add($"Calculated {calculated} of {expected.Count} supported producer formula cells.");
                foreach (var actual in document.InspectFormulas().Formulas) {
                    string key = actual.SheetName + "!" + actual.CellReference;
                    var original = expected[key];
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
            object? actualNamedValue = reopened.GetSheet("Expressions").ReadRange("C1:C1")[0, 0];
            Assert.Equal(expectedNamedValue?.GetType(), actualNamedValue?.GetType());
            Assert.True(EquivalentFormulaCache(Convert.ToString(expectedNamedValue, CultureInfo.InvariantCulture),
                Convert.ToString(actualNamedValue, CultureInfo.InvariantCulture)), "Saved named expression differs from Excel.");
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
