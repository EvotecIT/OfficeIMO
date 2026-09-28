using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> FixedDatePivotFilterCases() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "FixedDateFilters", "provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(path));
            foreach (JsonElement entry in manifest.RootElement.GetProperty("cases").EnumerateArray())
                yield return new object[] { entry.GetProperty("name").GetString()! };
        }

        [Theory]
        [MemberData(nameof(FixedDatePivotFilterCases))]
        public void PivotFixedDateFilter_MatchesExcelSavedViewFromImportedAndAuthoredPivots(string caseName) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "FixedDateFilters");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            JsonElement oracleCase = manifest.RootElement.GetProperty("cases").EnumerateArray()
                .Single(entry => entry.GetProperty("name").GetString() == caseName);
            string oraclePath = Path.Combine(directory, oracleCase.GetProperty("file").GetString()!);
            using (var stream = File.OpenRead(oraclePath))
            using (var sha = SHA256.Create()) {
                string actualHash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                Assert.Equal(oracleCase.GetProperty("sha256").GetString(), actualHash);
            }

            string expectedRange = oracleCase.GetProperty("outputRange").GetString()!;
            double expectedTotal = oracleCase.GetProperty("grandTotal").GetDouble();
            object?[,] expected;
            using (var excel = ExcelDocumentReader.Open(oraclePath))
                expected = excel.GetSheet("Grouped").ReadRange(expectedRange);

            string importedOutput = Path.Combine(_directoryWithFiles, "PivotDateImported-" + caseName + ".xlsx");
            using (var document = ExcelDocument.Load(oraclePath)) {
                var sheet = document.GetSheet("Grouped");
                var result = sheet.MaterializePivotTable("DatePivot");
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(expectedTotal, sheet.GetPivotData("DatePivot", "Metric").Value);
                document.Save(importedOutput);
            }
            AssertDatePivotView(expected, importedOutput, "Grouped", expectedRange);

            DateTime first = new(2025, 1, 15);
            DateTime middle = new(2025, 3, 20);
            DateTime last = new(2026, 1, 15);
            bool timed = oracleCase.GetProperty("timedSource").GetBoolean();
            bool blank = oracleCase.GetProperty("blankSource").GetBoolean();
            bool date1904 = oracleCase.GetProperty("date1904").GetBoolean();
            DateTime[] keys = { first, timed ? first.AddHours(8).AddMinutes(30) : first,
                timed ? middle.AddHours(12) : middle, last };
            string authoredOutput = Path.Combine(_directoryWithFiles, "PivotDateAuthored-" + caseName + ".xlsx");
            string actualRange = $"L4:M{expected.GetLength(0) + 3}";
            using (var document = ExcelDocument.Create()) {
                if (date1904) document.DateSystem = ExcelDateSystem.NineteenFour;
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "OrderDate");
                source.CellValue(1, 2, "Sales");
                double[] sales = { 10d, 5d, 20d, 30d };
                for (int index = 0; index < keys.Length; index++) {
                    int row = index + 2;
                    if (!blank || index != 1) {
                        source.CellValue(row, 1, keys[index]);
                        source.CellAt(row, 1).DateTime("yyyy-mm-dd");
                    }
                    source.CellValue(row, 2, sales[index]);
                }
                ExcelPivotFilter filter = caseName switch {
                    "equal" or "equal-times" or "equal-1904" or "equal-blank"
                        => ExcelPivotFilter.DateEquals("OrderDate", first),
                    "not-equal" or "not-equal-blank" => ExcelPivotFilter.DateNotEquals("OrderDate", first),
                    "before" => ExcelPivotFilter.DateOlderThan("OrderDate", last),
                    "before-equal" => ExcelPivotFilter.DateOlderThanOrEqual("OrderDate", last),
                    "after" or "after-times" => ExcelPivotFilter.DateNewerThan("OrderDate", first),
                    "after-equal" => ExcelPivotFilter.DateNewerThanOrEqual("OrderDate", first),
                    "between" or "between-times" or "between-1904"
                        => ExcelPivotFilter.DateBetween("OrderDate", first, middle),
                    "not-between" => ExcelPivotFilter.DateNotBetween("OrderDate", first, middle),
                    _ => throw new InvalidOperationException("Unexpected Excel date-filter oracle case.")
                };
                source.Pivot("A1:B5").Rows("OrderDate").Sum("Sales", "Metric")
                    .FieldNumberFormat("OrderDate", "yyyy-mm-dd")
                    .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("L4", "DatePivot");
                var result = source.MaterializePivotTable("DatePivot");
                Assert.Equal(actualRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(expectedTotal, source.GetPivotData("DatePivot", "Metric").Value);
                document.Save(authoredOutput);
            }
            AssertDatePivotView(expected, authoredOutput, "Source", actualRange);
        }

        private static void AssertDatePivotView(object?[,] expected, string path, string sheet, string range) {
            using var document = ExcelDocumentReader.Open(path);
            object?[,] actual = document.GetSheet(sheet).ReadRange(range);
            Assert.Equal(expected.GetLength(0), actual.GetLength(0));
            Assert.Equal(expected.GetLength(1), actual.GetLength(1));
            for (int row = 0; row < expected.GetLength(0); row++)
                for (int column = 0; column < expected.GetLength(1); column++)
                    AssertPivotLookupOracleValue(expected[row, column], actual[row, column]);
        }

        [Fact]
        public void PivotFixedDateFilter_RejectsNumericFieldWithoutChangingSavedView() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Source");
            sheet.CellValue(1, 1, "Item");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, 1d);
            sheet.CellValue(2, 2, 10d);
            sheet.Pivot("A1:B2").Rows("Item").Sum("Sales", "Metric")
                .Filter(ExcelPivotFilter.DateEquals("Item", new DateTime(2025, 1, 15)))
                .At("D4", "DatePivot");
            string before = sheet.WorksheetPart.Worksheet.OuterXml;
            Assert.Throws<NotSupportedException>(() => sheet.MaterializePivotTable("DatePivot"));
            Assert.Equal(before, sheet.WorksheetPart.Worksheet.OuterXml);
        }
    }
}
