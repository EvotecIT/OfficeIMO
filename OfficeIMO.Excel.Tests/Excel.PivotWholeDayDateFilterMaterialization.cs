using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> WholeDayPivotFilterCases() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "WholeDayDateFilters", "provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(path));
            foreach (JsonElement entry in manifest.RootElement.GetProperty("cases").EnumerateArray())
                yield return new object[] { entry.GetProperty("name").GetString()! };
        }

        [Theory]
        [MemberData(nameof(WholeDayPivotFilterCases))]
        public void PivotWholeDayDateFilter_MatchesExcelSavedViewFromImportedAndAuthoredPivots(string caseName) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "WholeDayDateFilters");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            JsonElement oracleCase = manifest.RootElement.GetProperty("cases").EnumerateArray()
                .Single(entry => entry.GetProperty("name").GetString() == caseName);
            Assert.True(oracleCase.GetProperty("wholeDay").GetBoolean());
            string oraclePath = Path.Combine(directory, oracleCase.GetProperty("file").GetString()!);
            using (var stream = File.OpenRead(oraclePath))
            using (var sha = SHA256.Create()) {
                string hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                Assert.Equal(oracleCase.GetProperty("sha256").GetString(), hash);
            }

            string expectedRange = oracleCase.GetProperty("outputRange").GetString()!;
            double expectedTotal = oracleCase.GetProperty("grandTotal").GetDouble();
            object?[,] expected;
            using (var excel = ExcelDocumentReader.Open(oraclePath))
                expected = excel.GetSheet("Grouped").ReadRange(expectedRange);

            string importedOutput = Path.Combine(_directoryWithFiles, "PivotWholeDayImported-" + caseName + ".xlsx");
            using (var document = ExcelDocument.Load(oraclePath)) {
                var sheet = document.GetSheet("Grouped");
                Assert.True(sheet.GetPivotTables().Single().Filters.Single().WholeDay);
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
            string authoredOutput = Path.Combine(_directoryWithFiles, "PivotWholeDayAuthored-" + caseName + ".xlsx");
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
                    "equal-times" or "equal-times-1904" or "equal-blank"
                        => ExcelPivotFilter.DateEquals("OrderDate", first, wholeDay: true),
                    "not-equal-times" or "not-equal-blank"
                        => ExcelPivotFilter.DateNotEquals("OrderDate", first, wholeDay: true),
                    "before-times" => ExcelPivotFilter.DateOlderThan("OrderDate", last, wholeDay: true),
                    "before-equal-times" => ExcelPivotFilter.DateOlderThanOrEqual("OrderDate", last, wholeDay: true),
                    "after-times" => ExcelPivotFilter.DateNewerThan("OrderDate", first, wholeDay: true),
                    "after-equal-times" => ExcelPivotFilter.DateNewerThanOrEqual("OrderDate", first, wholeDay: true),
                    "between-times" => ExcelPivotFilter.DateBetween("OrderDate", first, middle, wholeDay: true),
                    "not-between-times" => ExcelPivotFilter.DateNotBetween("OrderDate", first, middle, wholeDay: true),
                    _ => throw new InvalidOperationException("Unexpected Excel whole-day oracle case.")
                };
                Assert.True(filter.WholeDay);
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
            using (var document = ExcelDocument.Load(authoredOutput)) {
                var sheet = document.GetSheet("Source");
                Assert.True(sheet.GetPivotTables().Single().Filters.Single().WholeDay);
                var result = sheet.MaterializePivotTable("DatePivot");
                Assert.Equal(actualRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
            }
        }
    }
}
