using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> RelativeDatePivotFilterCases() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "RelativeDates", "provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(path));
            foreach (JsonElement entry in manifest.RootElement.GetProperty("cases").EnumerateArray())
                yield return new object[] { entry.GetProperty("name").GetString()! };
        }

        [Theory]
        [MemberData(nameof(RelativeDatePivotFilterCases))]
        public void PivotRelativeDate_MatchesExcelSavedViewAndFilterBounds(string caseName) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "RelativeDates");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            JsonElement oracleCase = manifest.RootElement.GetProperty("cases").EnumerateArray()
                .Single(entry => entry.GetProperty("name").GetString() == caseName);
            string oraclePath = Path.Combine(directory, oracleCase.GetProperty("file").GetString()!);
            using (var stream = File.OpenRead(oraclePath))
            using (var sha = SHA256.Create()) {
                string actualHash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                Assert.Equal(oracleCase.GetProperty("sha256").GetString(), actualHash);
            }
            DateTime referenceDate = DateTime.ParseExact(manifest.RootElement.GetProperty("referenceDate").GetString()!,
                "yyyy-MM-dd", CultureInfo.InvariantCulture);
            // Excel stores these as wall-clock dates; offset-bearing fixture values must not shift with the runner time zone.
            DateTime[] sourceDates = manifest.RootElement.GetProperty("sourceDates").EnumerateArray()
                .Select(value => DateTimeOffset.Parse(value.GetString()!, CultureInfo.InvariantCulture).DateTime).ToArray();
            string expectedRange = oracleCase.GetProperty("outputRange").GetString()!;
            double expectedTotal = oracleCase.GetProperty("grandTotal").GetDouble();
            string? siblingRange = oracleCase.GetProperty("siblingRange").GetString();
            double[] expectedBounds = ReadRelativeDateFilterBounds(oraclePath);
            object?[,] expected;
            object?[,]? siblingExpected = null;
            using (var excel = ExcelDocumentReader.Open(oraclePath)) {
                expected = excel.GetSheet("Grouped").ReadRange(expectedRange);
                if (siblingRange != null)
                    siblingExpected = excel.GetSheet("AllDates").ReadRange(siblingRange);
            }

            string importedOutput = Path.Combine(_directoryWithFiles, "PivotRelativeImported-" + caseName + ".xlsx");
            using (var document = ExcelDocument.Load(oraclePath)) {
                var sheet = document.GetSheet("Grouped");
                var result = sheet.MaterializePivotTable("DatePivot", referenceDate: referenceDate);
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(expectedTotal, sheet.GetPivotData("DatePivot", "Metric").Value);
                if (siblingRange != null) {
                    Assert.Equal(new[] { "DatePivot", "AllDatesPivot" }, result.AffectedPivotTables);
                    Assert.Equal(sourceDates.Length * (sourceDates.Length + 1) / 2d + 1000d,
                        document.GetSheet("AllDates").GetPivotData("AllDatesPivot", "Metric").Value);
                }
                document.Save(importedOutput);
            }
            AssertDatePivotView(expected, importedOutput, "Grouped", expectedRange);
            Assert.Equal(expectedBounds, ReadRelativeDateFilterBounds(importedOutput));
            if (siblingRange != null) {
                AssertDatePivotView(siblingExpected!, importedOutput, "AllDates", siblingRange);
                Assert.Equal(ReadPivotRowDateOrder(oraclePath, "AllDatesPivot"),
                    ReadPivotRowDateOrder(importedOutput, "AllDatesPivot"));
            }
            using (var document = ExcelDocument.Load(importedOutput)) {
                document.GetSheet("Grouped").MaterializePivotTable("DatePivot", referenceDate: referenceDate);
                document.Save();
            }
            AssertDatePivotView(expected, importedOutput, "Grouped", expectedRange);
            Assert.Equal(expectedBounds, ReadRelativeDateFilterBounds(importedOutput));
            if (siblingRange != null) {
                AssertDatePivotView(siblingExpected!, importedOutput, "AllDates", siblingRange);
                Assert.Equal(ReadPivotRowDateOrder(oraclePath, "AllDatesPivot"),
                    ReadPivotRowDateOrder(importedOutput, "AllDatesPivot"));
            }

            string authoredOutput = Path.Combine(_directoryWithFiles, "PivotRelativeAuthored-" + caseName + ".xlsx");
            string actualRange = $"L4:M{expected.GetLength(0) + 3}";
            using (var document = ExcelDocument.Create()) {
                if (oracleCase.GetProperty("date1904").GetBoolean())
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "OrderDate");
                source.CellValue(1, 2, "Sales");
                for (int index = 0; index < sourceDates.Length; index++) {
                    int row = index + 2;
                    source.CellValue(row, 1, sourceDates[index]);
                    source.CellAt(row, 1).DateTime("yyyy-mm-dd hh:mm");
                    source.CellValue(row, 2, (double)(index + 1));
                }
                source.CellValue(sourceDates.Length + 2, 2, 1000d);
                source.Pivot($"A1:B{sourceDates.Length + 2}").Rows("OrderDate").Sum("Sales", "Metric")
                    .FieldNumberFormat("OrderDate", "yyyy-mm-dd")
                    .Layout(ExcelPivotLayout.Tabular).Filter(CreateRelativeDateFilter(caseName)).At("L4", "DatePivot");
                var result = source.MaterializePivotTable("DatePivot", referenceDate: referenceDate);
                Assert.Equal(actualRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(expectedTotal, source.GetPivotData("DatePivot", "Metric").Value);
                document.Save(authoredOutput);
            }
            AssertDatePivotView(expected, authoredOutput, "Source", actualRange);
            Assert.Equal(expectedBounds, ReadRelativeDateFilterBounds(authoredOutput));
        }

        [Theory]
        [InlineData("today", "tomorrow", 1, 0, 0)]
        [InlineData("this-week", "next-week", 7, 0, 0)]
        [InlineData("this-month", "next-month", 0, 1, 0)]
        [InlineData("this-quarter", "next-quarter", 0, 3, 0)]
        [InlineData("this-year", "next-year", 0, 0, 1)]
        public void PivotRelativeDate_ReevaluatesAStaleSavedFilterAgainstANewReferenceDate(
            string original, string target, int days, int months, int years) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "RelativeDates");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            DateTime date = DateTime.ParseExact(manifest.RootElement.GetProperty("referenceDate").GetString()!,
                "yyyy-MM-dd", CultureInfo.InvariantCulture).AddDays(days).AddMonths(months).AddYears(years);
            JsonElement targetCase = manifest.RootElement.GetProperty("cases").EnumerateArray()
                .Single(entry => entry.GetProperty("name").GetString() == target);
            string expectedRange = targetCase.GetProperty("outputRange").GetString()!;
            string targetPath = Path.Combine(directory, targetCase.GetProperty("file").GetString()!);
            object?[,] expected;
            using (var oracle = ExcelDocumentReader.Open(targetPath))
                expected = oracle.GetSheet("Grouped").ReadRange(expectedRange);
            string output = Path.Combine(_directoryWithFiles, "PivotRelativeShifted-" + original + ".xlsx");
            using (var document = ExcelDocument.Load(Path.Combine(directory, original + ".xlsx"))) {
                var result = document.GetSheet("Grouped").MaterializePivotTable("DatePivot", referenceDate: date);
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                Assert.Equal(targetCase.GetProperty("grandTotal").GetDouble(),
                    document.GetSheet("Grouped").GetPivotData("DatePivot", "Metric").Value);
                document.Save(output);
            }
            AssertDatePivotView(expected, output, "Grouped", expectedRange);
            Assert.Equal(ReadRelativeDateFilterBounds(targetPath), ReadRelativeDateFilterBounds(output));
        }

        private static ExcelPivotFilter CreateRelativeDateFilter(string name) => name switch {
            "yesterday" => ExcelPivotFilter.DateYesterday("OrderDate"),
            "today" => ExcelPivotFilter.DateToday("OrderDate"),
            "tomorrow" => ExcelPivotFilter.DateTomorrow("OrderDate"),
            "last-week" => ExcelPivotFilter.DateLastWeek("OrderDate"),
            "this-week" => ExcelPivotFilter.DateThisWeek("OrderDate"),
            "next-week" => ExcelPivotFilter.DateNextWeek("OrderDate"),
            "last-month" => ExcelPivotFilter.DateLastMonth("OrderDate"),
            "this-month" or "this-month-1904" => ExcelPivotFilter.DateThisMonth("OrderDate"),
            "next-month" => ExcelPivotFilter.DateNextMonth("OrderDate"),
            "last-quarter" => ExcelPivotFilter.DateLastQuarter("OrderDate"),
            "this-quarter" => ExcelPivotFilter.DateThisQuarter("OrderDate"),
            "next-quarter" => ExcelPivotFilter.DateNextQuarter("OrderDate"),
            "last-year" => ExcelPivotFilter.DateLastYear("OrderDate"),
            "this-year" => ExcelPivotFilter.DateThisYear("OrderDate"),
            "next-year" => ExcelPivotFilter.DateNextYear("OrderDate"),
            "year-to-date" => ExcelPivotFilter.DateYearToDate("OrderDate"),
            _ => throw new ArgumentOutOfRangeException(nameof(name))
        };

        private static double[] ReadRelativeDateFilterBounds(string path) {
            using var package = SpreadsheetDocument.Open(path, false);
            var definition = package.WorkbookPart!.WorksheetParts.SelectMany(sheet => sheet.PivotTableParts)
                .Select(part => part.PivotTableDefinition)
                .Single(pivot => pivot?.Name?.Value == "DatePivot")!;
            var dynamic = definition.PivotFilters!.Elements<PivotFilter>().Single()
                .AutoFilter!.Elements<FilterColumn>().Single().GetFirstChild<DynamicFilter>()!;
            return new[] { dynamic.Val!.Value, dynamic.MaxVal!.Value };
        }
    }
}
