using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> CalendarPeriodPivotFilterCases() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "CalendarPeriods", "provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(path));
            foreach (JsonElement entry in manifest.RootElement.GetProperty("cases").EnumerateArray())
                yield return new object[] { entry.GetProperty("name").GetString()! };
        }

        [Theory]
        [MemberData(nameof(CalendarPeriodPivotFilterCases))]
        public void PivotCalendarPeriod_MatchesExcelSavedViewFromImportedAndAuthoredPivots(string caseName) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "CalendarPeriods");
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
            string? siblingRange = oracleCase.GetProperty("siblingRange").GetString();
            object?[,] expected;
            object?[,]? siblingExpected = null;
            using (var excel = ExcelDocumentReader.Open(oraclePath)) {
                expected = excel.GetSheet("Grouped").ReadRange(expectedRange);
                if (siblingRange != null)
                    siblingExpected = excel.GetSheet("AllDates").ReadRange(siblingRange);
            }

            string importedOutput = Path.Combine(_directoryWithFiles, "PivotCalendarImported-" + caseName + ".xlsx");
            using (var document = ExcelDocument.Load(oraclePath)) {
                var sheet = document.GetSheet("Grouped");
                var result = sheet.MaterializePivotTable("DatePivot");
                Assert.Equal(expectedRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid);
                if (siblingRange != null) {
                    Assert.Equal(new[] { "DatePivot", "AllDatesPivot" }, result.AffectedPivotTables);
                    Assert.Equal(1400d, document.GetSheet("AllDates").GetPivotData("AllDatesPivot", "Metric").Value);
                }
                Assert.Equal(expectedTotal, sheet.GetPivotData("DatePivot", "Metric").Value);
                document.Save(importedOutput);
            }
            AssertDatePivotView(expected, importedOutput, "Grouped", expectedRange);
            if (siblingRange != null)
                AssertDatePivotView(siblingExpected!, importedOutput, "AllDates", siblingRange);
            if (siblingRange != null)
                Assert.Equal(ReadPivotRowDateOrder(oraclePath, "AllDatesPivot"),
                    ReadPivotRowDateOrder(importedOutput, "AllDatesPivot"));
            using (var document = ExcelDocument.Load(importedOutput)) {
                document.GetSheet("Grouped").MaterializePivotTable("DatePivot");
                document.Save();
            }
            AssertDatePivotView(expected, importedOutput, "Grouped", expectedRange);
            if (siblingRange != null)
                AssertDatePivotView(siblingExpected!, importedOutput, "AllDates", siblingRange);
            if (siblingRange != null)
                Assert.Equal(ReadPivotRowDateOrder(oraclePath, "AllDatesPivot"),
                    ReadPivotRowDateOrder(importedOutput, "AllDatesPivot"));

            string authoredOutput = Path.Combine(_directoryWithFiles, "PivotCalendarAuthored-" + caseName + ".xlsx");
            string actualRange = $"L4:M{expected.GetLength(0) + 3}";
            using (var document = ExcelDocument.Create()) {
                if (oracleCase.GetProperty("date1904").GetBoolean())
                    document.DateSystem = ExcelDateSystem.NineteenFour;
                var source = document.AddWorksheet("Source");
                source.CellValue(1, 1, "OrderDate");
                source.CellValue(1, 2, "Sales");
                for (int index = 0; index < 24; index++) {
                    int row = index + 2;
                    source.CellValue(row, 1, new DateTime(index < 12 ? 2024 : 2025, index % 12 + 1, 1));
                    source.CellAt(row, 1).DateTime("yyyy-mm-dd");
                    source.CellValue(row, 2, (double)(index + 1));
                }
                source.CellValue(26, 1, new DateTime(2025, 1, 15, 8, 30, 0));
                source.CellAt(26, 1).DateTime("yyyy-mm-dd hh:mm");
                source.CellValue(26, 2, 100d);
                source.CellValue(27, 2, 1000d);
                JsonElement month = oracleCase.GetProperty("month");
                ExcelPivotFilter filter = month.ValueKind == JsonValueKind.Number
                    ? ExcelPivotFilter.DateMonth("OrderDate", month.GetInt32())
                    : ExcelPivotFilter.DateQuarter("OrderDate", oracleCase.GetProperty("quarter").GetInt32());
                source.Pivot("A1:B27").Rows("OrderDate").Sum("Sales", "Metric")
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
                document.GetSheet("Source").MaterializePivotTable("DatePivot");
                document.Save();
            }
            AssertDatePivotView(expected, authoredOutput, "Source", actualRange);
        }

        private static long?[] ReadPivotRowDateOrder(string path, string pivotName) {
            using var package = SpreadsheetDocument.Open(path, false);
            var part = package.WorkbookPart!.WorksheetParts.SelectMany(sheet => sheet.PivotTableParts)
                .Single(pivot => pivot.PivotTableDefinition?.Name?.Value == pivotName);
            var definition = part.PivotTableDefinition!;
            var fieldItems = definition.PivotFields!.Elements<PivotField>().First().Items!
                .Elements<Item>().ToArray();
            var sharedItems = part.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!
                .Elements<CacheField>().First().SharedItems!.ChildElements;
            return definition.RowItems!.Elements<RowItem>()
                .Where(item => item.ItemType?.Value != ItemValues.Grand)
                .Select(item => sharedItems[(int)fieldItems[(int)(item.Elements<MemberPropertyIndex>()
                    .First().Val?.Value ?? 0)].Index!.Value] is DateTimeItem date
                    ? (long?)date.Val!.Value.Ticks : null).ToArray();
        }
    }
}
