using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> AllErrorRankingCases() {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "AllErrorRanking", "provenance.json");
            using var manifest = JsonDocument.Parse(File.ReadAllText(path));
            foreach (JsonElement entry in manifest.RootElement.GetProperty("cases").EnumerateArray())
                yield return new object[] { entry.GetProperty("name").GetString()! };
        }

        [Theory]
        [MemberData(nameof(AllErrorRankingCases))]
        public void PivotAllErrorRanking_MatchesExcelProducedViewAfterPublicMaterializationAndReopen(string caseName) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", "AllErrorRanking");
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            JsonElement oracleCase = manifest.RootElement.GetProperty("cases").EnumerateArray()
                .Single(entry => entry.GetProperty("name").GetString() == caseName);
            string file = oracleCase.GetProperty("file").GetString()!;
            string sourcePath = Path.Combine(directory, file);
            using (var stream = File.OpenRead(sourcePath))
            using (var sha = SHA256.Create()) {
                string actualHash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                Assert.Equal(oracleCase.GetProperty("sha256").GetString(), actualHash);
            }

            string sourceRange = oracleCase.GetProperty("sourceRange").GetString()!;
            string expectedRange = oracleCase.GetProperty("outputRange").GetString()!;
            object?[,] input, expected;
            using (var excel = ExcelDocumentReader.Open(sourcePath)) {
                input = excel.GetSheet("Source").ReadRange(sourceRange);
                expected = excel.GetSheet("Grouped").ReadRange(expectedRange);
            }

            int filterType = oracleCase.GetProperty("filterType").GetInt32();
            double threshold = oracleCase.GetProperty("value").GetDouble();
            int aggregateFunction = oracleCase.GetProperty("aggregateFunction").GetInt32();
            ExcelPivotFilter filter = filterType switch {
                1 => ExcelPivotFilter.TopCount("Region", "Metric", (int)threshold),
                2 => ExcelPivotFilter.BottomCount("Region", "Metric", (int)threshold),
                3 => ExcelPivotFilter.TopPercent("Region", "Metric", (int)threshold),
                4 => ExcelPivotFilter.BottomPercent("Region", "Metric", (int)threshold),
                5 => ExcelPivotFilter.TopSum("Region", "Metric", threshold),
                6 => ExcelPivotFilter.BottomSum("Region", "Metric", threshold),
                _ => throw new InvalidOperationException("Unexpected Excel pivot filter type.")
            };
            string output = Path.Combine(_directoryWithFiles, "PivotAllError-" + caseName + ".xlsx");
            string actualRange = $"L4:M{expected.GetLength(0) + 3}";
            using (var document = ExcelDocument.Create(output)) {
                var source = document.AddWorksheet("Source");
                for (int row = 0; row < input.GetLength(0); row++) {
                    source.CellValue(row + 1, 1, Assert.IsType<string>(input[row, 0]));
                    if (row == 0) source.CellValue(1, 2, Assert.IsType<string>(input[row, 1]));
                    else source.CellError(row + 1, 2, Assert.IsType<string>(input[row, 1]));
                }
                var pivot = source.Pivot(sourceRange).Rows("Region");
                if (aggregateFunction == -4106) pivot.Average("Sales", "Metric");
                else if (aggregateFunction == -4157) pivot.Sum("Sales", "Metric");
                else throw new InvalidOperationException("Unexpected Excel pivot aggregate function.");
                pivot.Layout(ExcelPivotLayout.Tabular).Filter(filter).At("L4", "AllErrorPivot");
                var result = source.MaterializePivotTable("AllErrorPivot");
                Assert.Equal(actualRange, result.OutputRange);
                Assert.True(result.Mutation.PackageIsValid,
                    string.Join(Environment.NewLine, result.Mutation.Diagnostics.Select(d => d.Message)));
                Assert.Equal(expected[expected.GetLength(0) - 1, 1],
                    source.GetPivotData("AllErrorPivot", "Metric").Value);
                document.Save(new ExcelSaveOptions { ValidateOpenXml = true });
            }

            using (var reopened = ExcelDocument.Load(output)) {
                Assert.Equal(expected[expected.GetLength(0) - 1, 1],
                    reopened.GetSheet("Source").GetPivotData("AllErrorPivot", "Metric").Value);
            }
            using var actualDocument = ExcelDocumentReader.Open(output);
            object?[,] actual = actualDocument.GetSheet("Source").ReadRange(actualRange);
            for (int row = 0; row < expected.GetLength(0); row++) {
                for (int column = 0; column < expected.GetLength(1); column++) {
                    Assert.True(Equals(expected[row, column], actual[row, column]),
                        $"{caseName} {(char)('L' + column)}{row + 4}: Excel={expected[row, column] ?? "<blank>"}; OfficeIMO={actual[row, column] ?? "<blank>"}");
                }
            }
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PivotAllErrorRanking_LeavesUnqualifiedPercentAndSumDeferred(bool sum) {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source");
            source.CellValue(1, 1, "Region");
            source.CellValue(1, 2, "Sales");
            source.CellValue(2, 1, "Alpha");
            source.CellError(2, 2, "#N/A");
            source.CellValue(3, 1, "Bravo");
            source.CellError(3, 2, "#DIV/0!");
            ExcelPivotFilter filter = sum
                ? ExcelPivotFilter.TopSum("Region", "Metric", 1)
                : ExcelPivotFilter.TopPercent("Region", "Metric", 40);
            source.Pivot("A1:B3").Rows("Region").Sum("Sales", "Metric")
                .Layout(ExcelPivotLayout.Tabular).Filter(filter).At("L4", "FilteredPivot");

            Assert.Throws<NotSupportedException>(() => source.MaterializePivotTable("FilteredPivot"));
        }
    }
}
