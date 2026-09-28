using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("FixedDateFilters", 15)]
        [InlineData("CalendarPeriods", 18)]
        [InlineData("RelativeDates", 17)]
        public void PivotDateOracleCorpus_ManifestCoversEveryFixture(string corpus, int expectedCases) {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "Documents", "ExcelPivotCorpus", corpus);
            using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "provenance.json")));
            JsonElement[] cases = manifest.RootElement.GetProperty("cases").EnumerateArray().ToArray();
            Assert.Equal(expectedCases, cases.Length);
            Assert.Equal(expectedCases, cases.Select(entry => entry.GetProperty("name").GetString())
                .Distinct(StringComparer.Ordinal).Count());
            string[] indexed = cases.Select(entry => entry.GetProperty("file").GetString()!)
                .OrderBy(file => file, StringComparer.Ordinal).ToArray();
            string[] present = Directory.GetFiles(directory, "*.xlsx").Select(Path.GetFileName)
                .OrderBy(file => file, StringComparer.Ordinal).ToArray()!;
            Assert.Equal(indexed, present);
        }
    }
}
