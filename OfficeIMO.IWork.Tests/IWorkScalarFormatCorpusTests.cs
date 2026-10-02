using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkScalarFormatCorpusTests {
    [Fact]
    public void Native_selected_scalar_formats_match_independent_default_and_unassessed_metadata() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "cell-format-selectors.json")));
        using var dates = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "date-formats.json")));
        var qualifiedDateDeclarations = dates.RootElement.GetProperty("packages").EnumerateArray()
            .SelectMany(p => p.GetProperty("cells").EnumerateArray())
            .Select(c => c.GetProperty("formatHex").GetString()).ToHashSet();
        foreach (var package in manifest.RootElement.GetProperty("packages").EnumerateArray()) {
            string path = Path.Combine(root, package.GetProperty("source").GetString()!);
            Assert.Equal(package.GetProperty("sourceSha256").GetString(),
                Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
            var projection = IWorkSourceDocument.Open(path, new IWorkReadOptions { PreserveSourceRecords = false }).ReadNumbers();
            var cells = projection.Sheets.SelectMany(s => s.Tables).SelectMany(t => t.Cells).ToArray();
            Assert.Equal(package.GetProperty("selectorCounts").GetProperty("15").GetInt32(),
                cells.Count(c => c.NumberFormat?.Kind == IWorkNumberFormatKind.DateTime
                    || (c.UnsupportedFeatures & IWorkCellUnsupportedFeatures.DateFormat) != 0));
            Assert.Equal(package.GetProperty("selectorCounts").GetProperty("16").GetInt32(),
                cells.Count(c => c.NumberFormat?.Kind == IWorkNumberFormatKind.Duration
                    || (c.UnsupportedFeatures & IWorkCellUnsupportedFeatures.DurationFormat) != 0));
            foreach (var expected in package.GetProperty("cases").EnumerateArray()) {
                var table = projection.Sheets.Single(s => s.Name == expected.GetProperty("sheet").GetString())
                    .Tables.Single(t => t.Name == expected.GetProperty("table").GetString());
                var cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32())!;
                var feature = expected.GetProperty("selectorBit").GetInt32() switch {
                    15 => IWorkCellUnsupportedFeatures.DateFormat, 16 => IWorkCellUnsupportedFeatures.DurationFormat,
                    17 => IWorkCellUnsupportedFeatures.TextFormat, _ => IWorkCellUnsupportedFeatures.BooleanFormat
                };
                Assert.False(cell.HasDecodeError);
                bool qualified = expected.GetProperty("isDefaultScalarFormat").GetBoolean()
                    || expected.GetProperty("selectorBit").GetInt32() == 15
                        && qualifiedDateDeclarations.Contains(expected.GetProperty("formatHex").GetString())
                    || expected.GetProperty("selectorBit").GetInt32() == 16
                        && expected.GetProperty("formatHex").GetString() == "088c0238017802800102c00200";
                Assert.Equal(qualified
                    ? IWorkCellUnsupportedFeatures.None : feature, cell.UnsupportedFeatures & feature);
                if (!qualified)
                    Assert.Contains(projection.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
            }
        }
    }
}
