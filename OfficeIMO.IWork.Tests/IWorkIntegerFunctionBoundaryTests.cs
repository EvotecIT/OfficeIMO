using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_integer_function_boundaries_retain_caches_without_claiming_excel_compatibility() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(Fixture("numbers-parser/integer-function-boundaries.json")));
        JsonElement evidence = manifest.RootElement;
        string path = Fixture(evidence.GetProperty("source").GetString()!);
        Assert.Equal(evidence.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(path).ReadNumbers();
        IWorkTable table = projection.Sheets.Single(s => s.Name == evidence.GetProperty("sourceSheet").GetString())
            .Tables.Single(t => t.Name == evidence.GetProperty("sourceTable").GetString());
        foreach (JsonElement expected in evidence.GetProperty("cases").EnumerateArray()) {
            IWorkTableCell cell = table.GetCell(expected.GetProperty("row").GetInt32(), expected.GetProperty("column").GetInt32())!;
            // Names alone cannot qualify expressions whose retained native result differs
            // from the destination contract. A future mapping must keep this boundary gated.
            Assert.False(cell.FormulaIsComplete);
            Assert.True(cell.CachedValueIsComplete);
            double actual = Assert.IsType<double>(cell.Value);
            double cache = expected.GetProperty("nativeCachedValue").GetDouble();
            Assert.InRange(Math.Abs(actual / cache - 1), 0, 1e-14);
        }
    }
}
