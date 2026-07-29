using System.Security.Cryptography;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Hash-pinned copy of the public 65K sales fixtures used by Mark Pflug's benchmark repository.
/// Benchmark projects link this source file; it is not a runtime package.
/// </summary>
internal static class MarkPflug65KFixture {
    internal const string SourceCommit = "5e1113a1195bed985c10788a6b89caf551663bb1";
    internal const int ExpectedRows = 65_535;
    internal const int ExpectedColumns = 14;

    private const string SourceRoot =
        "https://raw.githubusercontent.com/MarkPflug/Benchmarks/" + SourceCommit +
        "/source/Benchmarks/Data/";

    private static readonly IReadOnlyDictionary<string, string> Hashes =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["65K_Records_Data.csv"] = "AC959F43CF1077B71D310E6E49E3C168BA63A448F1855D45F44E734273EBA490",
            ["65K_Records_Data.xlsx"] = "0F44D3E06454508DBD2CDBAF701B04160637162AB71471616D8ADC59D2EDD3A8",
            ["65K_Records_Data.xlsb"] = "9F03F160D32272CBE57D6023C73748D6C450783738FCD84CC552B03C00E23CC8"
        };

    internal static string Root =>
        Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA")
        ?? Path.Combine(Path.GetTempPath(), "OfficeIMO", "Benchmarks", "Fixtures", SourceCommit);

    internal static string CsvPath => Path.Combine(Root, "65K_Records_Data.csv");
    internal static string XlsxPath => Path.Combine(Root, "65K_Records_Data.xlsx");
    internal static string XlsbPath => Path.Combine(Root, "65K_Records_Data.xlsb");

    internal static void EnsureAuthentic() {
        Directory.CreateDirectory(Root);
        using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        foreach (KeyValuePair<string, string> fixture in Hashes) {
            string path = Path.Combine(Root, fixture.Key);
            if (!File.Exists(path) || !HashMatches(path, fixture.Value)) {
                Download(client, fixture.Key, path);
            }

            string actual = ComputeHash(path);
            if (!string.Equals(actual, fixture.Value, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException(
                    $"Fixture hash mismatch for {fixture.Key}: expected {fixture.Value}, got {actual}.");
            }
        }
    }

    internal static IReadOnlyDictionary<string, string> GetHashes() => Hashes;

    private static void Download(HttpClient client, string name, string path) {
        string temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".download";
        try {
            using HttpResponseMessage response = client.GetAsync(
                SourceRoot + name,
                HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
            using Stream input = response.Content.ReadAsStream();
            using (var output = new FileStream(
                       temporaryPath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.None,
                       81920,
                       FileOptions.SequentialScan)) {
                input.CopyTo(output);
            }

            File.Move(temporaryPath, path, overwrite: true);
        } finally {
            if (File.Exists(temporaryPath)) {
                File.Delete(temporaryPath);
            }
        }
    }

    private static bool HashMatches(string path, string expected) =>
        string.Equals(ComputeHash(path), expected, StringComparison.OrdinalIgnoreCase);

    private static string ComputeHash(string path) {
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            81920,
            FileOptions.SequentialScan);
        using SHA256 sha = SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(stream));
    }
}
