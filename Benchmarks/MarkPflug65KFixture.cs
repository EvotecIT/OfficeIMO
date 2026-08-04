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
    internal const long ExpectedCsvCharacters = 7_253_195;
    internal const ulong ExpectedCsvChecksum = 13_293_175_220_557_208_268UL;
    internal const ulong ExpectedExcelChecksum = 3_905_306_703_451_929_130UL;
    internal const string CsvFileName = "65K_Records_Data.csv";
    internal const string XlsFileName = "65K_Records_Data.xls";
    internal const string XlsxFileName = "65K_Records_Data.xlsx";
    internal const string XlsbFileName = "65K_Records_Data.xlsb";

    private const string SourceRoot =
        "https://raw.githubusercontent.com/MarkPflug/Benchmarks/" + SourceCommit +
        "/source/Benchmarks/Data/";

    private static readonly IReadOnlyDictionary<string, string> Hashes =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            [CsvFileName] = "AC959F43CF1077B71D310E6E49E3C168BA63A448F1855D45F44E734273EBA490",
            [XlsFileName] = "294102433C1527E5FCF48B6E4F1A707852633E623B8A20D184F5BAC926843DD6",
            [XlsxFileName] = "0F44D3E06454508DBD2CDBAF701B04160637162AB71471616D8ADC59D2EDD3A8",
            [XlsbFileName] = "9F03F160D32272CBE57D6023C73748D6C450783738FCD84CC552B03C00E23CC8"
        };

    internal static string Root =>
        Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA")
        ?? Path.Combine(Path.GetTempPath(), "OfficeIMO", "Benchmarks", "Fixtures", SourceCommit);

    internal static string CsvPath => Path.Combine(Root, CsvFileName);
    internal static string XlsPath => Path.Combine(Root, XlsFileName);
    internal static string XlsxPath => Path.Combine(Root, XlsxFileName);
    internal static string XlsbPath => Path.Combine(Root, XlsbFileName);

    internal static void EnsureAuthentic(string fixtureName) {
        if (!Hashes.TryGetValue(fixtureName, out string? expectedHash)) {
            throw new ArgumentException($"Unknown benchmark fixture '{fixtureName}'.", nameof(fixtureName));
        }

        Directory.CreateDirectory(Root);
        using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
        string path = Path.Combine(Root, fixtureName);
        if (!File.Exists(path) || !HashMatches(path, expectedHash)) {
            Download(client, fixtureName, path);
        }

        string actual = ComputeHash(path);
        if (!string.Equals(actual, expectedHash, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidDataException(
                $"Fixture hash mismatch for {fixtureName}: expected {expectedHash}, got {actual}.");
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
