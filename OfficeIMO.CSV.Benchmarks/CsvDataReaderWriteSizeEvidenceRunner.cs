using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.CSV.Benchmarks;

internal static class CsvDataReaderWriteSizeEvidenceRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static int Run(string[] args) {
        try {
            string? jsonPath = GetOption(args, "--json");
            int[] rowCounts = ParseRowCounts(GetOption(args, "--rows") ?? "25000,100000");
            var measurements = new List<CsvDataReaderWriteSizeMeasurement>();
            foreach (int rowCount in rowCounts) {
                foreach (CsvBenchmarkShape shape in Enum.GetValues<CsvBenchmarkShape>()) {
                    var benchmark = new CsvDataReaderWriteBenchmarks { RowCount = rowCount, Shape = shape };
                    IReadOnlyList<CsvDataReaderWriteOutputEvidence> outputs = benchmark.CaptureValidatedOutputSizes();
                    CsvDataReaderWriteOutputEvidence office = outputs.Single(value => value.Engine == "OfficeIMO");
                    CsvDataReaderWriteOutputEvidence parallel = outputs.Single(value => value.Engine == "OfficeIMO Parallel");
                    CsvDataReaderWriteOutputEvidence sylvan = outputs.Single(value => value.Engine == "Sylvan.Data.Csv");
                    if (office.Utf8Bytes != parallel.Utf8Bytes || office.Sha256 != parallel.Sha256) {
                        throw new InvalidOperationException($"Sequential and parallel OfficeIMO output differs for {shape}/{rowCount}.");
                    }
                    measurements.Add(new CsvDataReaderWriteSizeMeasurement(
                        rowCount, shape.ToString(), outputs, office.Utf8Bytes / (double)sylvan.Utf8Bytes));
                    Console.WriteLine(
                        $"{rowCount,7:N0} {shape,-9} OfficeIMO {office.Utf8Bytes,10:N0} bytes | " +
                        $"Sylvan {sylvan.Utf8Bytes,10:N0} bytes | ratio {office.Utf8Bytes / (double)sylvan.Utf8Bytes:F3}x | parallel exact");
                }
            }

            var report = new CsvDataReaderWriteSizeReport(
                DateTimeOffset.UtcNow, ResolveCommit(), ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription, RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(), Environment.ProcessorCount, measurements);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static int[] ParseRowCounts(string value) => value.Split(',', StringSplitOptions.RemoveEmptyEntries)
        .Select(item => int.TryParse(item, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException("--rows must contain positive comma-separated integers."))
        .Distinct().ToArray();

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            using Process process = Process.Start(GitStartInfo("rev-parse", "HEAD"))!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch { return "unknown"; }
    }

    private static bool ResolveSourceTreeDirty() {
        try {
            using Process tracked = Process.Start(GitStartInfo("diff", "--quiet", "HEAD", "--"))!;
            tracked.WaitForExit();
            if (tracked.ExitCode != 0) return true;
            using Process untracked = Process.Start(GitStartInfo("ls-files", "--others", "--exclude-standard"))!;
            string output = untracked.StandardOutput.ReadToEnd();
            untracked.WaitForExit();
            return untracked.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch { return true; }
    }

    private static ProcessStartInfo GitStartInfo(params string[] arguments) {
        var info = new ProcessStartInfo("git") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        return info;
    }
}

internal sealed record CsvDataReaderWriteOutputEvidence(
    string Engine,
    int Characters,
    int Utf8Bytes,
    int CarriageReturns,
    int LineFeeds,
    string Sha256) {
    internal static CsvDataReaderWriteOutputEvidence Create(string engine, string output) {
        byte[] bytes = Encoding.UTF8.GetBytes(output);
        return new CsvDataReaderWriteOutputEvidence(
            engine,
            output.Length,
            bytes.Length,
            output.Count(character => character == '\r'),
            output.Count(character => character == '\n'),
            Convert.ToHexString(SHA256.HashData(bytes)));
    }
}

internal sealed record CsvDataReaderWriteSizeMeasurement(
    int RowCount, string Shape, IReadOnlyList<CsvDataReaderWriteOutputEvidence> Outputs, double OfficeToSylvanSizeRatio);

internal sealed record CsvDataReaderWriteSizeReport(
    DateTimeOffset MeasuredAtUtc, string SourceCommit, bool SourceTreeDirty, string Framework,
    string OperatingSystem, string Architecture, int ProcessorCount,
    IReadOnlyList<CsvDataReaderWriteSizeMeasurement> Measurements);
