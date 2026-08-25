using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Word.Benchmarks;

internal static class WordOpenXmlEvidenceRunner {
    private static readonly string[] Workloads = ["CreateParagraph", "CreateReport", "Read", "Replace"];
    private static readonly string[] Implementations = ["OfficeIMO", "OpenXmlSdk"];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 3 || !int.TryParse(args[1], out int itemCount)) return 2;
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], itemCount, args[2]), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int Run(string[] args) {
        try {
            int repeat = GetPositiveIntOption(args, "--repeat", 1);
            string? jsonPath = GetOption(args, "--json");
            var measurements = new List<WordOpenXmlEvidenceMeasurement>();
            foreach (string workload in Workloads) {
                foreach (int itemCount in new[] { 100, 1000 }) {
                    foreach (string implementation in Implementations) {
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            WordOpenXmlEvidenceMeasurement measurement =
                                RunChildProbe(workload, itemCount, implementation) with { Iteration = iteration };
                            measurements.Add(measurement);
                            Console.WriteLine(
                                $"{workload,-15} {itemCount,4} {implementation,-10} #{iteration,-2} " +
                                $"{measurement.ElapsedMillisecondsPerOperation,9:F2} ms/op " +
                                $"{measurement.AllocatedBytesPerOperation / 1048576D,9:F2} MiB/op " +
                                $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                                $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                                $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                $"{measurement.InputBytes,9:N0}/{measurement.OutputBytes,9:N0} bytes");
                        }
                    }
                }
            }

            var report = new WordOpenXmlEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                measurements);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath);
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static WordOpenXmlEvidenceMeasurement Measure(
        string workload,
        int itemCount,
        string implementation) {
        if (!Workloads.Contains(workload, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(workload));
        }
        if (itemCount is not (100 or 1000)) throw new ArgumentOutOfRangeException(nameof(itemCount));
        if (!Implementations.Contains(implementation, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(implementation));
        }

        WordOpenXmlEvidenceOperation operation = CreateOperation(workload, itemCount, implementation);
        object warmupResult = operation.Invoke();
        operation.Validate(warmupResult);
        int operations = itemCount == 100 ? 8 : 2;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new WordManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < operations; index++) result = operation.Invoke();
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        operation.Validate(result!);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(warmupResult);
        GC.KeepAlive(result);
        return new WordOpenXmlEvidenceMeasurement(
            workload,
            itemCount,
            implementation,
            1,
            operation.InputBytes,
            operation.GetOutputBytes(result!),
            operations,
            stopwatch.Elapsed.TotalMilliseconds / operations,
            allocated / operations,
            retained,
            Math.Max(0, peakManaged - heapBefore),
            process.PeakWorkingSet64);
    }

    private static WordOpenXmlEvidenceOperation CreateOperation(
        string workload,
        int itemCount,
        string implementation) {
        bool office = implementation == "OfficeIMO";
        switch (workload) {
            case "CreateParagraph": {
                var benchmark = new WordCreateParagraphComparisonBenchmarks { ItemCount = itemCount };
                return new(
                    () => office ? benchmark.OfficeIMO() : WordOpenXmlEvidenceWorkloads.CreateParagraphs(itemCount),
                    result => WordBenchmarkCorpus.ValidateParagraphDocument((byte[])result, itemCount),
                    _ => ((byte[])_).Length,
                    0);
            }
            case "CreateReport": {
                var benchmark = new WordCreateReportComparisonBenchmarks { RowCount = itemCount };
                return new(
                    () => office ? benchmark.OfficeIMO() : WordOpenXmlEvidenceWorkloads.CreateReport(itemCount),
                    result => WordBenchmarkCorpus.ValidateReportDocument(
                        (byte[])result, itemCount, requireOfficeCompatibleDefaults: true),
                    result => ((byte[])result).Length,
                    0);
            }
            case "Read": {
                var benchmark = new WordReadComparisonBenchmarks { ItemCount = itemCount };
                benchmark.SetupOfficeAndOpenXml();
                WordReadObservation expected = WordBenchmarkCorpus.ObserveExpectedParagraphs(itemCount);
                return new(
                    () => office ? benchmark.OfficeIMO() : benchmark.OpenXmlSdk(),
                    result => {
                        if ((WordReadObservation)result != expected) {
                            throw new InvalidDataException("Read evidence observation did not match the corpus.");
                        }
                    },
                    _ => 0,
                    benchmark.InputBytes);
            }
            case "Replace": {
                var benchmark = new WordRichReplaceEvidenceWorkload(itemCount);
                int expectedStyles = WordBenchmarkCorpus.CountStyleDefinitions(benchmark.Fixture);
                return new(
                    () => office ? benchmark.OfficeIMO() : benchmark.OpenXmlSdk(),
                    result => {
                        byte[] payload = (byte[])result;
                        WordBenchmarkCorpus.ValidateReplacedDocument(payload, itemCount);
                        if (WordBenchmarkCorpus.CountStyleDefinitions(payload) != expectedStyles) {
                            throw new InvalidDataException("Replacement changed the rich input style catalog.");
                        }
                    },
                    result => ((byte[])result).Length,
                    benchmark.InputBytes);
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(workload));
        }
    }

    private static WordOpenXmlEvidenceMeasurement RunChildProbe(
        string workload,
        int itemCount,
        string implementation) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("No process path.");
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        foreach (string argument in new[] {
                     "--evidence-probe", workload, itemCount.ToString(System.Globalization.CultureInfo.InvariantCulture), implementation
                 }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<WordOpenXmlEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException("Probe returned no measurement.");
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static int GetPositiveIntOption(string[] args, string name, int fallback) =>
        int.TryParse(GetOption(args, name), out int value) && value > 0 ? value : fallback;

    private static string ResolveCommit() => RunGit("rev-parse", "HEAD") ?? "unknown";
    private static bool ResolveSourceTreeDirty() => !string.IsNullOrWhiteSpace(RunGit("status", "--porcelain"));

    private static string? RunGit(params string[] arguments) {
        var startInfo = new ProcessStartInfo {
            FileName = "git",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        using Process? process = Process.Start(startInfo);
        if (process == null) return null;
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode == 0 ? output.Trim() : null;
    }
}

internal sealed record WordOpenXmlEvidenceOperation(
    Func<object> Invoke,
    Action<object> Validate,
    Func<object, int> GetOutputBytes,
    int InputBytes);

internal sealed record WordOpenXmlEvidenceMeasurement(
    string Workload,
    int ItemCount,
    string Implementation,
    int Iteration,
    int InputBytes,
    int OutputBytes,
    int Operations,
    double ElapsedMillisecondsPerOperation,
    long AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record WordOpenXmlEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<WordOpenXmlEvidenceMeasurement> Measurements);

internal sealed class WordManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(false);
    private int _stopped;

    internal WordManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Word heap sampler" };
        _thread.Start();
    }

    internal long Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            Record();
        }
        return Interlocked.Read(ref _peakBytes);
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void Sample() {
        while (!_stop.Wait(1)) Record();
    }

    private void Record() {
        long observed = GC.GetTotalMemory(false);
        long current = Interlocked.Read(ref _peakBytes);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref _peakBytes, observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}
