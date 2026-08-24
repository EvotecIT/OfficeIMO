using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeIMO.Adf.Benchmarks;

internal static class AdfEvidenceRunner {
    private static readonly string[] Workloads = ["Parse", "RoundTrip"];
    private static readonly string[] Implementations = ["System.Text.Json", "OfficeIMO"];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 3) return 2;
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], args[1], args[2]), JsonOptions));
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
            var measurements = new List<AdfEvidenceMeasurement>();
            foreach (string workload in Workloads) {
                foreach (AdfBenchmarkScale scale in AdfBenchmarkCorpus.Scales) {
                    foreach (string implementation in Implementations) {
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            AdfEvidenceMeasurement measurement =
                                RunChildProbe(workload, scale.Name, implementation) with { Iteration = iteration };
                            measurements.Add(measurement);
                            Console.WriteLine(
                                $"{workload,-9} {scale.Name,-6} {implementation,-16} #{iteration,-2} " +
                                $"{measurement.ElapsedMillisecondsPerOperation,9:F3} ms/op " +
                                $"{measurement.AllocatedBytesPerOperation / 1048576D,8:F2} MiB/op " +
                                $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                                $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                                $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                $"{measurement.InputBytes,8:N0}/{measurement.OutputBytes,8:N0} bytes");
                        }
                    }
                }
            }

            var report = new AdfEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                "System.Text.Json is a narrower JSON-tree cost floor; OfficeIMO additionally creates a typed ADF model and validates ADF structure.",
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

    private static AdfEvidenceMeasurement Measure(string workload, string scaleName, string implementation) {
        if (!Workloads.Contains(workload, StringComparer.Ordinal)) throw new ArgumentOutOfRangeException(nameof(workload));
        if (!Implementations.Contains(implementation, StringComparer.Ordinal)) throw new ArgumentOutOfRangeException(nameof(implementation));

        AdfBenchmarkScale scale = AdfBenchmarkCorpus.Get(scaleName);
        string json = AdfBenchmarkCorpus.Create(scale);
        bool office = implementation == "OfficeIMO";
        Func<object> invoke = workload == "Parse"
            ? office
                ? () => AdfComparisonWorkflows.ParseOfficeIMO(json)
                : () => AdfComparisonWorkflows.ParsePlatform(json)
            : office
                ? () => AdfComparisonWorkflows.RoundTripOfficeIMO(json)
                : () => AdfComparisonWorkflows.RoundTripPlatform(json);
        Action<object> validate = workload == "Parse"
            ? office
                ? result => AdfComparisonValidation.ValidateOfficeParse(json, (AdfDocument)result)
                : result => AdfComparisonValidation.ValidatePlatformParse(json, (JsonNode)result)
            : result => AdfComparisonValidation.Inspect(json, (string)result, implementation);

        object? warmup = invoke();
        validate(warmup);
        warmup = null;
        int operations = scale.Name == "Small" ? 12 : 2;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new AdfManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < operations; index++) result = invoke();
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        validate(result!);

        long outputBytes = result is string output ? System.Text.Encoding.UTF8.GetByteCount(output) : 0;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        return new AdfEvidenceMeasurement(
            workload,
            scale.Name,
            implementation,
            1,
            System.Text.Encoding.UTF8.GetByteCount(json),
            outputBytes,
            operations,
            stopwatch.Elapsed.TotalMilliseconds / operations,
            allocated / operations,
            retained,
            Math.Max(0, peakManaged - heapBefore),
            process.PeakWorkingSet64);
    }

    private static AdfEvidenceMeasurement RunChildProbe(string workload, string scale, string implementation) {
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
        foreach (string argument in new[] { "--evidence-probe", workload, scale, implementation }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<AdfEvidenceMeasurement>(output, JsonOptions)
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

internal sealed record AdfEvidenceMeasurement(
    string Workload,
    string Scale,
    string Implementation,
    int Iteration,
    long InputBytes,
    long OutputBytes,
    int Operations,
    double ElapsedMillisecondsPerOperation,
    long AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record AdfEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    string ComparisonBoundary,
    IReadOnlyList<AdfEvidenceMeasurement> Measurements);

internal sealed class AdfManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(false);
    private int _stopped;

    internal AdfManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO ADF heap sampler" };
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
