using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.OpenDocument.Benchmarks.Comparisons;

internal static class OdsComparisonEvidenceRunner {
    private static readonly string[] Workloads = ["Create", "Read"];
    private static readonly string[] Implementations = ["OfficeIMO", "OpenStandardLibrary"];
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
            var measurements = new List<OdsComparisonEvidenceMeasurement>();
            foreach (string workload in Workloads) {
                foreach (OdsComparisonScale scale in OdsComparisonCorpus.Scales) {
                    foreach (string implementation in Implementations) {
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            OdsComparisonEvidenceMeasurement measurement =
                                RunChildProbe(workload, scale.Name, implementation) with { Iteration = iteration };
                            measurements.Add(measurement);
                            Console.WriteLine(
                                $"{workload,-6} {scale.Name,-6} {implementation,-19} #{iteration,-2} " +
                                $"{measurement.ElapsedMillisecondsPerOperation,9:F2} ms/op " +
                                $"{measurement.AllocatedBytesPerOperation / 1048576D,8:F2} MiB/op " +
                                $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                                $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                                $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                $"{measurement.InputBytes,8:N0}/{measurement.OutputBytes,8:N0} bytes");
                        }
                    }
                }
            }

            var report = new OdsComparisonEvidenceReport(
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

    private static OdsComparisonEvidenceMeasurement Measure(
        string workload,
        string scaleName,
        string implementation) {
        if (!Workloads.Contains(workload, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(workload));
        }
        if (!Implementations.Contains(implementation, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(implementation));
        }

        OdsComparisonScale scale = OdsComparisonCorpus.Get(scaleName);
        bool office = implementation == "OfficeIMO";
        long expected = checked((long)scale.Rows * scale.Columns * OdsComparisonCorpus.Cell(0, 0).Length);
        byte[]? fixture = workload == "Read" ? OdsComparisonWorkflows.CreateOfficeIMO(scale) : null;
        Func<object> invoke = workload == "Create"
            ? office
                ? () => OdsComparisonWorkflows.CreateOfficeIMO(scale)
                : () => OdsComparisonWorkflows.CreateOpenStandardLibrary(scale).GetAwaiter().GetResult()
            : office
                ? () => OdsComparisonWorkflows.ReadOfficeIMO(fixture!)
                : () => OdsComparisonWorkflows.ReadOpenStandardLibrary(fixture!).GetAwaiter().GetResult();
        Action<object> validate = workload == "Create"
            ? result => {
                OdsOutputEvidence evidence = OdsComparisonValidation.Inspect(implementation, scale, (byte[])result);
                if (evidence.RecordCount != checked((long)scale.Rows * scale.Columns) ||
                    evidence.ContentLength != expected) {
                    throw new InvalidDataException("Created ODS did not match the comparison corpus.");
                }
            }
            : result => {
                if ((long)result != expected) {
                    throw new InvalidDataException("ODS read checksum did not match the comparison corpus.");
                }
            };

        object warmupResult = invoke();
        validate(warmupResult);
        int operations = scale.Name == "Small" ? 4 : 2;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new OdsManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < operations; index++) result = invoke();
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        validate(result!);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(warmupResult);
        GC.KeepAlive(result);
        return new OdsComparisonEvidenceMeasurement(
            workload,
            scale.Name,
            implementation,
            1,
            fixture?.LongLength ?? 0,
            workload == "Create" ? ((byte[])result!).LongLength : 0,
            operations,
            stopwatch.Elapsed.TotalMilliseconds / operations,
            allocated / operations,
            retained,
            Math.Max(0, peakManaged - heapBefore),
            process.PeakWorkingSet64);
    }

    private static OdsComparisonEvidenceMeasurement RunChildProbe(
        string workload,
        string scale,
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
        foreach (string argument in new[] { "--evidence-probe", workload, scale, implementation }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<OdsComparisonEvidenceMeasurement>(output, JsonOptions)
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

internal sealed record OdsComparisonEvidenceMeasurement(
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

internal sealed record OdsComparisonEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<OdsComparisonEvidenceMeasurement> Measurements);

internal sealed class OdsManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(false);
    private int _stopped;

    internal OdsManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO ODS heap sampler" };
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
