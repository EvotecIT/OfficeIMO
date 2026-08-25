using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Provenance.Benchmarks;

internal static class ProvenanceEvidenceRunner {
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
            var measurements = new List<ProvenanceEvidenceMeasurement>();
            foreach (string format in ProvenanceBenchmarkCorpus.Formats) {
                foreach (string scale in ProvenanceBenchmarkCorpus.Scales) {
                    foreach (string operation in new[] { "Inspect", "Remove" }) {
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            ProvenanceEvidenceMeasurement measurement = RunChildProbe(format, scale, operation)
                                with { Iteration = iteration };
                            measurements.Add(measurement);
                            string output = measurement.OutputBytes is int outputBytes
                                ? $"{outputBytes,10:N0} output bytes"
                                : "       n/a output bytes";
                            Console.WriteLine(
                                $"{format,-5} {scale,-5} {operation,-7} #{iteration,-2} " +
                                $"{measurement.ElapsedMicrosecondsPerOperation,10:F2} us/op " +
                                $"{measurement.AllocatedBytesPerOperation / 1024D,10:F2} KiB/op " +
                                $"{measurement.RetainedManagedHeapGrowthBytes / 1024D,9:F2} KiB retained " +
                                $"{measurement.PeakManagedHeapGrowthBytes / 1024D,10:F2} KiB managed peak " +
                                $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                $"{measurement.InputBytes,10:N0} input bytes {output}");
                        }
                    }
                }
            }

            var report = new ProvenanceEvidenceReport(
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

    private static ProvenanceEvidenceMeasurement Measure(string format, string scale, string operation) {
        if (!ProvenanceBenchmarkCorpus.Formats.Contains(format, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(format));
        }
        if (!ProvenanceBenchmarkCorpus.Scales.Contains(scale, StringComparer.Ordinal)) {
            throw new ArgumentOutOfRangeException(nameof(scale));
        }
        if (operation is not ("Inspect" or "Remove")) throw new ArgumentOutOfRangeException(nameof(operation));

        ProvenanceBenchmarkFixture fixture = ProvenanceBenchmarkCorpus.Create(format, scale);
        ProvenanceBenchmarkValidation.Validate(fixture);
        int iterations = scale == "Small" ? 64 : 8;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new ProvenanceManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < iterations; index++) {
            result = operation == "Inspect"
                ? ProvenanceBenchmarkValidation.Inspect(fixture)
                : ProvenanceBenchmarkValidation.Remove(fixture);
        }
        stopwatch.Stop();
        long peakManagedHeap = sampler.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        int? outputBytes = ValidateResult(fixture, operation, result);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        return new ProvenanceEvidenceMeasurement(
            format,
            scale,
            operation,
            1,
            fixture.Asset.Length,
            outputBytes,
            iterations,
            stopwatch.Elapsed.TotalMicroseconds / iterations,
            allocatedBytes / iterations,
            retained,
            Math.Max(0, peakManagedHeap - heapBefore),
            process.PeakWorkingSet64);
    }

    private static int? ValidateResult(ProvenanceBenchmarkFixture fixture, string operation, object? result) {
        if (operation == "Inspect") {
            if (result is not OfficeProvenanceReport report ||
                !report.HasC2paManifest || report.Evidence.Count != 1 || !report.Evidence[0].IsStructurallyValid) {
                throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} inspection result was invalid.");
            }
            return null;
        }

        if (result is not OfficeProvenanceRemovalResult removal ||
            !removal.WasChanged || removal.Changes.Count != 1 || removal.After.HasC2paManifest) {
            throw new InvalidOperationException($"{fixture.Format}/{fixture.Scale} removal result was invalid.");
        }
        int outputBytes = removal.ToArray().Length;
        if (outputBytes != fixture.ExpectedOutputBytes) {
            throw new InvalidOperationException(
                $"{fixture.Format}/{fixture.Scale} output was {outputBytes} bytes, expected {fixture.ExpectedOutputBytes}.");
        }
        return outputBytes;
    }

    private static ProvenanceEvidenceMeasurement RunChildProbe(string format, string scale, string operation) {
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
        foreach (string argument in new[] { "--evidence-probe", format, scale, operation }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<ProvenanceEvidenceMeasurement>(output, JsonOptions)
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

internal sealed record ProvenanceEvidenceMeasurement(
    string Format,
    string Scale,
    string Operation,
    int Iteration,
    int InputBytes,
    int? OutputBytes,
    int Operations,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record ProvenanceEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<ProvenanceEvidenceMeasurement> Measurements);

internal sealed class ProvenanceManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
    private int _stopped;

    internal ProvenanceManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Provenance heap sampler" };
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
        long observed = GC.GetTotalMemory(forceFullCollection: false);
        long current = Interlocked.Read(ref _peakBytes);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref _peakBytes, observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}
