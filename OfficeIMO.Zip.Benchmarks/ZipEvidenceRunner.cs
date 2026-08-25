using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Zip.Benchmarks;

internal static class ZipEvidenceRunner {
    private static readonly string[] Engines = { "OfficeIMO", "System.IO.Compression" };
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 3) {
            Console.Error.WriteLine("Usage: --probe <OfficeIMO|System.IO.Compression> <Small|Normal|Large> <source.zip>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], args[1], args[2]), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        int repeat = GetPositiveIntOption(args, "--repeat", 3);
        ZipBenchmarkScale[] scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? ZipComparisonCorpus.Scales
            : new[] { ZipComparisonCorpus.Get(scaleFilter!) };
        var measurements = new List<ZipEvidenceMeasurement>();

        foreach (ZipBenchmarkScale scale in scales) {
            string sourcePath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Zip-{scale.Name}-{Guid.NewGuid():N}.zip");
            File.WriteAllBytes(sourcePath, ZipComparisonCorpus.CreateArchive(scale));
            try {
                foreach (string engine in Engines) {
                    for (int iteration = 1; iteration <= repeat; iteration++) {
                        ZipEvidenceMeasurement measurement = RunChildProbe(engine, scale.Name, sourcePath)
                            with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{engine,-21} {scale.Name,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedMilliseconds,9:F2} ms " +
                            $"{measurement.AllocatedBytes / 1048576D,9:F2} MiB alloc " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,9:F2} MiB managed-peak growth " +
                            $"{measurement.InputBytes / 1024D,9:F2} KiB input");
                    }
                }
            } finally {
                File.Delete(sourcePath);
            }
        }

        var report = new ZipEvidenceReport(
            DateTimeOffset.UtcNow,
            ResolveCommit(),
            ResolveSourceTreeDirty(),
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            Environment.ProcessorCount,
            measurements);
        if (!string.IsNullOrWhiteSpace(jsonPath)) {
            string fullPath = Path.GetFullPath(jsonPath!);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
            Console.WriteLine("Wrote " + fullPath);
        }
        return 0;
    }

    private static ZipEvidenceMeasurement Measure(string engine, string scaleName, string sourcePath) {
        string selectedEngine = Engines.FirstOrDefault(value => string.Equals(value, engine, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown ZIP benchmark engine: " + engine, nameof(engine));
        ZipBenchmarkScale scale = ZipComparisonCorpus.Get(scaleName);
        byte[] input = File.ReadAllBytes(sourcePath);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using var sampler = new ZipManagedHeapSampler();
        var stopwatch = Stopwatch.StartNew();
        object result = string.Equals(selectedEngine, "OfficeIMO", StringComparison.Ordinal)
            ? ZipComparisonWorkflows.TraverseOffice(input)
            : ZipComparisonWorkflows.TraversePlatform(input);
        stopwatch.Stop();
        long peakManagedHeap = sampler.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;

        ZipComparisonObservation observation = result is ZipTraversalResult office
            ? ZipComparisonWorkflows.Observe(office.Entries)
            : ZipComparisonWorkflows.Observe((IReadOnlyList<ZipProjectionDescriptor>)result);
        ZipComparisonReport expected = ZipComparisonValidation.Validate(scale, input);
        if (observation.EntryCount != expected.EntryCount
            || observation.TotalUncompressedBytes != expected.TotalUncompressedBytes
            || !string.Equals(observation.StructuralFingerprint, expected.StructuralFingerprint, StringComparison.Ordinal)) {
            throw new InvalidOperationException(selectedEngine + "/" + scale.Name + " failed exact projection validation.");
        }
        GC.KeepAlive(result);
        return new ZipEvidenceMeasurement(
            selectedEngine,
            scale.Name,
            1,
            input.LongLength,
            observation.EntryCount,
            stopwatch.Elapsed.TotalMilliseconds,
            allocatedBytes,
            Math.Max(0, peakManagedHeap - heapBefore),
            observation.StructuralFingerprint);
    }

    private static ZipEvidenceMeasurement RunChildProbe(string engine, string scale, string sourcePath) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
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
        foreach (string argument in new[] { "--probe", engine, scale, sourcePath }) startInfo.ArgumentList.Add(argument);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start ZIP benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {engine}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<ZipEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {engine}/{scale} returned no measurement.");
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
            throw new ArgumentException(name + " requires a value.");
        }
        return args[index + 1];
    }

    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) {
        string? value = GetOption(args, name);
        if (value == null) return defaultValue;
        return int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            using Process process = Process.Start(CreateGitStartInfo("rev-parse", "HEAD"))!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch {
            return "unknown";
        }
    }

    private static bool ResolveSourceTreeDirty() {
        try {
            using Process tracked = Process.Start(CreateGitStartInfo("diff", "--quiet", "HEAD", "--"))!;
            tracked.WaitForExit();
            if (tracked.ExitCode != 0) return true;
            using Process untracked = Process.Start(CreateGitStartInfo("ls-files", "--others", "--exclude-standard"))!;
            string output = untracked.StandardOutput.ReadToEnd();
            untracked.WaitForExit();
            return untracked.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch {
            return true;
        }
    }

    private static ProcessStartInfo CreateGitStartInfo(params string[] arguments) {
        var startInfo = new ProcessStartInfo {
            FileName = "git",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }
}

internal sealed class ZipManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes;
    private int _stopped;

    internal ZipManagedHeapSampler() {
        _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
        _thread = new Thread(SampleUntilStopped) { IsBackground = true, Name = "OfficeIMO.Zip managed heap sampler" };
        _thread.Start();
    }

    internal long Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentHeap();
        }
        return Interlocked.Read(ref _peakBytes);
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void SampleUntilStopped() {
        while (!_stop.Wait(1)) RecordCurrentHeap();
    }

    private void RecordCurrentHeap() {
        long observed = GC.GetTotalMemory(forceFullCollection: false);
        long current = Interlocked.Read(ref _peakBytes);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref _peakBytes, observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}

internal sealed record ZipEvidenceMeasurement(
    string Engine,
    string Scale,
    int Iteration,
    long InputBytes,
    int EntryCount,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakManagedHeapGrowthBytes,
    string StructuralFingerprint);

internal sealed record ZipEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<ZipEvidenceMeasurement> Measurements);
