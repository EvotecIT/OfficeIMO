using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Visio.Benchmarks;

internal static class VisioEvidenceRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 2) return 2;
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], args[1]), JsonOptions));
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
            var measurements = new List<VisioEvidenceMeasurement>();
            foreach (VisioBenchmarkScale scale in VisioBenchmarkCorpus.Scales) {
                foreach (string operation in new[] { "CreateSave", "LoadInspect" }) {
                    for (int iteration = 1; iteration <= repeat; iteration++) {
                        VisioEvidenceMeasurement measurement = RunChildProbe(scale.Name, operation)
                            with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{scale.Name,-6} {operation,-11} #{iteration,-2} " +
                            $"{measurement.ElapsedMillisecondsPerOperation,9:F2} ms/op " +
                            $"{measurement.AllocatedBytesPerOperation / 1048576D,9:F2} MiB/op " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                            $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                            $"{measurement.PackageBytes,10:N0} bytes");
                    }
                }
            }

            var report = new VisioEvidenceReport(
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

    private static VisioEvidenceMeasurement Measure(string scaleName, string operation) {
        VisioBenchmarkScale scale = VisioBenchmarkCorpus.Scales.Single(scale => scale.Name == scaleName);
        if (operation is not ("CreateSave" or "LoadInspect")) throw new ArgumentOutOfRangeException(nameof(operation));
        VisioBenchmarkFixture fixture = VisioBenchmarkCorpus.CreateFixture(scale);
        VisioBenchmarkValidation.ValidatePackage(fixture);
        int operations = scale.Name switch { "Small" => 32, "Normal" => 8, _ => 2 };

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new VisioManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < operations; index++) {
            result = operation == "CreateSave"
                ? VisioBenchmarkCorpus.CreateAndSave(scale)
                : VisioBenchmarkValidation.LoadAndInspect(fixture);
        }
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        if (operation == "CreateSave") {
            VisioBenchmarkValidation.ValidateBytes(scale, (byte[])result!);
        } else if (result is not VisioInspectionSnapshot snapshot || snapshot.Pages.Count != scale.PageCount) {
            throw new InvalidOperationException("Visio inspection evidence was invalid.");
        }
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        return new VisioEvidenceMeasurement(
            scale.Name,
            operation,
            1,
            scale.PageCount,
            scale.ShapeCount,
            scale.ConnectorCount,
            fixture.PackageBytes.Length,
            operations,
            stopwatch.Elapsed.TotalMilliseconds / operations,
            allocated / operations,
            retained,
            Math.Max(0, peakManaged - heapBefore),
            process.PeakWorkingSet64);
    }

    private static VisioEvidenceMeasurement RunChildProbe(string scale, string operation) {
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
        foreach (string argument in new[] { "--evidence-probe", scale, operation }) startInfo.ArgumentList.Add(argument);
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<VisioEvidenceMeasurement>(output, JsonOptions)
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
        var startInfo = new ProcessStartInfo { FileName = "git", RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        using Process? process = Process.Start(startInfo);
        if (process == null) return null;
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode == 0 ? output.Trim() : null;
    }
}

internal sealed record VisioEvidenceMeasurement(
    string Scale, string Operation, int Iteration, int PageCount, int ShapeCount,
    int ConnectorCount, int PackageBytes, int Operations,
    double ElapsedMillisecondsPerOperation, long AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes, long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record VisioEvidenceReport(
    DateTimeOffset CapturedAtUtc, string Commit, bool SourceTreeDirty,
    string Framework, string OperatingSystem, string Architecture,
    int ProcessorCount, int Repeat, IReadOnlyList<VisioEvidenceMeasurement> Measurements);

internal sealed class VisioManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(false);
    private int _stopped;

    internal VisioManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Visio heap sampler" };
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

    public void Dispose() { Stop(); _stop.Dispose(); }
    private void Sample() { while (!_stop.Wait(1)) Record(); }
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
