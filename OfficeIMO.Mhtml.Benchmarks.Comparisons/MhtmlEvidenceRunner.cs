using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using MimeKit;

namespace OfficeIMO.Mhtml.Benchmarks.Comparisons;

internal static class MhtmlEvidenceRunner {
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
            var measurements = new List<MhtmlEvidenceMeasurement>();
            foreach (string scale in MhtmlComparisonCorpus.ScaleNames) {
                foreach (string operation in new[] { "Read", "Write" }) {
                    foreach (string implementation in new[] { "OfficeIMO", "MimeKit" }) {
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            MhtmlEvidenceMeasurement measurement = RunChildProbe(scale, operation, implementation)
                                with { Iteration = iteration };
                            measurements.Add(measurement);
                            Console.WriteLine(
                                $"{scale,-6} {operation,-5} {implementation,-9} #{iteration,-2} " +
                                $"{measurement.ElapsedMillisecondsPerOperation,9:F2} ms/op " +
                                $"{measurement.AllocatedBytesPerOperation / 1048576D,9:F2} MiB/op " +
                                $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                                $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                                $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                $"{measurement.InputBytes,10:N0}/{measurement.OutputBytes,10:N0} bytes");
                        }
                    }
                }
            }

            var report = new MhtmlEvidenceReport(
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

    private static MhtmlEvidenceMeasurement Measure(string scaleName, string operation, string implementation) {
        if (operation is not ("Read" or "Write")) throw new ArgumentOutOfRangeException(nameof(operation));
        if (implementation is not ("OfficeIMO" or "MimeKit")) throw new ArgumentOutOfRangeException(nameof(implementation));
        MhtmlBenchmarkScale scale = MhtmlComparisonCorpus.Get(scaleName);
        MhtmlDocument officeDocument = MhtmlComparisonCorpus.CreateOfficeDocument(scale);
        using MimeMessage mimeMessage = MhtmlComparisonCorpus.CreateMimeMessage(scale);
        byte[] input = MhtmlComparisonCorpus.WriteMimeKit(mimeMessage);
        MhtmlComparisonValidation.Validate(scale.Name);
        int operations = scale.Name switch { "Small" => 32, "Normal" => 8, _ => 2 };

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new MhtmlManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < operations; index++) {
            result = (operation, implementation) switch {
                ("Read", "OfficeIMO") => MhtmlComparisonValidation.LoadOffice(input),
                ("Read", "MimeKit") => MhtmlComparisonValidation.LoadMimeKit(input),
                ("Write", "OfficeIMO") => officeDocument.ToBytes(),
                _ => MhtmlComparisonCorpus.WriteMimeKit(mimeMessage)
            };
        }
        stopwatch.Stop();
        long peakManaged = sampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;

        int outputBytes = 0;
        if (operation == "Write") {
            byte[] output = (byte[])result!;
            MhtmlComparisonValidation.ValidateOutput(scale, output);
            outputBytes = output.Length;
        } else if (implementation == "OfficeIMO") {
            if (result is not MhtmlDocument loaded || loaded.Resources.Count != scale.ResourceCount) {
                throw new InvalidOperationException("OfficeIMO MHTML read evidence was invalid.");
            }
        } else if (result is not MhtmlMimeKitProjection projection
                   || projection.Resources.Count != scale.ResourceCount) {
            throw new InvalidOperationException("MimeKit MHTML read evidence was invalid.");
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        var measurement = new MhtmlEvidenceMeasurement(
            scale.Name,
            operation,
            implementation,
            1,
            scale.ResourceCount,
            scale.ResourceCount * scale.ResourceBytes,
            input.Length,
            outputBytes,
            operations,
            stopwatch.Elapsed.TotalMilliseconds / operations,
            allocated / operations,
            retained,
            Math.Max(0, peakManaged - heapBefore),
            process.PeakWorkingSet64);
        if (result is IDisposable disposable) disposable.Dispose();
        return measurement;
    }

    private static MhtmlEvidenceMeasurement RunChildProbe(string scale, string operation, string implementation) {
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
        foreach (string argument in new[] { "--evidence-probe", scale, operation, implementation }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<MhtmlEvidenceMeasurement>(output, JsonOptions)
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

internal sealed record MhtmlEvidenceMeasurement(
    string Scale, string Operation, string Implementation, int Iteration,
    int ResourceCount, int DecodedResourceBytes, int InputBytes, int OutputBytes,
    int Operations, double ElapsedMillisecondsPerOperation,
    long AllocatedBytesPerOperation, long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes, long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record MhtmlEvidenceReport(
    DateTimeOffset CapturedAtUtc, string Commit, bool SourceTreeDirty,
    string Framework, string OperatingSystem, string Architecture,
    int ProcessorCount, int Repeat, IReadOnlyList<MhtmlEvidenceMeasurement> Measurements);

internal sealed class MhtmlManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(false);
    private int _stopped;

    internal MhtmlManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Mhtml heap sampler" };
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
