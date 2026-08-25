using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Security.Benchmarks;

internal static class SecurityCmsEvidenceRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 4) return 2;
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], args[1], args[2], args[3]), JsonOptions));
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
            var measurements = new List<SecurityCmsEvidenceMeasurement>();
            foreach (string scale in SecurityCmsBenchmarkCorpus.Scales) {
                foreach (string operation in new[] { "Sign", "Verify" }) {
                    string[] producers = operation == "Sign" ? ["Self"] : ["OfficeIMO", "Platform"];
                    foreach (string producer in producers) {
                        foreach (string engine in new[] { "OfficeIMO", "Platform" }) {
                            for (int iteration = 1; iteration <= repeat; iteration++) {
                                SecurityCmsEvidenceMeasurement measurement = RunChildProbe(engine, operation, scale, producer)
                                    with { Iteration = iteration };
                                measurements.Add(measurement);
                                Console.WriteLine(
                                    $"{operation,-6} {engine,-9} {scale,-6} {producer,-9} #{iteration,-2} " +
                                    $"{measurement.ElapsedMicrosecondsPerOperation,9:F2} us/op " +
                                    $"{measurement.AllocatedBytesPerOperation / 1024D,9:F2} KiB/op " +
                                    $"{measurement.RetainedManagedHeapGrowthBytes / 1024D,9:F2} KiB retained " +
                                    $"{measurement.PeakManagedHeapGrowthBytes / 1024D,9:F2} KiB managed peak " +
                                    $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                    $"{measurement.ArtifactBytes,6:N0} bytes");
                            }
                        }
                    }
                }
            }

            var report = new SecurityCmsEvidenceReport(
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

    private static SecurityCmsEvidenceMeasurement Measure(
        string engine,
        string operation,
        string scale,
        string producer) {
        if (engine is not ("OfficeIMO" or "Platform")) throw new ArgumentOutOfRangeException(nameof(engine));
        if (operation is not ("Sign" or "Verify")) throw new ArgumentOutOfRangeException(nameof(operation));
        using SecurityCmsBenchmarkFixture fixture = SecurityCmsBenchmarkCorpus.Create(scale);
        SecurityCmsValidationSnapshot validation = SecurityCmsBenchmarkValidation.Validate(fixture);
        if (operation == "Verify" && producer is not ("OfficeIMO" or "Platform")) {
            throw new ArgumentOutOfRangeException(nameof(producer));
        }
        byte[] signature = producer == "OfficeIMO" ? validation.OfficeSignature : validation.PlatformSignature;
        int iterations = scale switch { "Small" => 64, "Normal" => 32, _ => 8 };

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new SecurityCmsManagedHeapSampler();
        object? result = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < iterations; index++) {
            result = operation == "Sign"
                ? engine == "OfficeIMO"
                    ? SecurityCmsBenchmarkValidation.SignOffice(fixture)
                    : SecurityCmsBenchmarkValidation.SignPlatform(fixture)
                : engine == "OfficeIMO"
                    ? SecurityCmsBenchmarkValidation.VerifyOffice(signature, fixture.Content)
                    : SecurityCmsBenchmarkValidation.VerifyPlatform(signature, fixture.Content);
        }
        stopwatch.Stop();
        long peakManagedHeap = sampler.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        ValidateResult(engine, operation, result);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        GC.KeepAlive(result);
        return new SecurityCmsEvidenceMeasurement(
            engine,
            operation,
            scale,
            operation == "Sign" ? engine : producer,
            1,
            fixture.Content.Length,
            operation == "Sign" ? ((byte[])result!).Length : signature.Length,
            iterations,
            stopwatch.Elapsed.TotalMicroseconds / iterations,
            allocatedBytes / iterations,
            retained,
            Math.Max(0, peakManagedHeap - heapBefore),
            process.PeakWorkingSet64);
    }

    private static void ValidateResult(string engine, string operation, object? result) {
        if (operation == "Sign") {
            if (result is not byte[] signature || signature.Length == 0) throw new InvalidOperationException("No CMS output.");
            return;
        }
        if (engine == "OfficeIMO") {
            if (result is not CmsVerificationResult office || !office.IsCryptographicallyValid) {
                throw new InvalidOperationException("OfficeIMO verification failed.");
            }
        } else if (result is not PlatformCmsVerificationSnapshot platform ||
                   platform.SignerCount != 1 || !platform.UsageAccepted || !platform.SignedAttributesAccepted) {
            throw new InvalidOperationException("Platform verification failed.");
        }
    }

    private static SecurityCmsEvidenceMeasurement RunChildProbe(
        string engine,
        string operation,
        string scale,
        string producer) {
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
        foreach (string argument in new[] { "--evidence-probe", engine, operation, scale, producer }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe failed: {error}");
        return JsonSerializer.Deserialize<SecurityCmsEvidenceMeasurement>(output, JsonOptions)
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

internal sealed record SecurityCmsEvidenceMeasurement(
    string Engine,
    string Operation,
    string Scale,
    string Producer,
    int Iteration,
    int ContentBytes,
    int ArtifactBytes,
    int Operations,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record SecurityCmsEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<SecurityCmsEvidenceMeasurement> Measurements);

internal sealed class SecurityCmsManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
    private int _stopped;

    internal SecurityCmsManagedHeapSampler() {
        _thread = new Thread(Sample) { IsBackground = true, Name = "OfficeIMO.Security heap sampler" };
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
