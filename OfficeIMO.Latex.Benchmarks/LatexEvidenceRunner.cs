using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Latex.Benchmarks;

internal static class LatexEvidenceRunner {
    private static readonly string[] Operations = ["Parse", "ParseWrite"];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 2) {
            Console.Error.WriteLine("Usage: --evidence-probe <Parse|ParseWrite> <Small|Normal|Large>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(SelectOperation(args[0]), args[1]), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args, bool verifyBudgets) {
        try {
            string? operationFilter = GetOption(args, "--operation");
            string? scaleFilter = GetOption(args, "--scale");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 1);
            string[] operations = string.IsNullOrWhiteSpace(operationFilter)
                ? Operations
                : [SelectOperation(operationFilter!)];
            string[] scales = string.IsNullOrWhiteSpace(scaleFilter)
                ? LatexBenchmarkCorpus.Scales.ToArray()
                : [LatexBenchmarkCorpus.Get(scaleFilter!).Scale];
            LatexEvidenceBudgetManifest? manifest = verifyBudgets ? LoadBudgetManifest() : null;
            var measurements = new List<LatexEvidenceMeasurement>();
            var failures = new List<string>();
            foreach (string scale in scales) {
                LatexBenchmarkValidation.Validate(LatexBenchmarkCorpus.Get(scale));
                foreach (string operation in operations) {
                    for (int iteration = 1; iteration <= repeat; iteration++) {
                        LatexEvidenceMeasurement measurement = RunChildProbe(operation, scale) with { Iteration = iteration };
                        measurements.Add(measurement);
                        if (manifest != null) EvaluateBudget(manifest, measurement, failures);
                        Console.WriteLine(
                            $"{operation,-10} {scale,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedMilliseconds,9:F2} ms " +
                            $"{measurement.AllocatedBytes / 1048576D,9:F2} MiB alloc " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,9:F2} MiB retained " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,9:F2} MiB managed peak " +
                            $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,9:F2} MiB process peak " +
                            $"{measurement.OutputBytes / 1024D,9:F2} KiB output");
                    }
                }
            }
            var report = new LatexEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                measurements,
                failures);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            foreach (string failure in failures) Console.Error.WriteLine("BUDGET FAILURE: " + failure);
            return failures.Count == 0 ? 0 : 1;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static LatexEvidenceMeasurement Measure(string operation, string scale) {
        LatexBenchmarkFixture fixture = LatexBenchmarkCorpus.Get(scale);
        LatexBenchmarkValidation.Validate(fixture);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        using var sampler = new LatexManagedHeapSampler();
        var stopwatch = Stopwatch.StartNew();
        object result = string.Equals(operation, "Parse", StringComparison.Ordinal)
            ? LatexDocument.Parse(fixture.Source)
            : LatexDocument.Parse(fixture.Source).Document.ToLatex();
        stopwatch.Stop();
        long peakManagedHeap = sampler.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        ValidateMeasuredResult(operation, fixture, result);
        GC.KeepAlive(result);
        return new LatexEvidenceMeasurement(
            operation,
            fixture.Scale,
            1,
            Encoding.UTF8.GetByteCount(fixture.Source),
            result is string output ? Encoding.UTF8.GetByteCount(output) : 0,
            fixture.SectionCount,
            fixture.RecordCount,
            stopwatch.Elapsed.TotalMilliseconds,
            allocatedBytes,
            retained,
            Math.Max(0, peakManagedHeap - heapBefore),
            process.PeakWorkingSet64);
    }

    private static void ValidateMeasuredResult(string operation, LatexBenchmarkFixture fixture, object result) {
        if (result is LatexParseResult parsed) {
            if (parsed.HasErrors || !parsed.IsLossless || parsed.Document.Headings.Count != fixture.SectionCount) {
                throw new InvalidOperationException(operation + "/" + fixture.Scale + " failed measured parse validation.");
            }
            return;
        }
        if (result is not string output || !string.Equals(output, fixture.Source, StringComparison.Ordinal)) {
            throw new InvalidOperationException(operation + "/" + fixture.Scale + " failed measured output validation.");
        }
    }

    private static LatexEvidenceMeasurement RunChildProbe(string operation, string scale) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
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
        foreach (string argument in new[] { "--evidence-probe", operation, scale }) startInfo.ArgumentList.Add(argument);
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start LaTeX evidence probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {operation}/{scale} failed: {error}");
        return JsonSerializer.Deserialize<LatexEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {operation}/{scale} returned no measurement.");
    }

    private static LatexEvidenceBudgetManifest LoadBudgetManifest() {
        string path = Path.Combine(AppContext.BaseDirectory, "latex-performance-budgets.json");
        return JsonSerializer.Deserialize<LatexEvidenceBudgetManifest>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidOperationException("LaTeX performance budget manifest is invalid.");
    }

    private static void EvaluateBudget(
        LatexEvidenceBudgetManifest manifest,
        LatexEvidenceMeasurement measurement,
        ICollection<string> failures) {
        LatexEvidenceBudget? budget = manifest.Budgets.FirstOrDefault(item =>
            string.Equals(item.Operation, measurement.Operation, StringComparison.OrdinalIgnoreCase)
            && string.Equals(item.Scale, measurement.Scale, StringComparison.OrdinalIgnoreCase));
        if (budget == null) {
            failures.Add($"Missing budget for {measurement.Operation}/{measurement.Scale}.");
            return;
        }
        string lane = measurement.Operation + "/" + measurement.Scale;
        Check(measurement.ElapsedMilliseconds, budget.MaxElapsedMilliseconds, "elapsed ms");
        Check(measurement.AllocatedBytes, budget.MaxAllocatedBytes, "allocated bytes");
        Check(measurement.RetainedManagedHeapGrowthBytes, budget.MaxRetainedManagedHeapGrowthBytes, "retained bytes");
        Check(measurement.PeakManagedHeapGrowthBytes, budget.MaxPeakManagedHeapGrowthBytes, "managed peak bytes");
        Check(measurement.AbsoluteProcessPeakWorkingSetBytes, budget.MaxAbsoluteProcessPeakWorkingSetBytes, "process peak bytes");
        Check(measurement.OutputBytes, budget.MaxOutputBytes, "output bytes");
        void Check(double actual, double maximum, string metric) {
            if (actual > maximum) failures.Add($"{lane}: {metric} {actual:F0} > {maximum:F0}.");
        }
    }

    private static string SelectOperation(string value) => Operations.FirstOrDefault(
        operation => string.Equals(operation, value, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown LaTeX evidence operation: " + value);

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
        return value == null ? defaultValue : int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        return RunGit("rev-parse", "HEAD") ?? "unknown";
    }

    private static bool ResolveSourceTreeDirty() =>
        RunGit("status", "--porcelain", "--untracked-files=normal") is not "";

    private static string? RunGit(params string[] arguments) {
        try {
            var startInfo = new ProcessStartInfo {
                FileName = "git",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
            using Process process = Process.Start(startInfo)!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : null;
        } catch {
            return null;
        }
    }
}

internal sealed record LatexEvidenceMeasurement(
    string Operation,
    string Scale,
    int Iteration,
    long InputBytes,
    long OutputBytes,
    int SectionCount,
    int RecordCount,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record LatexEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<LatexEvidenceMeasurement> Measurements,
    IReadOnlyList<string> Failures);

internal sealed class LatexEvidenceBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<LatexEvidenceBudget> Budgets { get; set; } = new();
}

internal sealed class LatexEvidenceBudget {
    public string Operation { get; set; } = string.Empty;
    public string Scale { get; set; } = string.Empty;
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
    public long MaxRetainedManagedHeapGrowthBytes { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxAbsoluteProcessPeakWorkingSetBytes { get; set; }
    public long MaxOutputBytes { get; set; }
}

internal sealed class LatexManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
    private int _stopped;

    internal LatexManagedHeapSampler() {
        _thread = new Thread(SampleUntilStopped) { IsBackground = true, Name = "OfficeIMO.LaTeX heap sampler" };
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
