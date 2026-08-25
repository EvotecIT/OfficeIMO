using Markdig.Syntax;
using OfficeIMO.Markup;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Markup.Benchmarks;

internal static class OfficeMarkupEvidenceRunner {
    private const int WarmupDocuments = 32;
    private const string OfficeEngine = "OfficeIMO";
    private const string MarkdigEngine = "Markdig";
    private static readonly string[] Engines = [OfficeEngine, MarkdigEngine];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 3) {
            Console.Error.WriteLine("Usage: --evidence-probe <OfficeIMO|Markdig> <scale> <documents>");
            return 2;
        }

        try {
            if (!int.TryParse(args[2], out int documents) || documents <= 0) {
                throw new ArgumentException("documents must be a positive integer.");
            }
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0], args[1], documents), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args, bool verifyBudgets) {
        try {
            string? scaleFilter = GetOption(args, "--scale");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            int? requestedDocuments = GetOptionalPositiveIntOption(args, "--documents");
            OfficeMarkupEvidenceBudgetManifest? manifest = verifyBudgets ? LoadBudgetManifest() : null;
            string[] scales = string.IsNullOrWhiteSpace(scaleFilter)
                ? OfficeMarkupBenchmarkCorpus.Scales.ToArray()
                : [OfficeMarkupBenchmarkCorpus.Get(scaleFilter!).Scale];

            foreach (string scale in scales) {
                OfficeMarkupBenchmarkValidation.Validate(OfficeMarkupBenchmarkCorpus.Get(scale));
            }
            Console.WriteLine($"Validated equivalent semantic output for {scales.Length} scale(s).");

            var failures = new List<string>();
            var measurements = new List<OfficeMarkupEvidenceMeasurement>(scales.Length * repeat * Engines.Length);
            foreach (string scale in scales) {
                OfficeMarkupEvidenceBudget? budget = manifest?.Budgets.FirstOrDefault(item =>
                    string.Equals(item.Scale, scale, StringComparison.OrdinalIgnoreCase));
                if (verifyBudgets && budget == null) {
                    failures.Add("Missing budget for " + scale + ".");
                    continue;
                }
                int documents = requestedDocuments ?? budget?.Documents ?? ResolveAutomaticDocumentCount(scale);
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    foreach (string engine in Engines) {
                        OfficeMarkupEvidenceMeasurement measurement = RunChildProbe(engine, scale, documents)
                            with { Iteration = iteration };
                        measurements.Add(measurement);
                        if (verifyBudgets && engine == OfficeEngine) {
                            EvaluateAbsoluteBudget(budget!, measurement, failures);
                        }
                        Console.WriteLine(
                            $"{engine,-9} {scale,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedNanosecondsPerDocument / 1000D,10:F2} us/doc " +
                            $"{measurement.AllocatedBytesPerDocument / 1024D,9:F2} KiB alloc/doc " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                            $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak");
                    }
                }
            }

            IReadOnlyList<OfficeMarkupEvidenceSummary> summaries = BuildSummaries(scales, measurements);
            Console.WriteLine();
            Console.WriteLine("Median OfficeIMO / Markdig ratios (contender boundary: <= 2.00x; checked budgets may be tighter):");
            foreach (OfficeMarkupEvidenceSummary summary in summaries) {
                Console.WriteLine(
                    $"{summary.Scale,-6} {summary.ElapsedRatio,7:F2}x elapsed " +
                    $"{summary.AllocationRatio,7:F2}x allocation " +
                    $"{FormatOptionalRatio(summary.RetainedManagedRatio),9} retained " +
                    $"{FormatOptionalRatio(summary.PeakManagedHeapRatio),9} managed-peak " +
                    $"{FormatOptionalRatio(summary.ProcessPeakWorkingSetRatio),9} process-peak");
                if (verifyBudgets) {
                    OfficeMarkupEvidenceBudget budget = manifest!.Budgets.Single(item =>
                        string.Equals(item.Scale, summary.Scale, StringComparison.OrdinalIgnoreCase));
                    if (summary.ElapsedRatio > budget.MaxElapsedRatio) {
                        failures.Add($"{summary.Scale}: elapsed ratio {summary.ElapsedRatio:F2}x > {budget.MaxElapsedRatio:F2}x.");
                    }
                    if (summary.AllocationRatio > budget.MaxAllocationRatio) {
                        failures.Add($"{summary.Scale}: allocation ratio {summary.AllocationRatio:F2}x > {budget.MaxAllocationRatio:F2}x.");
                    }
                }
            }

            var report = new OfficeMarkupEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                requestedDocuments,
                repeat,
                scales,
                measurements,
                summaries,
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

    private static OfficeMarkupEvidenceMeasurement Measure(string engine, string scale, int documents) {
        string selectedEngine = Engines.FirstOrDefault(value => string.Equals(value, engine, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown engine: " + engine, nameof(engine));
        OfficeMarkupBenchmarkFixture fixture = OfficeMarkupBenchmarkCorpus.Get(scale);
        (OfficeMarkupParseResult office, MarkdownDocument markdig) = OfficeMarkupBenchmarkValidation.Validate(fixture);
        SemanticSnapshot officeSnapshot = OfficeMarkupBenchmarkValidation.CreateOfficeSnapshot(office);
        SemanticSnapshot markdigSnapshot = OfficeMarkupBenchmarkValidation.CreateMarkdigSnapshot(markdig);
        if (officeSnapshot != markdigSnapshot) throw new InvalidOperationException("Semantic validation changed before measurement.");

        for (int index = 0; index < WarmupDocuments; index++) {
            object warmup = Parse(selectedEngine, fixture.Source);
            GC.KeepAlive(warmup);
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        object? lastResult = null;
        var stopwatch = Stopwatch.StartNew();
        for (int index = 0; index < documents; index++) lastResult = Parse(selectedEngine, fixture.Source);
        stopwatch.Stop();
        long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        GC.KeepAlive(lastResult);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        var retainedResults = new object[documents];
        using var sampler = new OfficeMarkupMemorySampler(process);
        for (int index = 0; index < retainedResults.Length; index++) {
            retainedResults[index] = Parse(selectedEngine, fixture.Source);
        }
        OfficeMarkupMemoryPeak peak = sampler.Stop();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retainedManaged = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long absoluteProcessPeak = process.PeakWorkingSet64;
        GC.KeepAlive(retainedResults);

        return new OfficeMarkupEvidenceMeasurement(
            selectedEngine,
            fixture.Scale,
            1,
            documents,
            Encoding.UTF8.GetByteCount(fixture.Source),
            officeSnapshot.EventCount,
            officeSnapshot.Digest,
            stopwatch.Elapsed.TotalNanoseconds / documents,
            allocatedBytes / (double)documents,
            retainedManaged,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
            absoluteProcessPeak);
    }

    private static object Parse(string engine, string source) => engine == OfficeEngine
        ? OfficeMarkupParser.Parse(source, OfficeMarkupBenchmarkValidation.OfficeOptions)
        : global::Markdig.Markdown.Parse(source, OfficeMarkupBenchmarkValidation.MarkdigPipeline);

    private static OfficeMarkupEvidenceMeasurement RunChildProbe(string engine, string scale, int documents) {
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
        foreach (string argument in new[] { "--evidence-probe", engine, scale, documents.ToString(System.Globalization.CultureInfo.InvariantCulture) }) {
            startInfo.ArgumentList.Add(argument);
        }
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start evidence probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {engine}/{scale} failed: {error}");
        return JsonSerializer.Deserialize<OfficeMarkupEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {engine}/{scale} returned no measurement.");
    }

    private static IReadOnlyList<OfficeMarkupEvidenceSummary> BuildSummaries(
        IEnumerable<string> scales,
        IReadOnlyList<OfficeMarkupEvidenceMeasurement> measurements) {
        var summaries = new List<OfficeMarkupEvidenceSummary>();
        foreach (string scale in scales) {
            OfficeMarkupEvidenceMeasurement[] office = measurements.Where(value => value.Scale == scale && value.Engine == OfficeEngine).ToArray();
            OfficeMarkupEvidenceMeasurement[] markdig = measurements.Where(value => value.Scale == scale && value.Engine == MarkdigEngine).ToArray();
            if (office.Length == 0 || markdig.Length == 0) continue;
            if (office.Concat(markdig).Select(value => value.SemanticDigest).Distinct(StringComparer.Ordinal).Count() != 1
                || office.Concat(markdig).Select(value => value.SemanticEventCount).Distinct().Count() != 1
                || office.Concat(markdig).Select(value => value.InputBytesPerDocument).Distinct().Count() != 1) {
                throw new InvalidOperationException(scale + " probes did not observe identical input and semantic output.");
            }
            double officeElapsed = Median(office.Select(value => value.ElapsedNanosecondsPerDocument));
            double markdigElapsed = Median(markdig.Select(value => value.ElapsedNanosecondsPerDocument));
            double officeAllocated = Median(office.Select(value => value.AllocatedBytesPerDocument));
            double markdigAllocated = Median(markdig.Select(value => value.AllocatedBytesPerDocument));
            summaries.Add(new OfficeMarkupEvidenceSummary(
                scale,
                officeElapsed / markdigElapsed,
                officeAllocated / markdigAllocated,
                OptionalRatio(Median(office.Select(value => (double)value.RetainedManagedHeapGrowthBytes)), Median(markdig.Select(value => (double)value.RetainedManagedHeapGrowthBytes))),
                OptionalRatio(Median(office.Select(value => (double)value.PeakManagedHeapGrowthBytes)), Median(markdig.Select(value => (double)value.PeakManagedHeapGrowthBytes))),
                OptionalRatio(Median(office.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes)), Median(markdig.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes))),
                officeElapsed,
                markdigElapsed,
                officeAllocated,
                markdigAllocated));
        }
        return summaries;
    }

    private static void EvaluateAbsoluteBudget(
        OfficeMarkupEvidenceBudget budget,
        OfficeMarkupEvidenceMeasurement measurement,
        ICollection<string> failures) {
        string prefix = measurement.Scale + " #" + measurement.Iteration;
        if (measurement.ElapsedNanosecondsPerDocument > budget.MaxElapsedNanosecondsPerDocument) failures.Add($"{prefix}: elapsed {measurement.ElapsedNanosecondsPerDocument:F0} ns/doc > {budget.MaxElapsedNanosecondsPerDocument:F0}.");
        if (measurement.AllocatedBytesPerDocument > budget.MaxAllocatedBytesPerDocument) failures.Add($"{prefix}: allocation {measurement.AllocatedBytesPerDocument:F0} bytes/doc > {budget.MaxAllocatedBytesPerDocument}.");
        if (measurement.RetainedManagedHeapGrowthBytes > budget.MaxRetainedManagedHeapGrowthBytes) failures.Add($"{prefix}: retained heap {measurement.RetainedManagedHeapGrowthBytes} > {budget.MaxRetainedManagedHeapGrowthBytes}.");
        if (measurement.PeakManagedHeapGrowthBytes > budget.MaxPeakManagedHeapGrowthBytes) failures.Add($"{prefix}: managed peak {measurement.PeakManagedHeapGrowthBytes} > {budget.MaxPeakManagedHeapGrowthBytes}.");
        if (measurement.AbsoluteProcessPeakWorkingSetBytes > budget.MaxAbsoluteProcessPeakWorkingSetBytes) failures.Add($"{prefix}: process peak {measurement.AbsoluteProcessPeakWorkingSetBytes} > {budget.MaxAbsoluteProcessPeakWorkingSetBytes}.");
    }

    private static OfficeMarkupEvidenceBudgetManifest LoadBudgetManifest() {
        string path = Path.Combine(AppContext.BaseDirectory, "markup-performance-budgets.json");
        return JsonSerializer.Deserialize<OfficeMarkupEvidenceBudgetManifest>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidOperationException("Markup performance budget manifest is invalid.");
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0) throw new InvalidOperationException("Cannot calculate a median without measurements.");
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }

    private static double? OptionalRatio(double numerator, double denominator) => denominator > 0 ? numerator / denominator : null;
    private static string FormatOptionalRatio(double? ratio) => ratio.HasValue ? $"{ratio.Value:F2}x" : "n/a";
    private static int ResolveAutomaticDocumentCount(string scale) => scale switch { "Small" => 256, "Normal" => 128, _ => 16 };

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }

    private static int? GetOptionalPositiveIntOption(string[] args, string name) {
        string? value = GetOption(args, name);
        if (value == null) return null;
        return int.TryParse(value, out int parsed) && parsed > 0 ? parsed : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) => GetOptionalPositiveIntOption(args, name) ?? defaultValue;

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            using Process process = Process.Start(CreateGitStartInfo("rev-parse", "HEAD"))!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch { return "unknown"; }
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
        } catch { return true; }
    }

    private static ProcessStartInfo CreateGitStartInfo(params string[] arguments) {
        var startInfo = new ProcessStartInfo { FileName = "git", RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }
}

internal sealed class OfficeMarkupMemorySampler : IDisposable {
    private readonly Process _process;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakManagedHeapBytes;
    private long _peakWorkingSetBytes;
    private int _stopped;

    internal OfficeMarkupMemorySampler(Process process) {
        _process = process;
        _peakManagedHeapBytes = GC.GetTotalMemory(forceFullCollection: false);
        _process.Refresh();
        _peakWorkingSetBytes = _process.WorkingSet64;
        _thread = new Thread(SampleUntilStopped) { IsBackground = true, Name = "OfficeIMO.Markup memory sampler" };
        _thread.Start();
    }

    internal OfficeMarkupMemoryPeak Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentMemory();
        }
        return new OfficeMarkupMemoryPeak(Interlocked.Read(ref _peakManagedHeapBytes), Interlocked.Read(ref _peakWorkingSetBytes));
    }

    public void Dispose() { Stop(); _stop.Dispose(); }
    private void SampleUntilStopped() { while (!_stop.Wait(1)) RecordCurrentMemory(); }
    private void RecordCurrentMemory() {
        RecordPeak(ref _peakManagedHeapBytes, GC.GetTotalMemory(forceFullCollection: false));
        _process.Refresh();
        RecordPeak(ref _peakWorkingSetBytes, _process.WorkingSet64);
    }
    private static void RecordPeak(ref long peak, long observed) {
        long current = Interlocked.Read(ref peak);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref peak, observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}

internal readonly record struct OfficeMarkupMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);

internal sealed record OfficeMarkupEvidenceMeasurement(
    string Engine,
    string Scale,
    int Iteration,
    int DocumentCount,
    long InputBytesPerDocument,
    int SemanticEventCount,
    string SemanticDigest,
    double ElapsedNanosecondsPerDocument,
    double AllocatedBytesPerDocument,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record OfficeMarkupEvidenceSummary(
    string Scale,
    double ElapsedRatio,
    double AllocationRatio,
    double? RetainedManagedRatio,
    double? PeakManagedHeapRatio,
    double? ProcessPeakWorkingSetRatio,
    double OfficeElapsedNanosecondsPerDocument,
    double MarkdigElapsedNanosecondsPerDocument,
    double OfficeAllocatedBytesPerDocument,
    double MarkdigAllocatedBytesPerDocument);

internal sealed record OfficeMarkupEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int? RequestedDocumentsPerProbe,
    int Repeat,
    IReadOnlyList<string> ValidatedEquivalentScales,
    IReadOnlyList<OfficeMarkupEvidenceMeasurement> Measurements,
    IReadOnlyList<OfficeMarkupEvidenceSummary> Summaries,
    IReadOnlyList<string> Failures);

internal sealed class OfficeMarkupEvidenceBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<OfficeMarkupEvidenceBudget> Budgets { get; set; } = new();
}

internal sealed class OfficeMarkupEvidenceBudget {
    public string Scale { get; set; } = string.Empty;
    public int Documents { get; set; }
    public double MaxElapsedNanosecondsPerDocument { get; set; }
    public long MaxAllocatedBytesPerDocument { get; set; }
    public long MaxRetainedManagedHeapGrowthBytes { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxAbsoluteProcessPeakWorkingSetBytes { get; set; }
    public double MaxElapsedRatio { get; set; }
    public double MaxAllocationRatio { get; set; }
}
