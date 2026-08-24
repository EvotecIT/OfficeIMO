using Markdig;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class MarkdownParseEvidenceRunner {
    private const int WarmupDocuments = 128;
    private const int TargetInputBytesPerProbe = 8 * 1024 * 1024;
    private const string OfficeEngine = "OfficeIMO-Semantic";
    private const string MarkdigEngine = "Markdig";
    private static readonly string[] Engines = [OfficeEngine, MarkdigEngine];
    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 3) {
            Console.Error.WriteLine(
                "Usage: --parse-evidence-probe <OfficeIMO-Semantic|Markdig> <corpus> <documents>");
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

    internal static int RunEvidence(string[] args) {
        try {
            string? corpusFilter = GetOption(args, "--corpus");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            int? requestedDocuments = GetOptionalPositiveIntOption(args, "--documents");
            string[] corpusNames = ResolveCorpusNames(corpusFilter);

            foreach (string corpusName in corpusNames) {
                MarkdownBenchmarkValidation.AssertCommonMarkEquivalent(
                    corpusName,
                    MarkdownBenchmarkCorpus.Get(corpusName),
                    MarkdownReaderOptions.CreateCommonMarkProfile(),
                    MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions());
            }

            Console.WriteLine($"Validated equivalent CommonMark HTML for {corpusNames.Length} corpus/corpora.");
            var measurements = new List<MarkdownParseEvidenceMeasurement>(corpusNames.Length * repeat * Engines.Length);
            foreach (string corpusName in corpusNames) {
                int documents = requestedDocuments ?? ResolveAutomaticDocumentCount(corpusName);
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    foreach (string engine in Engines) {
                        MarkdownParseEvidenceMeasurement measurement = RunChildProbe(engine, corpusName, documents)
                            with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{engine,-18} {corpusName,-20} #{iteration,-2} " +
                            $"{measurement.ElapsedNanosecondsPerDocument / 1000D,10:F2} us/doc " +
                            $"{measurement.AllocatedBytesPerDocument / 1024D,9:F2} KiB alloc/doc " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained/batch " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                            $"{measurement.PeakWorkingSetGrowthBytes / 1048576D,8:F2} MiB WS peak");
                    }
                }
            }

            IReadOnlyList<MarkdownParseEvidenceSummary> summaries = BuildSummaries(corpusNames, measurements);
            Console.WriteLine();
            Console.WriteLine("Median OfficeIMO / Markdig ratios (target: <= 2.00x elapsed and allocation):");
            foreach (MarkdownParseEvidenceSummary summary in summaries) {
                Console.WriteLine(
                    $"{summary.Corpus,-20} {summary.ElapsedRatio,7:F2}x elapsed " +
                    $"{summary.AllocationRatio,7:F2}x allocation " +
                    $"{FormatOptionalRatio(summary.RetainedManagedRatio),9} retained " +
                    $"{FormatOptionalRatio(summary.PeakManagedHeapRatio),9} managed-peak " +
                    $"{FormatOptionalRatio(summary.ProcessPeakWorkingSetRatio),9} process-peak");
            }

            var report = new MarkdownParseEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                requestedDocuments,
                repeat,
                corpusNames,
                measurements,
                summaries);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(directory)) {
                    Directory.CreateDirectory(directory);
                }
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static MarkdownParseEvidenceMeasurement Measure(string engine, string corpusName, int documents) {
        string selectedEngine = Engines.FirstOrDefault(value => string.Equals(value, engine, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown Markdown parse benchmark engine: " + engine, nameof(engine));
        string selectedCorpus = MarkdownBenchmarkCorpus.AllNames.FirstOrDefault(
            value => string.Equals(value, corpusName, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown Markdown benchmark corpus: " + corpusName, nameof(corpusName));
        string markdown = MarkdownBenchmarkCorpus.Get(selectedCorpus);
        var officeOptions = MarkdownReaderOptions.CreateCommonMarkProfile();
        var officeHtmlOptions = MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions();
        MarkdownBenchmarkValidation.AssertCommonMarkEquivalent(selectedCorpus, markdown, officeOptions, officeHtmlOptions);

        string semanticHtml = MarkdownBenchmarkValidation.NormalizeHtml(
            Markdig.Markdown.ToHtml(markdown, MarkdigCommonMarkPipeline));
        string semanticFingerprint = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(semanticHtml)));
        long inputBytes = Encoding.UTF8.GetByteCount(markdown);
        long semanticHtmlBytes = Encoding.UTF8.GetByteCount(semanticHtml);

        for (int index = 0; index < WarmupDocuments; index++) {
            object warmup = Parse(selectedEngine, markdown, officeOptions);
            GC.KeepAlive(warmup);
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        var stopwatch = new Stopwatch();
        object? lastResult = null;
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();

        stopwatch.Restart();
        for (int index = 0; index < documents; index++) {
            lastResult = Parse(selectedEngine, markdown, officeOptions);
        }
        stopwatch.Stop();
        long allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        GC.KeepAlive(lastResult);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        var results = new object[documents];
        using var sampler = new MarkdownParseMemorySampler(process);
        for (int index = 0; index < results.Length; index++) {
            results[index] = Parse(selectedEngine, markdown, officeOptions);
        }
        MarkdownParseMemoryPeak peak = sampler.Stop();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retainedManagedHeapGrowth = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long absoluteProcessPeak = process.PeakWorkingSet64;
        GC.KeepAlive(results);

        return new MarkdownParseEvidenceMeasurement(
            selectedEngine,
            selectedCorpus,
            1,
            documents,
            inputBytes,
            semanticHtmlBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            stopwatch.Elapsed.TotalNanoseconds / documents,
            allocatedBytes,
            allocatedBytes / (double)documents,
            retainedManagedHeapGrowth,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
            absoluteProcessPeak,
            semanticFingerprint);
    }

    private static object Parse(string engine, string markdown, MarkdownReaderOptions officeOptions) =>
        string.Equals(engine, OfficeEngine, StringComparison.Ordinal)
            ? MarkdownReader.ParseSemantic(markdown, officeOptions)
            : Markdig.Markdown.Parse(markdown, MarkdigCommonMarkPipeline);

    private static MarkdownParseEvidenceMeasurement RunChildProbe(string engine, string corpus, int documents) {
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
        foreach (string argument in new[] {
            "--parse-evidence-probe",
            engine,
            corpus,
            documents.ToString(System.Globalization.CultureInfo.InvariantCulture)
        }) {
            startInfo.ArgumentList.Add(argument);
        }

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start Markdown parse benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {engine}/{corpus} failed: {error}");
        }
        return JsonSerializer.Deserialize<MarkdownParseEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {engine}/{corpus} returned no measurement.");
    }

    private static IReadOnlyList<MarkdownParseEvidenceSummary> BuildSummaries(
        IEnumerable<string> corpusNames,
        IReadOnlyList<MarkdownParseEvidenceMeasurement> measurements) {
        var summaries = new List<MarkdownParseEvidenceSummary>();
        foreach (string corpusName in corpusNames) {
            MarkdownParseEvidenceMeasurement[] office = measurements
                .Where(value => value.Corpus == corpusName && value.Engine == OfficeEngine)
                .ToArray();
            MarkdownParseEvidenceMeasurement[] markdig = measurements
                .Where(value => value.Corpus == corpusName && value.Engine == MarkdigEngine)
                .ToArray();

            double officeElapsed = Median(office.Select(value => value.ElapsedNanosecondsPerDocument));
            double markdigElapsed = Median(markdig.Select(value => value.ElapsedNanosecondsPerDocument));
            double officeAllocated = Median(office.Select(value => value.AllocatedBytesPerDocument));
            double markdigAllocated = Median(markdig.Select(value => value.AllocatedBytesPerDocument));
            double officeRetained = Median(office.Select(value => (double)value.RetainedManagedHeapGrowthBytes));
            double markdigRetained = Median(markdig.Select(value => (double)value.RetainedManagedHeapGrowthBytes));
            double officeManagedPeak = Median(office.Select(value => (double)value.PeakManagedHeapGrowthBytes));
            double markdigManagedPeak = Median(markdig.Select(value => (double)value.PeakManagedHeapGrowthBytes));
            double officeProcessPeak = Median(office.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes));
            double markdigProcessPeak = Median(markdig.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes));

            if (office.Concat(markdig).Select(value => value.SemanticFingerprint).Distinct(StringComparer.Ordinal).Count() != 1
                || office.Concat(markdig).Select(value => value.InputBytesPerDocument).Distinct().Count() != 1
                || office.Concat(markdig).Select(value => value.SemanticHtmlBytesPerDocument).Distinct().Count() != 1) {
                throw new InvalidOperationException(corpusName + " probes did not observe identical input and semantic output.");
            }

            summaries.Add(new MarkdownParseEvidenceSummary(
                corpusName,
                officeElapsed / markdigElapsed,
                officeAllocated / markdigAllocated,
                OptionalRatio(officeRetained, markdigRetained),
                OptionalRatio(officeManagedPeak, markdigManagedPeak),
                OptionalRatio(officeProcessPeak, markdigProcessPeak),
                officeElapsed,
                markdigElapsed,
                officeAllocated,
                markdigAllocated));
        }
        return summaries;
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0) {
            throw new InvalidOperationException("Cannot calculate a median without measurements.");
        }
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2D
            : ordered[middle];
    }

    private static double? OptionalRatio(double numerator, double denominator) => denominator > 0D
        ? numerator / denominator
        : null;

    private static string FormatOptionalRatio(double? ratio) => ratio.HasValue ? $"{ratio.Value:F2}x" : "n/a";

    private static string[] ResolveCorpusNames(string? corpusFilter) {
        if (string.IsNullOrWhiteSpace(corpusFilter)) {
            return MarkdownBenchmarkCorpus.Names.ToArray();
        }
        string selected = MarkdownBenchmarkCorpus.AllNames.FirstOrDefault(
            value => string.Equals(value, corpusFilter, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException(
                $"Unknown Markdown benchmark corpus '{corpusFilter}'. Valid values: " +
                string.Join(", ", MarkdownBenchmarkCorpus.AllNames));
        return [selected];
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) {
            return null;
        }
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
            throw new ArgumentException(name + " requires a value.");
        }
        return args[index + 1];
    }

    private static int? GetOptionalPositiveIntOption(string[] args, string name) {
        string? value = GetOption(args, name);
        if (value == null) {
            return null;
        }
        return int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) =>
        GetOptionalPositiveIntOption(args, name) ?? defaultValue;

    private static int ResolveAutomaticDocumentCount(string corpusName) {
        int inputBytes = Encoding.UTF8.GetByteCount(MarkdownBenchmarkCorpus.Get(corpusName));
        return Math.Clamp(TargetInputBytesPerProbe / Math.Max(1, inputBytes), 16, 256);
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) {
            return value;
        }
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
            if (tracked.ExitCode != 0) {
                return true;
            }
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
        foreach (string argument in arguments) {
            startInfo.ArgumentList.Add(argument);
        }
        return startInfo;
    }
}

internal sealed class MarkdownParseMemorySampler : IDisposable {
    private readonly Process _process;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakManagedHeapBytes;
    private long _peakWorkingSetBytes;
    private int _stopped;

    internal MarkdownParseMemorySampler(Process process) {
        _process = process;
        _peakManagedHeapBytes = GC.GetTotalMemory(forceFullCollection: false);
        _process.Refresh();
        _peakWorkingSetBytes = _process.WorkingSet64;
        _thread = new Thread(SampleUntilStopped) {
            IsBackground = true,
            Name = "OfficeIMO.Markdown parse memory sampler"
        };
        _thread.Start();
    }

    internal MarkdownParseMemoryPeak Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentMemory();
        }
        return new MarkdownParseMemoryPeak(
            Interlocked.Read(ref _peakManagedHeapBytes),
            Interlocked.Read(ref _peakWorkingSetBytes));
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void SampleUntilStopped() {
        while (!_stop.Wait(1)) {
            RecordCurrentMemory();
        }
    }

    private void RecordCurrentMemory() {
        RecordPeak(ref _peakManagedHeapBytes, GC.GetTotalMemory(forceFullCollection: false));
        _process.Refresh();
        RecordPeak(ref _peakWorkingSetBytes, _process.WorkingSet64);
    }

    private static void RecordPeak(ref long peak, long observed) {
        long current = Interlocked.Read(ref peak);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref peak, observed, current);
            if (prior == current) {
                return;
            }
            current = prior;
        }
    }
}

internal readonly record struct MarkdownParseMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);

internal sealed record MarkdownParseEvidenceMeasurement(
    string Engine,
    string Corpus,
    int Iteration,
    int DocumentCount,
    long InputBytesPerDocument,
    long SemanticHtmlBytesPerDocument,
    double ElapsedMilliseconds,
    double ElapsedNanosecondsPerDocument,
    long AllocatedBytes,
    double AllocatedBytesPerDocument,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    string SemanticFingerprint);

internal sealed record MarkdownParseEvidenceSummary(
    string Corpus,
    double ElapsedRatio,
    double AllocationRatio,
    double? RetainedManagedRatio,
    double? PeakManagedHeapRatio,
    double? ProcessPeakWorkingSetRatio,
    double OfficeElapsedNanosecondsPerDocument,
    double MarkdigElapsedNanosecondsPerDocument,
    double OfficeAllocatedBytesPerDocument,
    double MarkdigAllocatedBytesPerDocument);

internal sealed record MarkdownParseEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int? RequestedDocumentsPerProbe,
    int Repeat,
    IReadOnlyList<string> ValidatedEquivalentCorpora,
    IReadOnlyList<MarkdownParseEvidenceMeasurement> Measurements,
    IReadOnlyList<MarkdownParseEvidenceSummary> Summaries);
