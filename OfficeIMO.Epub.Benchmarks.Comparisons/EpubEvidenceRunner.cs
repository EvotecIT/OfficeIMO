using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Epub.Benchmarks.Comparisons;

internal static class EpubEvidenceRunner {
    private const int WarmupOperations = 256;
    private const string OfficeEngine = "OfficeIMO";
    private const string VersOneEngine = "VersOne.Epub+HAP";
    private static readonly string[] Engines = [OfficeEngine, VersOneEngine];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 4) {
            Console.Error.WriteLine(
                "Usage: --evidence-probe <OfficeIMO|VersOne.Epub+HAP> <Small|Normal> <operations> <retained-documents>");
            return 2;
        }

        try {
            int operations = ParsePositive(args[2], "operations");
            int retainedDocuments = ParsePositive(args[3], "retained-documents");
            Console.WriteLine(JsonSerializer.Serialize(
                Measure(args[0], args[1], operations, retainedDocuments), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        try {
            string? scaleFilter = GetOption(args, "--scale");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            EpubComparisonScale[] scales = string.IsNullOrWhiteSpace(scaleFilter)
                ? EpubComparisonCorpus.Scales.ToArray()
                : [EpubComparisonCorpus.Get(scaleFilter!)];

            foreach (EpubComparisonScale scale in scales) EpubComparisonValidation.Validate(scale.Name);
            Console.WriteLine($"Validated equivalent EPUB output for {scales.Length} scale(s).");

            var measurements = new List<EpubEvidenceMeasurement>(scales.Length * Engines.Length * repeat);
            foreach (EpubComparisonScale scale in scales) {
                (int operations, int retainedDocuments) = ResolveBatchSizes(scale.Name);
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    foreach (string engine in Engines) {
                        EpubEvidenceMeasurement measurement = RunChildProbe(
                            engine, scale.Name, operations, retainedDocuments) with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{engine,-20} {scale.Name,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedMicrosecondsPerOperation,10:F2} us/op " +
                            $"{measurement.AllocatedBytesPerOperation / 1048576D,9:F2} MiB alloc/op " +
                            $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                            $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                            $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak");
                    }
                }
            }

            IReadOnlyList<EpubEvidenceSummary> summaries = BuildSummaries(scales, measurements);
            Console.WriteLine();
            Console.WriteLine("Median OfficeIMO / VersOne ratios (target: <= 2.00x on every dimension):");
            foreach (EpubEvidenceSummary summary in summaries) {
                Console.WriteLine(
                    $"{summary.Scale,-6} {summary.ElapsedRatio,7:F2}x elapsed " +
                    $"{summary.AllocationRatio,7:F2}x allocation " +
                    $"{FormatOptionalRatio(summary.RetainedManagedRatio),9} retained " +
                    $"{FormatOptionalRatio(summary.PeakManagedHeapRatio),9} managed-peak " +
                    $"{FormatOptionalRatio(summary.ProcessPeakWorkingSetRatio),9} process-peak");
            }

            var report = new EpubEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                scales.Select(scale => scale.Name).ToArray(),
                measurements,
                summaries);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
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

    private static EpubEvidenceMeasurement Measure(
        string engine,
        string scaleName,
        int operations,
        int retainedDocuments) {
        string selectedEngine = Engines.FirstOrDefault(
            value => string.Equals(value, engine, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown EPUB benchmark engine: " + engine, nameof(engine));
        EpubComparisonScale scale = EpubComparisonCorpus.Get(scaleName);
        byte[] package = EpubComparisonCorpus.CreatePackage(scale);
        EpubReadEvidence observation = Inspect(selectedEngine, package);

        for (int index = 0; index < WarmupOperations; index++) {
            long warmup = Read(selectedEngine, package);
            GC.KeepAlive(warmup);
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        long checksum = 0;
        for (int index = 0; index < operations; index++) checksum ^= Read(selectedEngine, package);
        stopwatch.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        GC.KeepAlive(checksum);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        var retained = new object[retainedDocuments];
        using var sampler = new EpubMemorySampler(process);
        for (int index = 0; index < retained.Length; index++) retained[index] = Load(selectedEngine, package);
        EpubMemoryPeak peak = sampler.Stop();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retainedManaged = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long absoluteProcessPeak = process.PeakWorkingSet64;
        GC.KeepAlive(retained);

        return new EpubEvidenceMeasurement(
            selectedEngine,
            scale.Name,
            1,
            operations,
            retainedDocuments,
            package.LongLength,
            observation.ChapterCount,
            observation.ContentCharacters,
            observation.TextCharacters,
            stopwatch.Elapsed.TotalMilliseconds,
            stopwatch.Elapsed.TotalMicroseconds / operations,
            allocatedBytes,
            allocatedBytes / (double)operations,
            retainedManaged,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
            absoluteProcessPeak,
            observation.PathHash,
            observation.ContentHash,
            observation.TextHash);
    }

    private static long Read(string engine, byte[] package) =>
        string.Equals(engine, OfficeEngine, StringComparison.Ordinal)
            ? EpubComparisonWorkflows.ReadOfficeIMO(package)
            : EpubComparisonWorkflows.ReadVersOne(package);

    private static object Load(string engine, byte[] package) =>
        string.Equals(engine, OfficeEngine, StringComparison.Ordinal)
            ? EpubComparisonWorkflows.RetainOfficeIMO(package)
            : EpubComparisonWorkflows.RetainVersOne(package);

    private static EpubReadEvidence Inspect(string engine, byte[] package) =>
        string.Equals(engine, OfficeEngine, StringComparison.Ordinal)
            ? EpubComparisonWorkflows.InspectOfficeIMO(package)
            : EpubComparisonWorkflows.InspectVersOne(package);

    private static EpubEvidenceMeasurement RunChildProbe(
        string engine,
        string scale,
        int operations,
        int retainedDocuments) {
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
            "--evidence-probe",
            engine,
            scale,
            operations.ToString(System.Globalization.CultureInfo.InvariantCulture),
            retainedDocuments.ToString(System.Globalization.CultureInfo.InvariantCulture)
        }) startInfo.ArgumentList.Add(argument);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start EPUB benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {engine}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<EpubEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {engine}/{scale} returned no measurement.");
    }

    private static IReadOnlyList<EpubEvidenceSummary> BuildSummaries(
        IEnumerable<EpubComparisonScale> scales,
        IReadOnlyList<EpubEvidenceMeasurement> measurements) {
        var summaries = new List<EpubEvidenceSummary>();
        foreach (EpubComparisonScale scale in scales) {
            EpubEvidenceMeasurement[] office = measurements
                .Where(value => value.Scale == scale.Name && value.Engine == OfficeEngine).ToArray();
            EpubEvidenceMeasurement[] versOne = measurements
                .Where(value => value.Scale == scale.Name && value.Engine == VersOneEngine).ToArray();
            EpubEvidenceMeasurement[] combined = office.Concat(versOne).ToArray();
            if (combined.Select(Fingerprint).Distinct(StringComparer.Ordinal).Count() != 1
                || combined.Select(value => value.InputBytes).Distinct().Count() != 1) {
                throw new InvalidOperationException(scale.Name + " probes did not observe identical input and output.");
            }

            double officeElapsed = Median(office.Select(value => value.ElapsedMicrosecondsPerOperation));
            double versOneElapsed = Median(versOne.Select(value => value.ElapsedMicrosecondsPerOperation));
            double officeAllocated = Median(office.Select(value => value.AllocatedBytesPerOperation));
            double versOneAllocated = Median(versOne.Select(value => value.AllocatedBytesPerOperation));
            summaries.Add(new EpubEvidenceSummary(
                scale.Name,
                officeElapsed / versOneElapsed,
                officeAllocated / versOneAllocated,
                OptionalRatio(
                    Median(office.Select(value => (double)value.RetainedManagedHeapGrowthBytes)),
                    Median(versOne.Select(value => (double)value.RetainedManagedHeapGrowthBytes))),
                OptionalRatio(
                    Median(office.Select(value => (double)value.PeakManagedHeapGrowthBytes)),
                    Median(versOne.Select(value => (double)value.PeakManagedHeapGrowthBytes))),
                OptionalRatio(
                    Median(office.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes)),
                    Median(versOne.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes))),
                officeElapsed,
                versOneElapsed,
                officeAllocated,
                versOneAllocated));
        }
        return summaries;
    }

    private static string Fingerprint(EpubEvidenceMeasurement value) => string.Join(
        ":", value.ChapterCount, value.ContentCharacters, value.TextCharacters,
        value.PathHash, value.ContentHash, value.TextHash);

    private static (int Operations, int RetainedDocuments) ResolveBatchSizes(string scale) =>
        string.Equals(scale, "Normal", StringComparison.OrdinalIgnoreCase) ? (64, 4) : (256, 16);

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0) throw new InvalidOperationException("Cannot calculate a median without measurements.");
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }

    private static double? OptionalRatio(double numerator, double denominator) =>
        denominator > 0D ? numerator / denominator : null;

    private static string FormatOptionalRatio(double? ratio) => ratio.HasValue ? $"{ratio.Value:F2}x" : "n/a";

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
        return value == null ? defaultValue : ParsePositive(value, name);
    }

    private static int ParsePositive(string value, string name) =>
        int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");

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
