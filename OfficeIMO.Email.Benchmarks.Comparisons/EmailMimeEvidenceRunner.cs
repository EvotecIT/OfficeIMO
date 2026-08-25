using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using MimeKit;

namespace OfficeIMO.Email.Benchmarks.Comparisons;

internal static class EmailMimeEvidenceRunner {
    private const string OfficeEngine = "OfficeIMO";
    private const string MimeKitEngine = "MimeKit";
    private static readonly string[] Engines = [OfficeEngine, MimeKitEngine];
    private static readonly string[] Operations = ["Read", "Write"];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 6) {
            Console.Error.WriteLine(
                "Usage: --evidence-probe <Read|Write> <OfficeIMO|MimeKit> <Small|Normal> <input.eml> <operations> <retained-results>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(
                args[0],
                args[1],
                args[2],
                args[3],
                ParsePositive(args[4], "operations"),
                ParsePositive(args[5], "retained-results")), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        try {
            string? scaleFilter = GetOption(args, "--scale");
            string? operationFilter = GetOption(args, "--operation");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            string[] scales = string.IsNullOrWhiteSpace(scaleFilter)
                ? EmailMimeComparisonCorpus.ScaleNames.ToArray()
                : [EmailMimeComparisonCorpus.Get(scaleFilter!).Name];
            string[] operations = ResolveOperations(operationFilter);
            var validated = scales.Select(EmailMimeComparisonValidation.Validate).ToArray();
            Console.WriteLine($"Validated equivalent MIME output for {scales.Length} scale(s).");

            var measurements = new List<EmailMimeEvidenceMeasurement>(
                scales.Length * operations.Length * Engines.Length * repeat);
            foreach (string scaleName in scales) {
                EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(scaleName);
                string inputPath = Path.Combine(
                    Path.GetTempPath(), $"OfficeIMO-Email-{scale.Name}-{Guid.NewGuid():N}.eml");
                using (MimeMessage inputMessage = EmailMimeComparisonCorpus.CreateMimeMessage(scale)) {
                    File.WriteAllBytes(inputPath, EmailMimeComparisonCorpus.WriteMimeKit(inputMessage));
                }
                try {
                    foreach (string operation in operations) {
                        (int operationCount, int retainedResults) = ResolveBatchSizes(scale.Name);
                        for (int iteration = 1; iteration <= repeat; iteration++) {
                            foreach (string engine in Engines) {
                                EmailMimeEvidenceMeasurement measurement = RunChildProbe(
                                    operation, engine, scale.Name, inputPath, operationCount, retainedResults)
                                    with { Iteration = iteration };
                                measurements.Add(measurement);
                                Console.WriteLine(
                                    $"{operation,-5} {engine,-9} {scale.Name,-6} #{iteration,-2} " +
                                    $"{measurement.ElapsedMicrosecondsPerOperation,10:F2} us/op " +
                                    $"{measurement.AllocatedBytesPerOperation / 1024D,9:F2} KiB alloc/op " +
                                    $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,8:F2} MiB retained " +
                                    $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,8:F2} MiB managed peak " +
                                    $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,8:F2} MiB process peak " +
                                    $"{measurement.OutputBytes / 1024D,8:F2} KiB output");
                            }
                        }
                    }
                } finally {
                    File.Delete(inputPath);
                }
            }

            IReadOnlyList<EmailMimeEvidenceSummary> summaries = BuildSummaries(operations, scales, measurements);
            Console.WriteLine();
            Console.WriteLine("Median OfficeIMO / MimeKit ratios (target: <= 2.00x on every applicable dimension):");
            foreach (EmailMimeEvidenceSummary summary in summaries) {
                Console.WriteLine(
                    $"{summary.Operation,-5} {summary.Scale,-6} " +
                    $"{summary.ElapsedRatio,7:F2}x elapsed {summary.AllocationRatio,7:F2}x allocation " +
                    $"{FormatOptionalRatio(summary.RetainedManagedRatio),9} retained " +
                    $"{FormatOptionalRatio(summary.PeakManagedHeapRatio),9} managed-peak " +
                    $"{FormatOptionalRatio(summary.ProcessPeakWorkingSetRatio),9} process-peak " +
                    $"{FormatOptionalRatio(summary.OutputSizeRatio),9} output-size");
            }

            var report = new EmailMimeEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                validated.Select(value => value.Scale).ToArray(),
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

    private static EmailMimeEvidenceMeasurement Measure(
        string operation,
        string engine,
        string scaleName,
        string inputPath,
        int operations,
        int retainedResults) {
        string selectedOperation = ResolveOperations(operation).Single();
        string selectedEngine = Engines.FirstOrDefault(
            value => string.Equals(value, engine, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown MIME benchmark engine: " + engine, nameof(engine));
        EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(scaleName);
        byte[] input = File.ReadAllBytes(inputPath);
        object? preparedModel = selectedOperation == "Write" ? CreateModel(selectedEngine, scale) : null;
        try {
            EmailMimeObservation? inputObservation = selectedOperation == "Read"
                ? selectedEngine == OfficeEngine
                    ? EmailMimeComparisonValidation.ObserveOffice(input)
                    : EmailMimeComparisonValidation.ObserveMimeKit(input)
                : null;
            int warmups = scale.Name == "Normal" ? 128 : 512;
            for (int index = 0; index < warmups; index++) {
                object warmup = Execute(selectedOperation, selectedEngine, input, preparedModel);
                GC.KeepAlive(warmup);
            }

            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
            var stopwatch = Stopwatch.StartNew();
            object? lastResult = null;
            for (int index = 0; index < operations; index++) {
                lastResult = Execute(selectedOperation, selectedEngine, input, preparedModel);
            }
            stopwatch.Stop();
            long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
            GC.KeepAlive(lastResult);

            byte[] semanticSource = selectedOperation == "Write" ? (byte[])lastResult! : input;
            EmailMimeObservation observation = inputObservation ?? (selectedEngine == OfficeEngine
                ? EmailMimeComparisonValidation.ObserveOffice(semanticSource)
                : EmailMimeComparisonValidation.ObserveMimeKit(semanticSource));
            string fingerprint = Convert.ToHexString(SHA256.HashData(
                JsonSerializer.SerializeToUtf8Bytes(observation, JsonOptions)));

            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
            using Process process = Process.GetCurrentProcess();
            process.Refresh();
            long workingSetBefore = process.WorkingSet64;
            var retained = new object[retainedResults];
            using var sampler = new EmailMimeMemorySampler(process);
            for (int index = 0; index < retained.Length; index++) {
                retained[index] = selectedOperation == "Read"
                    ? RetainRead(selectedEngine, input)
                    : Execute(selectedOperation, selectedEngine, input, preparedModel);
            }
            EmailMimeMemoryPeak peak = sampler.Stop();
            GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
            long retainedManaged = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
            process.Refresh();
            long absoluteProcessPeak = process.PeakWorkingSet64;
            GC.KeepAlive(retained);
            foreach (object result in retained) {
                if (result is IDisposable disposable) disposable.Dispose();
            }

            return new EmailMimeEvidenceMeasurement(
                selectedOperation,
                selectedEngine,
                scale.Name,
                1,
                operations,
                retainedResults,
                input.LongLength,
                selectedOperation == "Write" ? semanticSource.LongLength : 0,
                stopwatch.Elapsed.TotalMilliseconds,
                stopwatch.Elapsed.TotalMicroseconds / operations,
                allocatedBytes,
                allocatedBytes / (double)operations,
                retainedManaged,
                Math.Max(0, peak.ManagedHeapBytes - heapBefore),
                Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
                absoluteProcessPeak,
                fingerprint);
        } finally {
            if (preparedModel is IDisposable disposable) disposable.Dispose();
        }
    }

    private static object CreateModel(string engine, EmailMimeBenchmarkScale scale) =>
        engine == OfficeEngine
            ? EmailMimeComparisonCorpus.CreateOfficeDocument(scale)
            : EmailMimeComparisonCorpus.CreateMimeMessage(scale);

    private static object Execute(string operation, string engine, byte[] input, object? preparedModel) {
        if (operation == "Read") {
            return engine == OfficeEngine
                ? EmailMimeComparisonValidation.ConsumeOffice(input)
                : EmailMimeComparisonValidation.ConsumeMimeKit(input);
        }
        return engine == OfficeEngine
            ? ((EmailDocument)preparedModel!).ToBytes(EmailFileFormat.Eml)
            : EmailMimeComparisonCorpus.WriteMimeKit((MimeMessage)preparedModel!);
    }

    private static EmailMimeRetainedProjection RetainRead(string engine, byte[] input) =>
        engine == OfficeEngine
            ? EmailMimeComparisonValidation.RetainOffice(input)
            : EmailMimeComparisonValidation.RetainMimeKit(input);

    private static EmailMimeEvidenceMeasurement RunChildProbe(
        string operation,
        string engine,
        string scale,
        string inputPath,
        int operations,
        int retainedResults) {
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
            "--evidence-probe", operation, engine, scale, inputPath,
            operations.ToString(System.Globalization.CultureInfo.InvariantCulture),
            retainedResults.ToString(System.Globalization.CultureInfo.InvariantCulture)
        }) startInfo.ArgumentList.Add(argument);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start MIME benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {operation}/{engine}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<EmailMimeEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {operation}/{engine}/{scale} returned no measurement.");
    }

    private static IReadOnlyList<EmailMimeEvidenceSummary> BuildSummaries(
        IEnumerable<string> operations,
        IEnumerable<string> scales,
        IReadOnlyList<EmailMimeEvidenceMeasurement> measurements) {
        var summaries = new List<EmailMimeEvidenceSummary>();
        foreach (string operation in operations) {
            foreach (string scale in scales) {
                EmailMimeEvidenceMeasurement[] office = measurements.Where(value =>
                    value.Operation == operation && value.Scale == scale && value.Engine == OfficeEngine).ToArray();
                EmailMimeEvidenceMeasurement[] mimeKit = measurements.Where(value =>
                    value.Operation == operation && value.Scale == scale && value.Engine == MimeKitEngine).ToArray();
                EmailMimeEvidenceMeasurement[] combined = office.Concat(mimeKit).ToArray();
                if (combined.Select(value => value.SemanticFingerprint).Distinct(StringComparer.Ordinal).Count() != 1
                    || combined.Select(value => value.InputBytes).Distinct().Count() != 1) {
                    throw new InvalidOperationException($"{operation}/{scale} probes did not observe equivalent results.");
                }
                summaries.Add(new EmailMimeEvidenceSummary(
                    operation,
                    scale,
                    Ratio(office, mimeKit, value => value.ElapsedMicrosecondsPerOperation),
                    Ratio(office, mimeKit, value => value.AllocatedBytesPerOperation),
                    OptionalRatio(office, mimeKit, value => value.RetainedManagedHeapGrowthBytes),
                    OptionalRatio(office, mimeKit, value => value.PeakManagedHeapGrowthBytes),
                    OptionalRatio(office, mimeKit, value => value.AbsoluteProcessPeakWorkingSetBytes),
                    operation == "Write" ? OptionalRatio(office, mimeKit, value => value.OutputBytes) : null));
            }
        }
        return summaries;
    }

    private static double Ratio(
        IEnumerable<EmailMimeEvidenceMeasurement> office,
        IEnumerable<EmailMimeEvidenceMeasurement> mimeKit,
        Func<EmailMimeEvidenceMeasurement, double> selector) =>
        Median(office.Select(selector)) / Median(mimeKit.Select(selector));

    private static double? OptionalRatio(
        IEnumerable<EmailMimeEvidenceMeasurement> office,
        IEnumerable<EmailMimeEvidenceMeasurement> mimeKit,
        Func<EmailMimeEvidenceMeasurement, double> selector) {
        double denominator = Median(mimeKit.Select(selector));
        return denominator > 0D ? Median(office.Select(selector)) / denominator : null;
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0) throw new InvalidOperationException("Cannot calculate a median without measurements.");
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }

    private static (int Operations, int RetainedResults) ResolveBatchSizes(string scale) =>
        scale == "Normal" ? (128, 16) : (512, 64);

    private static string[] ResolveOperations(string? operation) {
        if (string.IsNullOrWhiteSpace(operation)) return Operations;
        string selected = Operations.FirstOrDefault(value =>
            string.Equals(value, operation, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException("Unknown MIME benchmark operation: " + operation, nameof(operation));
        return [selected];
    }

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
            string output = ReadProcessOutput(untracked);
            return untracked.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch {
            return true;
        }
    }

    private static string ReadProcessOutput(Process process) {
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return output;
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
