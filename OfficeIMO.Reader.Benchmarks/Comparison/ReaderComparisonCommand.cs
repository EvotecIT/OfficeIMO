using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Zip;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader.Benchmarks.Comparison;

internal static class ReaderComparisonCommand {
    private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static async Task<int> RunAsync(string[] args, CancellationToken cancellationToken = default) {
        ComparisonCommandOptions options;
        try {
            options = ComparisonCommandOptions.Parse(args);
        } catch (ArgumentException ex) {
            Console.Error.WriteLine(ex.Message);
            WriteUsage();
            return 2;
        }

        if (options.ShowHelp) {
            WriteUsage();
            return 0;
        }

        string outputDirectory = Path.GetFullPath(options.OutputDirectory);
        string corpusDirectory = Path.Combine(outputDirectory, "corpus");
        string outputsDirectory = Path.Combine(outputDirectory, "outputs");
        Directory.CreateDirectory(corpusDirectory);
        Directory.CreateDirectory(outputsDirectory);

        IReadOnlyList<ReaderComparisonCase> cases = ReaderComparisonCorpus.Create();
        foreach (ReaderComparisonCase item in cases) {
            await File.WriteAllBytesAsync(Path.Combine(corpusDirectory, item.SourceName), item.Bytes, cancellationToken)
                .ConfigureAwait(false);
        }

        ReaderComparisonToolResult officeIMO = RunOfficeIMO(cases, outputsDirectory);
        PopulateOfficeProcessMeasurements(officeIMO);
        var tools = new List<ReaderComparisonToolResult> { officeIMO };
        if (!string.IsNullOrWhiteSpace(options.RunnerConfigPath)) {
            ReaderComparisonConfiguration configuration = await LoadConfigurationAsync(
                options.RunnerConfigPath!,
                cancellationToken).ConfigureAwait(false);
            foreach (ReaderComparisonRunnerConfiguration runner in configuration.Runners) {
                tools.Add(await RunExternalAsync(
                    runner,
                    cases,
                    corpusDirectory,
                    outputsDirectory,
                    cancellationToken).ConfigureAwait(false));
            }
        }

        var report = new ReaderComparisonReport {
            CreatedUtc = DateTimeOffset.UtcNow,
            SourceCommit = ResolveCommit(),
            SourceTreeDirty = ResolveSourceTreeDirty(),
            Runtime = RuntimeInformation.FrameworkDescription,
            OperatingSystem = RuntimeInformation.OSDescription,
            Tools = tools
        };
        string jsonPath = Path.Combine(outputDirectory, "reader-evidence.json");
        string markdownPath = Path.Combine(outputDirectory, "reader-evidence.md");
        await File.WriteAllTextAsync(jsonPath, JsonSerializer.Serialize(report, JsonOptions), cancellationToken)
            .ConfigureAwait(false);
        await File.WriteAllTextAsync(markdownPath, BuildMarkdownReport(report), cancellationToken)
            .ConfigureAwait(false);
        Console.WriteLine("Reader extraction evidence: " + markdownPath);
        return tools[0].Cases.All(item => item.Status == "success") ? 0 : 1;
    }

    internal static ReaderComparisonToolResult RunOfficeIMO(
        IReadOnlyList<ReaderComparisonCase> cases,
        string outputDirectory) {
        Directory.CreateDirectory(Path.Combine(outputDirectory, "officeimo"));
        OfficeDocumentReader reader = CreateReader();
        var results = new List<ReaderComparisonCaseResult>(cases.Count);
        foreach (ReaderComparisonCase item in cases) {
            results.Add(RunOfficeCase(reader, item, Path.Combine(outputDirectory, "officeimo", item.Id + ".md")));
        }
        return new ReaderComparisonToolResult {
            Tool = "OfficeIMO.Reader",
            ExecutionMode = "in-process",
            Status = "success",
            Cases = results
        };
    }

    internal static int RunOfficeProbe(string[] args) {
        if (args.Length != 1) {
            Console.Error.WriteLine("Usage: office-evidence-probe <case-id>");
            return 2;
        }

        try {
            ReaderComparisonCase item = ReaderComparisonCorpus.Create().Single(candidate =>
                string.Equals(candidate.Id, args[0], StringComparison.Ordinal));
            Console.WriteLine(JsonSerializer.Serialize(MeasureOfficeProbe(item), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static ReaderComparisonCaseResult RunOfficeCase(
        OfficeDocumentReader reader,
        ReaderComparisonCase item,
        string outputPath) {
        OfficeComparisonAttempt first = RunOfficeAttempt(reader, item);
        OfficeComparisonAttempt second = RunOfficeAttempt(reader, item);
        bool expectsRejection = ExpectsMalformedInputRejection(item);
        string firstStatus = ResolveOfficeStatus(expectsRejection, first.Rejected);
        string secondStatus = ResolveOfficeStatus(expectsRejection, second.Rejected);
        string status = firstStatus == "success" && secondStatus != "success" ? secondStatus : firstStatus;
        string? error = firstStatus == "success" ? null : first.Error;
        if (secondStatus != "success") {
            error = AppendError(error, "Repeat run " + secondStatus + ": " + (second.Error ?? "No error detail was provided."));
        }
        File.WriteAllText(outputPath, first.Markdown, new UTF8Encoding(false));
        IReadOnlyList<ReaderComparisonProbeResult> probes = ReaderComparisonScorer.ScoreOfficeDocument(
            first.Markdown,
            first.Document,
            item.Probes,
            first.Rejected);
        return BuildCaseResult(
            item,
            status,
            error,
            first.Markdown,
            firstStatus == secondStatus &&
                first.Rejected == second.Rejected &&
                string.Equals(first.Markdown, second.Markdown, StringComparison.Ordinal),
            (first.DurationMilliseconds + second.DurationMilliseconds) / 2d,
            first.AllocatedBytes,
            null,
            probes);
    }

    private static ReaderOfficeProbeMeasurement MeasureOfficeProbe(ReaderComparisonCase item) {
        OfficeDocumentReader reader = CreateReader();
        for (int index = 0; index < 2; index++) {
            OfficeComparisonAttempt warmup = RunOfficeAttempt(reader, item);
            GC.KeepAlive(warmup);
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        OfficeComparisonAttempt attempt;
        ReaderMemoryPeak peak;
        using (var sampler = new ReaderMemorySampler(process)) {
            attempt = RunOfficeAttempt(reader, item);
            peak = sampler.Stop();
        }
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long absoluteProcessPeak = process.PeakWorkingSet64;
        GC.KeepAlive(attempt);

        return new ReaderOfficeProbeMeasurement(
            item.Id,
            item.Bytes.LongLength,
            Encoding.UTF8.GetByteCount(attempt.Markdown),
            attempt.DurationMilliseconds,
            allocated,
            retained,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
            absoluteProcessPeak,
            Hash(attempt.Markdown));
    }

    private static void PopulateOfficeProcessMeasurements(ReaderComparisonToolResult tool) {
        foreach (ReaderComparisonCaseResult item in tool.Cases) {
            ReaderOfficeProbeMeasurement measurement = RunOfficeChildProbe(item.CaseId);
            if (!string.Equals(item.MarkdownSha256, measurement.MarkdownSha256, StringComparison.Ordinal)) {
                item.Status = "failed";
                item.Error = AppendError(item.Error, "The isolated memory probe produced different normalized output.");
            }
            item.RetainedManagedHeapGrowthBytes = measurement.RetainedManagedHeapGrowthBytes;
            item.PeakManagedHeapGrowthBytes = measurement.PeakManagedHeapGrowthBytes;
            item.PeakWorkingSetGrowthBytes = measurement.PeakWorkingSetGrowthBytes;
            item.PeakWorkingSetBytes = measurement.AbsoluteProcessPeakWorkingSetBytes;
        }
    }

    private static ReaderOfficeProbeMeasurement RunOfficeChildProbe(string caseId) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve the Reader benchmark process path.");
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
        startInfo.ArgumentList.Add("office-evidence-probe");
        startInfo.ArgumentList.Add(caseId);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start the Reader memory probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Reader memory probe '{caseId}' failed: {error}");
        }
        return JsonSerializer.Deserialize<ReaderOfficeProbeMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Reader memory probe '{caseId}' returned no measurement.");
    }

    private static OfficeComparisonAttempt RunOfficeAttempt(
        OfficeDocumentReader reader,
        ReaderComparisonCase item) {
        OfficeDocumentReadResult? document = null;
        string markdown = string.Empty;
        string? error = null;
        bool rejected = false;
        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        try {
            document = reader.ReadDocument(item.Bytes, item.SourceName, ComparisonReaderOptions());
            markdown = ToMarkdown(document);
            rejected = HasErrorDiagnostic(document);
        } catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or FormatException) {
            error = ex.Message;
            rejected = true;
        }
        stopwatch.Stop();
        return new OfficeComparisonAttempt(
            document,
            markdown,
            error,
            rejected,
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetAllocatedBytesForCurrentThread() - before);
    }

    internal static string ResolveOfficeStatus(bool expectsRejection, bool rejected) =>
        expectsRejection == rejected ? "success" : "failed";

    private static async Task<ReaderComparisonToolResult> RunExternalAsync(
        ReaderComparisonRunnerConfiguration runner,
        IReadOnlyList<ReaderComparisonCase> cases,
        string corpusDirectory,
        string outputsDirectory,
        CancellationToken cancellationToken) {
        string safeName = SafeName(runner.Name);
        string runnerDirectory = Path.Combine(outputsDirectory, safeName);
        Directory.CreateDirectory(runnerDirectory);
        var results = new List<ReaderComparisonCaseResult>(cases.Count);
        foreach (ReaderComparisonCase item in cases) {
            string inputPath = Path.Combine(corpusDirectory, item.SourceName);
            string outputPath = Path.Combine(runnerDirectory, item.Id + ".md");
            ReaderComparisonProcessOutput first = await ReaderComparisonProcessRunner.RunAsync(
                runner,
                inputPath,
                outputPath,
                cancellationToken).ConfigureAwait(false);
            bool expectsRejection = ExpectsMalformedInputRejection(item);
            ReaderComparisonProcessOutput second = ResolveExternalStatus(first, expectsRejection) == "success"
                ? await ReaderComparisonProcessRunner.RunAsync(
                    runner,
                    inputPath,
                    GetRepeatOutputPath(outputPath),
                    cancellationToken)
                    .ConfigureAwait(false)
                : first;
            (string caseStatus, string? error) = ResolveRepeatOutcome(first, second, expectsRejection);
            string firstStatus = ResolveExternalStatus(first, expectsRejection);
            string secondStatus = ResolveExternalStatus(second, expectsRejection);
            await File.WriteAllTextAsync(outputPath, first.Markdown, cancellationToken).ConfigureAwait(false);
            IReadOnlyList<ReaderComparisonProbeResult> probes = ReaderComparisonScorer.ScoreMarkdown(
                first.Markdown,
                item.Probes,
                first.Rejected);
            results.Add(BuildCaseResult(
                item,
                caseStatus,
                error,
                first.Markdown,
                !ReferenceEquals(first, second) &&
                    firstStatus == "success" && secondStatus == "success" &&
                    first.Rejected == second.Rejected &&
                    string.Equals(first.Markdown, second.Markdown, StringComparison.Ordinal),
                ReferenceEquals(first, second)
                    ? first.DurationMilliseconds
                    : (first.DurationMilliseconds + second.DurationMilliseconds) / 2d,
                null,
                MaxNullable(first.PeakWorkingSetBytes, second.PeakWorkingSetBytes),
                probes));
        }
        string status = results.All(item => item.Status == "unavailable") ? "unavailable" : "completed";
        return new ReaderComparisonToolResult {
            Tool = runner.Name,
            ExecutionMode = "external-process",
            Status = status,
            Cases = results
        };
    }

    internal static (string Status, string? Error) ResolveRepeatOutcome(
        ReaderComparisonProcessOutput first,
        ReaderComparisonProcessOutput second,
        bool expectsRejection = false) {
        string firstStatus = ResolveExternalStatus(first, expectsRejection);
        string secondStatus = ResolveExternalStatus(second, expectsRejection);
        string status = firstStatus == "success" && secondStatus != "success" ? secondStatus : firstStatus;
        string? error = firstStatus == "success" ? null : first.Error;
        if (!ReferenceEquals(first, second) && secondStatus != "success") {
            error = AppendError(error, "Repeat run " + secondStatus + ": " + (second.Error ?? "No error detail was provided."));
        }
        if (!ReferenceEquals(first, second) && expectsRejection && first.Rejected != second.Rejected) {
            status = "failed";
            error = AppendError(error, "Repeat run did not preserve the malformed-input rejection outcome.");
        }
        return (status, error);
    }

    internal static string ResolveExternalStatus(ReaderComparisonProcessOutput output, bool expectsRejection) =>
        expectsRejection && output.Rejected && output.Status == "failed" ? "success" : output.Status;

    internal static string GetRepeatOutputPath(string outputPath) {
        string extension = Path.GetExtension(outputPath);
        if (extension.Length == 0) return outputPath + ".repeat";
        return Path.Combine(
            Path.GetDirectoryName(outputPath) ?? string.Empty,
            Path.GetFileNameWithoutExtension(outputPath) + ".repeat" + extension);
    }

    private static bool ExpectsMalformedInputRejection(ReaderComparisonCase item) =>
        item.Probes.Any(probe => probe.Kind == ReaderComparisonProbeKind.RejectsMalformedInput);

    private static OfficeDocumentReader CreateReader() => new OfficeDocumentReaderBuilder()
        .AddCsvHandler()
        .AddEmailHandlers()
        .AddEpubHandler()
        .AddExcelHandler()
        .AddHtmlHandler()
        .AddMarkdownHandler()
        .AddPdfHandler()
        .AddPowerPointHandler()
        .AddWordHandler()
        .AddZipHandler()
        .Build();

    private static ReaderOptions ComparisonReaderOptions() => new ReaderOptions {
        ComputeHashes = false,
        MaxChars = 8_000,
        MaxTableRows = 500
    };

    private static string ToMarkdown(OfficeDocumentReadResult document) {
        if (!string.IsNullOrWhiteSpace(document.Markdown)) return Normalize(document.Markdown!);
        return Normalize(string.Join("\n\n", document.Chunks.Select(chunk => chunk.Markdown ?? chunk.Text)));
    }

    private static string Normalize(string markdown) =>
        markdown.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Trim() + "\n";

    private static bool HasErrorDiagnostic(OfficeDocumentReadResult document) =>
        document.Diagnostics.Any(diagnostic => string.Equals(
            diagnostic.Severity.ToString(),
            "Error",
            StringComparison.OrdinalIgnoreCase));

    private static ReaderComparisonCaseResult BuildCaseResult(
        ReaderComparisonCase item,
        string status,
        string? error,
        string markdown,
        bool deterministic,
        double durationMilliseconds,
        long? allocatedBytes,
        long? peakWorkingSetBytes,
        IReadOnlyList<ReaderComparisonProbeResult> probes) {
        int applied = probes.Count(probe => probe.Applied);
        int passed = probes.Count(probe => probe.Applied && probe.Passed);
        return new ReaderComparisonCaseResult {
            CaseId = item.Id,
            SourceName = item.SourceName,
            Status = status,
            Error = error,
            MarkdownSha256 = Hash(markdown),
            Deterministic = deterministic,
            DurationMilliseconds = durationMilliseconds,
            InputBytes = item.Bytes.LongLength,
            OutputBytes = Encoding.UTF8.GetByteCount(markdown),
            AllocatedBytes = allocatedBytes,
            PeakWorkingSetBytes = peakWorkingSetBytes,
            AppliedProbes = applied,
            PassedProbes = passed,
            Probes = probes
        };
    }

    private static async Task<ReaderComparisonConfiguration> LoadConfigurationAsync(
        string path,
        CancellationToken cancellationToken) {
        await using FileStream stream = File.OpenRead(path);
        ReaderComparisonConfiguration? configuration = await JsonSerializer.DeserializeAsync<ReaderComparisonConfiguration>(
            stream,
            JsonOptions,
            cancellationToken).ConfigureAwait(false);
        return configuration ?? throw new InvalidDataException("Runner configuration is empty.");
    }

    internal static string BuildMarkdownReport(ReaderComparisonReport report) {
        var builder = new StringBuilder();
        builder.AppendLine("# Reader extraction evidence").AppendLine();
        builder.Append("Generated: ").Append(report.CreatedUtc.ToString("O")).AppendLine("  ");
        builder.Append("Source commit: ").Append(report.SourceCommit).AppendLine("  ");
        builder.Append("Source tree dirty: ").Append(report.SourceTreeDirty ? "yes" : "no").AppendLine("  ");
        builder.Append("Runtime: ").Append(report.Runtime).AppendLine("  ");
        builder.Append("Operating system: ").Append(report.OperatingSystem).AppendLine().AppendLine();
        builder.AppendLine("Each runner is reported independently. Probe denominators are runner-specific: OfficeIMO-native tables, links, assets, and source locations do not apply to external Markdown runners. These sections are extraction evidence, not a performance leaderboard.").AppendLine();
        foreach (ReaderComparisonToolResult tool in report.Tools) {
            bool inProcess = string.Equals(tool.ExecutionMode, "in-process", StringComparison.Ordinal);
            builder.Append("## ").AppendLine(tool.Tool).AppendLine();
            builder.Append("Execution mode: ").Append(tool.ExecutionMode).AppendLine("  ");
            builder.Append("Runner status: ").Append(tool.Status).AppendLine().AppendLine();
            if (inProcess) {
                builder.AppendLine("| Case | Status | Passed / applicable probes | Deterministic | Diagnostic mean ms | Allocated bytes | Retained bytes | Managed peak growth | Process peak growth | Input bytes | Output bytes |");
                builder.AppendLine("| --- | --- | ---: | :---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |");
            } else {
                builder.AppendLine("| Case | Status | Passed / applicable probes | Deterministic | Diagnostic mean ms | Peak working set bytes | Input bytes | Output bytes |");
                builder.AppendLine("| --- | --- | ---: | :---: | ---: | ---: | ---: | ---: |");
            }
            foreach (ReaderComparisonCaseResult item in tool.Cases) {
                long? memory = inProcess ? item.AllocatedBytes : item.PeakWorkingSetBytes;
                builder.Append("| ").Append(Escape(item.CaseId))
                    .Append(" | ").Append(Escape(item.Status)).Append(" | ")
                    .Append(item.PassedProbes).Append('/').Append(item.AppliedProbes)
                    .Append(" | ").Append(item.Deterministic ? "yes" : "no")
                    .Append(" | ").Append(item.DurationMilliseconds.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" | ").Append(memory?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "n/a");
                if (inProcess) {
                    builder.Append(" | ").Append(item.RetainedManagedHeapGrowthBytes?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "n/a")
                        .Append(" | ").Append(item.PeakManagedHeapGrowthBytes?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "n/a")
                        .Append(" | ").Append(item.PeakWorkingSetGrowthBytes?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "n/a");
                }
                builder.Append(" | ").Append(item.InputBytes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" | ").Append(item.OutputBytes.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .AppendLine(" |");
            }
            builder.AppendLine();
        }
        builder.AppendLine("Duration and memory values describe each execution mode on this machine. OfficeIMO allocation is measured in process; its retained, managed-peak, and working-set-growth values come from an isolated child process per case. External-process peak working set is a different measurement and must not be compared with managed allocation. Output bytes measure normalized extracted Markdown, not a source-format round trip. Use BenchmarkDotNet lanes for release performance decisions.");
        return builder.ToString();
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string ResolveCommit() => RunGit("rev-parse", "HEAD") ?? "unknown";

    private static bool ResolveSourceTreeDirty() =>
        !string.IsNullOrWhiteSpace(RunGit("status", "--porcelain", "--untracked-files=no"));

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

    private static string SafeName(string value) {
        string safe = new string(value.ToLowerInvariant().Select(character =>
            char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');
        return string.IsNullOrWhiteSpace(safe) ? "runner" : safe;
    }

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);

    private static string AppendError(string? current, string addition) =>
        string.IsNullOrWhiteSpace(current) ? addition : current + " " + addition;

    private static long? MaxNullable(long? first, long? second) {
        if (!first.HasValue) return second;
        if (!second.HasValue) return first;
        return Math.Max(first.Value, second.Value);
    }

    private static void WriteUsage() {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run --project OfficeIMO.Reader.Benchmarks -c Release -f net8.0 -- evidence [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --output <path>         Report and corpus output directory.");
        Console.WriteLine("  --runners <json-path>   Optional direct-process runner configuration.");
        Console.WriteLine("  --help                  Show this help.");
    }

    private sealed class OfficeComparisonAttempt {
        internal OfficeComparisonAttempt(
            OfficeDocumentReadResult? document,
            string markdown,
            string? error,
            bool rejected,
            double durationMilliseconds,
            long allocatedBytes) {
            Document = document;
            Markdown = markdown;
            Error = error;
            Rejected = rejected;
            DurationMilliseconds = durationMilliseconds;
            AllocatedBytes = allocatedBytes;
        }

        internal OfficeDocumentReadResult? Document { get; }
        internal string Markdown { get; }
        internal string? Error { get; }
        internal bool Rejected { get; }
        internal double DurationMilliseconds { get; }
        internal long AllocatedBytes { get; }
    }

    private sealed class ComparisonCommandOptions {
        public string OutputDirectory { get; private set; } = Path.Combine("artifacts", "reader-evidence");
        public string? RunnerConfigPath { get; private set; }
        public bool ShowHelp { get; private set; }

        public static ComparisonCommandOptions Parse(string[] args) {
            var options = new ComparisonCommandOptions();
            for (int index = 0; index < args.Length; index++) {
                string argument = args[index];
                if (argument is "--help" or "-h") {
                    options.ShowHelp = true;
                } else if (argument == "--output") {
                    options.OutputDirectory = RequiredValue(args, ref index, argument);
                } else if (argument == "--runners") {
                    options.RunnerConfigPath = RequiredValue(args, ref index, argument);
                } else {
                    throw new ArgumentException("Unknown comparison option: " + argument);
                }
            }
            return options;
        }

        private static string RequiredValue(string[] args, ref int index, string option) {
            if (++index >= args.Length || string.IsNullOrWhiteSpace(args[index])) {
                throw new ArgumentException("Option " + option + " requires a value.");
            }
            return args[index];
        }
    }
}
