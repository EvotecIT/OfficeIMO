using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Zip;
using System.Diagnostics;
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

        var tools = new List<ReaderComparisonToolResult> {
            RunOfficeIMO(cases, outputsDirectory)
        };
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
            Runtime = RuntimeInformation.FrameworkDescription,
            OperatingSystem = RuntimeInformation.OSDescription,
            Tools = tools
        };
        string jsonPath = Path.Combine(outputDirectory, "reader-comparison.json");
        string markdownPath = Path.Combine(outputDirectory, "reader-comparison.md");
        await File.WriteAllTextAsync(jsonPath, JsonSerializer.Serialize(report, JsonOptions), cancellationToken)
            .ConfigureAwait(false);
        await File.WriteAllTextAsync(markdownPath, BuildMarkdownReport(report), cancellationToken)
            .ConfigureAwait(false);
        Console.WriteLine("Reader comparison report: " + markdownPath);
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
        return new ReaderComparisonToolResult { Tool = "OfficeIMO.Reader", Status = "success", Cases = results };
    }

    private static ReaderComparisonCaseResult RunOfficeCase(
        OfficeDocumentReader reader,
        ReaderComparisonCase item,
        string outputPath) {
        OfficeDocumentReadResult? first = null;
        string firstMarkdown = string.Empty;
        string secondMarkdown = string.Empty;
        string? error = null;
        bool rejected = false;
        long before = GC.GetAllocatedBytesForCurrentThread();
        var stopwatch = Stopwatch.StartNew();
        try {
            first = reader.ReadDocument(item.Bytes, item.SourceName, ComparisonReaderOptions());
            firstMarkdown = ToMarkdown(first);
            OfficeDocumentReadResult second = reader.ReadDocument(item.Bytes, item.SourceName, ComparisonReaderOptions());
            secondMarkdown = ToMarkdown(second);
            rejected = HasErrorDiagnostic(first);
        } catch (Exception ex) when (ex is InvalidDataException or NotSupportedException or FormatException) {
            error = ex.Message;
            rejected = true;
        }
        stopwatch.Stop();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        File.WriteAllText(outputPath, firstMarkdown, new UTF8Encoding(false));
        IReadOnlyList<ReaderComparisonProbeResult> probes = ReaderComparisonScorer.ScoreOfficeDocument(
            firstMarkdown,
            first,
            item.Probes,
            rejected);
        return BuildCaseResult(
            item,
            rejected && item.Probes.All(probe => probe.Kind != ReaderComparisonProbeKind.RejectsMalformedInput) ? "failed" : "success",
            error,
            firstMarkdown,
            string.Equals(firstMarkdown, secondMarkdown, StringComparison.Ordinal),
            stopwatch.Elapsed.TotalMilliseconds / 2d,
            allocated,
            null,
            probes);
    }

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
            ReaderComparisonProcessOutput second = first.Status == "success"
                ? await ReaderComparisonProcessRunner.RunAsync(runner, inputPath, outputPath + ".repeat", cancellationToken)
                    .ConfigureAwait(false)
                : first;
            await File.WriteAllTextAsync(outputPath, first.Markdown, cancellationToken).ConfigureAwait(false);
            IReadOnlyList<ReaderComparisonProbeResult> probes = ReaderComparisonScorer.ScoreMarkdown(
                first.Markdown,
                item.Probes,
                first.Rejected);
            results.Add(BuildCaseResult(
                item,
                first.Status,
                first.Error,
                first.Markdown,
                first.Status == "success" && second.Status == "success" &&
                    string.Equals(first.Markdown, second.Markdown, StringComparison.Ordinal),
                (first.DurationMilliseconds + second.DurationMilliseconds) / (ReferenceEquals(first, second) ? 1d : 2d),
                null,
                first.PeakWorkingSetBytes,
                probes));
        }
        string status = results.All(item => item.Status == "unavailable") ? "unavailable" : "completed";
        return new ReaderComparisonToolResult { Tool = runner.Name, Status = status, Cases = results };
    }

    private static OfficeDocumentReader CreateReader() => new OfficeDocumentReaderBuilder()
        .AddCsvHandler()
        .AddEpubHandler()
        .AddHtmlHandler()
        .AddPdfHandler()
        .AddZipHandler()
        .Build();

    private static ReaderOptions ComparisonReaderOptions() => new ReaderOptions {
        ComputeHashes = false,
        MaxChars = 8_000,
        MaxTableRows = 500,
        IncludePowerPointNotes = true,
        IncludeWordFootnotes = true
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

    private static string BuildMarkdownReport(ReaderComparisonReport report) {
        var builder = new StringBuilder();
        builder.AppendLine("# Reader comparison evidence").AppendLine();
        builder.Append("Generated: ").Append(report.CreatedUtc.ToString("O")).AppendLine("  ");
        builder.Append("Runtime: ").Append(report.Runtime).AppendLine("  ");
        builder.Append("Operating system: ").Append(report.OperatingSystem).AppendLine().AppendLine();
        builder.AppendLine("Scores count only probes applicable to that tool. OfficeIMO additionally reports rich tables, links, assets, and source locations; Markdown-only tools are not penalized for those native-result probes.").AppendLine();
        builder.AppendLine("| Tool | Case | Status | Semantic score | Deterministic | Mean ms | Allocated / peak bytes |");
        builder.AppendLine("| --- | --- | --- | ---: | :---: | ---: | ---: |");
        foreach (ReaderComparisonToolResult tool in report.Tools) {
            foreach (ReaderComparisonCaseResult item in tool.Cases) {
                long? memory = item.AllocatedBytes ?? item.PeakWorkingSetBytes;
                builder.Append("| ").Append(Escape(tool.Tool)).Append(" | ").Append(Escape(item.CaseId))
                    .Append(" | ").Append(Escape(item.Status)).Append(" | ")
                    .Append(item.PassedProbes).Append('/').Append(item.AppliedProbes)
                    .Append(" | ").Append(item.Deterministic ? "yes" : "no")
                    .Append(" | ").Append(item.DurationMilliseconds.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture))
                    .Append(" | ").Append(memory?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "n/a")
                    .AppendLine(" |");
            }
        }
        builder.AppendLine().AppendLine("Runtime and memory values are diagnostic evidence from this machine, not cross-machine benchmark claims. Use BenchmarkDotNet lanes for release performance decisions.");
        return builder.ToString();
    }

    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string SafeName(string value) {
        string safe = new string(value.ToLowerInvariant().Select(character =>
            char.IsLetterOrDigit(character) ? character : '-').ToArray()).Trim('-');
        return string.IsNullOrWhiteSpace(safe) ? "runner" : safe;
    }

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);

    private static void WriteUsage() {
        Console.WriteLine("Usage:");
        Console.WriteLine("  dotnet run --project OfficeIMO.Reader.Benchmarks -c Release -f net8.0 -- compare [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --output <path>         Report and corpus output directory.");
        Console.WriteLine("  --runners <json-path>   Optional direct-process runner configuration.");
        Console.WriteLine("  --help                  Show this help.");
    }

    private sealed class ComparisonCommandOptions {
        public string OutputDirectory { get; private set; } = Path.Combine("artifacts", "reader-comparison");
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