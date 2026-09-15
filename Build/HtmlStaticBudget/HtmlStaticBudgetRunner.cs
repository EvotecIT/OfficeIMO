using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Tests;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlStaticBudgetRunner {
    private static readonly TimeSpan WorkerTimeout = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan CancellationStartTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan CancellationPromptTimeout = TimeSpan.FromSeconds(3);
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    internal static async Task<int> RunAsync(string[] args) {
        ValidateArguments(args);
        if (args.Any(value => string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase))) {
            WriteHelp();
            return 0;
        }
        bool measureOnly = HasFlag(args, "--measure-only");
        int iterations = ReadIterations(args);
        string outputDirectory = ResolveOutputDirectory(args);
        using EvidenceOutputReservation reservation = EvidenceOutputReservation.Acquire(outputDirectory);
        string repositoryRoot = FindRepositoryRoot();
        GitSourceState? git = await SourceProvenanceReader.ReadGitStateAsync(repositoryRoot).ConfigureAwait(false);
        if (HasFlag(args, "--require-clean-source") && (git == null || !git.IsClean)) {
            throw new InvalidOperationException("Clean, commit-addressable OfficeIMO source is required for budget evidence.");
        }

        HtmlRenderingHeldOutCorpus corpus = HtmlRenderingHeldOutCorpus.Load();
        HtmlStaticBudgetConfiguration configuration = await ReadConfigurationAsync(corpus.RootPath).ConfigureAwait(false);
        HtmlStaticBudgetCeiling? ceiling = measureOnly ? null : ResolveCeiling(configuration, GetOsFamily(), corpus, iterations);
        HtmlStaticBudgetRun cold = await RunWorkerAsync("cold", 1, outputDirectory, repositoryRoot).ConfigureAwait(false);
        HtmlStaticBudgetRun warm = await RunWorkerAsync("warm", iterations, outputDirectory, repositoryRoot).ConfigureAwait(false);
        HtmlStaticBudgetCancellation cancellation = await ProbeCancellationAsync(corpus.Cases[0].Html).ConfigureAwait(false);

        int outputsPerIteration = cold.Iterations.Single().OutputCount;
        var source = new HtmlStaticBudgetSource(
            git?.Commit, git?.IsClean, corpus.Manifest.CorpusId, corpus.ManifestSha256,
            corpus.Cases.Count, outputsPerIteration);
        var failures = ceiling == null
            ? new List<string>()
            : Evaluate(ceiling, cold, warm, cancellation);
        var report = new HtmlStaticBudgetReport(
            1,
            DateTimeOffset.UtcNow,
            new HtmlStaticBudgetEnvironment(
                GetOsFamily(), RuntimeInformation.OSDescription, RuntimeInformation.OSArchitecture.ToString(),
                RuntimeInformation.ProcessArchitecture.ToString(), RuntimeInformation.FrameworkDescription,
                Environment.MachineName, Environment.ProcessorCount),
            source, ceiling, cold, warm, cancellation, failures);
        string reportPath = Path.Combine(outputDirectory, "html-static-budget.json");
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, JsonOptions)).ConfigureAwait(false);
        Console.WriteLine("HTML_STATIC_BUDGET_REPORT=" + reportPath);
        Console.WriteLine("HTML_STATIC_BUDGET_PLATFORM=" + report.Environment.OsFamily);
        Console.WriteLine("HTML_STATIC_BUDGET_COLD_MS=" + cold.ProcessElapsedMilliseconds.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture));
        Console.WriteLine("HTML_STATIC_BUDGET_WARM_MAX_MS=" + warm.Iterations.Max(item => item.ElapsedMilliseconds).ToString("0.0", System.Globalization.CultureInfo.InvariantCulture));
        Console.WriteLine("HTML_STATIC_BUDGET_PEAK_BYTES=" + Math.Max(cold.PeakWorkingSetBytes, warm.PeakWorkingSetBytes));
        Console.WriteLine("HTML_STATIC_BUDGET_COLD_WARM_DETERMINISTIC=" + SameFingerprintAcrossRuns(cold, warm));
        Console.WriteLine("HTML_STATIC_BUDGET_STATUS=" + (failures.Count == 0 ? (measureOnly ? "Measured" : "Passed") : "Failed"));
        foreach (string failure in failures) Console.Error.WriteLine(failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static async Task<HtmlStaticBudgetRun> RunWorkerAsync(
        string mode,
        int iterations,
        string outputDirectory,
        string repositoryRoot) {
        string resultPath = Path.Combine(outputDirectory, ".worker-" + mode + ".json");
        ProcessStartInfo startInfo = HtmlStaticBudgetWorker.CreateStartInfo(mode, iterations, resultPath, repositoryRoot);
        var processStopwatch = Stopwatch.StartNew();
        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Could not start the isolated HTML static budget worker.");
        Task<string> standardOutput = process.StandardOutput.ReadToEndAsync();
        Task<string> standardError = process.StandardError.ReadToEndAsync();
        await using var memory = new ProcessTreeMemorySampler(process);
        try {
            using var timeout = new CancellationTokenSource(WorkerTimeout);
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            try {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync().ConfigureAwait(false);
            } catch (InvalidOperationException) {
                // The worker exited while the timeout path took ownership.
            }
            throw new TimeoutException("The isolated HTML static budget worker exceeded ten minutes.");
        }
        processStopwatch.Stop();
        await memory.StopAsync().ConfigureAwait(false);
        string workerOutput = await standardOutput.ConfigureAwait(false);
        string workerError = await standardError.ConfigureAwait(false);
        if (process.ExitCode != 0 || !File.Exists(resultPath)) {
            throw new InvalidOperationException(
                $"The isolated HTML static budget worker exited with code {process.ExitCode}. " +
                (workerError.Length > 0 ? workerError : workerOutput));
        }
        IReadOnlyList<HtmlStaticBudgetIteration> results = await HtmlStaticBudgetWorker.ReadResultAsync(resultPath).ConfigureAwait(false);
        File.Delete(resultPath);
        HtmlPdfProcessTreeMemoryEvidence memoryEvidence = memory.CreateEvidence();
        return new HtmlStaticBudgetRun(
            mode, processStopwatch.Elapsed.TotalMilliseconds, memoryEvidence.PeakWorkingSetBytes,
            memoryEvidence.SampleCount, results,
            results.Select(item => item.FingerprintSha256).Distinct(StringComparer.Ordinal).Count() == 1);
    }

    private static async Task<HtmlStaticBudgetCancellation> ProbeCancellationAsync(string html) {
        var started = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        PdfResourcePolicy resourcePolicy = PdfResourcePolicy.CreateDefault();
        resourcePolicy.AllowRemoteResourceResolution = true;
        var options = new HtmlToPdfOptions {
            ResourcePolicy = resourcePolicy,
            ResourceResolver = async (_, cancellationToken) => {
                started.TrySetResult(true);
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ConfigureAwait(false);
                return null;
            }
        };
        int bodyEnd = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        const string probe = "<img src=\"https://officeimo.invalid/budget-cancellation.png\" alt=\"probe\">";
        string input = bodyEnd >= 0 ? html.Insert(bodyEnd, probe) : html + probe;
        Task<byte[]> conversion = HtmlConversionDocument.Parse(input).ToPdfBytesAsync(options, cancellation.Token);
        try {
            await started.Task.WaitAsync(CancellationStartTimeout).ConfigureAwait(false);
            var stopwatch = Stopwatch.StartNew();
            cancellation.Cancel();
            try {
                _ = await conversion.WaitAsync(CancellationPromptTimeout).ConfigureAwait(false);
                return new HtmlStaticBudgetCancellation("Failed", stopwatch.Elapsed.TotalMilliseconds, "The in-flight conversion completed instead of cancelling.");
            } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
                stopwatch.Stop();
                return new HtmlStaticBudgetCancellation("Passed", stopwatch.Elapsed.TotalMilliseconds, "Cancellation propagated through the asynchronous resource boundary.");
            } catch (TimeoutException) {
                return new HtmlStaticBudgetCancellation("Failed", stopwatch.Elapsed.TotalMilliseconds, "The in-flight conversion did not cancel promptly.");
            }
        } catch (TimeoutException) {
            cancellation.Cancel();
            return new HtmlStaticBudgetCancellation("Failed", CancellationStartTimeout.TotalMilliseconds, "The conversion did not reach the observable resource boundary.");
        }
    }

    private static List<string> Evaluate(
        HtmlStaticBudgetCeiling ceiling,
        HtmlStaticBudgetRun cold,
        HtmlStaticBudgetRun warm,
        HtmlStaticBudgetCancellation cancellation) {
        var failures = new List<string>();
        HtmlStaticBudgetIteration coldIteration = cold.Iterations.Single();
        Check(cold.ProcessElapsedMilliseconds, ceiling.ColdProcessElapsedMilliseconds, "cold process elapsed milliseconds", failures);
        Check(coldIteration.ManagedAllocatedBytes, ceiling.ColdManagedAllocatedBytes, "cold managed allocated bytes", failures);
        foreach (HtmlStaticBudgetIteration item in warm.Iterations) {
            Check(item.ElapsedMilliseconds, ceiling.WarmIterationElapsedMilliseconds, $"warm iteration {item.Number} elapsed milliseconds", failures);
            Check(item.ManagedAllocatedBytes, ceiling.WarmManagedAllocatedBytes, $"warm iteration {item.Number} managed allocated bytes", failures);
            Check(item.OutputBytes, ceiling.OutputBytesPerIteration, $"warm iteration {item.Number} output bytes", failures);
        }
        Check(coldIteration.OutputBytes, ceiling.OutputBytesPerIteration, "cold output bytes", failures);
        Check(Math.Max(cold.PeakWorkingSetBytes, warm.PeakWorkingSetBytes), ceiling.PeakWorkingSetBytes, "peak working-set bytes", failures);
        if (!cold.Deterministic || !warm.Deterministic) failures.Add("Output fingerprints were not deterministic within a worker run.");
        if (!SameFingerprintAcrossRuns(cold, warm)) failures.Add("The cold output fingerprint differed from one or more warm output fingerprints.");
        if (!string.Equals(cancellation.Status, "Passed", StringComparison.Ordinal)) failures.Add("Cancellation probe failed: " + cancellation.Detail);
        Check(cancellation.ElapsedMilliseconds, ceiling.CancellationElapsedMilliseconds, "cancellation elapsed milliseconds", failures);
        return failures;
    }

    private static bool SameFingerprintAcrossRuns(HtmlStaticBudgetRun cold, HtmlStaticBudgetRun warm) {
        string coldFingerprint = cold.Iterations.Single().FingerprintSha256;
        return warm.Iterations.All(iteration => string.Equals(
            coldFingerprint,
            iteration.FingerprintSha256,
            StringComparison.Ordinal));
    }

    private static void Check(double actual, double maximum, string name, ICollection<string> failures) {
        if (actual > maximum) failures.Add($"{name} {actual:0.0} exceeded ceiling {maximum:0.0}.");
    }

    private static async Task<HtmlStaticBudgetConfiguration> ReadConfigurationAsync(string corpusRoot) {
        string path = Path.Combine(corpusRoot, "budgets.json");
        await using FileStream stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<HtmlStaticBudgetConfiguration>(stream, JsonOptions).ConfigureAwait(false)
            ?? throw new InvalidDataException("The HTML static budget configuration was empty.");
    }

    private static HtmlStaticBudgetCeiling ResolveCeiling(
        HtmlStaticBudgetConfiguration configuration,
        string osFamily,
        HtmlRenderingHeldOutCorpus corpus,
        int iterations) {
        if (configuration.SchemaVersion != 1 || configuration.CorpusId != corpus.Manifest.CorpusId
            || configuration.ManifestSha256 != corpus.ManifestSha256 || configuration.WarmIterations != iterations) {
            throw new InvalidDataException("The HTML static budget configuration does not match this corpus or execution policy.");
        }
        return configuration.Platforms.Single(item => string.Equals(item.OsFamily, osFamily, StringComparison.OrdinalIgnoreCase));
    }

    private static int ReadIterations(string[] args) {
        string? value = ReadOption(args, "--iterations");
        if (value == null) return 3;
        if (int.TryParse(value, out int iterations) && iterations is >= 2 and <= 10) return iterations;
        throw new ArgumentException("--iterations must be between 2 and 10.");
    }

    private static string ResolveOutputDirectory(string[] args) {
        string? configured = ReadOption(args, "--output");
        if (!string.IsNullOrWhiteSpace(configured)) return Path.GetFullPath(configured);
        string run = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss.fff", System.Globalization.CultureInfo.InvariantCulture)
            + "-" + Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return Path.Combine(Path.GetTempPath(), "OfficeIMO", "HtmlStaticBudget", run);
    }

    private static void ValidateArguments(string[] args) {
        for (int index = 1; index < args.Length; index++) {
            string argument = args[index];
            if (argument.Equals("--help", StringComparison.OrdinalIgnoreCase)
                || argument.Equals("--measure-only", StringComparison.OrdinalIgnoreCase)
                || argument.Equals("--require-clean-source", StringComparison.OrdinalIgnoreCase)) continue;
            if (argument.Equals("--output", StringComparison.OrdinalIgnoreCase)
                || argument.Equals("--iterations", StringComparison.OrdinalIgnoreCase)) {
                if (++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException(argument + " requires a value.");
                continue;
            }
            throw new ArgumentException("Unknown html-static-budget option: " + argument);
        }
    }

    private static bool HasFlag(string[] args, string flag) => args.Any(value => value.Equals(flag, StringComparison.OrdinalIgnoreCase));
    private static string? ReadOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) if (args[index].Equals(option, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        return null;
    }
    private static string GetOsFamily() => OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsLinux() ? "Linux" : OperatingSystem.IsMacOS() ? "macOS" : "Unknown";
    private static string FindRepositoryRoot() {
        for (string? current = Path.GetFullPath(Directory.GetCurrentDirectory()); !string.IsNullOrWhiteSpace(current); current = Directory.GetParent(current)?.FullName) {
            if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }
    private static void WriteHelp() {
        Console.WriteLine("html-static-budget [--output <new-directory>] [--iterations 2-10] [--measure-only] [--require-clean-source]");
        Console.WriteLine("Measures the frozen H4/v2 corpus in fresh cold and warmed worker processes, including all declared static output contracts, allocations, process-tree peak memory, output bytes, determinism, and cancellation.");
    }
}
