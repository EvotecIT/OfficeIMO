using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures each declared H10 PDF intent and editable target in a fresh child process.
/// Conversion time and allocations come from the owner runner; process-tree peak includes startup and conversion.
/// </summary>
internal static class H10BudgetRunner {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static async Task<int> RunAsync(string[] args) {
        Dictionary<string, string> options = ParseOptions(args);
        if (options.ContainsKey("help")) {
            Console.WriteLine("--case <predeclared-id> --mhtml <archive> --output <new-directory> --expected-sha256 <hash> [--require-clean-source] [--ceilings <json>]");
            return 0;
        }
        string caseId = Required(options, "case");
        string archivePath = Path.GetFullPath(Required(options, "mhtml"));
        string outputPath = Path.GetFullPath(Required(options, "output"));
        string expectedSha = Required(options, "expected-sha256");
        if (Directory.Exists(outputPath) || File.Exists(outputPath))
            throw new IOException("The evidence output directory must be new.");
        if (!File.Exists(archivePath)) throw new FileNotFoundException("Frozen MHTML was not found.", archivePath);
        string repoRoot = FindRepositoryRoot();
        string manifestPath = Path.Combine(repoRoot, "OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json");
        string manifestSha = await HashFileAsync(manifestPath).ConfigureAwait(false);
        using JsonDocument manifest = JsonDocument.Parse(await File.ReadAllTextAsync(manifestPath).ConfigureAwait(false));
        JsonElement page = manifest.RootElement.GetProperty("pages").EnumerateArray().FirstOrDefault(item =>
            item.GetProperty("id").GetString() == caseId);
        if (page.ValueKind == JsonValueKind.Undefined ||
            !page.TryGetProperty("pdfIntents", out JsonElement intents) ||
            !page.TryGetProperty("editableTargets", out JsonElement targets))
            throw new InvalidDataException("The case must predeclare PDF intents and editable targets.");
        string[] pdfIntents = intents.EnumerateArray().Select(item => item.GetProperty("intent").GetString()!).ToArray();
        string[] editableTargets = targets.EnumerateArray().Select(item => item.GetProperty("target").GetString()!.ToLowerInvariant()).ToArray();
        if (pdfIntents.Length == 0 || editableTargets.Length == 0 ||
            pdfIntents.Distinct(StringComparer.Ordinal).Count() != pdfIntents.Length ||
            editableTargets.Distinct(StringComparer.Ordinal).Count() != editableTargets.Length ||
            pdfIntents.Any(item => item is not ("print-reflow" or "screen-media-pagination" or "screen-snapshot-pagination")) ||
            editableTargets.Any(item => item is not ("word" or "excel" or "powerpoint" or "onenote" or "rtf" or "markdown")))
            throw new InvalidDataException("The case declares an unsupported or duplicate operation.");

        long maximumBytes = manifest.RootElement.GetProperty("capture").GetProperty("maximumArchiveBytes").GetInt64();
        long actualBytes = new FileInfo(archivePath).Length;
        if (actualBytes == 0 || actualBytes > maximumBytes) throw new InvalidDataException("The archive exceeds the corpus input bound or is empty.");
        string archiveSha = await HashFileAsync(archivePath).ConfigureAwait(false);
        if (!archiveSha.Equals(expectedSha, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The archive SHA-256 does not match the predeclared input.");
        GitSourceState? source = await SourceProvenanceReader.ReadGitStateAsync(repoRoot).ConfigureAwait(false);
        string? runnerVersion = typeof(H10BudgetRunner).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (options.ContainsKey("require-clean-source") &&
            (source == null || !source.IsClean || runnerVersion == null ||
             !runnerVersion.EndsWith("+" + source.Commit, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("Budget evidence requires clean source and a runner rebuilt from its exact head.");

        string pdfRunner = Path.Combine(repoRoot, "OfficeIMO.Pdf.Benchmarks.Comparisons/bin/Release/net10.0/OfficeIMO.Pdf.Benchmarks.Comparisons.dll");
        string editableRunner = Path.Combine(repoRoot, "Build/HtmlEditableEvidence/bin/Release/net10.0/OfficeIMO.Html.EditableEvidence.dll");
        if (!File.Exists(pdfRunner) || !File.Exists(editableRunner))
            throw new FileNotFoundException("Build both Release evidence runners before measuring H10.");
        Directory.CreateDirectory(outputPath);
        var operations = new List<H10BudgetOperation>();
        var failures = new List<string>();
        foreach (string intent in pdfIntents) {
            string isolatedIntent = intent switch {
                "print-reflow" => "print",
                "screen-media-pagination" => "screen-media",
                _ => "screen-snapshot"
            };
            string directory = Path.Combine(outputPath, "pdf-" + isolatedIntent);
            var arguments = new List<string> {
                "html-mhtml-evidence", "--mhtml", archivePath, "--output", directory,
                "--isolated-officeimo-intent", isolatedIntent
            };
            if (options.ContainsKey("require-clean-source")) arguments.Add("--require-clean-source");
            await MeasureAsync("pdf-" + isolatedIntent, "pdf", pdfRunner, arguments,
                Path.Combine(directory, "html-mhtml-evidence.json"), operations, failures, repoRoot).ConfigureAwait(false);
        }
        foreach (string target in editableTargets) {
            string directory = Path.Combine(outputPath, "editable-" + target);
            var arguments = new List<string> {
                "--case", caseId, "--mhtml", archivePath, "--output", directory,
                "--target", target, "--expected-sha256", archiveSha
            };
            if (options.ContainsKey("require-clean-source")) arguments.Add("--require-clean-source");
            await MeasureAsync("editable-" + target, "editable", editableRunner, arguments,
                Path.Combine(directory, target, "report.json"), operations, failures, repoRoot).ConfigureAwait(false);
        }

        H10BudgetPlatform? ceiling = null;
        if (options.TryGetValue("ceilings", out string? ceilingPath)) {
            var configuration = JsonSerializer.Deserialize<H10BudgetConfiguration>(
                await File.ReadAllTextAsync(Path.GetFullPath(ceilingPath)).ConfigureAwait(false), JsonOptions)
                ?? throw new InvalidDataException("H10 budget configuration is empty.");
            if (configuration.SchemaVersion != 1 ||
                !configuration.SourceSha256.Equals(archiveSha, StringComparison.OrdinalIgnoreCase) ||
                !configuration.ManifestSha256.Equals(manifestSha, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("H10 budget configuration does not match the frozen source and selection manifest.");
            ceiling = configuration.Platforms.Single(item => item.CaseId == caseId && item.OsFamily == OsFamily());
            EvaluateCeilings(ceiling, operations, failures);
        }
        var report = new {
            schemaVersion = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            caseId,
            sourceUrl = page.GetProperty("url").GetString(),
            sourceCommit = source?.Commit,
            sourceClean = source?.IsClean,
            runnerVersion,
            manifestSha256 = manifestSha,
            archivePath,
            archiveBytes = actualBytes,
            archiveSha256 = archiveSha,
            osFamily = OsFamily(),
            osDescription = RuntimeInformation.OSDescription,
            framework = RuntimeInformation.FrameworkDescription,
            processorCount = Environment.ProcessorCount,
            machineName = Environment.MachineName,
            memoryScope = "fresh child process tree per declared operation, including startup and conversion",
            conversionScope = "layout and PDF conversion of preloaded MHTML for PDF; MHTML load, visible-content preparation, save and reopen for editable",
            ceiling,
            operations,
            failures
        };
        string reportPath = Path.Combine(outputPath, "h10-budget.json");
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, JsonOptions)).ConfigureAwait(false);
        Console.WriteLine("H10_BUDGET_REPORT=" + reportPath);
        Console.WriteLine("H10_BUDGET_STATUS=" + (failures.Count > 0 ? "Failed" : ceiling == null ? "Measured" : "Passed"));
        foreach (string failure in failures) Console.Error.WriteLine(failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static async Task MeasureAsync(string name, string kind, string runner, List<string> arguments,
        string reportPath, ICollection<H10BudgetOperation> operations,
        ICollection<string> failures, string repoRoot) {
        var start = new ProcessStartInfo("dotnet") {
            WorkingDirectory = repoRoot,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        start.ArgumentList.Add(runner);
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Could not start an H10 evidence worker.");
        var timer = Stopwatch.StartNew();
        Task<string> stdout = process.StandardOutput.ReadToEndAsync();
        Task<string> stderr = process.StandardError.ReadToEndAsync();
        await using var memory = new ProcessTreeMemorySampler(process);
        try {
            using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(10));
            await process.WaitForExitAsync(timeout.Token).ConfigureAwait(false);
        } catch (OperationCanceledException) {
            try {
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync().ConfigureAwait(false);
            } catch (InvalidOperationException) {
                // The worker exited while the timeout path took ownership.
            }
            failures.Add(name + ": isolated evidence worker exceeded ten minutes.");
        }
        timer.Stop();
        await memory.StopAsync().ConfigureAwait(false);
        string workerOutput = await stdout.ConfigureAwait(false);
        string workerError = await stderr.ConfigureAwait(false);
        if (process.ExitCode != 0 || !File.Exists(reportPath)) {
            failures.Add(name + ": evidence worker failed: " + (workerError.Length > 0 ? workerError : workerOutput));
            return;
        }
        try {
            using JsonDocument report = JsonDocument.Parse(await File.ReadAllTextAsync(reportPath).ConfigureAwait(false));
            JsonElement root = report.RootElement;
            bool passed;
            bool hasLoss;
            double elapsed;
            long allocated;
            long outputBytes;
            string? outputSha;
            int? pages;
            if (kind == "pdf") {
                JsonElement result = root.GetProperty("operations").EnumerateArray().Single();
                if (root.GetProperty("failures").GetArrayLength() != 0 ||
                    result.GetProperty("intent").GetString() != "officeimo-" + name[4..])
                    throw new InvalidDataException("The isolated PDF worker returned an unexpected operation.");
                passed = true;
                hasLoss = result.GetProperty("conversionReport").GetProperty("hasLoss").GetBoolean();
                elapsed = result.GetProperty("elapsedMilliseconds").GetDouble();
                allocated = result.GetProperty("managedAllocatedBytes").GetInt64();
                outputBytes = result.GetProperty("bytes").GetInt64();
                outputSha = result.GetProperty("sha256").GetString();
                pages = result.GetProperty("pageCount").GetInt32();
            } else {
                passed = root.GetProperty("Passed").GetBoolean();
                hasLoss = root.GetProperty("ReportHasLoss").GetBoolean();
                elapsed = root.GetProperty("ElapsedMs").GetDouble();
                allocated = root.GetProperty("AllocatedBytes").GetInt64();
                string artifact = root.GetProperty("Artifact").GetString()!;
                outputBytes = root.GetProperty("ArtifactBytes").GetInt64();
                outputSha = await HashFileAsync(artifact).ConfigureAwait(false);
                pages = null;
            }
            HtmlPdfProcessTreeMemoryEvidence sample = memory.CreateEvidence();
            if (!passed || sample.SampleCount == 0 || sample.PeakWorkingSetBytes <= 0)
                throw new InvalidDataException("The worker did not pass its artifact contract or memory could not be observed.");
            operations.Add(new H10BudgetOperation(name, kind, passed, hasLoss, elapsed, allocated,
                timer.Elapsed.TotalMilliseconds, sample, outputBytes, outputSha, pages, reportPath));
            Console.WriteLine($"{name}: {elapsed:0.0} ms conversion, {allocated} allocated bytes, {sample.PeakWorkingSetBytes} peak bytes, loss={hasLoss}");
        } catch (Exception exception) {
            failures.Add(name + ": could not validate evidence report: " + exception.Message);
        }
    }

    private static void EvaluateCeilings(H10BudgetPlatform ceiling,
        IReadOnlyList<H10BudgetOperation> operations, ICollection<string> failures) {
        if (ceiling.Operations.Count != operations.Count ||
            ceiling.Operations.Select(item => item.Name).Order().SequenceEqual(operations.Select(item => item.Name).Order()) == false)
            throw new InvalidDataException("The platform budget does not cover the exact declared operations.");
        foreach (H10BudgetOperation operation in operations) {
            H10BudgetLimit limit = ceiling.Operations.Single(item => item.Name == operation.Name);
            if (limit.ConversionMilliseconds <= 0 || limit.ManagedAllocatedBytes <= 0 || limit.PeakWorkingSetBytes <= 0)
                throw new InvalidDataException("H10 budget ceilings must be positive.");
            if (operation.ConversionMilliseconds > limit.ConversionMilliseconds)
                failures.Add(operation.Name + ": conversion time exceeded the platform ceiling.");
            if (operation.ManagedAllocatedBytes > limit.ManagedAllocatedBytes)
                failures.Add(operation.Name + ": allocations exceeded the platform ceiling.");
            if (operation.ProcessTreeMemory.PeakWorkingSetBytes > limit.PeakWorkingSetBytes)
                failures.Add(operation.Name + ": process-tree peak memory exceeded the platform ceiling.");
        }
    }

    private static Dictionary<string, string> ParseOptions(string[] args) {
        var options = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < args.Length; index++) {
            if (!args[index].StartsWith("--", StringComparison.Ordinal)) throw new ArgumentException("Expected a named option.");
            string key = args[index][2..];
            if (key is "help" or "require-clean-source") {
                if (!options.TryAdd(key, "true")) throw new ArgumentException("Duplicate option: " + key);
            } else if (key is "case" or "mhtml" or "output" or "expected-sha256" or "ceilings") {
                if (++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal) ||
                    !options.TryAdd(key, args[index])) throw new ArgumentException("Missing or duplicate option: " + key);
            } else throw new ArgumentException("Unknown option: " + key);
        }
        return options;
    }

    private static string Required(IReadOnlyDictionary<string, string> options, string name) =>
        options.TryGetValue(name, out string? value) ? value : throw new ArgumentException("Missing --" + name);

    private static string FindRepositoryRoot() {
        for (DirectoryInfo? directory = new(Directory.GetCurrentDirectory()); directory != null; directory = directory.Parent)
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) return directory.FullName;
        throw new DirectoryNotFoundException("OfficeIMO repository root was not found.");
    }

    private static string OsFamily() => OperatingSystem.IsWindows() ? "Windows" : OperatingSystem.IsLinux() ? "Linux" : OperatingSystem.IsMacOS() ? "macOS" : "Unknown";

    private static async Task<string> HashFileAsync(string path) {
        await using FileStream stream = File.OpenRead(path);
        return Convert.ToHexString(await SHA256.HashDataAsync(stream).ConfigureAwait(false)).ToLowerInvariant();
    }
}
