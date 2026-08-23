using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using HtmlTinkerX;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlPdfEvidenceRunner {
    private const int MinimumIterations = 2;
    private const int MaximumIterations = 10;
    private static readonly TimeSpan WorkerTimeout = TimeSpan.FromMinutes(5);

    internal static async Task<int> RunAsync(string[] args) {
        ValidateArguments(args);
        if (args.Any(value => string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase))) {
            WriteHelp();
            return 0;
        }

        PdfBenchmarkScale scale = ReadScale(args);
        int iterations = ReadIterations(args);
        string outputDirectory = ResolveOutputDirectory(args);
        using EvidenceOutputReservation outputReservation = EvidenceOutputReservation.Acquire(outputDirectory);
        string repositoryRoot = FindRepositoryRoot();
        HtmlPdfEvidenceProvenance provenance = await ReadProvenanceAsync(repositoryRoot).ConfigureAwait(false);
        if (args.Any(value => string.Equals(value, "--require-clean-source", StringComparison.OrdinalIgnoreCase))) {
            ValidateCleanSource(provenance);
        }
        ExternalPdfRasterizer? rasterizer = await ExternalPdfRasterizer.FindAsync().ConfigureAwait(false);
        if (rasterizer == null && args.Any(value => string.Equals(value, "--require-external-rasterizer", StringComparison.OrdinalIgnoreCase))) {
            throw new InvalidOperationException("--require-external-rasterizer was specified, but pdftoppm was not found on PATH.");
        }

        PdfBenchmarkScenario scenario = PdfBenchmarkScenario.Get(scale);
        string html = PdfHtmlScenarioBuilder.Create(scenario);
        string htmlPath = Path.Combine(outputDirectory, "scenario.html");
        await File.WriteAllTextAsync(htmlPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)).ConfigureAwait(false);

        var engineReports = new List<HtmlPdfEngineEvidence>();
        foreach (HtmlPdfComparisonEngine engine in HtmlPdfComparisonRenderers.AllEngines) {
            engineReports.Add(await RunEngineAsync(engine, html, scenario, iterations, outputDirectory, rasterizer).ConfigureAwait(false));
        }

        var report = new HtmlPdfEvidenceReport(
            SchemaVersion: 2,
            GeneratedUtc: DateTimeOffset.UtcNow,
            Scale: scale.ToString(),
            Iterations: iterations,
            Environment: new HtmlPdfEvidenceEnvironment(
                RuntimeInformation.OSDescription,
                RuntimeInformation.OSArchitecture.ToString(),
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.Version.ToString(),
                RuntimeInformation.FrameworkDescription,
                rasterizer?.Identity),
            Provenance: provenance,
            Input: new HtmlPdfEvidenceInput(
                Path.GetFileName(htmlPath),
                Encoding.UTF8.GetByteCount(html),
                Sha256(Encoding.UTF8.GetBytes(html)),
                scenario.PageCount,
                scenario.PageCount),
            Engines: engineReports);

        string reportPath = Path.Combine(outputDirectory, "html-pdf-evidence.json");
        var jsonOptions = new JsonSerializerOptions {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, jsonOptions)).ConfigureAwait(false);

        bool cancellationPassed = engineReports
            .Where(engine => engine.Cancellation.ApiSupportsCancellation)
            .All(engine => string.Equals(engine.Cancellation.Status, "Passed", StringComparison.Ordinal));
        bool processTreeMemoryPassed = engineReports.All(engine => engine.MemoryComparable);
        Console.WriteLine("HTML_PDF_EVIDENCE_REPORT=" + reportPath);
        Console.WriteLine("HTML_PDF_EVIDENCE_ENGINES=" + engineReports.Count);
        Console.WriteLine("HTML_PDF_EVIDENCE_OUTPUTS=" + engineReports.Sum(engine => engine.Outputs.Count));
        Console.WriteLine("HTML_PDF_EVIDENCE_CANCELLATION=" + (cancellationPassed ? "Passed" : "Failed"));
        Console.WriteLine("HTML_PDF_EVIDENCE_PROCESS_TREE_MEMORY=" + (processTreeMemoryPassed ? "Passed" : "Failed"));
        return cancellationPassed && processTreeMemoryPassed ? 0 : 1;
    }

    private static async Task<HtmlPdfEngineEvidence> RunEngineAsync(
        HtmlPdfComparisonEngine engine,
        string html,
        PdfBenchmarkScenario scenario,
        int iterations,
        string outputDirectory,
        ExternalPdfRasterizer? rasterizer) {
        var outputs = new List<HtmlPdfOutputEvidence>(iterations);
        for (int iteration = 1; iteration <= iterations; iteration++) {
            outputs.Add(await RenderOnceAsync(
                engine,
                iteration,
                outputDirectory,
                scenario,
                rasterizer).ConfigureAwait(false));
        }

        HtmlPdfCancellationEvidence cancellation = engine switch {
            HtmlPdfComparisonEngine.OfficeIMO => await ProbeOfficeImoCancellationAsync(html).ConfigureAwait(false),
            HtmlPdfComparisonEngine.Chromium => await ProbeChromiumCancellationAsync(html).ConfigureAwait(false),
            _ => new HtmlPdfCancellationEvidence(
                ApiSupportsCancellation: false,
                Status: "Unsupported",
                Detail: "The compared public conversion entry point does not accept a CancellationToken.")
        };
        string[] externalVisualHashes = outputs
            .Select(output => output.ExternalVisual?.Sha256)
            .Where(hash => hash != null)
            .Cast<string>()
            .ToArray();

        return new HtmlPdfEngineEvidence(
            Engine: engine.ToString(),
            Owner: Owner(engine),
            AssemblyVersion: AssemblyVersion(engine),
            ExecutionKind: engine == HtmlPdfComparisonEngine.Chromium
                ? "Fresh worker process; Chromium through HtmlTinkerX"
                : "Fresh managed worker process",
            Cancellation: cancellation,
            Determinism: new HtmlPdfDeterminismEvidence(
                ExactBytesIdentical: outputs.Select(output => output.Sha256).Distinct(StringComparer.Ordinal).Count() == 1,
                SemanticOutputIdentical: outputs.Select(output => output.SemanticSha256).Distinct(StringComparer.Ordinal).Count() == 1,
                ManagedVisualPreviewIdentical: outputs.Select(output => output.ManagedVisual.Sha256).Distinct(StringComparer.Ordinal).Count() == 1,
                ExternalVisualPreviewIdentical: externalVisualHashes.Length == outputs.Count
                    ? externalVisualHashes.Distinct(StringComparer.Ordinal).Count() == 1
                    : null,
                UniqueByteHashCount: outputs.Select(output => output.Sha256).Distinct(StringComparer.Ordinal).Count(),
                UniqueSemanticHashCount: outputs.Select(output => output.SemanticSha256).Distinct(StringComparer.Ordinal).Count(),
                UniqueManagedVisualHashCount: outputs.Select(output => output.ManagedVisual.Sha256).Distinct(StringComparer.Ordinal).Count(),
                UniqueExternalVisualHashCount: externalVisualHashes.Length == outputs.Count
                    ? externalVisualHashes.Distinct(StringComparer.Ordinal).Count()
                    : null),
            MemoryScope: "Fresh worker process tree sampled from process start through renderer shutdown; the evidence coordinator is excluded.",
            MemoryComparable: outputs.All(output =>
                output.ProcessTreeMemory.SampleCount > 0 &&
                output.ProcessTreeMemory.MinimumObservedProcessCount > 0),
            Outputs: outputs);
    }

    private static async Task<HtmlPdfOutputEvidence> RenderOnceAsync(
        HtmlPdfComparisonEngine engine,
        int iteration,
        string outputDirectory,
        PdfBenchmarkScenario scenario,
        ExternalPdfRasterizer? rasterizer) {
        string fileName = EngineFileName(engine) + "-" + iteration.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + ".pdf";
        string outputPath = Path.Combine(outputDirectory, fileName);
        string workerResultPath = Path.Combine(
            outputDirectory,
            "." + Path.GetFileNameWithoutExtension(fileName) + "-worker.json");
        ProcessStartInfo startInfo = HtmlPdfEvidenceWorker.CreateStartInfo(
            engine,
            Path.Combine(outputDirectory, "scenario.html"),
            outputPath,
            workerResultPath,
            FindRepositoryRoot());
        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Could not start the isolated {engine} evidence worker.");
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
                // The process exited while the timeout path was taking ownership.
            }
            throw new TimeoutException(
                $"The isolated {engine} evidence worker exceeded the {WorkerTimeout.TotalMinutes:0}-minute limit.");
        }
        await memory.StopAsync().ConfigureAwait(false);
        string workerOutput = await standardOutput.ConfigureAwait(false);
        string workerError = await standardError.ConfigureAwait(false);
        if (process.ExitCode != 0) {
            throw new InvalidOperationException(
                $"The isolated {engine} evidence worker exited with code {process.ExitCode}. " +
                FormatWorkerDiagnostics(workerOutput, workerError));
        }
        if (!File.Exists(outputPath) || !File.Exists(workerResultPath)) {
            throw new InvalidDataException(
                $"The isolated {engine} evidence worker did not produce its PDF and result metadata. " +
                FormatWorkerDiagnostics(workerOutput, workerError));
        }

        byte[] bytes = await File.ReadAllBytesAsync(outputPath).ConfigureAwait(false);
        HtmlPdfWorkerResult workerResult = await HtmlPdfEvidenceWorker.ReadResultAsync(workerResultPath).ConfigureAwait(false);
        File.Delete(workerResultPath);

        PdfReadObservation observation = PdfBenchmarkValidation.ValidateGenerated(bytes, scenario, engine.ToString());
        PdfBenchmarkValidation.ValidateTaggedStructure(bytes, engine.ToString(), scenario);
        PdfDocumentInfo info = OfficeIMO.Pdf.PdfDocument.Open(bytes).Inspect();
        PdfTaggedContentInfo tagged = info.TaggedContent
            ?? throw new InvalidDataException($"{engine} did not expose tagged-content evidence.");

        PdfPageRenderResult visual = OfficeIMO.Pdf.PdfDocument.Open(bytes).Read.RenderPages(
            "1",
            new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Png,
                Dpi = 120D,
                ContinueOnError = false,
                MaxPages = 1
            }).Single();
        byte[] visualBytes = visual.Bytes
            ?? throw new InvalidDataException($"{engine} page-one visual preview did not render.");
        string visualFileName = Path.GetFileNameWithoutExtension(fileName) + "-page-1.png";
        await File.WriteAllBytesAsync(Path.Combine(outputDirectory, visualFileName), visualBytes).ConfigureAwait(false);
        HtmlPdfVisualEvidence? externalVisual = rasterizer == null
            ? null
            : await rasterizer.RenderFirstPageAsync(
                outputPath,
                outputDirectory,
                Path.GetFileNameWithoutExtension(fileName) + "-page-1-poppler").ConfigureAwait(false);

        var structureTypeCounts = tagged.StructureTypeCounts
            .OrderBy(pair => pair.Key, StringComparer.Ordinal)
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        string semantic = string.Join("|", new[] {
            observation.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            observation.TextLength.ToString(System.Globalization.CultureInfo.InvariantCulture),
            observation.ReportMarkerCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            observation.CharacterChecksum.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Sha256(Encoding.UTF8.GetBytes(observation.NormalizedText)),
            info.HasTaggedContent.ToString(),
            tagged.Marked.ToString(),
            info.CatalogLanguage ?? string.Empty,
            tagged.LanguageElementCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            tagged.StructureElementCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            tagged.MarkedContentReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            tagged.ParentTreeEntryCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            tagged.HasDocumentStructureElement.ToString(),
            tagged.FiguresHaveAlternateText.ToString(),
            string.Join(",", structureTypeCounts.Select(pair => pair.Key + ":" + pair.Value))
        });

        return new HtmlPdfOutputEvidence(
            Iteration: iteration,
            RelativePath: fileName,
            DurationMilliseconds: workerResult.DurationMilliseconds,
            SizeBytes: bytes.LongLength,
            Sha256: Sha256(bytes),
            SemanticSha256: Sha256(Encoding.UTF8.GetBytes(semantic)),
            ManagedAllocatedBytes: workerResult.ManagedAllocatedBytes,
            ProcessTreeMemory: memory.CreateEvidence(),
            Contract: new HtmlPdfContractEvidence(
                observation.PageCount,
                observation.TextLength,
                observation.ReportMarkerCount,
                observation.CharacterChecksum,
                info.HasTaggedContent,
                tagged.Marked == true,
                info.CatalogLanguage,
                tagged.LanguageElementCount,
                tagged.StructureElementCount,
                tagged.MarkedContentReferenceCount,
                tagged.ParentTreeEntryCount,
                tagged.HasDocumentStructureElement,
                tagged.FiguresHaveAlternateText,
                structureTypeCounts),
            ManagedVisual: new HtmlPdfVisualEvidence(
                Renderer: "OfficeIMO.Pdf managed page renderer",
                RelativePath: visualFileName,
                PageNumber: visual.PageNumber,
                Width: visual.Width,
                Height: visual.Height,
                SizeBytes: visualBytes.LongLength,
                Sha256: Sha256(visualBytes),
                Diagnostics: visual.Diagnostics),
            ExternalVisual: externalVisual);
    }

    private static string FormatWorkerDiagnostics(string standardOutput, string standardError) {
        string combined = string.Join(
            " | ",
            new[] { standardError.Trim(), standardOutput.Trim() }.Where(value => value.Length > 0));
        if (combined.Length == 0) return "No worker diagnostics were written.";
        const int maximumLength = 2000;
        return combined.Length <= maximumLength ? combined : combined[..maximumLength] + "...";
    }

    private static async Task<HtmlPdfCancellationEvidence> ProbeChromiumCancellationAsync(string html) {
        try {
            await using HtmlBrowserSession session = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
            await HtmlPdfComparisonRenderers.PrepareChromiumPageAsync(session, html).ConfigureAwait(false);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            try {
                _ = await HtmlPdfComparisonRenderers.CaptureChromiumPageAsync(session, cancellation.Token).ConfigureAwait(false);
                return new HtmlPdfCancellationEvidence(true, "Failed", "A pre-cancelled Chromium PDF request completed instead of cancelling.");
            } catch (OperationCanceledException) {
                return new HtmlPdfCancellationEvidence(true, "Passed", "A pre-cancelled Chromium PDF request was rejected through HtmlTinkerX.");
            }
        } catch (Exception exception) {
            return new HtmlPdfCancellationEvidence(true, "Failed", exception.GetType().Name + ": " + exception.Message);
        }
    }

    private static PdfBenchmarkScale ReadScale(string[] args) {
        string? value = ReadOption(args, "--scale");
        if (value == null) return PdfBenchmarkScale.Easy;
        if (Enum.TryParse(value, ignoreCase: true, out PdfBenchmarkScale scale) && Enum.IsDefined(scale)) return scale;
        throw new ArgumentException("--scale must be Easy, Medium, or High.");
    }

    private static int ReadIterations(string[] args) {
        string? value = ReadOption(args, "--iterations");
        if (value == null) return 3;
        if (int.TryParse(value, out int iterations) && iterations >= MinimumIterations && iterations <= MaximumIterations) return iterations;
        throw new ArgumentException($"--iterations must be between {MinimumIterations} and {MaximumIterations}.");
    }

    private static string ResolveOutputDirectory(string[] args) {
        string? configured = ReadOption(args, "--output");
        if (!string.IsNullOrWhiteSpace(configured)) return Path.GetFullPath(configured);
        string runDirectory = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss.fff", System.Globalization.CultureInfo.InvariantCulture) +
            "-" + Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "-" + Guid.NewGuid().ToString("N")[..8];
        return Path.Combine(FindRepositoryRoot(), "Ignore", "Benchmarks", "HtmlPdfEvidence", runDirectory);
    }

    private static string? ReadOption(string[] args, string option) {
        for (int index = 1; index < args.Length; index++) {
            if (!string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) continue;
            if (index == args.Length - 1 || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
                throw new ArgumentException(option + " requires a value.");
            }
            return args[index + 1];
        }
        return null;
    }

    private static void ValidateArguments(string[] args) {
        for (int index = 1; index < args.Length; index++) {
            string argument = args[index];
            if (string.Equals(argument, "--help", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "--require-external-rasterizer", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "--require-clean-source", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }
            if (string.Equals(argument, "--output", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "--scale", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "--iterations", StringComparison.OrdinalIgnoreCase)) {
                if (index == args.Length - 1 || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
                    throw new ArgumentException(argument + " requires a value.");
                }
                index++;
                continue;
            }
            throw new ArgumentException("Unknown html-evidence option: " + argument);
        }
    }

    private static string FindRepositoryRoot() {
        foreach (string seed in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() }) {
            string? current = Path.GetFullPath(seed);
            while (!string.IsNullOrWhiteSpace(current)) {
                if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
                current = Directory.GetParent(current)?.FullName;
            }
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static string Owner(HtmlPdfComparisonEngine engine) => engine switch {
        HtmlPdfComparisonEngine.OfficeIMO => "OfficeIMO.Html.Pdf",
        HtmlPdfComparisonEngine.PeachPDF => "PeachPDF",
        HtmlPdfComparisonEngine.ITextPdfHtml => "iText pdfHTML",
        HtmlPdfComparisonEngine.Chromium => "HtmlTinkerX",
        _ => throw new ArgumentOutOfRangeException(nameof(engine), engine, null)
    };

    private static string EngineFileName(HtmlPdfComparisonEngine engine) => engine switch {
        HtmlPdfComparisonEngine.OfficeIMO => "officeimo",
        HtmlPdfComparisonEngine.PeachPDF => "peachpdf",
        HtmlPdfComparisonEngine.ITextPdfHtml => "itext-pdfhtml",
        HtmlPdfComparisonEngine.Chromium => "chromium-htmltinkerx",
        _ => throw new ArgumentOutOfRangeException(nameof(engine), engine, null)
    };

    private static string AssemblyVersion(HtmlPdfComparisonEngine engine) {
        Assembly assembly = engine switch {
            HtmlPdfComparisonEngine.OfficeIMO => typeof(OfficeIMO.Html.Pdf.HtmlPdfSaveOptions).Assembly,
            HtmlPdfComparisonEngine.PeachPDF => typeof(PeachPDF.PdfGenerator).Assembly,
            HtmlPdfComparisonEngine.ITextPdfHtml => typeof(iText.Html2pdf.HtmlConverter).Assembly,
            HtmlPdfComparisonEngine.Chromium => typeof(HtmlBrowser).Assembly,
            _ => throw new ArgumentOutOfRangeException(nameof(engine), engine, null)
        };
        return assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? assembly.GetName().Version?.ToString()
            ?? "unknown";
    }

    private static async Task<HtmlPdfEvidenceProvenance> ReadProvenanceAsync(string repositoryRoot) {
        GitSourceState? officeImo = await SourceProvenanceReader.ReadGitStateAsync(repositoryRoot).ConfigureAwait(false);
        string? compiledHtmlTinkerXProjectPath = Assembly.GetExecutingAssembly()
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(attribute => string.Equals(
                attribute.Key,
                "HtmlTinkerXSourceProjectPath",
                StringComparison.Ordinal))
            ?.Value;
        string? environmentHtmlTinkerXProjectPath = Environment.GetEnvironmentVariable("HTMLTINKERX_PROJECT_PATH");
        ValidateCompiledSourceSelection(compiledHtmlTinkerXProjectPath, environmentHtmlTinkerXProjectPath);
        bool htmlTinkerXSourceConfigured = !string.IsNullOrWhiteSpace(compiledHtmlTinkerXProjectPath);
        GitSourceState? htmlTinkerX = !htmlTinkerXSourceConfigured
            ? null
            : await SourceProvenanceReader.ReadGitStateAsync(Path.GetDirectoryName(Path.GetFullPath(compiledHtmlTinkerXProjectPath!))!).ConfigureAwait(false);
        return new HtmlPdfEvidenceProvenance(
            OfficeIMO: new HtmlPdfSourceReference(
                Kind: "source",
                Version: AssemblyVersion(HtmlPdfComparisonEngine.OfficeIMO),
                Commit: officeImo?.Commit,
                WorktreeClean: officeImo?.IsClean),
            HtmlTinkerX: new HtmlPdfSourceReference(
                Kind: htmlTinkerXSourceConfigured ? "source" : "package",
                Version: AssemblyVersion(HtmlPdfComparisonEngine.Chromium),
                Commit: htmlTinkerX?.Commit,
                WorktreeClean: htmlTinkerX?.IsClean));
    }

    private static void ValidateCompiledSourceSelection(string? compiledPath, string? environmentPath) {
        if (string.IsNullOrWhiteSpace(compiledPath)) {
            if (!string.IsNullOrWhiteSpace(environmentPath)) {
                throw new InvalidOperationException(
                    "HTMLTINKERX_PROJECT_PATH is set at runtime, but this executable was not compiled against HtmlTinkerX source.");
            }
            return;
        }
        if (string.IsNullOrWhiteSpace(environmentPath)) return;

        string authoritativePath = Path.GetFullPath(compiledPath);
        string runtimePath = Path.GetFullPath(environmentPath);
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        if (!string.Equals(authoritativePath, runtimePath, comparison)) {
            throw new InvalidOperationException(
                $"Runtime HTMLTINKERX_PROJECT_PATH '{runtimePath}' does not match compiled source '{authoritativePath}'.");
        }
    }

    private static void ValidateCleanSource(HtmlPdfEvidenceProvenance provenance) {
        if (provenance.OfficeIMO.Commit == null || provenance.OfficeIMO.WorktreeClean != true) {
            throw new InvalidOperationException("Clean, commit-addressable OfficeIMO source is required for this evidence run.");
        }
        if (string.Equals(provenance.HtmlTinkerX.Kind, "source", StringComparison.Ordinal) &&
            (provenance.HtmlTinkerX.Commit == null || provenance.HtmlTinkerX.WorktreeClean != true)) {
            throw new InvalidOperationException("Clean, commit-addressable HtmlTinkerX source is required for this evidence run.");
        }
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static void WriteHelp() {
        Console.WriteLine("html-evidence --output <directory> [--scale Easy|Medium|High] [--iterations 2-10] [--require-external-rasterizer] [--require-clean-source]");
        Console.WriteLine("Generates equivalent four-engine PDFs and machine-readable correctness, tagging, size, repeatability, cancellation, and isolated process-tree memory evidence.");
    }

    private static async Task<HtmlPdfCancellationEvidence> ProbeOfficeImoCancellationAsync(string html) {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try {
            _ = await OfficeImoPdfGenerator.GenerateHtmlAsync(
                html,
                cancellationToken: cancellation.Token).ConfigureAwait(false);
            return new HtmlPdfCancellationEvidence(true, "Failed", "A pre-cancelled OfficeIMO HTML-to-PDF request completed instead of cancelling.");
        } catch (OperationCanceledException) {
            return new HtmlPdfCancellationEvidence(true, "Passed", "A pre-cancelled OfficeIMO HTML-to-PDF request was rejected through ToPdfAsync.");
        } catch (Exception exception) {
            return new HtmlPdfCancellationEvidence(true, "Failed", exception.GetType().Name + ": " + exception.Message);
        }
    }
}
