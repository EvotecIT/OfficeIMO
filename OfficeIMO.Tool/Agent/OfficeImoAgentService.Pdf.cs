using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal async Task<AgentPdfWorkflowResult> PdfAsync(string path, string outputPath, string operation,
        PdfWorkflowSettings settings, int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) {
        settings.Validate();
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        operation = operation.ToLowerInvariant();
        OfficeWorkflowOperation kind = operation switch {
            "extract" => OfficeWorkflowOperation.ExtractPages,
            "decrypt" => OfficeWorkflowOperation.RemovePdfProtection,
            "flatten" => OfficeWorkflowOperation.ScanCleanup,
            "optimize" => OfficeWorkflowOperation.Optimize,
            "sanitize" => OfficeWorkflowOperation.Sanitize,
            _ => throw new AgentUsageException("PDF operation must be extract, decrypt, flatten, optimize, or sanitize.")
        };
        if (kind == OfficeWorkflowOperation.ExtractPages && settings.Pages is null) throw new AgentUsageException("Extraction requires an explicit page selection.");
        if (kind is not (OfficeWorkflowOperation.ExtractPages or OfficeWorkflowOperation.ScanCleanup) && settings.Pages is not null)
            throw new AgentUsageException("Page selection is available only for extract and flatten.");
        if (kind == OfficeWorkflowOperation.ScanCleanup && !settings.AcknowledgeRasterOutput)
            throw new AgentUsageException("Acknowledge raster output: only rendered appearances are copied; native text, forms, links, signatures, and attachments are omitted.");
        string input = ResolvePdfInput(path);
        string destination = PreparePdfOutput(outputPath, [input], isDirectory: false, settings.Overwrite, operation, maxOutputCharacters);
        string? password = PdfPassword(settings);
        if (kind == OfficeWorkflowOperation.RemovePdfProtection && password is null)
            throw new AgentUsageException("Decryption requires an owner password environment variable.");
        var request = new OfficeWorkflowRequest {
            Operation = kind, InputPath = input, InputStream = PdfInput(input), OutputPath = destination,
            PdfPassword = password, PdfOwnerPassword = kind == OfficeWorkflowOperation.RemovePdfProtection ? password : null,
            PageSelector = kind == OfficeWorkflowOperation.ExtractPages ? settings.Selector() : null,
            MaximumExtractedPages = settings.MaximumPages,
            ScanCleanup = kind == OfficeWorkflowOperation.ScanCleanup ? new OfficeScanCleanupOptions {
                AcknowledgeRasterOutput = true, PageSelector = settings.Selector(), Preparation = OcrOptions(settings)
            } : null,
            ConflictPolicy = settings.ConflictPolicy, Limits = settings.Limits(), PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, [input])
        };
        OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(request, cancellationToken: cancellationToken).ConfigureAwait(false);
        return PdfResult(operation, result.Status, result.FailureKind, result.OutputPath, result.OutputBytes,
            result.OutputPath is null ? [] : [new AgentPdfArtifact { Path = result.OutputPath, SizeBytes = result.OutputBytes }],
            result.Diagnostics, maxOutputCharacters);
    }

    internal async Task<AgentPdfWorkflowResult> SplitPdfAsync(string path, string outputDirectory, int pagesPerDocument,
        PdfWorkflowSettings settings, int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) {
        settings.Validate(); maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        if (pagesPerDocument is < 1 or > 10_000) throw new AgentUsageException("pagesPerDocument must be from 1 through 10000.");
        string input = ResolvePdfInput(path);
        string destination = PreparePdfOutput(outputDirectory, [input], true, settings.Overwrite, "split", maxOutputCharacters);
        PdfSplitWorkflowResult result = await new OfficeWorkflowRunner().SplitPdfAsync(new PdfSplitWorkflowRequest {
            InputPath = input, InputStream = PdfInput(input), OutputDirectory = destination,
            PagesPerDocument = pagesPerDocument, MaximumParts = settings.MaximumPages, PdfPassword = PdfPassword(settings),
            ConflictPolicy = settings.ConflictPolicy, Limits = settings.Limits(), PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, [input])
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        return PdfResult("split", result.Status, OfficeWorkflowFailureKind.None, result.Succeeded ? destination : null,
            result.Files.Sum(file => file.SizeBytes), result.Files.Select(file => new AgentPdfArtifact {
                Path = file.Path, SizeBytes = file.SizeBytes, FirstSourcePage = file.FirstSourcePage, PageCount = file.PageCount
            }).ToArray(), result.Diagnostics, maxOutputCharacters);
    }

    internal async Task<AgentPdfWorkflowResult> SearchablePdfAsync(string path, string outputPath, string providerId,
        PdfWorkflowSettings settings, string? language = null, double minimumConfidence = 0.5,
        int maxOutputCharacters = 4000, CancellationToken cancellationToken = default) {
        settings.Validate(); maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        if (!double.IsFinite(minimumConfidence) || minimumConfidence is < 0 or > 1) throw new AgentUsageException("minimumConfidence must be from 0 through 1.");
        if (language is { Length: > 256 }) throw new AgentUsageException("language cannot exceed 256 characters.");
        string input = ResolvePdfInput(path);
        string destination = PreparePdfOutput(outputPath, [input], false, settings.Overwrite, "ocr", maxOutputCharacters);
        string? password = PdfPassword(settings);
        IOcrEngine engine = _pdfOcrCatalog.Create(providerId, _pdfOcrProviderOptions);
        try {
            var options = OcrOptions(settings);
            options.Language = language; options.MinimumConfidence = minimumConfidence;
            PdfSearchableWorkflowResult result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new PdfSearchableWorkflowRequest {
                InputPath = input, InputStream = PdfInput(input), OutputPath = destination, PdfPassword = password,
                Ocr = options, PageSelector = settings.Selector(), ConflictPolicy = settings.ConflictPolicy,
                Limits = settings.Limits(), PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, [input])
            }, engine, cancellationToken).ConfigureAwait(false);
            long bytes = result.OutputPath is null ? 0 : new FileInfo(result.OutputPath).Length;
            return PdfResult("ocr", result.Status, OfficeWorkflowFailureKind.None, result.OutputPath, bytes,
                result.OutputPath is null ? [] : [new AgentPdfArtifact { Path = result.OutputPath, SizeBytes = bytes }],
                result.Diagnostics, maxOutputCharacters, result.AddedWordCount);
        } finally {
            if (engine is IAsyncDisposable asynchronous) await asynchronous.DisposeAsync().ConfigureAwait(false);
            else if (engine is IDisposable disposable) disposable.Dispose();
        }
    }

    internal IReadOnlyList<OcrEngineDescriptor> PdfOcrProviders() => _pdfOcrCatalog.Discover();

    private string? PdfPassword(PdfWorkflowSettings settings) {
        if (settings.PasswordEnvironmentVariable is { } name && !_pdfPasswordEnvironmentVariables.Contains(name))
            throw new AgentUsageException("The selected PDF password environment variable is not admitted by this host.");
        return settings.Password();
    }

    internal AgentPdfOcrProvidersResult PdfOcrProviders(int maximumCharacters) {
        maximumCharacters = ValidateOutputBudget(maximumCharacters);
        var all = _pdfOcrCatalog.Discover();
        var ids = all.Take(25).Select(provider => provider.Id).ToList();
        var result = new AgentPdfOcrProvidersResult { ProviderCount = all.Count, ProviderIds = ids, Truncated = all.Count > ids.Count };
        while (AgentJson.Measure(result) > maximumCharacters) { ids.RemoveAt(ids.Count - 1); result.Truncated = true; }
        return result;
    }

    private static PdfOcrMergeOptions OcrOptions(PdfWorkflowSettings settings) => new() {
        Dpi = settings.Dpi, MaxPages = settings.MaximumPages, MaxPixelsPerPage = settings.MaximumPixelsPerPage,
        MaxRenderedBytesPerPage = Math.Min(25L * 1024 * 1024, settings.MaximumOutputBytes),
        MaxDiagnosticsPerPage = 100, MaxDiagnosticCharactersPerPage = 16_000
    };

    private string ResolvePdfInput(string path) {
        string input = _pathPolicy.ResolveInput(path);
        if (!File.Exists(input) || !Path.GetExtension(input).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new AgentUsageException("Select one local PDF file.");
        return input;
    }

    private OfficeWorkflowStreamInput PdfInput(string input) => new(Path.GetFileName(input), token => {
        token.ThrowIfCancellationRequested();
        return Task.FromResult<Stream>(File.OpenRead(_pathPolicy.ResolveInput(input)));
    });

    private string PreparePdfOutput(string path, IReadOnlyList<string> sources, bool isDirectory, bool overwrite,
        string operation, int maximumCharacters, string extension = ".pdf") {
        string destination = _pathPolicy.ResolveOutput(path);
        if (!isDirectory && !Path.GetExtension(destination).Equals(extension, StringComparison.OrdinalIgnoreCase))
            throw new AgentUsageException("Output must have the " + extension + " extension.");
        foreach (string source in sources) {
            if (OfficeImoToolPathSafety.PathsEqual(source, destination) ||
                (isDirectory && OfficeImoToolPathSafety.IsSameOrChildPath(destination, source)))
                throw new AgentUsageException("Choose an output separate from every source; an output folder cannot contain a source.");
        }
        if (!overwrite && (File.Exists(destination) || Directory.Exists(destination)))
            throw new AgentUsageException("Output already exists. Choose a new path or explicitly enable overwrite.");
        // Required fields must fit before any operation writes output. Artifact samples and diagnostic codes are trimmed later.
        var minimal = new AgentPdfWorkflowResult { Operation = operation, Status = "Completed", FailureKind = "OperationFailed", Succeeded = true, OutputPath = destination, Truncated = true };
        if (AgentJson.Measure(minimal) + 64 > maximumCharacters)
            throw new AgentUsageException("Increase maxOutputCharacters to include the required output path.");
        return destination;
    }

    private static AgentPdfWorkflowResult PdfResult(string operation, OfficeWorkflowStatus status,
        OfficeWorkflowFailureKind failureKind, string? outputPath, long bytes, IReadOnlyList<AgentPdfArtifact> artifacts,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, int maximumCharacters, int words = 0, string? summary = null) {
        var sampledArtifacts = artifacts.Take(25).ToList();
        var sampledDiagnostics = diagnostics.Take(10).Select(item => new AgentPdfDiagnostic {
            Code = AgentJson.Limit(item.Code, 80), Severity = item.Severity.ToString()
        }).ToList();
        var result = new AgentPdfWorkflowResult {
            Summary = summary is null ? null : AgentJson.Limit(summary, 320),
            Operation = operation, Status = status.ToString(), Succeeded = status == OfficeWorkflowStatus.Completed,
            FailureKind = status == OfficeWorkflowStatus.Failed && failureKind == OfficeWorkflowFailureKind.None ? "OperationFailed" : failureKind.ToString(),
            OutputPath = outputPath, OutputBytes = bytes, ArtifactCount = artifacts.Count, AddedWordCount = words,
            DiagnosticCount = diagnostics.Count, Artifacts = sampledArtifacts, Diagnostics = sampledDiagnostics,
            Truncated = sampledArtifacts.Count < artifacts.Count || sampledDiagnostics.Count < diagnostics.Count
        };
        while (AgentJson.Measure(result) > maximumCharacters) {
            result.Truncated = true;
            if (sampledArtifacts.Count > 0) sampledArtifacts.RemoveAt(sampledArtifacts.Count - 1);
            else if (sampledDiagnostics.Count > 0) sampledDiagnostics.RemoveAt(sampledDiagnostics.Count - 1);
            else if (result.Summary is not null) result.Summary = null;
            else throw new AgentUsageException("Increase maxOutputCharacters for this workflow report.");
        }
        return result;
    }

    private sealed class PdfRootPublicationGuard(AgentPathPolicy policy, IReadOnlyList<string> sources, Action<CancellationToken>? validateSourceIds = null) : IOfficeWorkflowPublicationGuard, IOfficeWorkflowStagingGuard {
        public ValueTask EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            _ = policy.ResolveOutput(directory);
            return ValueTask.CompletedTask;
        }
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            validateSourceIds?.Invoke(cancellationToken);
            string destination = policy.ResolveOutput(path);
            foreach (string source in sources) {
                string current = policy.ResolveInput(source);
                if (OfficeImoToolPathSafety.PathsEqual(current, destination) ||
                    (isDirectory && OfficeImoToolPathSafety.IsSameOrChildPath(destination, current))) return ValueTask.FromResult(false);
            }
            return ValueTask.FromResult(true);
        }
    }
}
