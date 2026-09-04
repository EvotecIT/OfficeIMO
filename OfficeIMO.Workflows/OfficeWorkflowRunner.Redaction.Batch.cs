using System.Text.Json;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <inheritdoc />
    public async Task<PdfRedactionBatchResult> RunRedactionBatchAsync(
        PdfRedactionBatchRequest request,
        IProgress<OfficeWorkflowProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(request);
        PdfRedactionWorkflowRequest[] items = await BuildBatchItemsAsync(request, cancellationToken).ConfigureAwait(false);
        PdfRedactionBatchResult result;
        if (request.PublicationPolicy == PdfRedactionBatchPublicationPolicy.AtomicAll) {
            return await RunAtomicRedactionBatchAsync(items, Path.GetFullPath(request.ManifestPath), request.Limits.MaximumEvidenceBytes, request.ConflictPolicy, progress, cancellationToken).ConfigureAwait(false);
        } else {
            result = await RunContinuePerItemBatchAsync(items, progress, cancellationToken).ConfigureAwait(false);
        }

        byte[] manifest = JsonSerializer.SerializeToUtf8Bytes(new PdfRedactionBatchRecord(result), PdfRedactionWorkflowJsonContext.Default.PdfRedactionBatchRecord);
        if (manifest.LongLength > request.Limits.MaximumEvidenceBytes) throw new InvalidOperationException("Privacy-safe redaction batch manifest exceeds the configured evidence-byte limit.");
        try {
            PublishPreparedFiles(new[] { new PreparedFile(Path.GetFullPath(request.ManifestPath), manifest) }, request.ConflictPolicy, cancellationToken);
            return result;
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException and not OperationCanceledException) {
            return new PdfRedactionBatchResult(OfficeWorkflowStatus.Failed, result.Items, result.PublishedAtomically, "Batch items completed, but the consolidated privacy-safe manifest could not be published: " + exception.GetType().Name + ".");
        }
    }

    private static async Task<PdfRedactionWorkflowRequest[]> BuildBatchItemsAsync(PdfRedactionBatchRequest batch, CancellationToken cancellationToken) {
        ValidateBatchRequest(batch);
        string inputRoot = Path.GetFullPath(batch.InputRoot);
        string evidenceRoot = Path.GetFullPath(batch.EvidenceRoot);
        string? outputRoot = NormalizeOptionalPath(batch.OutputRoot);
        string? decisionsRoot = NormalizeOptionalPath(batch.DecisionsRoot);
        string manifestPath = Path.GetFullPath(batch.ManifestPath);
        var protectedPaths = batch.ProtectedInputPaths.Select(Path.GetFullPath).ToArray();

        string[] sourcePaths;
        if (batch.InputPaths.Count == 0) {
            SearchOption searchOption = batch.RecurseSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            sourcePaths = Directory.EnumerateFiles(inputRoot, batch.SearchPattern, searchOption)
                .Where(static path => string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                .Select(Path.GetFullPath)
                .OrderBy(path => Path.GetRelativePath(inputRoot, path), StringComparer.Ordinal)
                .ToArray();
        } else {
            sourcePaths = batch.InputPaths
                .Select(relativePath => ResolveRelativePath(inputRoot, relativePath, "input"))
                .OrderBy(path => Path.GetRelativePath(inputRoot, path), StringComparer.Ordinal)
                .ToArray();
        }

        if (sourcePaths.Length > batch.Limits.MaximumBatchItems) throw new InvalidOperationException($"The redaction batch selected {sourcePaths.Length} items, above the configured {batch.Limits.MaximumBatchItems}-item limit.");
        EnsurePortableUniquePaths(sourcePaths, "Batch inputs");
        var requests = new PdfRedactionWorkflowRequest[sourcePaths.Length];
        var destinations = new List<string>(sourcePaths.Length * 2 + 1) { manifestPath };
        for (int index = 0; index < sourcePaths.Length; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string sourcePath = sourcePaths[index];
            if (!File.Exists(sourcePath)) throw new FileNotFoundException("A selected redaction batch input was not found.", sourcePath);
            if (!string.Equals(Path.GetExtension(sourcePath), ".pdf", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Every redaction batch input must be a PDF.");
            string relativePath = Path.GetRelativePath(inputRoot, sourcePath);
            string relativeDirectory = Path.GetDirectoryName(relativePath) ?? string.Empty;
            string stem = Path.GetFileNameWithoutExtension(relativePath);
            string evidencePath = Path.Combine(evidenceRoot, relativeDirectory, stem + batch.EvidenceSuffix);
            string? outputPath = outputRoot is null ? null : Path.Combine(outputRoot, relativeDirectory, stem + batch.OutputSuffix);
            string? decisionsPath = decisionsRoot is null ? null : Path.Combine(decisionsRoot, relativeDirectory, stem + batch.DecisionsSuffix);
            PdfRedactionDecisionManifest? decisions = decisionsPath is null
                ? null
                : await ReadDecisionManifestAsync(decisionsPath, cancellationToken).ConfigureAwait(false);
            destinations.Add(evidencePath);
            if (outputPath is not null && batch.Mode == PdfRedactionWorkflowMode.ApplyAndVerify) destinations.Add(outputPath);
            var itemProtectedPaths = protectedPaths.Append(sourcePath);
            if (decisionsPath is not null) itemProtectedPaths = itemProtectedPaths.Append(decisionsPath);
            requests[index] = new PdfRedactionWorkflowRequest {
                Id = "batch-" + ComputeSha256(System.Text.Encoding.UTF8.GetBytes(relativePath.Replace('\\', '/')))[..24],
                Mode = batch.Mode,
                InputPath = sourcePath,
                OutputPath = outputPath,
                EvidencePath = evidencePath,
                ProtectedInputPaths = itemProtectedPaths.ToArray(),
                Recipe = batch.Recipe,
                Decisions = decisions,
                OcrEngine = batch.OcrEngine,
                OcrOptions = batch.OcrOptions,
                OwnerPassword = batch.OwnerPassword,
                OutputEncryption = batch.OutputEncryption,
                OutputSigner = batch.OutputSigner,
                OutputSignatureOptions = batch.OutputSignatureOptions,
                OutputSignatureValidator = batch.OutputSignatureValidator,
                ExternalValidators = batch.ExternalValidators,
                ConflictPolicy = batch.ConflictPolicy,
                Limits = batch.Limits
            };
        }
        EnsurePortableUniquePaths(destinations, "Batch destinations");
        EnsureDestinationsOutsideInputs(destinations, sourcePaths, protectedPaths);
        return requests;
    }

    private static async Task<PdfRedactionDecisionManifest> ReadDecisionManifestAsync(string path, CancellationToken cancellationToken) {
        byte[] bytes = await ReadFileBoundedAsync(path, 8L * 1024L * 1024L, cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Deserialize(bytes, PdfRedactionWorkflowJsonContext.Default.PdfRedactionDecisionManifest)
            ?? throw new JsonException("Redaction decision manifest was empty.");
    }

    private async Task<PdfRedactionBatchResult> RunContinuePerItemBatchAsync(
        PdfRedactionWorkflowRequest[] items,
        IProgress<OfficeWorkflowProgress>? progress,
        CancellationToken cancellationToken) {
        if (items.Length == 0) return new PdfRedactionBatchResult(OfficeWorkflowStatus.Completed, Array.Empty<PdfRedactionWorkflowResult>(), false, "The redaction batch was empty.");
        int maximumConcurrency = items.Min(static item => item.Limits.MaximumConcurrency);
        var results = new PdfRedactionWorkflowResult[items.Length];
        using var gate = new SemaphoreSlim(maximumConcurrency, maximumConcurrency);
        Task[] tasks = items.Select((item, index) => Task.Run(async () => {
            await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try {
                var itemProgress = progress is null ? null : new InlineProgress<OfficeWorkflowProgress>(value => progress.Report(
                    new OfficeWorkflowProgress(value.RequestId, value.Stage, $"{index + 1} of {items.Length} · {value.Message}", value.Fraction, (index + value.Fraction) / items.Length)));
                results[index] = await RunRedactionAsync(item, itemProgress, cancellationToken).ConfigureAwait(false);
            } finally {
                gate.Release();
            }
        }, CancellationToken.None)).ToArray();
        try {
            await Task.WhenAll(tasks).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            for (int index = 0; index < results.Length; index++) results[index] ??= FailedRedactionResult(items[index], OfficeWorkflowStatus.Cancelled, "Redaction batch cancelled.", "Cancelled", OfficeWorkflowDiagnosticSeverity.Information);
        }
        OfficeWorkflowStatus status = cancellationToken.IsCancellationRequested
            ? OfficeWorkflowStatus.Cancelled
            : results.All(static result => result.Succeeded) ? OfficeWorkflowStatus.Completed : OfficeWorkflowStatus.Failed;
        int succeeded = results.Count(static result => result.Succeeded);
        return new PdfRedactionBatchResult(status, results, false, $"Published {succeeded} of {results.Length} redaction workflow item(s) independently.");
    }

    private static void ValidateBatchRequest(PdfRedactionBatchRequest batch) {
        if (!string.Equals(batch.Schema, PdfRedactionBatchRequest.CurrentSchema, StringComparison.Ordinal)) throw new ArgumentException("Unsupported redaction batch request schema.");
        if (!Enum.IsDefined(batch.Mode) || !Enum.IsDefined(batch.PublicationPolicy) || !Enum.IsDefined(batch.ConflictPolicy)) throw new ArgumentException("The redaction batch contains an undefined mode or policy.");
        if (batch.ConflictPolicy == OfficeWorkflowConflictPolicy.Rename) throw new ArgumentException("Redaction batches require Fail or Replace conflict policy.");
        if (string.IsNullOrWhiteSpace(batch.InputRoot) || !Directory.Exists(batch.InputRoot)) throw new DirectoryNotFoundException("Redaction batch input root was not found.");
        if (batch.InputPaths is null || batch.ProtectedInputPaths is null || batch.ExternalValidators is null || batch.Recipe is null || batch.Limits is null) throw new ArgumentException("Redaction batch collections, recipe, and limits cannot be null.");
        if (string.IsNullOrWhiteSpace(batch.EvidenceRoot) || string.IsNullOrWhiteSpace(batch.ManifestPath)) throw new ArgumentException("Redaction batches require EvidenceRoot and ManifestPath.");
        if (batch.Mode != PdfRedactionWorkflowMode.PlanOnly && (string.IsNullOrWhiteSpace(batch.OutputRoot) || string.IsNullOrWhiteSpace(batch.DecisionsRoot))) throw new ArgumentException("Apply and verify batches require OutputRoot and DecisionsRoot.");
        ValidateBatchSuffix(batch.OutputSuffix, nameof(batch.OutputSuffix));
        ValidateBatchSuffix(batch.EvidenceSuffix, nameof(batch.EvidenceSuffix));
        ValidateBatchSuffix(batch.DecisionsSuffix, nameof(batch.DecisionsSuffix));
        if (string.IsNullOrWhiteSpace(batch.SearchPattern) || batch.SearchPattern.IndexOfAny(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }) >= 0 || batch.SearchPattern.Contains("..", StringComparison.Ordinal)) throw new ArgumentException("Redaction batch SearchPattern must be one file-name pattern without directory traversal.");
        string inputRoot = Path.GetFullPath(batch.InputRoot);
        foreach (string generatedRoot in new[] { batch.OutputRoot, batch.EvidenceRoot }.Where(static path => !string.IsNullOrWhiteSpace(path))!) {
            if (IsPathWithin(inputRoot, Path.GetFullPath(generatedRoot!))) throw new ArgumentException("Redaction output and evidence roots cannot be inside the input root.");
        }
    }

    private static void ValidateBatchSuffix(string suffix, string parameterName) {
        if (string.IsNullOrWhiteSpace(suffix) || suffix.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || suffix.IndexOfAny(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, '*', '?' }) >= 0) throw new ArgumentException("Batch suffixes must be non-empty file-name suffixes.", parameterName);
    }

    private static string ResolveRelativePath(string root, string relativePath, string kind) {
        if (string.IsNullOrWhiteSpace(relativePath) || Path.IsPathRooted(relativePath)) throw new ArgumentException($"Batch {kind} paths must be non-empty and relative to their configured root.");
        string resolved = Path.GetFullPath(Path.Combine(root, relativePath));
        if (!IsPathWithin(root, resolved)) throw new ArgumentException($"Batch {kind} path escapes its configured root.");
        return resolved;
    }

    private static bool IsPathWithin(string root, string candidate) {
        string relative = Path.GetRelativePath(root, candidate);
        return !Path.IsPathRooted(relative) && relative != ".." && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal) && !relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal);
    }

    private static void EnsurePortableUniquePaths(IEnumerable<string> paths, string kind) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string path in paths) {
            string key = Path.GetFullPath(path).Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).TrimEnd(Path.DirectorySeparatorChar);
            if (!seen.Add(key)) throw new ArgumentException(kind + " collide under portable case-insensitive path semantics.");
        }
    }

    private static void EnsureDestinationsOutsideInputs(IEnumerable<string> destinations, IEnumerable<string> inputs, IEnumerable<string> protectedInputs) {
        string[] forbidden = inputs.Concat(protectedInputs).Select(Path.GetFullPath).ToArray();
        foreach (string destination in destinations) {
            if (forbidden.Any(path => OfficeWorkflowPathIdentity.AreEquivalentWithPortableFallback(destination, path))) throw new ArgumentException("A redaction batch destination cannot replace a source, decision, manifest definition, or protected host input.");
        }
    }
}
