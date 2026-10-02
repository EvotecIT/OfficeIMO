using System.Text.Json;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Incremental conversion through the existing routes, with optional verified checkpoints.</summary>
internal static partial class OfficeConversionBatchExecutor {
    /// <summary>Runs a bounded local batch, reusing only hash-verified completed work under the same captured configuration.</summary>
    public static async Task<OfficeConversionBatchResult> RunAsync(IOfficeWorkflowRunner runner, OfficeConversionBatchRequest request,
        IProgress<OfficeConversionBatchItemResult>? progress = null, CancellationToken cancellationToken = default,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        OfficeConversionBatchRequest settings = Snapshot(request);
        cancellationToken.ThrowIfCancellationRequested();
        if (publicationGuard != null &&
            (!await publicationGuard.CanPublishAsync(settings.OutputDirectory, true, cancellationToken).ConfigureAwait(false) ||
             (settings.CheckpointDirectory != null && !await publicationGuard.CanPublishAsync(settings.CheckpointDirectory, true, cancellationToken).ConfigureAwait(false))))
            throw new UnauthorizedAccessException("Batch output or checkpoint directory is protected by the host publication policy.");
        Directory.CreateDirectory(settings.OutputDirectory);
        using var lease = OpenCheckpoint(settings);
        string configuration = settings.CheckpointDirectory == null ? string.Empty : CaptureConfiguration(settings);
        if (settings.CheckpointDirectory != null) {
            string planPath = Path.Combine(settings.CheckpointDirectory, "plan.json");
            OfficeConversionBatchPlan? plan = ReadState(planPath, OfficeConversionBatchJsonContext.Default.OfficeConversionBatchPlan);
            if (plan != null && (plan.Schema != Schema || plan.ConfigurationSha256 != configuration))
                throw new InvalidDataException("Conversion settings or checkpoint compatibility changed. Select a new job destination and checkpoint; execution limits may be changed in place.");
            if (plan == null) WriteState(planPath, new OfficeConversionBatchPlan(Schema, configuration), OfficeConversionBatchJsonContext.Default.OfficeConversionBatchPlan);
        }
        long discovered = 0, selected = 0, completed = 0, reused = 0, failed = 0, skipped = 0;
        bool cancelled = false;
        try {
            await Parallel.ForEachAsync(SelectInputs(settings, cancellationToken),
                new ParallelOptions { MaxDegreeOfParallelism = settings.MaximumConcurrency, CancellationToken = cancellationToken }, async (input, token) => {
                    if (Interlocked.Increment(ref discovered) > settings.MaximumFiles) throw new InvalidDataException("Conversion selection exceeds its file limit.");
                    OfficeWorkflowRoute? route = SelectRoute(settings, input);
                    if (route == null) {
                        Interlocked.Increment(ref skipped);
                        progress?.Report(new(input, null, null, false, "No selected executable conversion route.",
                            [new OfficeWorkflowDiagnostic("ConversionInputSkipped", "The file does not match the selected formats and executable target routes.", OfficeWorkflowDiagnosticSeverity.Information, "selection")], true));
                        return;
                    }
                    Interlocked.Increment(ref selected);
                    string relative = settings.InputDirectory == null ? Path.GetFileName(input) : Path.GetRelativePath(settings.InputDirectory, input);
                    string output = Path.Combine(settings.OutputDirectory, relative + settings.TargetExtension);
                    OfficeConversionBatchItemResult item;
                    if (settings.CheckpointDirectory == null) {
                        OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
                            InputPath = input, OutputPath = output, Operation = OfficeWorkflowOperation.Convert,
                            ConversionRouteId = route.Id, ConversionOptions = settings.ConversionOptions.ForRoute(route.Id),
                            OutputProfile = settings.OutputProfile, ConflictPolicy = settings.ConflictPolicy, PublicationGuard = publicationGuard, PdfPassword = settings.PdfPassword,
                            Limits = new OfficeWorkflowLimits { MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes }
                        }, cancellationToken: token).ConfigureAwait(false);
                        item = new(input, result.OutputPath, result.Status, false, result.Summary, result.Diagnostics);
                    } else {
                        string key = Hash(OfficePathIdentity.GetPathIdentityKey(input));
                        string receiptPath = Path.Combine(settings.CheckpointDirectory, "items", key[..2], key + ".json");
                        item = await RunItemAsync(runner, settings, configuration, route.Id, input, output, receiptPath, token, publicationGuard).ConfigureAwait(false);
                    }
                    if (item.Status == OfficeWorkflowStatus.Completed) {
                        Interlocked.Increment(ref completed);
                        if (item.Reused) Interlocked.Increment(ref reused);
                    } else if (item.Status == OfficeWorkflowStatus.Cancelled) cancelled = true;
                    else Interlocked.Increment(ref failed);
                    progress?.Report(item);
                }).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { cancelled = true; }
        return new OfficeConversionBatchResult(selected, completed, reused, failed, cancelled, skipped);
    }

    private static async Task<OfficeConversionBatchItemResult> RunItemAsync(IOfficeWorkflowRunner runner, OfficeConversionBatchRequest settings, string configuration, string routeId,
        string input, string output, string receiptPath, CancellationToken token, IOfficeWorkflowPublicationGuard? publicationGuard) {
        string? inputHash = null;
        string? stagingPath = null;
        bool pendingRecorded = false;
        try {
            EnsureNoLinks(input); EnsureNoLinks(output);
            string sourceRoot = settings.InputDirectory ?? Path.GetDirectoryName(input)!;
            inputHash = await HashFileAsync(input, sourceRoot, settings.MaximumInputBytes, token).ConfigureAwait(false);
            string? resourceRoot = GetResourceRoot(settings, routeId, input);
            if (resourceRoot != null) {
                resourceRoot = Path.GetFullPath(resourceRoot);
                if (!OfficePathIdentity.IsSameOrDescendant(resourceRoot, sourceRoot))
                    throw new ArgumentException("Checkpoint resources must remain inside the selected source root.");
            }
            string? resourceIdentity = resourceRoot == null ? null : await CaptureResourceIdentityAsync(resourceRoot, input, settings.MaximumInputBytes, token).ConfigureAwait(false);
            configuration = Hash(configuration + CaptureRenderingConfiguration(settings, routeId) + resourceIdentity);
            OfficeConversionBatchReceipt? receipt = ReadState(receiptPath, OfficeConversionBatchJsonContext.Default.OfficeConversionBatchReceipt);
            if (receipt != null) {
                if (receipt.Schema != Schema || receipt.ConfigurationSha256 != configuration ||
                    receipt.Status is not OfficeWorkflowStatus.Completed and not OfficeWorkflowStatus.Failed ||
                    (receipt.InputSha256 != inputHash && !(settings.RetryFailed && receipt.Status == OfficeWorkflowStatus.Failed)))
                    throw new InvalidDataException("Batch input or checkpoint changed; no output was replaced.");
                if (receipt.Status == OfficeWorkflowStatus.Completed) {
                    await CompletePublicationAsync(settings, input, output, receiptPath, receipt, token, publicationGuard, resourceRoot, resourceIdentity).ConfigureAwait(false);
                    return new(input, output, OfficeWorkflowStatus.Completed, true, "Verified recorded batch artifact and completed publication.",
                        RestoreDiagnostics(receipt));
                }
                if (!settings.RetryFailed) return new(input, output, receipt.Status, false, receipt.Summary,
                    RestoreDiagnostics(receipt));
            }
            if (File.Exists(output)) throw new InvalidDataException("An output exists without a verified completion receipt. Check it before retrying.");
            string stageId = Guid.NewGuid().ToString("N");
            stagingPath = GetPendingStagePath(output, stageId);
            OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
                InputPath = input, OutputPath = stagingPath, Operation = OfficeWorkflowOperation.Convert,
                InputStream = routeId is "html-pdf" or "markdown-pdf" ? null : new OfficeWorkflowStreamInput(Path.GetFileName(input), readToken => {
                    readToken.ThrowIfCancellationRequested();
                    return Task.FromResult<Stream>(OfficePathIdentity.OpenRegularFileForRead(input, sourceRoot, 81920));
                }, inputHash),
                ConversionRouteId = routeId,
                ConversionOptions = settings.ConversionOptions.ForRoute(routeId), OutputProfile = settings.OutputProfile, ConflictPolicy = OfficeWorkflowConflictPolicy.Fail, PdfPassword = settings.PdfPassword,
                Limits = new OfficeWorkflowLimits { MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes },
                PublicationGuard = new BatchPublicationGuard(settings, input, inputHash, publicationGuard, output, stagingPath, resourceRoot, resourceIdentity)
            }, cancellationToken: token).ConfigureAwait(false);
            if (result.Status == OfficeWorkflowStatus.Cancelled) return new(input, output, result.Status, false, result.Summary, result.Diagnostics);
            string? outputHash = result.Succeeded
                ? await HashFileAsync(stagingPath, settings.OutputDirectory, settings.MaximumOutputBytes, CancellationToken.None).ConfigureAwait(false) : null;
            if (result.Succeeded) {
                // Persist the artifact before its durable publication intent. Both names
                // are siblings, so final publication remains one filesystem move.
                using var staged = new FileStream(stagingPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                staged.Flush(flushToDisk: true);
            }
            var completedReceipt = new OfficeConversionBatchReceipt(Schema, configuration, inputHash, outputHash, result.Status,
                result.Summary.Length > 4096 ? result.Summary[..4096] : result.Summary,
                result.Diagnostics.Where(IsRecordedDiagnostic)
                    .Take(32).Select(diagnostic => new OfficeConversionBatchStoredDiagnostic(diagnostic.Code,
                        BoundMessage(diagnostic.Message), diagnostic.Severity, diagnostic.Stage,
                        diagnostic.Details.TryGetValue("source", out string? source) ? BoundMessage(source) : null,
                        diagnostic.Details.TryGetValue("lossKind", out string? loss) ? loss : null)).ToArray(),
                result.Diagnostics.Count(IsRecordedDiagnostic), result.Succeeded ? stageId : null);
            StoreReceipt(receiptPath, completedReceipt);
            pendingRecorded = result.Succeeded;
            if (result.Succeeded)
                await CompletePublicationAsync(settings, input, output, receiptPath, completedReceipt, token, publicationGuard, resourceRoot, resourceIdentity).ConfigureAwait(false);
            return new(input, output, result.Status, false, result.Summary, result.Diagnostics);
        } catch (OperationCanceledException) when (token.IsCancellationRequested) {
            return new(input, output, OfficeWorkflowStatus.Cancelled, false, "Batch item cancelled; completed outputs remain available.", []);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            // Never replace an earlier receipt after changed-source or output-integrity failures.
            return new(input, output, OfficeWorkflowStatus.Failed, false, error.Message, []);
        } finally {
            if (!pendingRecorded && stagingPath != null && File.Exists(stagingPath)) {
                EnsureNoLinks(stagingPath);
                File.Delete(stagingPath);
            }
        }
    }

    private static string BoundMessage(string message) => message.Length > 1024 ? message[..1024] : message;
    private static bool IsRecordedDiagnostic(OfficeWorkflowDiagnostic diagnostic) =>
        diagnostic.Severity != OfficeWorkflowDiagnosticSeverity.Information ||
        (diagnostic.Details.TryGetValue("lossKind", out string? loss) && loss != nameof(OfficeConversionLossKind.None));

    private static void StoreReceipt(string path, OfficeConversionBatchReceipt receipt) {
        // Escaping can multiply non-ASCII/control text size. Bound the encoded checkpoint,
        // not just character counts, so diagnostics cannot strand a published output.
        while (receipt.Diagnostics.Length > 0 && JsonSerializer.SerializeToUtf8Bytes(receipt,
            OfficeConversionBatchJsonContext.Default.OfficeConversionBatchReceipt).Length > MaximumStateBytes)
            receipt = receipt with { Diagnostics = receipt.Diagnostics[..^1] };
        WriteState(path, receipt, OfficeConversionBatchJsonContext.Default.OfficeConversionBatchReceipt);
    }

    private static IReadOnlyList<OfficeWorkflowDiagnostic> RestoreDiagnostics(OfficeConversionBatchReceipt receipt) {
        var diagnostics = receipt.Diagnostics.Select(finding => {
            var details = new Dictionary<string, string>();
            if (finding.Source != null) details["source"] = finding.Source;
            if (finding.LossKind != null) details["lossKind"] = finding.LossKind;
            return new OfficeWorkflowDiagnostic(finding.Code, finding.Message, finding.Severity, finding.Stage, details);
        }).ToList();
        if (receipt.DiagnosticCount > receipt.Diagnostics.Length)
            diagnostics.Add(new OfficeWorkflowDiagnostic("RecordedDiagnosticsTruncated",
                $"The checkpoint retains {receipt.Diagnostics.Length} of {receipt.DiagnosticCount} diagnostics.", OfficeWorkflowDiagnosticSeverity.Warning, "resume"));
        return diagnostics;
    }

    private sealed class BatchPublicationGuard(OfficeConversionBatchRequest settings, string input, string hash,
        IOfficeWorkflowPublicationGuard? hostGuard, string output, string stagingPath, string? resourceRoot, string? resourceIdentity) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) {
            EnsureNoLinks(absoluteDestination);
            EnsureNoLinks(output);
            await VerifyResourcesAsync(resourceRoot, input, resourceIdentity, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false);
            return !isDirectory && absoluteDestination == stagingPath && OfficePathIdentity.IsSameOrDescendant(absoluteDestination, settings.OutputDirectory) &&
                (settings.InputDirectory == null || !OfficePathIdentity.IsSameOrDescendant(absoluteDestination, settings.InputDirectory)) &&
                (hostGuard == null ||
                    (await hostGuard.CanPublishAsync(absoluteDestination, false, cancellationToken).ConfigureAwait(false) &&
                     await hostGuard.CanPublishAsync(output, false, cancellationToken).ConfigureAwait(false))) &&
                await HashFileAsync(input, settings.InputDirectory ?? Path.GetDirectoryName(input)!, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false) == hash;
        }
    }
}
