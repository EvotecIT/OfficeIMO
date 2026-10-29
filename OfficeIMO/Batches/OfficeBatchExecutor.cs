using System.Text.Json;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO;

internal static partial class OfficeBatchExecutor {
    /// <summary>Runs a bounded local batch, reusing only hash-verified completed work under the same captured configuration.</summary>
    public static async Task<OfficeBatchSummary> RunAsync(IOfficeConversionExecutor runner, OfficeBatchOptions request,
        IProgress<OfficeBatchItemReport>? progress = null, CancellationToken cancellationToken = default,
        IOfficePublicationGuard? publicationGuard = null) {
        OfficeConversionExecutionRoute[] routes = runner.ConversionRoutes.Where(route => route.CanExecute).ToArray();
        OfficeBatchOptions settings = Snapshot(request, routes);
        cancellationToken.ThrowIfCancellationRequested();
        if (publicationGuard != null &&
            (!await publicationGuard.CanPublishAsync(settings.OutputDirectory, true, cancellationToken).ConfigureAwait(false) ||
             (settings.CheckpointDirectory != null && !await publicationGuard.CanPublishAsync(settings.CheckpointDirectory, true, cancellationToken).ConfigureAwait(false))))
            throw new UnauthorizedAccessException("Batch output or checkpoint directory is protected by the host publication policy.");
        if (settings.ResumeFrom != null && publicationGuard != null &&
            !await publicationGuard.CanPublishAsync(settings.ResumeFrom, false, cancellationToken).ConfigureAwait(false))
            throw new UnauthorizedAccessException("Checkpoint manifest is protected by the host publication policy.");
        Directory.CreateDirectory(settings.OutputDirectory);
        EnsureNoLinks(settings.OutputDirectory);
        string outputPhysicalRoot = OfficePathIdentity.ResolvePhysicalPath(settings.OutputDirectory);
        using var lease = OpenCheckpoint(settings);
        string configuration = settings.CheckpointDirectory == null ? string.Empty : CaptureConfiguration(settings);
        if (settings.CheckpointDirectory != null) {
            string planPath = settings.ResumeFrom!;
            OfficeBatchCheckpointPlan? plan = ReadState(planPath, OfficeBatchStateJsonContext.Default.OfficeBatchCheckpointPlan);
            if (plan != null && (plan.Schema != Schema || plan.ConfigurationSha256 != configuration))
                throw new InvalidDataException("Conversion settings or checkpoint compatibility changed. Select a new job destination and checkpoint; execution limits may be changed in place.");
            if (plan == null) WriteState(planPath, new OfficeBatchCheckpointPlan(Schema, configuration), OfficeBatchStateJsonContext.Default.OfficeBatchCheckpointPlan);
        }
        long discovered = 0, selected = 0, completed = 0, reused = 0, failed = 0, skipped = 0;
        bool cancelled = false;
        string? failureSummary = null;
        try {
            await OfficeRuntimeCompatibility.ForEachAsync(SelectInputs(settings, cancellationToken),
                new ParallelOptions { MaxDegreeOfParallelism = settings.MaxParallelism, CancellationToken = cancellationToken }, async (input, token) => {
                    if (Interlocked.Increment(ref discovered) > settings.MaximumFiles) throw new InvalidDataException("Conversion selection exceeds its file limit.");
                    OfficeConversionExecutionRoute? route = SelectRoute(settings, input, routes);
                    if (route == null) {
                        Interlocked.Increment(ref skipped);
                        progress?.Report(new(input, null, null, false, "No selected executable conversion route.",
                            [new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionInputSkipped, OfficeDiagnosticSeverity.Information, "The file does not match the selected formats and executable target routes.", format: "selection")], true));
                        return;
                    }
                    Interlocked.Increment(ref selected);
                    string relative = settings.InputDirectory == null ? Path.GetFileName(input) : OfficeRuntimeCompatibility.GetRelativePath(settings.InputDirectory, input);
                    string output = Path.Combine(settings.OutputDirectory, relative + settings.TargetExtension);
                    OfficeBatchItemReport item;
                    if (settings.CheckpointDirectory == null) {
                        try {
                            EnsureNoLinks(input); EnsureNoLinks(output);
                            ValidateSelectedInput(runner, settings, route, input);
                            OfficeBatchItemReport? skip = OfficeBatchOutputSelection.ObserveSkip(input, output, settings.SkipExistingOutputs, settings.SkipUpToDate);
                            if (skip != null) item = skip;
                            else {
                                OfficeConversionExecutionResult result = await runner.RunAsync(new OfficeConversionExecutionRequest {
                                    InputPath = input, OutputPath = output,
                                    ConversionRouteId = route.Id, ConversionOptions = settings.ConversionOptions.ForRoute(route),
                                    OutputProfile = settings.OutputProfile, ConflictPolicy = settings.ConflictPolicy,
                                    PublicationGuard = new DirectBatchPublicationGuard(outputPhysicalRoot, publicationGuard),
                                    Limits = new OfficeOperationLimits { MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes,
                                        MaximumXmlCharactersInPart = settings.MaximumXmlCharactersInPart }
                                }, cancellationToken: token).ConfigureAwait(false);
                                item = new(input, result.OutputPath, result.Status, false, result.Summary, result.Diagnostics) { Report = result.Report };
                            }
                        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException and not OperationCanceledException) {
                            item = new(input, null, OfficeBatchItemStatus.Failed, false, error.Message, []);
                        }
                    } else {
                        string key = Hash(OfficePathIdentity.GetPathIdentityKey(input));
                        string receiptPath = Path.Combine(settings.CheckpointDirectory, "items", key.Substring(0, 2), key + ".json");
                        item = await RunItemAsync(runner, settings, configuration, route, input, output, receiptPath, token, publicationGuard).ConfigureAwait(false);
                    }
                    if (item.Skipped) Interlocked.Increment(ref skipped);
                    else if (item.Status == OfficeBatchItemStatus.Completed) {
                        Interlocked.Increment(ref completed);
                        if (item.Reused) Interlocked.Increment(ref reused);
                    } else if (item.Status == OfficeBatchItemStatus.Cancelled) cancelled = true;
                    else Interlocked.Increment(ref failed);
                    progress?.Report(item);
                }).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { cancelled = true; }
        catch (Exception error) when (error is InvalidDataException or IOException or UnauthorizedAccessException) {
            failureSummary = BoundMessage(error.Message);
        }
        // A terminal loop failure cancels sibling work through its internal token.
        // Report that as caller cancellation only when the caller also cancelled.
        cancelled &= failureSummary == null || cancellationToken.IsCancellationRequested;
        return new OfficeBatchSummary(selected, completed, reused, failed, cancelled, skipped) { FailureSummary = failureSummary };
    }

    private sealed class DirectBatchPublicationGuard(string outputPhysicalRoot, IOfficePublicationGuard? hostGuard)
        : IOfficePublicationGuard, IOfficeStagingGuard {
        public Task EnsureStagingDirectoryAllowedAsync(string directory, CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            EnsureNoLinks(directory);
            if (!OfficePathIdentity.IsSameOrDescendant(directory, outputPhysicalRoot))
                throw new UnauthorizedAccessException("Batch staging directory escaped the selected output root.");
            return Task.CompletedTask;
        }
        public async Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) {
            EnsureNoLinks(absoluteDestination);
            return OfficePathIdentity.IsSameOrDescendant(absoluteDestination, outputPhysicalRoot) &&
                (hostGuard == null || await hostGuard.CanPublishAsync(absoluteDestination, isDirectory, cancellationToken).ConfigureAwait(false));
        }
    }

    private static async Task<OfficeBatchItemReport> RunItemAsync(IOfficeConversionExecutor runner, OfficeBatchOptions settings, string configuration, OfficeConversionExecutionRoute route,
        string input, string output, string receiptPath, CancellationToken token, IOfficePublicationGuard? publicationGuard) {
        string? inputHash = null;
        string? stagingPath = null;
        bool pendingRecorded = false;
        string routeId = route.Id;
        try {
            if (OfficeConverter.FindExecutionRoute(routeId)?.Checkpointable != true)
                throw new NotSupportedException("Registered conversion routes require an ordinary batch without checkpoints; their runtime configuration cannot be fingerprinted.");
            EnsureNoLinks(input); EnsureNoLinks(output);
            ValidateSelectedInput(runner, settings, route, input);
            if (OfficeBatchOutputSelection.ObserveSkip(input, output, settings.SkipExistingOutputs, settings.SkipUpToDate) is { } skip)
                return skip;
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
            OfficeBatchReceipt? receipt = ReadState(receiptPath, OfficeBatchStateJsonContext.Default.OfficeBatchReceipt);
            if (receipt != null) {
                if (receipt.Schema != Schema || receipt.ConfigurationSha256 != configuration ||
                    receipt.Status is not OfficeBatchItemStatus.Completed and not OfficeBatchItemStatus.Failed ||
                    (receipt.InputSha256 != inputHash && !(settings.RetryFailed && receipt.Status == OfficeBatchItemStatus.Failed)))
                    throw new InvalidDataException("Batch input or checkpoint changed; no output was replaced.");
                if (receipt.Status == OfficeBatchItemStatus.Completed) {
                    await CompletePublicationAsync(settings, input, output, receiptPath, receipt, token, publicationGuard, resourceRoot, resourceIdentity).ConfigureAwait(false);
                    return new(input, output, OfficeBatchItemStatus.Completed, true, "Verified recorded batch artifact and completed publication.",
                        RestoreDiagnostics(receipt));
                }
                if (!settings.RetryFailed) return new(input, output, receipt.Status, false, receipt.Summary,
                    RestoreDiagnostics(receipt));
            }
            if (File.Exists(output)) throw new InvalidDataException("An output exists without a verified completion receipt. Check it before retrying.");
            string stageId = Guid.NewGuid().ToString("N");
            stagingPath = GetPendingStagePath(output, stageId);
            OfficeConversionExecutionResult result = await runner.RunAsync(new OfficeConversionExecutionRequest {
                InputPath = input, OutputPath = stagingPath,
                InputStream = routeId is "html-pdf" or "markdown-pdf" ? null : new OfficeConversionExecutionInput(Path.GetFileName(input), readToken => {
                    readToken.ThrowIfCancellationRequested();
                    return Task.FromResult<Stream>(OfficePathIdentity.OpenRegularFileForRead(input, sourceRoot, 81920));
                }, inputHash),
                ConversionRouteId = routeId,
                ConversionOptions = settings.ConversionOptions.ForRoute(routeId), OutputProfile = settings.OutputProfile, ConflictPolicy = OfficePublicationPolicy.Fail,
                Limits = new OfficeOperationLimits { MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes,
                    MaximumXmlCharactersInPart = settings.MaximumXmlCharactersInPart },
                PublicationGuard = new BatchPublicationGuard(settings, input, inputHash, publicationGuard, output, stagingPath, resourceRoot, resourceIdentity)
            }, cancellationToken: token).ConfigureAwait(false);
            OfficeConversionReport? publicReport = result.Report == null ? null : new OfficeConversionReport(
                result.Report.Source, result.Report.Target, result.Succeeded ? output : null,
                result.Report.ItemsWritten, result.Report.Elapsed, result.Report.Diagnostics);
            if (result.Status == OfficeBatchItemStatus.Cancelled) return new(input, output, result.Status, false, result.Summary, result.Diagnostics) { Report = publicReport };
            string? outputHash = result.Succeeded
                ? await HashFileAsync(stagingPath, settings.OutputDirectory, settings.MaximumOutputBytes, CancellationToken.None).ConfigureAwait(false) : null;
            if (result.Succeeded) {
                // Persist the artifact before its durable publication intent. Both names
                // are siblings, so final publication remains one filesystem move.
                using var staged = new FileStream(stagingPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                staged.Flush(flushToDisk: true);
            }
            var completedReceipt = new OfficeBatchReceipt(Schema, configuration, inputHash, outputHash, result.Status,
                result.Summary.Length > 4096 ? result.Summary.Substring(0, 4096) : result.Summary,
                result.Diagnostics.Where(IsRecordedDiagnostic)
                    .Take(32).Select(diagnostic => new OfficeBatchStoredDiagnostic(diagnostic.Code,
                        BoundMessage(diagnostic.Message), diagnostic.Severity, diagnostic.Format,
                        diagnostic.Format == null ? null : BoundMessage(diagnostic.Format),
                        diagnostic.LossKind.ToString(), diagnostic.Location == null ? null : BoundMessage(diagnostic.Location))).ToArray(),
                result.Diagnostics.Count(IsRecordedDiagnostic), result.Succeeded ? stageId : null);
            StoreReceipt(receiptPath, completedReceipt);
            pendingRecorded = result.Succeeded;
            if (result.Succeeded)
                await CompletePublicationAsync(settings, input, output, receiptPath, completedReceipt, token, publicationGuard, resourceRoot, resourceIdentity).ConfigureAwait(false);
            return new(input, output, result.Status, false, result.Summary, result.Diagnostics) { Report = publicReport };
        } catch (OperationCanceledException) when (token.IsCancellationRequested) {
            return new(input, output, OfficeBatchItemStatus.Cancelled, false, "Batch item cancelled; completed outputs remain available.", []);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            // Never replace an earlier receipt after changed-source or output-integrity failures.
            return new(input, output, OfficeBatchItemStatus.Failed, false, error.Message, []);
        } finally {
            if (!pendingRecorded && stagingPath != null && File.Exists(stagingPath)) {
                EnsureNoLinks(stagingPath);
                File.Delete(stagingPath);
            }
        }
    }

    private static string BoundMessage(string message) => message.Length > 1024 ? message.Substring(0, 1024) : message;
    private static bool IsRecordedDiagnostic(OfficeDiagnostic diagnostic) =>
        diagnostic.Severity != OfficeDiagnosticSeverity.Information ||
        diagnostic.LossKind != OfficeConversionLossKind.None;

    private static void StoreReceipt(string path, OfficeBatchReceipt receipt) {
        // Escaping can multiply non-ASCII/control text size. Bound the encoded checkpoint,
        // not just character counts, so diagnostics cannot strand a published output.
        while (receipt.Diagnostics.Length > 0 && JsonSerializer.SerializeToUtf8Bytes(receipt,
            OfficeBatchStateJsonContext.Default.OfficeBatchReceipt).Length > MaximumStateBytes)
            receipt = receipt with { Diagnostics = receipt.Diagnostics.Take(receipt.Diagnostics.Length - 1).ToArray() };
        WriteState(path, receipt, OfficeBatchStateJsonContext.Default.OfficeBatchReceipt);
    }

    private static IReadOnlyList<OfficeDiagnostic> RestoreDiagnostics(OfficeBatchReceipt receipt) {
        var diagnostics = receipt.Diagnostics.Select(finding => new OfficeDiagnostic(finding.Code, finding.Severity,
            finding.Message, finding.Location, finding.Source ?? finding.Stage,
            Enum.TryParse(finding.LossKind, out OfficeConversionLossKind loss) ? loss : OfficeConversionLossKind.None)).ToList();
        if (receipt.DiagnosticCount > receipt.Diagnostics.Length)
            diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.RecordedDiagnosticsTruncated, OfficeDiagnosticSeverity.Warning,
                $"The checkpoint retains {receipt.Diagnostics.Length} of {receipt.DiagnosticCount} diagnostics.", format: "resume"));
        return diagnostics;
    }

    private sealed class BatchPublicationGuard(OfficeBatchOptions settings, string input, string hash,
        IOfficePublicationGuard? hostGuard, string output, string stagingPath, string? resourceRoot, string? resourceIdentity) : IOfficePublicationGuard {
        public async Task<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) {
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
