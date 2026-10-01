using System.Text.Json;
using OfficeIMO.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Restartable DOC/DOCX/TXT archive conversion. Outputs are per-item, source-bound and never overwritten.</summary>
public static partial class OfficePdfArchiveWorkflow {
    /// <summary>Runs a bounded local archive, reusing only hash-verified completed work under the same captured configuration.</summary>
    public static async Task<OfficePdfArchiveResult> RunAsync(OfficePdfArchiveRequest request,
        IProgress<OfficePdfArchiveItemResult>? progress = null, CancellationToken cancellationToken = default,
        IOfficeWorkflowPublicationGuard? publicationGuard = null) {
        OfficePdfArchiveRequest settings = Snapshot(request);
        cancellationToken.ThrowIfCancellationRequested();
        if (publicationGuard != null &&
            (!await publicationGuard.CanPublishAsync(settings.OutputDirectory, true, cancellationToken).ConfigureAwait(false) ||
             !await publicationGuard.CanPublishAsync(settings.CheckpointDirectory, true, cancellationToken).ConfigureAwait(false)))
            throw new UnauthorizedAccessException("Archive output or checkpoint directory is protected by the host publication policy.");
        Directory.CreateDirectory(settings.OutputDirectory);
        Directory.CreateDirectory(settings.CheckpointDirectory);
        string lockPath = Path.Combine(settings.CheckpointDirectory, "archive.lock");
        EnsureNoLinks(lockPath);
        using var lease = new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        OfficePdfArchiveRequest identity = settings with { MaximumConcurrency = 1, RetryFailed = false };
        string configuration = Hash(JsonSerializer.Serialize(identity, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveRequest) +
            typeof(OfficePdfArchiveWorkflow).Assembly.ManifestModule.ModuleVersionId +
            typeof(PdfDocument).Assembly.ManifestModule.ModuleVersionId + typeof(OfficeIMO.Word.WordDocument).Assembly.ManifestModule.ModuleVersionId +
            typeof(OfficeIMO.Word.Pdf.LegacyDocPdfConverter).Assembly.ManifestModule.ModuleVersionId +
            OfficePathIdentity.GetPathIdentityKey(settings.InputDirectory));
        string planPath = Path.Combine(settings.CheckpointDirectory, "plan.json");
        OfficePdfArchivePlan? plan = ReadState(planPath, OfficePdfArchiveJsonContext.Default.OfficePdfArchivePlan);
        if (plan != null && (plan.Schema != Schema || plan.ConfigurationSha256 != configuration))
            throw new InvalidDataException("Archive configuration or engine changed. Use a new checkpoint and output directory.");
        if (plan == null) WriteState(planPath, new OfficePdfArchivePlan(Schema, configuration), OfficePdfArchiveJsonContext.Default.OfficePdfArchivePlan);
        long selected = 0, completed = 0, reused = 0, failed = 0;
        bool cancelled = false;
        try {
            await Parallel.ForEachAsync(Discover(settings.InputDirectory, cancellationToken),
                new ParallelOptions { MaxDegreeOfParallelism = settings.MaximumConcurrency, CancellationToken = cancellationToken }, async (input, token) => {
                    if (Interlocked.Increment(ref selected) > settings.MaximumFiles) throw new InvalidDataException("Archive selection exceeds its file limit.");
                    string relative = Path.GetRelativePath(settings.InputDirectory, input);
                    string output = Path.Combine(settings.OutputDirectory, relative + ".pdf");
                    string key = Hash(relative.Replace(Path.DirectorySeparatorChar, '/'));
                    string receiptPath = Path.Combine(settings.CheckpointDirectory, "items", key[..2], key + ".json");
                    OfficePdfArchiveItemResult item = await RunItemAsync(settings, configuration, input, output, receiptPath, token, publicationGuard).ConfigureAwait(false);
                    if (item.Status == OfficeWorkflowStatus.Completed) {
                        Interlocked.Increment(ref completed);
                        if (item.Reused) Interlocked.Increment(ref reused);
                    } else if (item.Status == OfficeWorkflowStatus.Cancelled) cancelled = true;
                    else Interlocked.Increment(ref failed);
                    progress?.Report(item);
                }).ConfigureAwait(false);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { cancelled = true; }
        if (selected == 0 && !cancelled) throw new InvalidDataException("The archive contains no DOC, DOCX or TXT files.");
        return new OfficePdfArchiveResult(selected, completed, reused, failed, cancelled);
    }

    private static async Task<OfficePdfArchiveItemResult> RunItemAsync(OfficePdfArchiveRequest settings, string configuration,
        string input, string output, string receiptPath, CancellationToken token, IOfficeWorkflowPublicationGuard? publicationGuard) {
        string? inputHash = null;
        try {
            EnsureNoLinks(input); EnsureNoLinks(output);
            inputHash = await HashFileAsync(input, settings.InputDirectory, settings.MaximumInputBytes, token).ConfigureAwait(false);
            OfficePdfArchiveReceipt? receipt = ReadState(receiptPath, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveReceipt);
            if (receipt != null) {
                if (receipt.Schema != Schema || receipt.ConfigurationSha256 != configuration ||
                    receipt.Status is not OfficeWorkflowStatus.Completed and not OfficeWorkflowStatus.Failed ||
                    (receipt.InputSha256 != inputHash && !(settings.RetryFailed && receipt.Status == OfficeWorkflowStatus.Failed)))
                    throw new InvalidDataException("Archive input or checkpoint changed; no output was replaced.");
                if (receipt.Status == OfficeWorkflowStatus.Completed) {
                    if (receipt.OutputSha256 == null || !File.Exists(output) ||
                        await HashFileAsync(output, settings.OutputDirectory, settings.MaximumOutputBytes, token).ConfigureAwait(false) != receipt.OutputSha256)
                        throw new InvalidDataException("Completed archive output is missing or changed; no automatic retry is allowed.");
                    return new(input, output, OfficeWorkflowStatus.Completed, true, "Verified previously completed output.",
                        RestoreDiagnostics(receipt));
                }
                if (!settings.RetryFailed) return new(input, output, receipt.Status, true, receipt.Summary,
                    RestoreDiagnostics(receipt));
            }
            if (File.Exists(output)) throw new InvalidDataException("An output exists without a verified completion receipt. Check it before retrying.");
            string extension = Path.GetExtension(input).ToLowerInvariant();
            var options = new OfficeWorkflowConversionOptions {
                LegacyDocLossPolicy = settings.AllowLegacyImportLoss && extension == ".doc" ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block,
                PlainText = extension == ".txt" ? new PdfPlainTextOptions {
                    EncodingName = settings.TextEncoding, TabSize = settings.TabSize,
                    MaximumCharacters = settings.MaximumTextCharacters, MaximumPages = settings.MaximumTextPages
                } : null
            };
            OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
                InputPath = input, OutputPath = output, Operation = OfficeWorkflowOperation.Convert,
                InputStream = new OfficeWorkflowStreamInput(Path.GetFileName(input), readToken => {
                    readToken.ThrowIfCancellationRequested();
                    return Task.FromResult<Stream>(OfficePathIdentity.OpenRegularFileForRead(input, settings.InputDirectory, 81920));
                }, inputHash),
                ConversionRouteId = extension == ".doc" ? "doc-pdf" : extension == ".txt" ? "txt-pdf" : "docx-pdf",
                ConversionOptions = options, ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
                Limits = new OfficeWorkflowLimits { MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes },
                PublicationGuard = new ArchivePublicationGuard(settings, input, inputHash, publicationGuard)
            }, cancellationToken: token).ConfigureAwait(false);
            if (result.Status == OfficeWorkflowStatus.Cancelled) return new(input, output, result.Status, false, result.Summary, result.Diagnostics);
            string? outputHash = result.Succeeded
                ? await HashFileAsync(output, settings.OutputDirectory, settings.MaximumOutputBytes, CancellationToken.None).ConfigureAwait(false) : null;
            StoreReceipt(receiptPath, new OfficePdfArchiveReceipt(Schema, configuration, inputHash, outputHash, result.Status,
                result.Summary.Length > 4096 ? result.Summary[..4096] : result.Summary,
                result.Diagnostics.Where(IsRecordedDiagnostic)
                    .Take(32).Select(diagnostic => new OfficePdfArchiveStoredDiagnostic(diagnostic.Code,
                        BoundMessage(diagnostic.Message), diagnostic.Severity, diagnostic.Stage,
                        diagnostic.Details.TryGetValue("source", out string? source) ? BoundMessage(source) : null,
                        diagnostic.Details.TryGetValue("lossKind", out string? loss) ? loss : null)).ToArray(),
                result.Diagnostics.Count(IsRecordedDiagnostic)));
            return new(input, output, result.Status, false, result.Summary, result.Diagnostics);
        } catch (OperationCanceledException) when (token.IsCancellationRequested) {
            return new(input, output, OfficeWorkflowStatus.Cancelled, false, "Archive item cancelled; completed outputs remain available.", []);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            // Never replace an earlier receipt after changed-source or output-integrity failures.
            return new(input, output, OfficeWorkflowStatus.Failed, false, error.Message, []);
        }
    }

    private static string BoundMessage(string message) => message.Length > 1024 ? message[..1024] : message;
    private static bool IsRecordedDiagnostic(OfficeWorkflowDiagnostic diagnostic) =>
        diagnostic.Severity != OfficeWorkflowDiagnosticSeverity.Information ||
        (diagnostic.Details.TryGetValue("lossKind", out string? loss) && loss != nameof(OfficeConversionLossKind.None));

    private static void StoreReceipt(string path, OfficePdfArchiveReceipt receipt) {
        // Escaping can multiply non-ASCII/control text size. Bound the encoded checkpoint,
        // not just character counts, so diagnostics cannot strand a published output.
        while (receipt.Diagnostics.Length > 0 && JsonSerializer.SerializeToUtf8Bytes(receipt,
            OfficePdfArchiveJsonContext.Default.OfficePdfArchiveReceipt).Length > MaximumStateBytes)
            receipt = receipt with { Diagnostics = receipt.Diagnostics[..^1] };
        WriteState(path, receipt, OfficePdfArchiveJsonContext.Default.OfficePdfArchiveReceipt);
    }

    private static IReadOnlyList<OfficeWorkflowDiagnostic> RestoreDiagnostics(OfficePdfArchiveReceipt receipt) {
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

    private sealed class ArchivePublicationGuard(OfficePdfArchiveRequest settings, string input, string hash, IOfficeWorkflowPublicationGuard? hostGuard) : IOfficeWorkflowPublicationGuard {
        public async ValueTask<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) {
            EnsureNoLinks(absoluteDestination);
            return !isDirectory && OfficePathIdentity.IsSameOrDescendant(absoluteDestination, settings.OutputDirectory) &&
                !OfficePathIdentity.IsSameOrDescendant(absoluteDestination, settings.InputDirectory) &&
                (hostGuard == null || await hostGuard.CanPublishAsync(absoluteDestination, false, cancellationToken).ConfigureAwait(false)) &&
                await HashFileAsync(input, settings.InputDirectory, settings.MaximumInputBytes, cancellationToken).ConfigureAwait(false) == hash;
        }
    }
}
