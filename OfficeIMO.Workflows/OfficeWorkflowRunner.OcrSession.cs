using OfficeIMO.Internal;
using OfficeIMO.Ocr;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    /// <summary>Runs an ordered OCR session with one caller-owned engine and at most one active recognition/review.</summary>
    /// <remarks>Requests are snapshotted before execution. Every selected source is protected from every output,
    /// destinations must be distinct, and failures do not repeat completed items. Cancellation returns explicit
    /// cancelled results for unstarted items; it does not roll back already published outputs. An uncertain
    /// output stops the remaining items. Pass previously completed output locations in protectedOutputPaths
    /// when retrying a subset so deferred provider destinations cannot overwrite those outputs. Pass every
    /// retained session source through protectedInputs, including provider access, when executing only a subset.</remarks>
    public async Task<IReadOnlyList<OfficeOcrSessionResult>> RunOcrSessionAsync(IEnumerable<OfficeOcrSessionRequest> requests,
        IOcrEngine engine, IProgress<OfficeWorkflowProgress>? progress = null,
        IProgress<OfficeOcrSessionResult>? resultsProgress = null, CancellationToken cancellationToken = default,
        IReadOnlyList<string>? protectedOutputPaths = null, IReadOnlyList<OfficeWorkflowProtectedSource>? protectedInputs = null) {
        ArgumentNullException.ThrowIfNull(requests);
        ArgumentNullException.ThrowIfNull(engine);
        if (protectedOutputPaths?.Count > MaximumBatchRequestCount) throw new ArgumentException("The protected output list exceeds the session item limit.");
        string[] retainedOutputs = protectedOutputPaths?.Select(OfficeStorageIdentity.Normalize).ToArray() ?? [];
        if (protectedInputs?.Count > MaximumBatchRequestCount) throw new ArgumentException("The protected input list exceeds the session item limit.");
        var retainedInputs = protectedInputs?.Select(input => {
            ArgumentNullException.ThrowIfNull(input);
            return (Location: input.Location, Stream: input.InputStream);
        }).ToArray() ?? [];
        var batch = new List<OfficeOcrSessionRequest>();
        foreach (var request in requests) {
            cancellationToken.ThrowIfCancellationRequested();
            if (batch.Count >= MaximumBatchRequestCount) throw new ArgumentException("The OCR session exceeds the batch item limit.");
            ArgumentNullException.ThrowIfNull(request);
            batch.Add(SnapshotOcrSessionItem(request));
        }
        if (batch.Select(item => item.Id).Distinct(StringComparer.Ordinal).Count() != batch.Count)
            throw new ArgumentException("OCR session item ids must be unique.");
        var sources = batch.Select(item => (
            Location: OfficeStorageIdentity.Normalize(item.Pdf?.InputPath ?? item.Image!.InputPath),
            Stream: item.Pdf?.InputStream ?? item.Image?.InputStream)).Concat(retainedInputs).ToArray();
        string[] protectedSources = sources.Select(item => item.Location).Distinct(StringComparer.Ordinal).ToArray();
        var accesses = sources.Where(item => item.Stream is not null)
            .GroupBy(item => item.Location, StringComparer.Ordinal)
            .Select(group => new WorkflowSourceAccess(group.Key, group.First().Stream!)).ToArray();
        var destinations = batch.Select(item => OfficeStorageIdentity.Normalize(item.Pdf?.OutputPath ?? item.Image!.OutputPath)).ToArray();
        for (int index = 0; index < destinations.Length; index++) {
            for (int previous = 0; previous < index; previous++) {
                if (OfficeStorageIdentity.AreEquivalent(destinations[index], destinations[previous]))
                    throw new ArgumentException("OCR session outputs must have distinct destinations.");
            }
        }
        var results = new List<OfficeOcrSessionResult>(batch.Count);
        bool stoppedForUncertainOutput = false;
        for (int index = 0; index < batch.Count; index++) {
            var item = batch[index];
            if (cancellationToken.IsCancellationRequested || stoppedForUncertainOutput) {
                results.Add(new(item.Id, OfficeWorkflowStatus.Cancelled, null,
                    stoppedForUncertainOutput ? "Not started because an earlier output needs checking. Completed outputs were retained."
                        : "Cancelled before this item started; completed outputs were retained.", []));
                resultsProgress?.Report(results[results.Count - 1]);
                continue;
            }
            progress?.Report(new OfficeWorkflowProgress(item.Id, "execute", "Recognizing and reviewing the next input.", 0, (double)index / batch.Count));
            bool dispatched = false;
            try {
                if (item.Pdf is { } pdf) {
                    var distinct = new DistinctWorkflowOutputPublicationGuard(pdf.PublicationGuard,
                        () => retainedOutputs.Concat(results.Where(result => result.OutputPath is not null).Select(result => result.OutputPath!)));
                    pdf.PublicationGuard = new WorkflowScopedSourcePublicationGuard(distinct,
                        protectedSources, accesses, pdf.OutputStream, allowMissingLocalSources: true);
                    dispatched = true;
                    var result = await MakePdfSearchableAsync(pdf, engine, cancellationToken).ConfigureAwait(false);
                    results.Add(new(item.Id, result.Status, result.OutputPath, result.Summary, result.Diagnostics, result.Recovery));
                } else {
                    var image = item.Image!;
                    var distinct = new DistinctWorkflowOutputPublicationGuard(image.PublicationGuard,
                        () => retainedOutputs.Concat(results.Where(result => result.OutputPath is not null).Select(result => result.OutputPath!)));
                    image.PublicationGuard = new WorkflowScopedSourcePublicationGuard(distinct,
                        protectedSources, accesses, image.OutputStream, allowMissingLocalSources: true);
                    dispatched = true;
                    var result = await RecognizeImageAsync(image, engine, cancellationToken).ConfigureAwait(false);
                    results.Add(new(item.Id, result.Status, result.OutputPath, result.Summary, result.Diagnostics, result.Recovery));
                }
            } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                results.Add(new(item.Id, dispatched ? OfficeWorkflowStatus.Unconfirmed
                    : cancellationToken.IsCancellationRequested ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed,
                    null, error.Message, [new OfficeWorkflowDiagnostic("OcrSessionItemFailed", error.Message, OfficeWorkflowDiagnosticSeverity.Error)]));
            }
            stoppedForUncertainOutput = results[results.Count - 1].Status == OfficeWorkflowStatus.Unconfirmed;
            resultsProgress?.Report(results[results.Count - 1]);
            progress?.Report(new OfficeWorkflowProgress(item.Id, "complete", results[results.Count - 1].Summary, 1, (double)(index + 1) / batch.Count));
        }
        return results.AsReadOnly();
    }

    private static OfficeOcrSessionRequest SnapshotOcrSessionItem(OfficeOcrSessionRequest item) {
        if (item.Pdf is { } pdf) return new(item.Id, new PdfSearchableWorkflowRequest {
            Id = pdf.Id, InputPath = pdf.InputPath, InputStream = pdf.InputStream, OutputPath = pdf.OutputPath,
            OutputStream = pdf.OutputStream, ConflictPolicy = pdf.ConflictPolicy, PublicationGuard = pdf.PublicationGuard,
            Limits = pdf.Limits.CloneAndValidate(), PdfPassword = pdf.PdfPassword, Ocr = pdf.Ocr.Clone(), ReviewAsync = pdf.ReviewAsync
        });
        var image = item.Image!;
        return new(item.Id, new ImageOcrWorkflowRequest {
            InputPath = image.InputPath, InputStream = image.InputStream, OutputPath = image.OutputPath,
            OutputStream = image.OutputStream, ConflictPolicy = image.ConflictPolicy, PublicationGuard = image.PublicationGuard,
            Limits = image.Limits.CloneAndValidate(), Ocr = image.Ocr.Clone(), ReviewAsync = image.ReviewAsync
        });
    }
}
