using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>What making a document searchable added: the pages that gained text and their recognized words.</summary>
internal sealed record PdfWorkspaceSearchableResult(int ModifiedPages, int Words);

internal sealed partial class PdfWorkspace {
    /// <summary>
    /// Recognizes the pages' text with <paramref name="engine"/> and adds it as invisible text over the page images, as
    /// one undoable change. Words over text the page already has are left out, so digital pages stay as they are.
    /// It works on a snapshot of the current revision and fails if the document changes meanwhile.
    /// </summary>
    internal async Task<PdfWorkspaceSearchableResult> MakeSearchableAsync(IOcrEngine engine, PdfOcrMergeOptions options,
        CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(engine);
        ArgumentNullException.ThrowIfNull(options);
        if (!CanEditPageContent) throw new NotSupportedException("Text can't be added to this document's pages safely.");
        var capturedOptions = options.Clone();
        // The native host owns the recognizer context only until this operation returns.
        capturedOptions.AwaitProviderSettlement = true;
        long revision;
        byte[] bytes;
        await _operationGate.WaitAsync(token).ConfigureAwait(false);
        try {
            ThrowIfDisposed(); token.ThrowIfCancellationRequested();
            revision = Revision;
            bytes = _bytes;
        } finally { _operationGate.Release(); }
        PdfSearchableOcrResult result = await RunNonDetachableCpuWorkAsync(
            () => LoadDocument(bytes).MakeSearchableAsync(engine, capturedOptions, token).GetAwaiter().GetResult(), token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        if (!result.WasModified) return new(0, 0);
        byte[] searchable = result.Document.ToBytes();
        await MutateBytesAsync(PdfWorkspaceOperationKind.AddedContent, $"Made {result.ModifiedPages.Count} pages searchable", result.ModifiedPages,
            _ => Revision == revision ? searchable : throw new InvalidOperationException("The document changed during recognition. Make it searchable again."),
            token, progress).ConfigureAwait(false);
        return new(result.ModifiedPages.Count, result.AddedWordCount);
    }
}
