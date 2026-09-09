using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed class PreparedTextEdit {
    internal required long Revision { get; init; }
    internal required string SourceHash { get; init; }
    internal required byte[] SourceBytes { get; init; }
    internal required byte[] OutputBytes { get; init; }
    internal required IReadOnlyList<string> Warnings { get; init; }
    internal required int AffectedCount { get; init; }
    internal required int[] Pages { get; init; }
}

internal sealed partial class PdfWorkspace {
    internal Task<PdfTextMatch> InspectSelectedTextAsync(PdfEditorSelection selection, CancellationToken token) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        return RunCancellableCpuWorkAsync(() => ResolveTextMatch(snapshot, selection), token);
    }

    internal Task<IReadOnlyList<PdfTextMatch>> FindReplacementMatchesAsync(string text, bool matchCase,
        bool wholeWords, CancellationToken token) {
        PdfDocument snapshot = CreateDocumentSnapshot();
        return RunCancellableCpuWorkAsync(() => snapshot.Text.Find(text,
            new PdfTextSearchOptions { MatchCase = matchCase, WholeWords = wholeWords }), token);
    }

    internal Task<PreparedTextEdit> PreviewSelectedTextAsync(PdfEditorSelection selection, string replacement,
        PdfTextEditOptions options, CancellationToken token) => PrepareTextEditAsync(document => {
            PdfTextMatch match = ResolveTextMatch(document, selection);
            return document.Text.Replace(match, replacement, options);
        }, [selection.PageNumber], token);

    internal Task<PreparedTextEdit> PreviewTextReplacementsAsync(string find, string replacement, bool matchCase,
        bool wholeWords, IReadOnlyCollection<int> indexes, PdfTextEditOptions options, int[] pages, CancellationToken token) {
        int[] selection = indexes.ToArray();
        return PrepareTextEditAsync(document => document.Text.ReplaceSelected(find, replacement, selection,
            new PdfTextSearchOptions { MatchCase = matchCase, WholeWords = wholeWords }, options), pages, token);
    }

    private Task<PreparedTextEdit> PrepareTextEditAsync(Func<PdfDocument, PdfTextEditResult> edit, int[] pages, CancellationToken token) {
        byte[] source = CopyBytes();
        long revision = Revision;
        return RunCancellableCpuWorkAsync(() => {
            token.ThrowIfCancellationRequested();
            PdfTextEditResult result = edit(LoadDocument(source));
            byte[] output = result.Document.ToBytes();
            token.ThrowIfCancellationRequested();
            return new PreparedTextEdit {
                Revision = revision, SourceHash = Convert.ToHexString(SHA256.HashData(source)),
                SourceBytes = source, OutputBytes = output, Warnings = result.Warnings,
                AffectedCount = result.AffectedCount, Pages = pages.Distinct().Order().ToArray()
            };
        }, token);
    }

    internal Task ApplyPreparedTextEditAsync(PreparedTextEdit prepared, CancellationToken token,
        IProgress<PdfWorkspaceProgress>? progress = null) => MutateBytesAsync(PdfWorkspaceOperationKind.TextEdit,
        "Applied reviewed text replacements", prepared.Pages, bytes => {
            if (Revision != prepared.Revision || !string.Equals(Convert.ToHexString(SHA256.HashData(bytes)), prepared.SourceHash, StringComparison.Ordinal)) {
                throw new InvalidOperationException("The document changed after the text preview. Review the replacement again.");
            }
            return prepared.OutputBytes.ToArray();
        }, token, progress);

    internal Task<(byte[] Before, byte[] After)> RenderTextPreviewAsync(PreparedTextEdit prepared, int page, CancellationToken token) =>
        RunNonDetachableCpuWorkAsync(() => RenderTextPreviewPairAsync(prepared, page, token).GetAwaiter().GetResult(), token);

    private async Task<(byte[] Before, byte[] After)> RenderTextPreviewPairAsync(PreparedTextEdit prepared, int page, CancellationToken token) {
        var options = new PdfImageExportOptions { Scale = 2, ThumbnailMaxDimension = 1600, MaximumOutputCount = 1 };
        async Task<byte[]> RenderAsync(byte[] bytes) {
            IReadOnlyList<OfficeImageExportResult> result = await LoadDocument(bytes).ToImages(options)
                .Pages(PdfPageSelection.From(page)).AsPng().ExportAsync(token).ConfigureAwait(false);
            return result.Single().Bytes;
        }
        byte[] before = await RenderAsync(prepared.SourceBytes).ConfigureAwait(false);
        byte[] after = await RenderAsync(prepared.OutputBytes).ConfigureAwait(false);
        return (before, after);
    }
}
