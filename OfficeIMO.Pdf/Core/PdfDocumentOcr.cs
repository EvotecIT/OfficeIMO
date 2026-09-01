using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>Caller-provider OCR enrichment for a loaded PDF document.</summary>
public sealed class PdfDocumentOcr {
    private readonly PdfDocument _document;

    internal PdfDocumentOcr(PdfDocument document) => _document = document;

    /// <summary>
    /// Renders the selected pages, invokes the caller-owned OCR provider, and merges accepted words
    /// into the same logical result contract returned by <see cref="PdfDocument.Read"/>.
    /// </summary>
    public Task<PdfOcrMergeResult> ReadAsync(
        IPdfOcrProvider provider,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) =>
        PdfOcr.RecognizeAndMergeAsync(
            _document.GetBytesForOperation(),
            provider,
            options,
            _document.ReadOptions,
            cancellationToken);

    /// <summary>
    /// Recognizes selected pages and returns a PDF with geometry-aligned invisible text for accepted OCR words.
    /// Pages without accepted OCR content are left unchanged. Existing digital signatures may block the required
    /// full rewrite through the ordinary PDF mutation policy.
    /// </summary>
    public async Task<PdfSearchableOcrResult> MakeSearchableAsync(
        IPdfOcrProvider provider,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        PdfOcrMergeResult ocr = await ReadAsync(provider, options, cancellationToken).ConfigureAwait(false);
        int[] modifiedPages = ocr.Pages
            .Where(static page => page.Words.Count > 0)
            .Select(static page => page.PageNumber)
            .Distinct()
            .ToArray();
        if (modifiedPages.Length == 0) {
            return new PdfSearchableOcrResult(_document, ocr, Array.Empty<int>());
        }

        var wordsByPage = ocr.Pages
            .Where(static page => page.Words.Count > 0)
            .GroupBy(static page => page.PageNumber)
            .ToDictionary(static pages => pages.Key, static pages => pages.First().Words);
        string pageSelector = string.Join(",", modifiedPages.Select(static page => page.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PdfDocument searchable = _document.Stamp.Content(
            (canvas, context) => {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<PdfRecognizedWord> words = wordsByPage[context.PageNumber];
                for (int i = 0; i < words.Count; i++) {
                    PdfRecognizedWord word = words[i];
                    canvas.SearchableText(word.Text, word.X, word.Y, word.Width, word.Height);
                }
            },
            new PdfCanvasStampOptions().UseTargetPages(pageSelector),
            _document.ReadOptions);
        return new PdfSearchableOcrResult(searchable, ocr, Array.AsReadOnly(modifiedPages));
    }
}
