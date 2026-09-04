using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Optional OCR operations for loaded PDF documents.</summary>
public static class PdfOcrExtensions {
    /// <summary>
    /// Renders selected pages, invokes an engine-neutral OCR provider, and merges accepted spans into the
    /// same logical result contract returned by <see cref="PdfDocument.Read"/>.
    /// </summary>
    public static Task<PdfOcrMergeResult> ReadWithOcrAsync(
        this PdfDocument document,
        IOcrEngine engine,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return PdfOcr.RecognizeAndMergeAsync(
            document.GetBytesForOperation(cancellationToken),
            engine,
            options,
            document.ReadOptions,
            cancellationToken);
    }

    /// <summary>
    /// Returns a PDF with geometry-aligned invisible text for accepted OCR spans. Pages without accepted OCR
    /// content remain unchanged. Existing digital signatures may block the required full rewrite.
    /// </summary>
    public static async Task<PdfSearchableOcrResult> MakeSearchableAsync(
        this PdfDocument document,
        IOcrEngine engine,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        PdfOcrMergeOptions effectiveOptions = options?.Clone() ?? new PdfOcrMergeOptions();
        PdfPageSelection? selection = effectiveOptions.ReadOptions.PageSelection;
        if (selection != null) {
            int pageCount = document.Inspect(document.ReadOptions, cancellationToken).PageCount;
            int[] uniquePages = selection
                .ToPageNumbers(pageCount, nameof(options))
                .Distinct()
                .ToArray();
            effectiveOptions.ReadOptions = PdfReadOptions.WithPageSelection(
                effectiveOptions.ReadOptions,
                PdfPageSelection.From(uniquePages));
        }

        PdfOcrMergeResult ocr = await document.ReadWithOcrAsync(engine, effectiveOptions, cancellationToken).ConfigureAwait(false);
        int[] modifiedPages = ocr.Pages
            .Where(static page => page.Words.Count > 0)
            .Select(static page => page.PageNumber)
            .Distinct()
            .ToArray();
        if (modifiedPages.Length == 0) {
            return new PdfSearchableOcrResult(document, ocr, Array.Empty<int>());
        }

        var wordsByPage = ocr.Pages
            .Where(static page => page.Words.Count > 0)
            .GroupBy(static page => page.PageNumber)
            .ToDictionary(static pages => pages.Key, static pages => pages.First().Words);
        string pageSelector = string.Join(",", modifiedPages.Select(static page => page.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PdfDocument searchable = document.Stamp.Content(
            (canvas, context) => {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<PdfRecognizedWord> words = wordsByPage[context.PageNumber];
                PdfLogicalPage canonicalPage = ocr.Document.Pages.First(page => page.PageNumber == context.PageNumber);
                IReadOnlyList<PdfRecognizedWord> logicalWords = PdfOcrLogicalDocumentBuilder.OrderWordsForLogicalReading(
                    words,
                    canonicalPage,
                    effectiveOptions.ReadOptions.LayoutOptions.ReadingDirection,
                    cancellationToken);
                for (int index = 0; index < logicalWords.Count; index++) {
                    PdfRecognizedWord word = logicalWords[index];
                    canvas.SearchableText(word.Text, word.X, word.Y, word.Width, word.Height);
                }
            },
            new PdfCanvasStampOptions().UseTargetPages(pageSelector),
            document.ReadOptions);
        return new PdfSearchableOcrResult(searchable, ocr, Array.AsReadOnly(modifiedPages));
    }
}
