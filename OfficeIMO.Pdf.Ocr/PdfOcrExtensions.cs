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
        var review = await document.PrepareSearchableOcrAsync(engine, options, cancellationToken).ConfigureAwait(false);
        var result = review.ApplyAll(cancellationToken);
        return result.WasModified ? result : new PdfSearchableOcrResult(document, result.Ocr, result.ModifiedPages, result.WrittenWords);
    }

    /// <summary>Recognizes a private source snapshot and returns review evidence without modifying or publishing a PDF.</summary>
    public static async Task<PdfSearchableOcrReview> PrepareSearchableOcrAsync(
        this PdfDocument document,
        IOcrEngine engine,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument snapshot = PdfDocument.Load(document.GetBytesForOperation(cancellationToken), document.ReadOptions);
        PdfOcrMergeOptions effectiveOptions = options?.Clone() ?? new PdfOcrMergeOptions();
        PdfPageSelection? selection = effectiveOptions.ReadOptions.PageSelection;
        if (selection != null) {
            int pageCount = snapshot.Inspect(snapshot.ReadOptions, cancellationToken).PageCount;
            int[] uniquePages = selection
                .ToPageNumbers(pageCount, nameof(options))
                .Distinct()
                .ToArray();
            effectiveOptions.ReadOptions = PdfReadOptions.WithPageSelection(
                effectiveOptions.ReadOptions,
                PdfPageSelection.From(uniquePages));
        }

        PdfOcrMergeResult ocr = await snapshot.ReadWithOcrAsync(engine, effectiveOptions, cancellationToken).ConfigureAwait(false);
        return new PdfSearchableOcrReview(snapshot, effectiveOptions, ocr);
    }
}
