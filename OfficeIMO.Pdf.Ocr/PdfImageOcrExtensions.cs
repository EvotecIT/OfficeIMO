using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Recognizes standalone raster images through the canonical positioned-text reconstruction pipeline.</summary>
/// <remarks>Recognition uses the source resolution, rather than rendering a PDF intermediate. The resulting
/// logical document can be passed to existing Word, Excel, HTML and other format adapters. Animated and
/// multi-page images are rejected; supply separate static sources when every frame or page is required.</remarks>
public static class PdfImageOcrExtensions {
    /// <summary>Recognizes a captured image and reconstructs editable text and tables without writing an output file.</summary>
    /// <remarks>Null options enable layout reconstruction. Supplied options retain their explicit settings.
    /// Recognition errors and low-confidence exclusions remain in the result and require caller review.</remarks>
    public static async Task<PdfOcrMergeResult> ReadWithOcrAsync(
        this PdfImageDocumentSource image, IOcrEngine engine, PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) =>
        (await image.PrepareSearchableOcrAsync(engine, options, cancellationToken).ConfigureAwait(false)).Ocr;

    /// <summary>Recognizes an image and retains a private image-page snapshot for review and searchable PDF generation.</summary>
    public static Task<PdfSearchableOcrReview> PrepareSearchableOcrAsync(
        this PdfImageDocumentSource image, IOcrEngine engine, PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) =>
        PdfOcr.RecognizeImageAsync(image, engine, options, cancellationToken);

    /// <summary>Creates an in-memory searchable PDF preserving the image and adding all accepted OCR words.</summary>
    /// <remarks>Inspect the returned recognition evidence before publishing; successful composition does not
    /// establish recognition accuracy. No source file is changed.</remarks>
    public static async Task<PdfSearchableOcrResult> MakeSearchableAsync(
        this PdfImageDocumentSource image, IOcrEngine engine, PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) =>
        (await image.PrepareSearchableOcrAsync(engine, options, cancellationToken).ConfigureAwait(false))
            .ApplyAll(cancellationToken);
}
