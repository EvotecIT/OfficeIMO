using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Single owner for parser-backed semantic reconstruction.</summary>
internal static class PdfDocumentReadEngine {
    internal static PdfDocumentReadResult Read(
        PdfDocument source,
        PdfReadOptions options,
        CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();

        return Read(source.GetReadDocument(), options, cancellationToken);
    }

    internal static PdfDocumentReadResult Read(
        PdfReadDocument document,
        PdfReadOptions options,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        int[] pageNumbers = options.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(options.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(options.Pipeline);
        IReadOnlyList<PdfUnderstandingPageResult> analyses = new PdfUnderstandingPipeline(
                options.LayoutOptions,
                pipelineOptions)
            .RunPages(document, pageNumbers, cancellationToken);
        if (options.Profile == PdfReadProfile.Structured) {
            analyses = PdfDocumentSemanticEnricher.Enrich(
                document,
                pageNumbers,
                analyses,
                pipelineOptions.MaxRegionsPerPage,
                pipelineOptions.MaxDocumentWorkUnits,
                cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();

        return PdfDocumentReadResult.FromPageNumbers(
            document,
            options.LayoutOptions,
            pageNumbers,
            analyses,
            options.Profile,
            cancellationToken);
    }
}
