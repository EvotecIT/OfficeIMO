using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Single owner for parser-backed semantic reconstruction.</summary>
internal static class PdfDocumentReadEngine {
    internal static PdfLogicalPage ReadLogicalPage(
        PdfReadDocument document,
        int pageNumber,
        IReadOnlyDictionary<int, IReadOnlyList<PdfLogicalFormWidget>> formWidgetsByPageNumber,
        CancellationToken cancellationToken) {
        document.DemandContentExtraction("logical object");
        cancellationToken.ThrowIfCancellationRequested();
        var layout = new PdfTextLayoutOptions();
        var pipeline = new PdfUnderstandingPipeline(layout, new PdfUnderstandingPipelineOptions { MaxPages = 1 });
        PdfUnderstandingPageResult analysis = pipeline.RunPages(document, new[] { pageNumber }, cancellationToken)[0];
        formWidgetsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfLogicalFormWidget>? pageFormWidgets);
        PdfLogicalPage page = PdfLogicalPage.From(document, document.Pages[pageNumber - 1], pageNumber,
            layout, pageFormWidgets, analysis, cancellationToken);
        analysis.CompleteOperation();
        return page;
    }

    internal static PdfDocumentReadResult Read(
        PdfDocument source,
        PdfReadOptions options,
        CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();

        return Read(source.GetReadDocument(source.ReadOptions, cancellationToken), options, cancellationToken);
    }

    internal static PdfDocumentReadResult Read(
        PdfReadDocument document,
        PdfReadOptions options,
        CancellationToken cancellationToken = default) =>
        Read(document, options, out _, cancellationToken);

    internal static PdfDocumentReadResult Read(
        PdfReadDocument document,
        PdfReadOptions options,
        out IReadOnlyList<PdfUnderstandingPageResult> pageAnalyses,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        int[] pageNumbers = options.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(options.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(options.Pipeline);
        pageAnalyses = new PdfUnderstandingPipeline(
                options.LayoutOptions,
                pipelineOptions)
            .RunPages(document, pageNumbers, cancellationToken);
        IReadOnlyList<PdfUnderstandingPageResult> analyses = pageAnalyses;
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
        PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(
            document,
            options.LayoutOptions,
            pageNumbers,
            analyses,
            options.Profile,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        for (int pageIndex = 0; pageIndex < analyses.Count; pageIndex++) {
            analyses[pageIndex].CompleteOperation();
        }
        return result;
    }
}
