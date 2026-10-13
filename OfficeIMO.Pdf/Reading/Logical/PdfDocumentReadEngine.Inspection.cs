using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentReadEngine {
    internal static PdfInspectionReadResult ReadForInspection(PdfReadDocument document, PdfDocumentInfo sourceInfo,
        PdfReadOptions options, CancellationToken cancellationToken) {
        int[] selected = options.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(options.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(options.Pipeline);
        if (selected.Length > pipelineOptions.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, pipelineOptions.MaxPages, selected.Length);
        }
        var pipeline = new PdfUnderstandingPipeline(options.LayoutOptions, pipelineOptions);
        var numbers = new List<int>();
        var analyses = new List<PdfUnderstandingPageResult>();
        var failures = new Dictionary<int, PdfPageReadOutcome>();
        try {
            for (int index = 0; index < selected.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int pageNumber = selected[index];
                string[] unsupported = document.Pages[pageNumber - 1].GetUnsupportedContentStreamFilters(cancellationToken).ToArray();
                cancellationToken.ThrowIfCancellationRequested();
                if (unsupported.Length > 0) {
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index,
                        PdfPageReadStatus.UnsupportedContentStreamFilter, unsupportedFilters: Array.AsReadOnly(unsupported)));
                    continue;
                }
                try {
                    PdfUnderstandingPageResult analysis = pipeline.RunPages(document, new[] { pageNumber }, cancellationToken)[0];
                    numbers.Add(pageNumber);
                    analyses.Add(analysis);
                } catch (PdfUnsupportedTextMappingException failure) {
                    cancellationToken.ThrowIfCancellationRequested();
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index,
                        PdfPageReadStatus.UnmappableText, mappingFailure: failure));
                }
            }
            int[] successfulNumbers = numbers.ToArray();
            IReadOnlyList<PdfUnderstandingPageResult> enriched = options.Profile == PdfReadProfile.Structured
                ? PdfDocumentSemanticEnricher.Enrich(document, successfulNumbers, analyses,
                    pipelineOptions.MaxRegionsPerPage, pipelineOptions.MaxDocumentWorkUnits, cancellationToken)
                : analyses;
            PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(document, options.LayoutOptions,
                successfulNumbers, enriched, options.Profile, cancellationToken, preserveDocumentWideObjects: true);
            var outcomes = new List<PdfPageReadOutcome>(selected.Length);
            int successfulIndex = 0;
            for (int index = 0; index < selected.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                outcomes.Add(failures.TryGetValue(index, out PdfPageReadOutcome? failure)
                    ? failure : new PdfPageReadOutcome(selected[index], index, PdfPageReadStatus.Read,
                        result.Pages[successfulIndex++]));
            }
            return new PdfInspectionReadResult(sourceInfo, document.RepairReport, result, outcomes.AsReadOnly());
        } finally {
            foreach (PdfUnderstandingPageResult analysis in analyses) analysis.CompleteOperation();
        }
    }
}
