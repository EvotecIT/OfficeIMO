using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfDocumentReadEngine {
    internal static PdfInspectionReadResult ReadForInspection(PdfReadDocument document, PdfDocumentInfo sourceInfo,
        PdfLoadOptions options, bool recoverDecodableText, CancellationToken cancellationToken) {
        int[] selected = options.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(options.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(options.Pipeline);
        if (selected.Length > pipelineOptions.MaxPages) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, pipelineOptions.MaxPages, selected.Length);
        }
        var pipeline = new PdfUnderstandingPipeline(options.LayoutOptions, pipelineOptions, options.IncludeImages);
        var numbers = new List<int>();
        var analyses = new List<PdfUnderstandingPageResult>();
        var failures = new Dictionary<int, PdfPageReadOutcome>();
        var partialFailures = recoverDecodableText ? new PdfUnsupportedTextMappingException?[selected.Length] : null;
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
                    PdfUnderstandingPageResult analysis = pipeline.RunInspectionPage(document.Pages[pageNumber - 1], pageNumber,
                        recoverDecodableText, cancellationToken, out PdfUnsupportedTextMappingException? mappingFailure);
                    if (mappingFailure is not null && !analysis.DecodedRuns.Any(static run => !string.IsNullOrWhiteSpace(run.Text))) {
                        analysis.CompleteOperation();
                        failures.Add(index, new PdfPageReadOutcome(pageNumber, index, PdfPageReadStatus.UnmappableText,
                            mappingFailure: mappingFailure));
                        continue;
                    }
                    numbers.Add(pageNumber);
                    analyses.Add(analysis);
                    if (mappingFailure is not null) partialFailures![index] = mappingFailure;
                } catch (PdfUnsupportedTextMappingException failure) {
                    cancellationToken.ThrowIfCancellationRequested();
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index,
                        PdfPageReadStatus.UnmappableText, mappingFailure: failure));
                }
            }
            int[] projectedNumbers = numbers.ToArray();
            IReadOnlyList<PdfUnderstandingPageResult> enriched = options.Profile == PdfReadProfile.Structured
                ? PdfDocumentSemanticEnricher.Enrich(document, projectedNumbers, analyses,
                    pipelineOptions.MaxRegionsPerPage, pipelineOptions.MaxDocumentWorkUnits, cancellationToken)
                : analyses;
            PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(document, options.LayoutOptions,
                projectedNumbers, enriched, options.Profile, cancellationToken, preserveDocumentWideObjects: true,
                includeImages: options.IncludeImages, includeFormWidgets: options.IncludeFormWidgets,
                recoverDecodableText: recoverDecodableText);
            var outcomes = new List<PdfPageReadOutcome>(selected.Length);
            int projectedIndex = 0;
            for (int index = 0; index < selected.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (failures.TryGetValue(index, out PdfPageReadOutcome? failure)) {
                    outcomes.Add(failure);
                    continue;
                }
                PdfUnsupportedTextMappingException? partialFailure = partialFailures?[index];
                outcomes.Add(new PdfPageReadOutcome(selected[index], index,
                    partialFailure is null ? PdfPageReadStatus.Read : PdfPageReadStatus.UnmappableText,
                    result.Pages[projectedIndex++], mappingFailure: partialFailure));
            }
            return new PdfInspectionReadResult(sourceInfo, document.RepairReport, result, outcomes.AsReadOnly());
        } finally {
            foreach (PdfUnderstandingPageResult analysis in analyses) analysis.CompleteOperation();
        }
    }
}
