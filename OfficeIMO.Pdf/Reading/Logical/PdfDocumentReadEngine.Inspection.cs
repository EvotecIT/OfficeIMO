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
        var partialFailures = recoverDecodableText
            ? new (PdfPageReadStatus Status, PdfUnsupportedTextMappingException? MappingFailure)[selected.Length] : null;
        PdfReadPage[]? inspectionPages = null;
        try {
            for (int index = 0; index < selected.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int pageNumber = selected[index];
                PdfReadPage sourcePage = document.Pages[pageNumber - 1];
                string[] unsupported = sourcePage.InspectContentStreams(recoverDecodableText, cancellationToken,
                    out PdfReadPage? inspectionPage).ToArray();
                cancellationToken.ThrowIfCancellationRequested();
                if (unsupported.Length > 0) {
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index,
                        PdfPageReadStatus.UnsupportedContentStreamFilter, unsupportedFilters: Array.AsReadOnly(unsupported)));
                    continue;
                }
                if (inspectionPage is null) {
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index, PdfPageReadStatus.DamagedContentStream));
                    continue;
                }
                bool damaged = !ReferenceEquals(sourcePage, inspectionPage);
                try {
                    PdfUnderstandingPageResult analysis = pipeline.RunInspectionPage(inspectionPage, pageNumber,
                        recoverDecodableText, cancellationToken, out PdfUnsupportedTextMappingException? mappingFailure);
                    PdfPageReadStatus status = damaged ? PdfPageReadStatus.DamagedContentStream
                        : mappingFailure is not null ? PdfPageReadStatus.UnmappableText : PdfPageReadStatus.Read;
                    if (status != PdfPageReadStatus.Read && !analysis.DecodedRuns.Any(static run => !string.IsNullOrWhiteSpace(run.Text))) {
                        analysis.CompleteOperation();
                        failures.Add(index, new PdfPageReadOutcome(pageNumber, index, status, mappingFailure: mappingFailure));
                        continue;
                    }
                    numbers.Add(pageNumber);
                    analyses.Add(analysis);
                    if (status != PdfPageReadStatus.Read) partialFailures![index] = (status, mappingFailure);
                    if (damaged) {
                        inspectionPages ??= document.Pages.ToArray();
                        inspectionPages[pageNumber - 1] = inspectionPage;
                    }
                } catch (PdfUnsupportedTextMappingException failure) {
                    cancellationToken.ThrowIfCancellationRequested();
                    failures.Add(index, new PdfPageReadOutcome(pageNumber, index,
                        damaged ? PdfPageReadStatus.DamagedContentStream : PdfPageReadStatus.UnmappableText,
                        mappingFailure: failure));
                }
            }
            int[] projectedNumbers = numbers.ToArray();
            PdfReadDocument inspectionDocument = inspectionPages is null ? document : document.WithInspectionPages(inspectionPages);
            IReadOnlyList<PdfUnderstandingPageResult> enriched = options.Profile == PdfReadProfile.Structured
                ? PdfDocumentSemanticEnricher.Enrich(inspectionDocument, projectedNumbers, analyses, options.LayoutOptions,
                    pipelineOptions.MaxRegionsPerPage, pipelineOptions.MaxDocumentWorkUnits, cancellationToken)
                : analyses;
            PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(inspectionDocument, options.LayoutOptions,
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
                var outcome = partialFailures?[index] ?? default;
                outcomes.Add(new PdfPageReadOutcome(selected[index], index, outcome.Status,
                    result.Pages[projectedIndex++], mappingFailure: outcome.MappingFailure));
            }
            return new PdfInspectionReadResult(sourceInfo, document.RepairReport, result, outcomes.AsReadOnly());
        } finally {
            foreach (PdfUnderstandingPageResult analysis in analyses) analysis.CompleteOperation();
        }
    }
}
