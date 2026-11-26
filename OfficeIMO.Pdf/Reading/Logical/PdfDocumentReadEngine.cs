namespace OfficeIMO.Pdf;

/// <summary>Single owner for parser-backed semantic reconstruction.</summary>
internal static partial class PdfDocumentReadEngine {
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
            layout, pageFormWidgets, analysis, cancellationToken: cancellationToken);
        analysis.CompleteOperation();
        return page;
    }

    internal static PdfDocumentReadResult Read(
        PdfDocument source,
        PdfLoadOptions options,
        CancellationToken cancellationToken) {
        Guard.NotNull(source, nameof(source));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();

        return Read(source.GetReadDocument(GetSourceReadOptions(source, options), cancellationToken), options, cancellationToken);
    }

    /// <summary>
    /// Pages are parsed with the settings the document was opened with. Artifact text changes which content a page
    /// yields, so a read that asks for it uses the source parsed once with artifact text included.
    /// </summary>
    internal static PdfLoadOptions GetSourceReadOptions(PdfDocument source, PdfLoadOptions? options) =>
        options?.IncludeArtifactText == true ? source.ArtifactTextReadOptions : source.ReadOptions;

    internal static PdfDocumentReadResult Read(
        PdfReadDocument document,
        PdfLoadOptions options,
        CancellationToken cancellationToken = default) =>
        Read(document, options, out _, cancellationToken);

    internal static PdfDocumentReadResult Read(
        PdfReadDocument document,
        PdfLoadOptions options,
        out IReadOnlyList<PdfUnderstandingPageResult> pageAnalyses,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        int[] pageNumbers = options.PageSelection?.ToPageNumbers(document.Pages.Count, nameof(options.PageSelection))
            ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        PdfUnderstandingPipelineOptions pipelineOptions = PdfUnderstandingPipelineOptions.Resolve(options.Pipeline);
        var pipeline = new PdfUnderstandingPipeline(options.LayoutOptions, pipelineOptions, options.IncludeImages);
        bool structured = options.Profile == PdfReadProfile.Structured;
        // The structured read takes a page's tables from its tagged structure when that structure passes its checks,
        // so painted borders there would be read only to be discarded. Those pages skip the border pass; any whose
        // tagged tables then fall back, or that also hold a table found in their geometry, are analysed again with it.
        PdfDocumentSemanticEnricher.TaggedTableReads? taggedTables = structured
            ? PdfDocumentSemanticEnricher.PrepareTaggedTables(document, pipelineOptions.MaxDocumentWorkUnits, cancellationToken)
            : null;
        pageAnalyses = pipeline.RunPages(document, pageNumbers, taggedTables?.Pages, cancellationToken);
        IReadOnlyList<PdfUnderstandingPageResult> analyses = pageAnalyses;
        if (structured) {
            if (taggedTables is not null) {
                // Pages whose tagged tables fail, or that none supplies, are known before enrichment: the reads
                // depend only on the pages' text, which the second analysis keeps, so enrichment reuses them.
                bool[] fallBack = PdfDocumentSemanticEnricher.ReadTaggedTables(
                    document, pageNumbers, pageAnalyses, taggedTables, pipelineOptions.MaxDocumentWorkUnits, cancellationToken);
                pageAnalyses = Redo(pageAnalyses, pageIndex => fallBack[pageIndex]) ?? pageAnalyses;
            }
            analyses = Enrich(pageAnalyses);
            if (taggedTables is not null) {
                // A page keeps the skip only when every table on it came from tags; a table found in the geometry
                // beside them may still need its borders.
                IReadOnlyList<PdfUnderstandingPageResult> enriched = analyses;
                IReadOnlyList<PdfUnderstandingPageResult>? redone = Redo(pageAnalyses, pageIndex =>
                    enriched[pageIndex].TableCandidates.Any(static table => table.TaggedSegment is null));
                if (redone is not null) {
                    pageAnalyses = redone;
                    analyses = Enrich(pageAnalyses);
                }
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocumentReadResult result = PdfDocumentReadResult.FromPageNumbers(
            document,
            options.LayoutOptions,
            pageNumbers,
            analyses,
            options.Profile,
            cancellationToken,
            includeImages: options.IncludeImages,
            includeFormWidgets: options.IncludeFormWidgets);
        cancellationToken.ThrowIfCancellationRequested();
        for (int pageIndex = 0; pageIndex < analyses.Count; pageIndex++) {
            analyses[pageIndex].CompleteOperation();
        }
        return result;

        // Analyses again, with the painted-border pass, the pages that skipped it and match; null when none does.
        static IReadOnlyList<PdfUnderstandingPageResult>? Redo(IReadOnlyList<PdfUnderstandingPageResult> pages, Func<int, bool> needed) {
            PdfUnderstandingPageResult[]? redone = null;
            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
                if (pages[pageIndex].RedoWithPaintedBorders is not { } redo || !needed(pageIndex)) continue;
                redone ??= pages.ToArray();
                redone[pageIndex] = redo();
            }
            return redone is null ? null : Array.AsReadOnly(redone);
        }

        IReadOnlyList<PdfUnderstandingPageResult> Enrich(IReadOnlyList<PdfUnderstandingPageResult> pages) =>
            PdfDocumentSemanticEnricher.Enrich(
                document,
                pageNumbers,
                pages,
                options.LayoutOptions,
                pipelineOptions.MaxRegionsPerPage,
                pipelineOptions.MaxDocumentWorkUnits,
                cancellationToken,
                taggedTables);
    }
}
