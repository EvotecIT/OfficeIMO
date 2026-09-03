using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Configures independently replaceable PDF understanding stages.</summary>
public sealed class PdfUnderstandingPipelineOptions {
    /// <summary>Default maximum selected pages processed by one semantic read.</summary>
    public const int DefaultMaxPages = 1000;

    /// <summary>Creates the built-in structured semantic profile.</summary>
    public static PdfUnderstandingPipelineOptions Structured() => new PdfUnderstandingPipelineOptions {
        GlyphDecoding = PdfAdvancedUnderstandingStages.GlyphDecoding,
        WordGrouping = PdfAdvancedUnderstandingStages.WordGrouping,
        LineGrouping = PdfAdvancedUnderstandingStages.LineGrouping,
        TableDetection = PdfAdvancedUnderstandingStages.TableDetection,
        PageSegmentation = PdfAdvancedUnderstandingStages.PageSegmentation,
        ReadingOrder = PdfAdvancedUnderstandingStages.ReadingOrder,
        SemanticClassification = PdfAdvancedUnderstandingStages.SemanticClassification,
        ImageRegionDetection = PdfAdvancedUnderstandingStages.ImageRegionDetection
    };

    /// <summary>Glyph/text decoding stage.</summary>
    public IPdfGlyphDecodingStage? GlyphDecoding { get; set; }
    /// <summary>Word grouping stage.</summary>
    public IPdfWordGroupingStage? WordGrouping { get; set; }
    /// <summary>Line grouping stage.</summary>
    public IPdfLineGroupingStage? LineGrouping { get; set; }
    /// <summary>Table detection stage run before general page segmentation.</summary>
    public IPdfTableDetectionStage? TableDetection { get; set; }
    /// <summary>Page segmentation stage.</summary>
    public IPdfPageSegmentationStage? PageSegmentation { get; set; }
    /// <summary>Reading-order stage.</summary>
    public IPdfReadingOrderStage? ReadingOrder { get; set; }
    /// <summary>Semantic classification stage.</summary>
    public IPdfSemanticClassificationStage? SemanticClassification { get; set; }
    /// <summary>Image-region and figure-caption association stage.</summary>
    public IPdfImageRegionDetectionStage? ImageRegionDetection { get; set; }
    /// <summary>Maximum selected pages processed by one run.</summary>
    public int MaxPages { get; set; } = DefaultMaxPages;
    /// <summary>Maximum decoded text runs retained for one page.</summary>
    public int MaxRunsPerPage { get; set; } = 100_000;
    /// <summary>Maximum decoded text characters retained for one page.</summary>
    public int MaxTextCharactersPerPage { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum grouped words retained for one page.</summary>
    public int MaxWordsPerPage { get; set; } = 100_000;
    /// <summary>Maximum grouped lines retained for one page.</summary>
    public int MaxLinesPerPage { get; set; } = 50_000;
    /// <summary>Maximum table candidates retained for one page.</summary>
    public int MaxTableCandidatesPerPage { get; set; } = 1_024;
    /// <summary>Maximum regions and semantic elements retained for one page.</summary>
    public int MaxRegionsPerPage { get; set; } = 10_000;
    /// <summary>Maximum positioned image regions retained for one page.</summary>
    public int MaxImageRegionsPerPage { get; set; } = 10_000;
    /// <summary>Maximum viable image-caption association edges retained before deterministic matching.</summary>
    public int MaxImageCaptionCandidatesPerPage { get; set; } = 100_000;
    /// <summary>Maximum comparison and traversal work performed by built-in stages for one page.</summary>
    public long MaxWorkUnitsPerPage { get; set; } = 10_000_000;
    /// <summary>Maximum comparison and traversal work performed by document-wide semantic enrichment.</summary>
    public long MaxDocumentWorkUnits { get; set; } = 10_000_000;

    internal static PdfUnderstandingPipelineOptions Resolve(PdfUnderstandingPipelineOptions? overrides) {
        PdfUnderstandingPipelineOptions source = overrides ?? new PdfUnderstandingPipelineOptions();
        return new PdfUnderstandingPipelineOptions {
            GlyphDecoding = source.GlyphDecoding ?? PdfAdvancedUnderstandingStages.GlyphDecoding,
            WordGrouping = source.WordGrouping ?? PdfAdvancedUnderstandingStages.WordGrouping,
            LineGrouping = source.LineGrouping ?? PdfAdvancedUnderstandingStages.LineGrouping,
            TableDetection = source.TableDetection ?? PdfAdvancedUnderstandingStages.TableDetection,
            PageSegmentation = source.PageSegmentation ?? PdfAdvancedUnderstandingStages.PageSegmentation,
            ReadingOrder = source.ReadingOrder ?? PdfAdvancedUnderstandingStages.ReadingOrder,
            SemanticClassification = source.SemanticClassification ?? PdfAdvancedUnderstandingStages.SemanticClassification,
            ImageRegionDetection = source.ImageRegionDetection ?? PdfAdvancedUnderstandingStages.ImageRegionDetection,
            MaxPages = source.MaxPages,
            MaxRunsPerPage = source.MaxRunsPerPage,
            MaxTextCharactersPerPage = source.MaxTextCharactersPerPage,
            MaxWordsPerPage = source.MaxWordsPerPage,
            MaxLinesPerPage = source.MaxLinesPerPage,
            MaxTableCandidatesPerPage = source.MaxTableCandidatesPerPage,
            MaxRegionsPerPage = source.MaxRegionsPerPage,
            MaxImageRegionsPerPage = source.MaxImageRegionsPerPage,
            MaxImageCaptionCandidatesPerPage = source.MaxImageCaptionCandidatesPerPage,
            MaxWorkUnitsPerPage = source.MaxWorkUnitsPerPage,
            MaxDocumentWorkUnits = source.MaxDocumentWorkUnits
        };
    }
}

/// <summary>Runs a bounded, typed, pluggable PDF text-understanding pipeline.</summary>
internal sealed class PdfUnderstandingPipeline {
    private readonly IPdfGlyphDecodingStage _glyphDecoding;
    private readonly IPdfWordGroupingStage _wordGrouping;
    private readonly IPdfLineGroupingStage _lineGrouping;
    private readonly IPdfTableDetectionStage _tableDetection;
    private readonly IPdfPageSegmentationStage _pageSegmentation;
    private readonly IPdfReadingOrderStage _readingOrder;
    private readonly IPdfSemanticClassificationStage _semanticClassification;
    private readonly IPdfImageRegionDetectionStage _imageRegionDetection;
    private readonly PdfTextLayoutOptions _layout;
    private readonly int _maxPages;
    private readonly PdfUnderstandingPipelineOptions _limits;
    private readonly bool _restrictLogicalProjectionToReadingOrder;

    /// <summary>Creates a pipeline by overlaying caller stages on the canonical structured stage set.</summary>
    internal PdfUnderstandingPipeline(PdfTextLayoutOptions layout, PdfUnderstandingPipelineOptions? options = null) {
        PdfUnderstandingPipelineOptions effective = PdfUnderstandingPipelineOptions.Resolve(options);
        _glyphDecoding = effective.GlyphDecoding!;
        _wordGrouping = effective.WordGrouping!;
        _lineGrouping = effective.LineGrouping!;
        _tableDetection = effective.TableDetection!;
        _pageSegmentation = effective.PageSegmentation!;
        _readingOrder = effective.ReadingOrder!;
        _semanticClassification = effective.SemanticClassification!;
        _imageRegionDetection = effective.ImageRegionDetection!;
        _restrictLogicalProjectionToReadingOrder =
            !ReferenceEquals(_wordGrouping, PdfAdvancedUnderstandingStages.WordGrouping) ||
            !ReferenceEquals(_lineGrouping, PdfAdvancedUnderstandingStages.LineGrouping) ||
            !ReferenceEquals(_tableDetection, PdfAdvancedUnderstandingStages.TableDetection) ||
            !ReferenceEquals(_pageSegmentation, PdfAdvancedUnderstandingStages.PageSegmentation) ||
            !ReferenceEquals(_readingOrder, PdfAdvancedUnderstandingStages.ReadingOrder);
        _layout = layout ?? throw new ArgumentNullException(nameof(layout));
        _maxPages = effective.MaxPages;
        _limits = effective;
        if (_maxPages <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxPages, "Maximum pages must be positive.");
        ValidateLimit(effective.MaxRunsPerPage, nameof(effective.MaxRunsPerPage));
        ValidateLimit(effective.MaxTextCharactersPerPage, nameof(effective.MaxTextCharactersPerPage));
        ValidateLimit(effective.MaxWordsPerPage, nameof(effective.MaxWordsPerPage));
        ValidateLimit(effective.MaxLinesPerPage, nameof(effective.MaxLinesPerPage));
        ValidateLimit(effective.MaxTableCandidatesPerPage, nameof(effective.MaxTableCandidatesPerPage));
        ValidateLimit(effective.MaxRegionsPerPage, nameof(effective.MaxRegionsPerPage));
        ValidateLimit(effective.MaxImageRegionsPerPage, nameof(effective.MaxImageRegionsPerPage));
        ValidateLimit(effective.MaxImageCaptionCandidatesPerPage, nameof(effective.MaxImageCaptionCandidatesPerPage));
        ValidateLimit(effective.MaxWorkUnitsPerPage, nameof(effective.MaxWorkUnitsPerPage));
        ValidateLimit(effective.MaxDocumentWorkUnits, nameof(effective.MaxDocumentWorkUnits));
    }

    internal IReadOnlyList<PdfUnderstandingPageResult> RunPages(
        PdfReadDocument document,
        int[] pageNumbers,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        if (pageNumbers.Length > _maxPages) throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, _maxPages, pageNumbers.Length);
        var pages = new List<PdfUnderstandingPageResult>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[i];
            pages.Add(RunPage(document.Pages[pageNumber - 1], pageNumber, cancellationToken));
        }
        return pages.AsReadOnly();
    }

    internal PdfUnderstandingPageResult RunPositionedPage(
        PdfReadPage page,
        int pageNumber,
        IReadOnlyList<PdfTextSpan> runs,
        IReadOnlyList<PdfUnderstandingWord> words,
        IReadOnlyList<PdfUnderstandingLine> lines,
        Type sourceProviderType,
        CancellationToken cancellationToken) {
        Guard.NotNull(page, nameof(page));
        Guard.NotNull(runs, nameof(runs));
        Guard.NotNull(words, nameof(words));
        Guard.NotNull(lines, nameof(lines));
        Guard.NotNull(sourceProviderType, nameof(sourceProviderType));
#pragma warning disable CA1512 // ThrowIfNegativeOrZero is unavailable on every target framework.
        if (pageNumber <= 0) throw new ArgumentOutOfRangeException(nameof(pageNumber));
#pragma warning restore CA1512
        var context = CreateContext(page, pageNumber, cancellationToken);
        EnsureCount(runs.Count, _limits.MaxRunsPerPage);
        EnsureTextCharacters(runs.Select(static run => run.Text), _limits.MaxTextCharactersPerPage);
        EnsureCount(words.Count, _limits.MaxWordsPerPage);
        EnsureTextCharacters(words.Select(static word => word.Text), _limits.MaxTextCharactersPerPage);
        EnsureCount(lines.Count, _limits.MaxLinesPerPage);
        context.DecodedRuns = runs;
        var trace = new List<PdfUnderstandingStageTrace>(8) {
            new PdfUnderstandingStageTrace("positioned-word-input", sourceProviderType, runs.Count, words.Count),
            new PdfUnderstandingStageTrace("positioned-line-input", sourceProviderType, words.Count, lines.Count)
        };
        IReadOnlyList<PdfUnderstandingTableCandidate> tableCandidates = NotNull(
            _tableDetection.DetectTables(context, lines),
            nameof(IPdfTableDetectionStage));
        return RunFromLines(context, runs, words, lines, tableCandidates, _tableDetection.GetType(), trace, cancellationToken);
    }

    private PdfUnderstandingPageResult RunPage(PdfReadPage page, int pageNumber, CancellationToken cancellationToken) {
        PdfUnderstandingPageContext context = CreateContext(page, pageNumber, cancellationToken);
        var trace = new List<PdfUnderstandingStageTrace>(8);
        IReadOnlyList<PdfTextSpan> runs = NotNull(_glyphDecoding.Decode(context), nameof(IPdfGlyphDecodingStage));
        cancellationToken.ThrowIfCancellationRequested();
        EnsureCount(runs.Count, _limits.MaxRunsPerPage);
        EnsureTextCharacters(runs.Select(static run => run.Text), _limits.MaxTextCharactersPerPage);
        runs = TextLayoutEngine.FilterIgnoredPageBands(
            runs,
            context.Page,
            _layout,
            context.ConsumeWork,
            context.ThrowIfCancellationRequested);
        EnsureCount(runs.Count, _limits.MaxRunsPerPage);
        EnsureTextCharacters(runs.Select(static run => run.Text), _limits.MaxTextCharactersPerPage);
        context.DecodedRuns = runs;
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("glyph-decoding", _glyphDecoding.GetType(), 0, runs.Count));
        IReadOnlyList<PdfUnderstandingWord> words = NotNull(_wordGrouping.GroupWords(context, runs), nameof(IPdfWordGroupingStage));
        EnsureCount(words.Count, _limits.MaxWordsPerPage);
        EnsureTextCharacters(words.Select(static word => word.Text), _limits.MaxTextCharactersPerPage);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("word-grouping", _wordGrouping.GetType(), runs.Count, words.Count));
        IReadOnlyList<PdfUnderstandingLine> lines = NotNull(_lineGrouping.GroupLines(context, words), nameof(IPdfLineGroupingStage));
        EnsureCount(lines.Count, _limits.MaxLinesPerPage);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("line-grouping", _lineGrouping.GetType(), words.Count, lines.Count));
        IReadOnlyList<PdfUnderstandingTableCandidate> tableCandidates = NotNull(
            _tableDetection.DetectTables(context, lines),
            nameof(IPdfTableDetectionStage));
        return RunFromLines(context, runs, words, lines, tableCandidates, _tableDetection.GetType(), trace, cancellationToken);
    }

    private PdfUnderstandingPageContext CreateContext(
        PdfReadPage page,
        int pageNumber,
        CancellationToken cancellationToken) {
        var context = new PdfUnderstandingPageContext(
            page,
            pageNumber,
            _layout,
            _limits.MaxTextCharactersPerPage,
            _limits.MaxWordsPerPage,
            _limits.MaxWorkUnitsPerPage,
        cancellationToken) {
            MaxTableCandidatesPerPage = _limits.MaxTableCandidatesPerPage,
            MaxImageCaptionCandidatesPerPage = _limits.MaxImageCaptionCandidatesPerPage
        };
        context.ImagePlacements = page.GetImagePlacements(pageNumber);
        EnsureCount(context.ImagePlacements.Count, _limits.MaxImageRegionsPerPage);
        return context;
    }

    private PdfUnderstandingPageResult RunFromLines(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfTextSpan> runs,
        IReadOnlyList<PdfUnderstandingWord> words,
        IReadOnlyList<PdfUnderstandingLine> lines,
        IReadOnlyList<PdfUnderstandingTableCandidate> tableCandidates,
        Type tableProviderType,
        List<PdfUnderstandingStageTrace> trace,
        CancellationToken cancellationToken) {
        EnsureCount(tableCandidates.Count, _limits.MaxTableCandidatesPerPage);
        EnsureTableCandidateArtifacts(
            context,
            tableCandidates,
            _limits.MaxLinesPerPage,
            _limits.MaxWordsPerPage,
            _limits.MaxTextCharactersPerPage);
        context.TableCandidates = tableCandidates;
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("table-detection", tableProviderType, lines.Count, tableCandidates.Count));
        IReadOnlyList<PdfUnderstandingRegion> regions = NotNull(_pageSegmentation.Segment(context, lines), nameof(IPdfPageSegmentationStage));
        EnsureCount(regions.Count, _limits.MaxRegionsPerPage);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("page-segmentation", _pageSegmentation.GetType(), lines.Count, regions.Count));
        IReadOnlyList<PdfUnderstandingRegion> ordered = NotNull(_readingOrder.Order(context, regions), nameof(IPdfReadingOrderStage));
        EnsureCount(ordered.Count, _limits.MaxRegionsPerPage);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("reading-order", _readingOrder.GetType(), regions.Count, ordered.Count));
        IReadOnlyList<PdfReadingOrderEvidence> readingOrderEvidence = BuildReadingOrderEvidence(ordered, _readingOrder.GetType());
        IReadOnlyList<PdfUnderstandingSemanticElement> elements = NotNull(_semanticClassification.Classify(context, ordered), nameof(IPdfSemanticClassificationStage));
        EnsureCount(elements.Count, _limits.MaxRegionsPerPage);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace("semantic-classification", _semanticClassification.GetType(), ordered.Count, elements.Count));
        IReadOnlyList<PdfUnderstandingImageRegion> imageRegions = NotNull(
            _imageRegionDetection.Detect(context, elements),
            nameof(IPdfImageRegionDetectionStage));
        EnsureCount(imageRegions.Count, _limits.MaxImageRegionsPerPage);
        EnsureImageRegionArtifacts(context, imageRegions, elements);
        cancellationToken.ThrowIfCancellationRequested();
        trace.Add(new PdfUnderstandingStageTrace(
            "image-region-detection",
            _imageRegionDetection.GetType(),
            context.ImagePlacements.Count,
            imageRegions.Count));
        return new PdfUnderstandingPageResult(
            context.PageNumber,
            runs,
            words,
            lines,
            regions,
            ordered,
            readingOrderEvidence,
            elements,
            trace.AsReadOnly(),
            context.ConsumeWork,
            context.ThrowIfCancellationRequested,
            context.CompleteOperation,
            logicalProjectionLines: null,
            restrictLogicalProjectionToReadingOrder: _restrictLogicalProjectionToReadingOrder,
            tableCandidates: tableCandidates,
            imagePlacements: context.ImagePlacements,
            imageRegions: imageRegions);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfReadingOrderEvidence> BuildReadingOrderEvidence(IReadOnlyList<PdfUnderstandingRegion> ordered, Type providerType) {
        var result = new PdfReadingOrderEvidence[ordered.Count];
        for (int i = 0; i < ordered.Count; i++) {
            var evidence = new[] {
                new PdfInferenceEvidence(
                    "reading-order.provider",
                    "Reading order was produced by " + providerType.FullName + ".",
                    0.5D)
            };
            // Direction and bidi metadata are not yet part of positioned-text contracts. Do not
            // manufacture confidence by comparing a provider result with an LTR-only geometry rule.
            result[i] = new PdfReadingOrderEvidence(i, ordered[i], ordered[i].Confidence, evidence);
        }
        return Array.AsReadOnly(result);
    }

    private static IReadOnlyList<T> NotNull<T>(IReadOnlyList<T>? value, string stage) => value ?? throw new InvalidOperationException(stage + " returned null.");

    private static void ValidateLimit(int value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
    }

    private static void ValidateLimit(long value, string name) {
        if (value <= 0) throw new ArgumentOutOfRangeException(name);
    }

    private static void EnsureCount(int actual, int maximum) {
        if (actual > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, actual);
    }

    private static void EnsureTextCharacters(IEnumerable<string?> values, int maximum) {
        long total = 0;
        foreach (string? value in values) {
            total = checked(total + (value?.Length ?? 0));
            if (total > maximum) throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, total);
        }
    }

    private static void EnsureImageRegionArtifacts(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingImageRegion> imageRegions,
        IReadOnlyList<PdfUnderstandingSemanticElement> semanticElements) {
        if (imageRegions.Count != context.ImagePlacements.Count) {
            throw new InvalidOperationException(
                nameof(IPdfImageRegionDetectionStage) + " must return exactly one region per visible image placement.");
        }
        var expectedPlacements = new HashSet<PdfImagePlacement>(context.ImagePlacements);
        var observedPlacements = new HashSet<PdfImagePlacement>();
        var expectedSemanticElements = new HashSet<PdfUnderstandingSemanticElement>(semanticElements);
        for (int index = 0; index < imageRegions.Count; index++) {
            context.ConsumeWork();
            PdfUnderstandingImageRegion region = imageRegions[index]
                ?? throw new InvalidOperationException(nameof(IPdfImageRegionDetectionStage) + " returned a null image region.");
            if (!expectedPlacements.Contains(region.Placement) || !observedPlacements.Add(region.Placement)) {
                throw new InvalidOperationException(
                    nameof(IPdfImageRegionDetectionStage) + " must represent each input image placement exactly once.");
            }
            if (region.Caption is not null && !expectedSemanticElements.Contains(region.Caption)) {
                throw new InvalidOperationException(
                    nameof(IPdfImageRegionDetectionStage) + " returned a caption outside the semantic-stage result.");
            }
        }
    }

    private static void EnsureTableCandidateArtifacts(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingTableCandidate> candidates,
        int maximumLines,
        int maximumCells,
        int maximumCharacters) {
        long rows = 0;
        long cells = 0;
        long sourceLines = 0;
        long characters = 0;
        for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
            context.ConsumeWork();
            PdfUnderstandingTableCandidate candidate = candidates[candidateIndex]
                ?? throw new InvalidOperationException(nameof(IPdfTableDetectionStage) + " returned a null table candidate.");
            AddCharacters(candidate.DetectionKind);
            EnsureCount(candidate.Columns.Count, maximumCells);
            context.ConsumeWork(candidate.Columns.Count);
            rows = AddAndEnsure(rows, candidate.Rows.Count, maximumLines);
            sourceLines = AddAndEnsure(sourceLines, candidate.SourceLines.Count, maximumLines);
            context.ConsumeWork(candidate.SourceLines.Count + 1L);
            for (int rowIndex = 0; rowIndex < candidate.Rows.Count; rowIndex++) {
                context.ConsumeWork();
                IReadOnlyList<string> row = candidate.Rows[rowIndex];
                cells = AddAndEnsure(cells, row.Count, maximumCells);
                context.ConsumeWork(row.Count + 1L);
                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) AddCharacters(row[cellIndex]);
            }
            for (int evidenceIndex = 0; evidenceIndex < candidate.Evidence.Count; evidenceIndex++) {
                context.ConsumeWork();
                AddCharacters(candidate.Evidence[evidenceIndex].Code);
                AddCharacters(candidate.Evidence[evidenceIndex].Message);
            }
        }

        void AddCharacters(string? value) {
            characters = AddAndEnsure(characters, value?.Length ?? 0, maximumCharacters);
        }
    }

    private static long AddAndEnsure(long current, long additional, long maximum) {
        long total;
        try {
            total = checked(current + additional);
        } catch (OverflowException) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, long.MaxValue);
        }
        if (total > maximum) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, maximum, total);
        }
        return total;
    }
}
