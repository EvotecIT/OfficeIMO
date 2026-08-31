using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Configures independently replaceable PDF understanding stages.</summary>
public sealed class PdfUnderstandingPipelineOptions {
    /// <summary>Creates the built-in structured semantic profile.</summary>
    public static PdfUnderstandingPipelineOptions Structured() => new PdfUnderstandingPipelineOptions {
        GlyphDecoding = PdfAdvancedUnderstandingStages.GlyphDecoding,
        WordGrouping = PdfAdvancedUnderstandingStages.WordGrouping,
        LineGrouping = PdfAdvancedUnderstandingStages.LineGrouping,
        PageSegmentation = PdfAdvancedUnderstandingStages.PageSegmentation,
        ReadingOrder = PdfAdvancedUnderstandingStages.ReadingOrder,
        SemanticClassification = PdfAdvancedUnderstandingStages.SemanticClassification
    };

    /// <summary>Glyph/text decoding stage.</summary>
    public IPdfGlyphDecodingStage? GlyphDecoding { get; set; }
    /// <summary>Word grouping stage.</summary>
    public IPdfWordGroupingStage? WordGrouping { get; set; }
    /// <summary>Line grouping stage.</summary>
    public IPdfLineGroupingStage? LineGrouping { get; set; }
    /// <summary>Page segmentation stage.</summary>
    public IPdfPageSegmentationStage? PageSegmentation { get; set; }
    /// <summary>Reading-order stage.</summary>
    public IPdfReadingOrderStage? ReadingOrder { get; set; }
    /// <summary>Semantic classification stage.</summary>
    public IPdfSemanticClassificationStage? SemanticClassification { get; set; }
    /// <summary>Maximum selected pages processed by one run.</summary>
    public int MaxPages { get; set; } = 1000;
    /// <summary>Maximum decoded text runs retained for one page.</summary>
    public int MaxRunsPerPage { get; set; } = 100_000;
    /// <summary>Maximum decoded text characters retained for one page.</summary>
    public int MaxTextCharactersPerPage { get; set; } = 4 * 1024 * 1024;
    /// <summary>Maximum grouped words retained for one page.</summary>
    public int MaxWordsPerPage { get; set; } = 100_000;
    /// <summary>Maximum grouped lines retained for one page.</summary>
    public int MaxLinesPerPage { get; set; } = 50_000;
    /// <summary>Maximum regions and semantic elements retained for one page.</summary>
    public int MaxRegionsPerPage { get; set; } = 10_000;
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
            PageSegmentation = source.PageSegmentation ?? PdfAdvancedUnderstandingStages.PageSegmentation,
            ReadingOrder = source.ReadingOrder ?? PdfAdvancedUnderstandingStages.ReadingOrder,
            SemanticClassification = source.SemanticClassification ?? PdfAdvancedUnderstandingStages.SemanticClassification,
            MaxPages = source.MaxPages,
            MaxRunsPerPage = source.MaxRunsPerPage,
            MaxTextCharactersPerPage = source.MaxTextCharactersPerPage,
            MaxWordsPerPage = source.MaxWordsPerPage,
            MaxLinesPerPage = source.MaxLinesPerPage,
            MaxRegionsPerPage = source.MaxRegionsPerPage,
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
    private readonly IPdfPageSegmentationStage _pageSegmentation;
    private readonly IPdfReadingOrderStage _readingOrder;
    private readonly IPdfSemanticClassificationStage _semanticClassification;
    private readonly PdfTextLayoutOptions _layout;
    private readonly int _maxPages;
    private readonly PdfUnderstandingPipelineOptions _limits;

    /// <summary>Creates a pipeline by overlaying caller stages on the canonical structured stage set.</summary>
    internal PdfUnderstandingPipeline(PdfTextLayoutOptions layout, PdfUnderstandingPipelineOptions? options = null) {
        PdfUnderstandingPipelineOptions effective = PdfUnderstandingPipelineOptions.Resolve(options);
        _glyphDecoding = effective.GlyphDecoding!;
        _wordGrouping = effective.WordGrouping!;
        _lineGrouping = effective.LineGrouping!;
        _pageSegmentation = effective.PageSegmentation!;
        _readingOrder = effective.ReadingOrder!;
        _semanticClassification = effective.SemanticClassification!;
        _layout = layout ?? throw new ArgumentNullException(nameof(layout));
        _maxPages = effective.MaxPages;
        _limits = effective;
        if (_maxPages <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxPages, "Maximum pages must be positive.");
        ValidateLimit(effective.MaxRunsPerPage, nameof(effective.MaxRunsPerPage));
        ValidateLimit(effective.MaxTextCharactersPerPage, nameof(effective.MaxTextCharactersPerPage));
        ValidateLimit(effective.MaxWordsPerPage, nameof(effective.MaxWordsPerPage));
        ValidateLimit(effective.MaxLinesPerPage, nameof(effective.MaxLinesPerPage));
        ValidateLimit(effective.MaxRegionsPerPage, nameof(effective.MaxRegionsPerPage));
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

    private PdfUnderstandingPageResult RunPage(PdfReadPage page, int pageNumber, CancellationToken cancellationToken) {
        var context = new PdfUnderstandingPageContext(
            page,
            pageNumber,
            _layout,
            _limits.MaxTextCharactersPerPage,
            _limits.MaxWordsPerPage,
            _limits.MaxWorkUnitsPerPage,
            cancellationToken);
        var trace = new List<PdfUnderstandingStageTrace>(6);
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
        return new PdfUnderstandingPageResult(pageNumber, runs, words, lines, regions, ordered, readingOrderEvidence, elements, trace.AsReadOnly());
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfReadingOrderEvidence> BuildReadingOrderEvidence(IReadOnlyList<PdfUnderstandingRegion> ordered, Type providerType) {
        var result = new PdfReadingOrderEvidence[ordered.Count];
        for (int i = 0; i < ordered.Count; i++) {
            bool geometryConsistent = i == 0 || ordered[i - 1].YTop >= ordered[i].YTop || ordered[i - 1].XStart <= ordered[i].XStart;
            double confidence = PdfInference.Clamp((ordered[i].Confidence * 0.8D) + (geometryConsistent ? 0.2D : 0D));
            var evidence = new[] {
                new PdfInferenceEvidence("reading-order.provider", "Reading order was produced by " + providerType.FullName + ".", 0.5D),
                new PdfInferenceEvidence(geometryConsistent ? "reading-order.geometry-consistent" : "reading-order.geometry-conflict", geometryConsistent ? "The position is consistent with top-to-bottom, left-to-right geometry." : "The position conflicts with simple top-to-bottom, left-to-right geometry.", geometryConsistent ? 0.5D : -0.5D)
            };
            result[i] = new PdfReadingOrderEvidence(i, ordered[i], confidence, evidence);
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
}
