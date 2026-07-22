using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Configures independently replaceable PDF understanding stages.</summary>
public sealed class PdfUnderstandingPipelineOptions {
    /// <summary>Creates the built-in advanced layout profile while preserving caller-provided layout options.</summary>
    public static PdfUnderstandingPipelineOptions Advanced(PdfTextLayoutOptions? layout = null) => new PdfUnderstandingPipelineOptions {
        WordGrouping = PdfAdvancedUnderstandingStages.WordGrouping,
        LineGrouping = PdfAdvancedUnderstandingStages.LineGrouping,
        PageSegmentation = PdfAdvancedUnderstandingStages.PageSegmentation,
        ReadingOrder = PdfAdvancedUnderstandingStages.ReadingOrder,
        SemanticClassification = PdfAdvancedUnderstandingStages.SemanticClassification,
        Layout = layout ?? new PdfTextLayoutOptions()
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
    /// <summary>Options used by the default fast layout stages.</summary>
    public PdfTextLayoutOptions Layout { get; set; } = new PdfTextLayoutOptions();
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
}

/// <summary>Runs a bounded, typed, pluggable PDF text-understanding pipeline.</summary>
public sealed class PdfUnderstandingPipeline {
    private readonly IPdfGlyphDecodingStage _glyphDecoding;
    private readonly IPdfWordGroupingStage _wordGrouping;
    private readonly IPdfLineGroupingStage _lineGrouping;
    private readonly IPdfPageSegmentationStage _pageSegmentation;
    private readonly IPdfReadingOrderStage _readingOrder;
    private readonly IPdfSemanticClassificationStage _semanticClassification;
    private readonly PdfTextLayoutOptions _layout;
    private readonly int _maxPages;
    private readonly PdfUnderstandingPipelineOptions _limits;

    /// <summary>Creates a pipeline, using the lightweight built-in stage for each unspecified slot.</summary>
    public PdfUnderstandingPipeline(PdfUnderstandingPipelineOptions? options = null) {
        PdfUnderstandingPipelineOptions effective = options ?? new PdfUnderstandingPipelineOptions();
        _glyphDecoding = effective.GlyphDecoding ?? PdfFastUnderstandingStages.GlyphDecoding;
        _wordGrouping = effective.WordGrouping ?? PdfFastUnderstandingStages.WordGrouping;
        _lineGrouping = effective.LineGrouping ?? PdfFastUnderstandingStages.LineGrouping;
        _pageSegmentation = effective.PageSegmentation ?? PdfFastUnderstandingStages.PageSegmentation;
        _readingOrder = effective.ReadingOrder ?? PdfFastUnderstandingStages.ReadingOrder;
        _semanticClassification = effective.SemanticClassification ?? PdfFastUnderstandingStages.SemanticClassification;
        _layout = effective.Layout ?? throw new ArgumentNullException(nameof(options), "Layout options cannot be null.");
        _maxPages = effective.MaxPages;
        _limits = effective;
        if (_maxPages <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxPages, "Maximum pages must be positive.");
        ValidateLimit(effective.MaxRunsPerPage, nameof(effective.MaxRunsPerPage));
        ValidateLimit(effective.MaxTextCharactersPerPage, nameof(effective.MaxTextCharactersPerPage));
        ValidateLimit(effective.MaxWordsPerPage, nameof(effective.MaxWordsPerPage));
        ValidateLimit(effective.MaxLinesPerPage, nameof(effective.MaxLinesPerPage));
        ValidateLimit(effective.MaxRegionsPerPage, nameof(effective.MaxRegionsPerPage));
    }

    /// <summary>Runs all stages for all pages or a caller-ordered page selection.</summary>
    public PdfUnderstandingResult Run(PdfReadDocument document, PdfPageSelection? selection = null, CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        int[] pageNumbers = selection?.ToPageNumbers(document.Pages.Count, nameof(selection)) ?? Enumerable.Range(1, document.Pages.Count).ToArray();
        if (pageNumbers.Length > _maxPages) throw PdfReadLimitException.Create(PdfReadLimitKind.Pages, _maxPages, pageNumbers.Length);
        var pages = new List<PdfUnderstandingPageResult>(pageNumbers.Length);
        for (int i = 0; i < pageNumbers.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            int pageNumber = pageNumbers[i];
            pages.Add(RunPage(document.Pages[pageNumber - 1], pageNumber, cancellationToken));
        }
        return new PdfUnderstandingResult(pages.AsReadOnly());
    }

    /// <summary>Runs all stages for pages resolved by a document-relative selector.</summary>
    public PdfUnderstandingResult Run(PdfReadDocument document, PdfPageSelector selector, CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(selector, nameof(selector));
        return Run(document, selector.ResolveSelection(document.Pages.Count), cancellationToken);
    }

    private PdfUnderstandingPageResult RunPage(PdfReadPage page, int pageNumber, CancellationToken cancellationToken) {
        var context = new PdfUnderstandingPageContext(
            page,
            pageNumber,
            _layout,
            _limits.MaxTextCharactersPerPage,
            _limits.MaxWordsPerPage);
        var trace = new List<PdfUnderstandingStageTrace>(6);
        IReadOnlyList<PdfTextSpan> runs = NotNull(_glyphDecoding.Decode(context), nameof(IPdfGlyphDecodingStage));
        EnsureCount(runs.Count, _limits.MaxRunsPerPage);
        EnsureTextCharacters(runs.Select(static run => run.Text), _limits.MaxTextCharactersPerPage);
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
        trace.Add(new PdfUnderstandingStageTrace("reading-order", _readingOrder.GetType(), regions.Count, ordered.Count));
        IReadOnlyList<PdfReadingOrderEvidence> readingOrderEvidence = BuildReadingOrderEvidence(ordered, _readingOrder.GetType());
        IReadOnlyList<PdfUnderstandingSemanticElement> elements = NotNull(_semanticClassification.Classify(context, ordered), nameof(IPdfSemanticClassificationStage));
        EnsureCount(elements.Count, _limits.MaxRegionsPerPage);
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

internal static class PdfFastUnderstandingStages {
    internal static readonly IPdfGlyphDecodingStage GlyphDecoding = new FastGlyphDecodingStage();
    internal static readonly IPdfWordGroupingStage WordGrouping = new FastWordGroupingStage();
    internal static readonly IPdfLineGroupingStage LineGrouping = new FastLineGroupingStage();
    internal static readonly IPdfPageSegmentationStage PageSegmentation = new FastPageSegmentationStage();
    internal static readonly IPdfReadingOrderStage ReadingOrder = new FastReadingOrderStage();
    internal static readonly IPdfSemanticClassificationStage SemanticClassification = new FastSemanticClassificationStage();

    private sealed class FastGlyphDecodingStage : IPdfGlyphDecodingStage {
        public IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context) => context.Page.GetTextSpans();
    }

    private sealed class FastWordGroupingStage : IPdfWordGroupingStage {
        public IReadOnlyList<PdfUnderstandingWord> GroupWords(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs) {
            var words = new List<PdfUnderstandingWord>();
            for (int i = 0; i < runs.Count; i++) {
                PdfTextSpan run = runs[i];
                string text = run.Text ?? string.Empty;
                int cursor = 0;
                while (cursor < text.Length) {
                    while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) cursor++;
                    int start = cursor;
                    while (cursor < text.Length && !char.IsWhiteSpace(text[cursor])) cursor++;
                    if (cursor <= start) continue;
                    if (words.Count >= context.MaxWordsPerPage) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts, context.MaxWordsPerPage, words.Count + 1L);
                    }
                    double perCharacter = run.Advance > 0D && text.Length > 0 ? run.Advance / text.Length : run.FontSize * 0.55D;
                    double xStart = run.X + (start * perCharacter);
                    double xEnd = xStart + ((cursor - start) * perCharacter);
                    bool splitRun = start > 0 || cursor < text.Length;
                    words.Add(new PdfUnderstandingWord(
                        text.Substring(start, cursor - start), xStart, xEnd, run.Y, run.FontSize, run.RotationDegrees, new[] { run },
                        splitRun ? 0.85D : 0.97D,
                        new[] { new PdfInferenceEvidence(splitRun ? "word.whitespace-split" : "word.single-run", splitRun ? "The word was split from a decoded run at whitespace." : "The decoded run maps directly to one word.", splitRun ? 0.6D : 0.9D) }));
                }
            }
            return words.Count == 0 ? Array.Empty<PdfUnderstandingWord>() : words.AsReadOnly();
        }
    }

    private sealed class FastLineGroupingStage : IPdfLineGroupingStage {
        public IReadOnlyList<PdfUnderstandingLine> GroupLines(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingWord> words) {
            var ordered = words.OrderByDescending(static word => word.BaselineY).ThenBy(static word => word.XStart).ToList();
            var groups = new List<List<PdfUnderstandingWord>>();
            for (int i = 0; i < ordered.Count; i++) {
                PdfUnderstandingWord word = ordered[i];
                List<PdfUnderstandingWord>? match = null;
                for (int j = groups.Count - 1; j >= 0; j--) {
                    PdfUnderstandingWord anchor = groups[j][0];
                    double tolerance = Math.Min(context.LayoutOptions.LineMergeMaxPoints, Math.Max(anchor.FontSize, word.FontSize) * context.LayoutOptions.LineMergeToleranceEm);
                    if (Math.Abs(anchor.BaselineY - word.BaselineY) <= tolerance && Math.Abs(anchor.RotationDegrees - word.RotationDegrees) <= 2D) { match = groups[j]; break; }
                    if (anchor.BaselineY - word.BaselineY > context.LayoutOptions.LineMergeMaxPoints * 2D) break;
                }
                if (match == null) { match = new List<PdfUnderstandingWord>(); groups.Add(match); }
                match.Add(word);
            }
            var lines = groups.Select(CreateLine).ToList();
            lines.Sort(static (left, right) => { int y = right.BaselineY.CompareTo(left.BaselineY); return y != 0 ? y : left.XStart.CompareTo(right.XStart); });
            return lines.Count == 0 ? Array.Empty<PdfUnderstandingLine>() : lines.AsReadOnly();
        }

        private static PdfUnderstandingLine CreateLine(List<PdfUnderstandingWord> group) {
            PdfUnderstandingWord[] words = group.OrderBy(static word => word.XStart).ToArray();
            double baselineSpread = words.Max(static word => word.BaselineY) - words.Min(static word => word.BaselineY);
            double rotationSpread = words.Max(static word => word.RotationDegrees) - words.Min(static word => word.RotationDegrees);
            double confidence = PdfInference.Clamp(words.Average(static word => word.Confidence) - Math.Min(0.35D, baselineSpread * 0.05D) - Math.Min(0.35D, rotationSpread * 0.05D));
            return new PdfUnderstandingLine(words, confidence, new[] {
                new PdfInferenceEvidence("line.baseline-spread", "Grouped-word baseline spread is " + baselineSpread.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " points.", baselineSpread <= 1D ? 0.8D : -Math.Min(1D, baselineSpread / 10D)),
                new PdfInferenceEvidence("line.rotation-spread", "Grouped-word rotation spread is " + rotationSpread.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " degrees.", rotationSpread <= 1D ? 0.8D : -Math.Min(1D, rotationSpread / 10D))
            });
        }
    }

    private sealed class FastPageSegmentationStage : IPdfPageSegmentationStage {
        public IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines) {
            var regions = new List<PdfUnderstandingRegion>();
            var current = new List<PdfUnderstandingLine>();
            for (int i = 0; i < lines.Count; i++) {
                PdfUnderstandingLine line = lines[i];
                if (current.Count > 0) {
                    PdfUnderstandingLine previous = current[current.Count - 1];
                    double gap = previous.BaselineY - line.BaselineY;
                    double maximumGap = Math.Max(previous.FontSize, line.FontSize) * 1.8D;
                    bool overlaps = line.XStart <= previous.XEnd + 12D && line.XEnd >= previous.XStart - 12D;
                    double smallerFont = Math.Min(previous.FontSize, line.FontSize);
                    double fontRatio = smallerFont > 0D ? Math.Max(previous.FontSize, line.FontSize) / smallerFont : 1D;
                    if (gap < 0D || gap > maximumGap || !overlaps || fontRatio > 1.15D || Math.Abs(previous.RotationDegrees - line.RotationDegrees) > 2D) {
                        regions.Add(CreateRegion(current)); current.Clear();
                    }
                }
                current.Add(line);
            }
            if (current.Count > 0) regions.Add(CreateRegion(current));
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
        }

        private static PdfUnderstandingRegion CreateRegion(List<PdfUnderstandingLine> lines) {
            double largestGap = 0D;
            for (int i = 1; i < lines.Count; i++) largestGap = Math.Max(largestGap, lines[i - 1].BaselineY - lines[i].BaselineY);
            double confidence = PdfInference.Clamp(lines.Average(static line => line.Confidence) - Math.Min(0.25D, largestGap / 100D));
            return new PdfUnderstandingRegion(lines.ToArray(), confidence, new[] {
                new PdfInferenceEvidence("region.vertical-continuity", "Largest adjacent baseline gap is " + largestGap.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + " points.", largestGap <= 24D ? 0.7D : -Math.Min(1D, largestGap / 100D)),
                new PdfInferenceEvidence("region.line-count", "The region contains " + lines.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " line(s).", lines.Count > 1 ? 0.6D : 0.2D)
            });
        }
    }

    private sealed class FastReadingOrderStage : IPdfReadingOrderStage {
        public IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions) =>
            regions.OrderByDescending(static region => region.YTop).ThenBy(static region => region.XStart).ToArray();
    }

    private sealed class FastSemanticClassificationStage : IPdfSemanticClassificationStage {
        public IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions) {
            double[] sizes = orderedRegions.SelectMany(static region => region.Lines).Select(static line => line.FontSize).OrderBy(static size => size).ToArray();
            double median = sizes.Length == 0 ? 0D : sizes.Length % 2 == 1
                ? sizes[sizes.Length / 2]
                : (sizes[(sizes.Length / 2) - 1] + sizes[sizes.Length / 2]) / 2D;
            var result = new List<PdfUnderstandingSemanticElement>(orderedRegions.Count);
            for (int i = 0; i < orderedRegions.Count; i++) {
                PdfUnderstandingRegion region = orderedRegions[i];
                (PdfUnderstandingSemanticKind kind, double confidence, PdfInferenceEvidence evidence) = ClassifyRegion(context, region, median);
                result.Add(new PdfUnderstandingSemanticElement(region, kind, confidence, new[] { evidence }));
            }
            return result.AsReadOnly();
        }

        private static (PdfUnderstandingSemanticKind Kind, double Confidence, PdfInferenceEvidence Evidence) ClassifyRegion(PdfUnderstandingPageContext context, PdfUnderstandingRegion region, double median) {
            if (context.LayoutOptions.IgnoreHeaderHeight > 0D && region.YTop >= context.Height - context.LayoutOptions.IgnoreHeaderHeight) return (PdfUnderstandingSemanticKind.Header, 0.9D, new PdfInferenceEvidence("semantic.header-band", "The region falls inside the configured header band.", 0.9D));
            if (context.LayoutOptions.IgnoreFooterHeight > 0D && region.YBottom <= context.LayoutOptions.IgnoreFooterHeight) return (PdfUnderstandingSemanticKind.Footer, 0.9D, new PdfInferenceEvidence("semantic.footer-band", "The region falls inside the configured footer band.", 0.9D));
            string text = region.Text.TrimStart();
            if ((text.Length > 0 && (text[0] == '-' || text[0] == '*' || text[0] == '•')) || StartsWithNumberedMarker(text)) return (PdfUnderstandingSemanticKind.ListItem, 0.9D, new PdfInferenceEvidence("semantic.list-marker", "The region begins with a recognized bullet or numbered marker.", 0.9D));
            double largest = region.Lines.Count == 0 ? 0D : region.Lines.Max(static line => line.FontSize);
            if (median > 0D && largest >= median * 1.2D) return (PdfUnderstandingSemanticKind.Heading, PdfInference.Clamp(0.65D + Math.Min(0.3D, (largest / median - 1.2D) * 0.5D)), new PdfInferenceEvidence("semantic.font-size-heading", "The largest font is materially larger than the page median.", 0.8D));
            return (PdfUnderstandingSemanticKind.Paragraph, 0.7D, new PdfInferenceEvidence("semantic.body-font", "The region uses the page's body-font range and has no stronger semantic marker.", 0.5D));
        }

        private static bool StartsWithNumberedMarker(string text) {
            int index = 0;
            while (index < text.Length && char.IsDigit(text[index])) index++;
            return index > 0 && index < text.Length && (text[index] == '.' || text[index] == ')');
        }
    }
}
