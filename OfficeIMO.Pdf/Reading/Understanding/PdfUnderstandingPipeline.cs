using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Configures independently replaceable PDF understanding stages.</summary>
public sealed class PdfUnderstandingPipelineOptions {
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
        if (_maxPages <= 0) throw new ArgumentOutOfRangeException(nameof(options), effective.MaxPages, "Maximum pages must be positive.");
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
            pages.Add(RunPage(document.Pages[pageNumber - 1], pageNumber));
        }
        return new PdfUnderstandingResult(pages.AsReadOnly());
    }

    private PdfUnderstandingPageResult RunPage(PdfReadPage page, int pageNumber) {
        var context = new PdfUnderstandingPageContext(page, pageNumber, _layout);
        var trace = new List<PdfUnderstandingStageTrace>(6);
        IReadOnlyList<PdfTextSpan> runs = NotNull(_glyphDecoding.Decode(context), nameof(IPdfGlyphDecodingStage));
        trace.Add(new PdfUnderstandingStageTrace("glyph-decoding", _glyphDecoding.GetType(), 0, runs.Count));
        IReadOnlyList<PdfUnderstandingWord> words = NotNull(_wordGrouping.GroupWords(context, runs), nameof(IPdfWordGroupingStage));
        trace.Add(new PdfUnderstandingStageTrace("word-grouping", _wordGrouping.GetType(), runs.Count, words.Count));
        IReadOnlyList<PdfUnderstandingLine> lines = NotNull(_lineGrouping.GroupLines(context, words), nameof(IPdfLineGroupingStage));
        trace.Add(new PdfUnderstandingStageTrace("line-grouping", _lineGrouping.GetType(), words.Count, lines.Count));
        IReadOnlyList<PdfUnderstandingRegion> regions = NotNull(_pageSegmentation.Segment(context, lines), nameof(IPdfPageSegmentationStage));
        trace.Add(new PdfUnderstandingStageTrace("page-segmentation", _pageSegmentation.GetType(), lines.Count, regions.Count));
        IReadOnlyList<PdfUnderstandingRegion> ordered = NotNull(_readingOrder.Order(context, regions), nameof(IPdfReadingOrderStage));
        trace.Add(new PdfUnderstandingStageTrace("reading-order", _readingOrder.GetType(), regions.Count, ordered.Count));
        IReadOnlyList<PdfUnderstandingSemanticElement> elements = NotNull(_semanticClassification.Classify(context, ordered), nameof(IPdfSemanticClassificationStage));
        trace.Add(new PdfUnderstandingStageTrace("semantic-classification", _semanticClassification.GetType(), ordered.Count, elements.Count));
        return new PdfUnderstandingPageResult(pageNumber, runs, words, lines, regions, ordered, elements, trace.AsReadOnly());
    }

    private static IReadOnlyList<T> NotNull<T>(IReadOnlyList<T>? value, string stage) => value ?? throw new InvalidOperationException(stage + " returned null.");
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
                    double perCharacter = run.Advance > 0D && text.Length > 0 ? run.Advance / text.Length : run.FontSize * 0.55D;
                    double xStart = run.X + (start * perCharacter);
                    double xEnd = xStart + ((cursor - start) * perCharacter);
                    words.Add(new PdfUnderstandingWord(text.Substring(start, cursor - start), xStart, xEnd, run.Y, run.FontSize, run.RotationDegrees, new[] { run }));
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
            var lines = groups.Select(static group => new PdfUnderstandingLine(group.OrderBy(static word => word.XStart).ToArray())).ToList();
            lines.Sort(static (left, right) => { int y = right.BaselineY.CompareTo(left.BaselineY); return y != 0 ? y : left.XStart.CompareTo(right.XStart); });
            return lines.Count == 0 ? Array.Empty<PdfUnderstandingLine>() : lines.AsReadOnly();
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
                        regions.Add(new PdfUnderstandingRegion(current.ToArray())); current.Clear();
                    }
                }
                current.Add(line);
            }
            if (current.Count > 0) regions.Add(new PdfUnderstandingRegion(current.ToArray()));
            return regions.Count == 0 ? Array.Empty<PdfUnderstandingRegion>() : regions.AsReadOnly();
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
                PdfUnderstandingSemanticKind kind = ClassifyRegion(context, region, median);
                result.Add(new PdfUnderstandingSemanticElement(region, kind));
            }
            return result.AsReadOnly();
        }

        private static PdfUnderstandingSemanticKind ClassifyRegion(PdfUnderstandingPageContext context, PdfUnderstandingRegion region, double median) {
            if (context.LayoutOptions.IgnoreHeaderHeight > 0D && region.YTop >= context.Height - context.LayoutOptions.IgnoreHeaderHeight) return PdfUnderstandingSemanticKind.Header;
            if (context.LayoutOptions.IgnoreFooterHeight > 0D && region.YBottom <= context.LayoutOptions.IgnoreFooterHeight) return PdfUnderstandingSemanticKind.Footer;
            string text = region.Text.TrimStart();
            if ((text.Length > 0 && (text[0] == '-' || text[0] == '*' || text[0] == '•')) || StartsWithNumberedMarker(text)) return PdfUnderstandingSemanticKind.ListItem;
            double largest = region.Lines.Count == 0 ? 0D : region.Lines.Max(static line => line.FontSize);
            return median > 0D && largest >= median * 1.2D ? PdfUnderstandingSemanticKind.Heading : PdfUnderstandingSemanticKind.Paragraph;
        }

        private static bool StartsWithNumberedMarker(string text) {
            int index = 0;
            while (index < text.Length && char.IsDigit(text[index])) index++;
            return index > 0 && index < text.Length && (text[index] == '.' || text[index] == ')');
        }
    }
}
