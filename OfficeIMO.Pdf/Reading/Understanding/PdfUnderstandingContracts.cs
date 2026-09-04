using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Semantic classifications produced by a PDF understanding pipeline.</summary>
public enum PdfUnderstandingSemanticKind {
    /// <summary>Ordinary paragraph or body region.</summary>
    Paragraph,
    /// <summary>Heading-like region.</summary>
    Heading,
    /// <summary>List-item-like region.</summary>
    ListItem,
    /// <summary>Repeated or page-edge header region.</summary>
    Header,
    /// <summary>Repeated or page-edge footer region.</summary>
    Footer,
    /// <summary>Caption-like region.</summary>
    Caption,
    /// <summary>Footnote-like region near the bottom of a page.</summary>
    Footnote,
    /// <summary>Table-like region.</summary>
    Table,
    /// <summary>Region not classified by the active strategy.</summary>
    Unknown
}

/// <summary>Page and option context shared by understanding stages.</summary>
public sealed class PdfUnderstandingPageContext {
    private readonly PdfUnderstandingWorkBudget _workBudget;

    internal PdfUnderstandingPageContext(PdfReadPage page, int pageNumber, PdfTextLayoutOptions options,
        int maxTextCharactersPerPage, int maxWordsPerPage,
        long maxWorkUnitsPerPage = 10_000_000,
        CancellationToken cancellationToken = default) {
        Page = page;
        PageNumber = pageNumber;
        LayoutOptions = options;
        (double width, double height) = page.GetPageSize();
        Width = width;
        Height = height;
        MaxTextCharactersPerPage = maxTextCharactersPerPage;
        MaxWordsPerPage = maxWordsPerPage;
        _workBudget = new PdfUnderstandingWorkBudget(maxWorkUnitsPerPage, cancellationToken);
    }

    /// <summary>Parsed source page.</summary>
    public PdfReadPage Page { get; }
    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }
    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }
    /// <summary>Layout options supplied to the pipeline.</summary>
    public PdfTextLayoutOptions LayoutOptions { get; }
    /// <summary>Maximum decoded text characters accepted for this page.</summary>
    public int MaxTextCharactersPerPage { get; }
    /// <summary>Maximum word artifacts accepted for this page.</summary>
    public int MaxWordsPerPage { get; }
    /// <summary>Maximum table candidates accepted for this page.</summary>
    public int MaxTableCandidatesPerPage { get; internal set; } = 1_024;
    /// <summary>Maximum viable image-caption association edges retained before deterministic matching.</summary>
    public int MaxImageCaptionCandidatesPerPage { get; internal set; } = 100_000;
    /// <summary>Cancellation token observed by built-in and cooperative custom stages.</summary>
    public CancellationToken CancellationToken => _workBudget.CancellationToken;
    /// <summary>Maximum comparison and traversal work units available to this page.</summary>
    public long MaxWorkUnitsPerPage => _workBudget.Maximum;
    /// <summary>Work units consumed so far by semantic reconstruction for this page.</summary>
    public long WorkUnitsConsumed => _workBudget.Consumed;
    /// <summary>Charges comparison or traversal work and observes cancellation.</summary>
    public void ConsumeWork(long units = 1) => _workBudget.Consume(units);
    /// <summary>Throws when semantic reconstruction has been cancelled.</summary>
    public void ThrowIfCancellationRequested() => _workBudget.ThrowIfCancellationRequested();
    internal void CompleteOperation() => _workBudget.CompleteOperation();
    internal IReadOnlyList<PdfTextSpan> DecodedRuns { get; set; } = Array.Empty<PdfTextSpan>();
    /// <summary>Table candidates available to page segmentation and later stages.</summary>
    public IReadOnlyList<PdfUnderstandingTableCandidate> TableCandidates { get; internal set; } = Array.Empty<PdfUnderstandingTableCandidate>();
    /// <summary>Image placement invocations available to semantic stages in stable paint order.</summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements { get; internal set; } = Array.Empty<PdfImagePlacement>();
}

/// <summary>One decoded word candidate with source-run traceability.</summary>
public sealed class PdfUnderstandingWord {
    /// <summary>Creates a positioned word artifact for a custom grouping stage.</summary>
    public PdfUnderstandingWord(string text, double xStart, double xEnd, double baselineY, double fontSize, double rotationDegrees, IReadOnlyList<PdfTextSpan> sourceRuns, double confidence = 1D, IEnumerable<PdfInferenceEvidence>? evidence = null, double? advance = null, PdfLogicalVisualBounds? visualBounds = null, int? sourceSequence = null) {
        Guard.NotNull(text, nameof(text)); Guard.NotNull(sourceRuns, nameof(sourceRuns));
        Text = text; XStart = xStart; XEnd = xEnd; BaselineY = baselineY; FontSize = fontSize; RotationDegrees = rotationDegrees; SourceRuns = sourceRuns; Confidence = PdfInference.Clamp(confidence); Evidence = PdfInference.Snapshot(evidence); Advance = advance; VisualBounds = visualBounds; SourceSequence = sourceSequence;
    }
    /// <summary>Decoded word text.</summary>
    public string Text { get; }
    /// <summary>Left bound in PDF points.</summary>
    public double XStart { get; }
    /// <summary>Right bound in PDF points.</summary>
    public double XEnd { get; }
    /// <summary>Baseline Y in PDF points.</summary>
    public double BaselineY { get; }
    /// <summary>Representative font size.</summary>
    public double FontSize { get; }
    /// <summary>Baseline rotation in degrees.</summary>
    public double RotationDegrees { get; }
    /// <summary>Distance occupied along the baseline, including vertical and rotated baselines, when known.</summary>
    public double? Advance { get; }
    /// <summary>Direct top-left visual bounds, when supplied by a positioned source such as OCR.</summary>
    public PdfLogicalVisualBounds? VisualBounds { get; }
    /// <summary>Original zero-based source order, when supplied by the decoder or positioned provider.</summary>
    public int? SourceSequence { get; }
    /// <summary>Decoded source runs that produced this word.</summary>
    public IReadOnlyList<PdfTextSpan> SourceRuns { get; }
    /// <summary>Normalized grouping confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting this word grouping.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>One grouped text line.</summary>
public sealed class PdfUnderstandingLine {
    /// <summary>Creates a line from words in local reading order.</summary>
    public PdfUnderstandingLine(IReadOnlyList<PdfUnderstandingWord> words, double? confidence = null, IEnumerable<PdfInferenceEvidence>? evidence = null)
        : this(words, JoinWords(words), confidence, evidence) {
    }

    internal PdfUnderstandingLine(
        IReadOnlyList<PdfUnderstandingWord> words,
        string text,
        double? confidence,
        IEnumerable<PdfInferenceEvidence>? evidence,
        PdfLogicalContentSourceKind sourceKind = PdfLogicalContentSourceKind.Native,
        int? sourceSequence = null,
        string? blockId = null,
        string? paragraphId = null,
        string? lineId = null,
        PdfLogicalVisualBounds? visualBounds = null) {
        Guard.NotNull(words, nameof(words));
        Guard.NotNull(text, nameof(text));
        if (words.Count == 0) throw new ArgumentException("A line requires at least one word.", nameof(words));
        Words = words;
        Text = text;
        XStart = words.Min(static word => word.XStart);
        XEnd = words.Max(static word => word.XEnd);
        BaselineY = words.Average(static word => word.BaselineY);
        FontSize = words.Max(static word => word.FontSize);
        RotationDegrees = words.Average(static word => word.RotationDegrees);
        Confidence = PdfInference.Clamp(confidence ?? words.Average(static word => word.Confidence));
        Evidence = PdfInference.Snapshot(evidence);
        SourceKind = sourceKind;
        SourceSequence = sourceSequence;
        BlockId = blockId;
        ParagraphId = paragraphId;
        LineId = lineId;
        VisualBounds = visualBounds;
    }

    private static string JoinWords(IReadOnlyList<PdfUnderstandingWord> words) {
        Guard.NotNull(words, nameof(words));
        return string.Join(" ", words.Select(static word => word.Text));
    }
    /// <summary>Words in line order.</summary>
    public IReadOnlyList<PdfUnderstandingWord> Words { get; }
    /// <summary>Normalized line text.</summary>
    public string Text { get; }
    /// <summary>Left bound.</summary>
    public double XStart { get; }
    /// <summary>Right bound.</summary>
    public double XEnd { get; }
    /// <summary>Representative baseline.</summary>
    public double BaselineY { get; }
    /// <summary>Largest font size in the line.</summary>
    public double FontSize { get; }
    /// <summary>Representative baseline rotation.</summary>
    public double RotationDegrees { get; }
    /// <summary>Normalized line-grouping confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting this line grouping.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
    /// <summary>Whether the line came from native PDF text or accepted OCR geometry.</summary>
    public PdfLogicalContentSourceKind SourceKind { get; }
    /// <summary>Original source reading position, when supplied.</summary>
    public int? SourceSequence { get; }
    /// <summary>Provider block identifier, when supplied.</summary>
    public string? BlockId { get; }
    /// <summary>Provider paragraph identifier, when supplied.</summary>
    public string? ParagraphId { get; }
    /// <summary>Provider line identifier, when supplied.</summary>
    public string? LineId { get; }
    /// <summary>Direct top-left visual bounds, when supplied.</summary>
    public PdfLogicalVisualBounds? VisualBounds { get; }
}

/// <summary>One page-segmentation region containing related lines.</summary>
public sealed class PdfUnderstandingRegion {
    /// <summary>Creates a segmented region from lines in local reading order.</summary>
    public PdfUnderstandingRegion(IReadOnlyList<PdfUnderstandingLine> lines, double? confidence = null, IEnumerable<PdfInferenceEvidence>? evidence = null) {
        Guard.NotNull(lines, nameof(lines));
        if (lines.Count == 0) throw new ArgumentException("A region requires at least one line.", nameof(lines));
        Lines = lines;
        Text = string.Join(" ", lines.Select(static line => line.Text));
        XStart = lines.Min(static line => line.XStart);
        XEnd = lines.Max(static line => line.XEnd);
        YTop = lines.Max(static line => line.BaselineY);
        YBottom = lines.Min(static line => line.BaselineY);
        Confidence = PdfInference.Clamp(confidence ?? lines.Average(static line => line.Confidence));
        Evidence = PdfInference.Snapshot(evidence);
    }
    /// <summary>Lines in local region order.</summary>
    public IReadOnlyList<PdfUnderstandingLine> Lines { get; }
    /// <summary>Normalized region text.</summary>
    public string Text { get; }
    /// <summary>Left bound.</summary>
    public double XStart { get; }
    /// <summary>Right bound.</summary>
    public double XEnd { get; }
    /// <summary>Top baseline.</summary>
    public double YTop { get; }
    /// <summary>Bottom baseline.</summary>
    public double YBottom { get; }
    /// <summary>Normalized segmentation confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting this region segmentation.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>Semantic classification of one ordered region.</summary>
public sealed class PdfUnderstandingSemanticElement {
    /// <summary>Creates a semantic classification for a region without an explicit hierarchy level.</summary>
    public PdfUnderstandingSemanticElement(
        PdfUnderstandingRegion region,
        PdfUnderstandingSemanticKind kind,
        double confidence,
        IEnumerable<PdfInferenceEvidence>? evidence)
        : this(region, kind, confidence, evidence, level: null) { }

    /// <summary>Creates a semantic classification for a region.</summary>
    public PdfUnderstandingSemanticElement(PdfUnderstandingRegion region, PdfUnderstandingSemanticKind kind, double confidence = 0.5D, IEnumerable<PdfInferenceEvidence>? evidence = null, int? level = null) { Guard.NotNull(region, nameof(region)); if (level <= 0) throw new ArgumentOutOfRangeException(nameof(level)); Region = region; Kind = kind; Confidence = PdfInference.Clamp(confidence); Evidence = PdfInference.Snapshot(evidence); Level = level; }
    /// <summary>Classified region.</summary>
    public PdfUnderstandingRegion Region { get; }
    /// <summary>Semantic kind selected by the active stage.</summary>
    public PdfUnderstandingSemanticKind Kind { get; }
    /// <summary>Normalized classification confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the semantic classification.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
    /// <summary>Best-evidence hierarchy level for headings or list items, when known.</summary>
    public int? Level { get; }
}

/// <summary>One positioned image region with an optional associated caption.</summary>
public sealed class PdfUnderstandingImageRegion {
    /// <summary>Creates an image region for a custom detection stage.</summary>
    public PdfUnderstandingImageRegion(
        PdfImagePlacement placement,
        PdfUnderstandingSemanticElement? caption = null,
        double confidence = 0.5D,
        IEnumerable<PdfInferenceEvidence>? evidence = null,
        bool isFigure = false,
        string? alternativeText = null) {
        Guard.NotNull(placement, nameof(placement));
        if (caption is not null && caption.Kind != PdfUnderstandingSemanticKind.Caption) {
            throw new ArgumentException("An associated image caption must have Caption semantics.", nameof(caption));
        }
        Placement = placement;
        Caption = caption;
        IsFigure = isFigure || caption is not null;
        AlternativeText = string.IsNullOrWhiteSpace(alternativeText) ? null : alternativeText;
        Confidence = PdfInference.Clamp(confidence);
        Evidence = PdfInference.Snapshot(evidence);
    }

    /// <summary>Source image placement represented by this region.</summary>
    public PdfImagePlacement Placement { get; }

    /// <summary>Caption associated by structural or geometric evidence, when available.</summary>
    public PdfUnderstandingSemanticElement? Caption { get; }

    /// <summary>True when a caption association or tagged-PDF Figure role establishes a figure.</summary>
    public bool IsFigure { get; }

    /// <summary>Tagged-PDF alternate text for the figure, when present.</summary>
    public string? AlternativeText { get; }

    /// <summary>Normalized image-region and association confidence from 0 to 1.</summary>
    public double Confidence { get; }

    /// <summary>Stable evidence supporting the image region and optional caption association.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>Trace record proving which stage implementation produced an artifact set.</summary>
public sealed class PdfUnderstandingStageTrace {
    internal PdfUnderstandingStageTrace(string stage, Type providerType, int inputCount, int outputCount) { Stage = stage; ProviderType = providerType; InputCount = inputCount; OutputCount = outputCount; }
    /// <summary>Stable stage name.</summary>
    public string Stage { get; }
    /// <summary>Concrete provider type.</summary>
    public Type ProviderType { get; }
    /// <summary>Input artifact count.</summary>
    public int InputCount { get; }
    /// <summary>Output artifact count.</summary>
    public int OutputCount { get; }
}

/// <summary>All intermediate and final artifacts for one page.</summary>
public sealed class PdfUnderstandingPageResult {
    private readonly Action? _completeOperation;

    internal PdfUnderstandingPageResult(
        int pageNumber,
        IReadOnlyList<PdfTextSpan> runs,
        IReadOnlyList<PdfUnderstandingWord> words,
        IReadOnlyList<PdfUnderstandingLine> lines,
        IReadOnlyList<PdfUnderstandingRegion> regions,
        IReadOnlyList<PdfUnderstandingRegion> readingOrder,
        IReadOnlyList<PdfReadingOrderEvidence> readingOrderEvidence,
        IReadOnlyList<PdfUnderstandingSemanticElement> elements,
        IReadOnlyList<PdfUnderstandingStageTrace> trace,
        Action<long>? consumeWork = null,
        Action? cancellationCheck = null,
        Action? completeOperation = null,
        IReadOnlyList<PdfUnderstandingLine>? logicalProjectionLines = null,
        bool restrictLogicalProjectionToReadingOrder = false,
        IReadOnlyList<PdfUnderstandingTableCandidate>? tableCandidates = null,
        IReadOnlyList<PdfImagePlacement>? imagePlacements = null,
        IReadOnlyList<PdfUnderstandingImageRegion>? imageRegions = null) {
        PageNumber = pageNumber;
        DecodedRuns = runs;
        Words = words;
        Lines = lines;
        Regions = regions;
        ReadingOrder = readingOrder;
        ReadingOrderEvidence = readingOrderEvidence;
        Elements = elements;
        Trace = trace;
        TableCandidates = tableCandidates ?? Array.Empty<PdfUnderstandingTableCandidate>();
        ImagePlacements = imagePlacements ?? Array.Empty<PdfImagePlacement>();
        ImageRegions = imageRegions ?? Array.Empty<PdfUnderstandingImageRegion>();
        LogicalProjectionLines = logicalProjectionLines ?? CollectLogicalProjectionLines(readingOrder);
        RestrictLogicalProjectionToReadingOrder = restrictLogicalProjectionToReadingOrder;
        ConsumeWork = consumeWork;
        CancellationCheck = cancellationCheck;
        _completeOperation = completeOperation;
    }
    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Decoded text runs.</summary>
    public IReadOnlyList<PdfTextSpan> DecodedRuns { get; }
    /// <summary>Grouped words.</summary>
    public IReadOnlyList<PdfUnderstandingWord> Words { get; }
    /// <summary>Grouped lines.</summary>
    public IReadOnlyList<PdfUnderstandingLine> Lines { get; }
    /// <summary>Page-segmentation regions.</summary>
    public IReadOnlyList<PdfUnderstandingRegion> Regions { get; }
    /// <summary>Regions in inferred reading order.</summary>
    public IReadOnlyList<PdfUnderstandingRegion> ReadingOrder { get; }
    /// <summary>Confidence and evidence for every inferred reading-order position.</summary>
    public IReadOnlyList<PdfReadingOrderEvidence> ReadingOrderEvidence { get; }
    /// <summary>Semantically classified ordered regions.</summary>
    public IReadOnlyList<PdfUnderstandingSemanticElement> Elements { get; }
    /// <summary>Tables recovered before general page segmentation.</summary>
    public IReadOnlyList<PdfUnderstandingTableCandidate> TableCandidates { get; }
    /// <summary>Image placement invocations recovered once for semantic analysis and logical projection.</summary>
    public IReadOnlyList<PdfImagePlacement> ImagePlacements { get; }
    /// <summary>Positioned image regions and their evidence-backed caption associations.</summary>
    public IReadOnlyList<PdfUnderstandingImageRegion> ImageRegions { get; }
    /// <summary>Stage execution trace.</summary>
    public IReadOnlyList<PdfUnderstandingStageTrace> Trace { get; }
    internal Action<long>? ConsumeWork { get; }
    internal Action? CancellationCheck { get; }
    /// <summary>Pre-enrichment line sequence retained by caller-supplied structural stages for logical projection.</summary>
    internal IReadOnlyList<PdfUnderstandingLine> LogicalProjectionLines { get; }
    /// <summary>Whether caller-supplied structural stages make the retained sequence an extraction boundary.</summary>
    internal bool RestrictLogicalProjectionToReadingOrder { get; }
    internal void CompleteOperation() => _completeOperation?.Invoke();

    internal PdfUnderstandingPageResult WithAdditionalTableCandidates(
        IReadOnlyList<PdfUnderstandingTableCandidate> candidates) {
        if (candidates.Count == 0) return this;
        return new PdfUnderstandingPageResult(
            PageNumber,
            DecodedRuns,
            Words,
            Lines,
            Regions,
            ReadingOrder,
            ReadingOrderEvidence,
            Elements,
            Trace,
            ConsumeWork,
            CancellationCheck,
            _completeOperation,
            LogicalProjectionLines,
            RestrictLogicalProjectionToReadingOrder,
            TableCandidates.Concat(candidates).ToArray(),
            ImagePlacements,
            ImageRegions);
    }

    private static IReadOnlyList<PdfUnderstandingLine> CollectLogicalProjectionLines(
        IReadOnlyList<PdfUnderstandingRegion> readingOrder) {
        var result = new List<PdfUnderstandingLine>();
        var seen = new HashSet<PdfUnderstandingLine>();
        for (int regionIndex = 0; regionIndex < readingOrder.Count; regionIndex++) {
            IReadOnlyList<PdfUnderstandingLine> lines = readingOrder[regionIndex].Lines;
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                if (seen.Add(lines[lineIndex])) result.Add(lines[lineIndex]);
            }
        }
        return result.Count == 0 ? Array.Empty<PdfUnderstandingLine>() : result.AsReadOnly();
    }

    internal static PdfUnderstandingPageResult Empty(int pageNumber) => new PdfUnderstandingPageResult(
        pageNumber,
        Array.Empty<PdfTextSpan>(),
        Array.Empty<PdfUnderstandingWord>(),
        Array.Empty<PdfUnderstandingLine>(),
        Array.Empty<PdfUnderstandingRegion>(),
        Array.Empty<PdfUnderstandingRegion>(),
        Array.Empty<PdfReadingOrderEvidence>(),
        Array.Empty<PdfUnderstandingSemanticElement>(),
        Array.Empty<PdfUnderstandingStageTrace>());
}

/// <summary>Decodes page glyph/text content into positioned runs.</summary>
public interface IPdfGlyphDecodingStage {
    /// <summary>Decodes the page into positioned text runs.</summary>
    IReadOnlyList<PdfTextSpan> Decode(PdfUnderstandingPageContext context);
}
/// <summary>Groups decoded runs into words.</summary>
public interface IPdfWordGroupingStage {
    /// <summary>Groups decoded runs into word artifacts.</summary>
    IReadOnlyList<PdfUnderstandingWord> GroupWords(PdfUnderstandingPageContext context, IReadOnlyList<PdfTextSpan> runs);
}
/// <summary>Groups words into lines.</summary>
public interface IPdfLineGroupingStage {
    /// <summary>Groups words into line artifacts.</summary>
    IReadOnlyList<PdfUnderstandingLine> GroupLines(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingWord> words);
}
/// <summary>Detects table candidates before general page segmentation.</summary>
public interface IPdfTableDetectionStage {
    /// <summary>Returns table candidates and the source lines owned by each table.</summary>
    IReadOnlyList<PdfUnderstandingTableCandidate> DetectTables(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines);
}
/// <summary>Segments page lines into regions.</summary>
public interface IPdfPageSegmentationStage {
    /// <summary>Segments lines into page regions.</summary>
    IReadOnlyList<PdfUnderstandingRegion> Segment(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingLine> lines);
}
/// <summary>Orders page regions for reading.</summary>
public interface IPdfReadingOrderStage {
    /// <summary>Returns the regions in inferred reading order.</summary>
    IReadOnlyList<PdfUnderstandingRegion> Order(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> regions);
}
/// <summary>Classifies ordered regions semantically.</summary>
public interface IPdfSemanticClassificationStage {
    /// <summary>Classifies the ordered regions.</summary>
    IReadOnlyList<PdfUnderstandingSemanticElement> Classify(PdfUnderstandingPageContext context, IReadOnlyList<PdfUnderstandingRegion> orderedRegions);
}
/// <summary>Detects positioned image regions and associates classified captions.</summary>
public interface IPdfImageRegionDetectionStage {
    /// <summary>Returns one region per image placement invocation with an optional caption association.</summary>
    IReadOnlyList<PdfUnderstandingImageRegion> Detect(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingSemanticElement> semanticElements);
}
