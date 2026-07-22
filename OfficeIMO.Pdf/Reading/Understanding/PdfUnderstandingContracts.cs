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
    internal PdfUnderstandingPageContext(PdfReadPage page, int pageNumber, PdfTextLayoutOptions options,
        int maxTextCharactersPerPage, int maxWordsPerPage) {
        Page = page;
        PageNumber = pageNumber;
        LayoutOptions = options;
        (double width, double height) = page.GetPageSize();
        Width = width;
        Height = height;
        MaxTextCharactersPerPage = maxTextCharactersPerPage;
        MaxWordsPerPage = maxWordsPerPage;
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
}

/// <summary>One decoded word candidate with source-run traceability.</summary>
public sealed class PdfUnderstandingWord {
    /// <summary>Creates a positioned word artifact for a custom grouping stage.</summary>
    public PdfUnderstandingWord(string text, double xStart, double xEnd, double baselineY, double fontSize, double rotationDegrees, IReadOnlyList<PdfTextSpan> sourceRuns, double confidence = 1D, IEnumerable<PdfInferenceEvidence>? evidence = null) {
        Guard.NotNull(text, nameof(text)); Guard.NotNull(sourceRuns, nameof(sourceRuns));
        Text = text; XStart = xStart; XEnd = xEnd; BaselineY = baselineY; FontSize = fontSize; RotationDegrees = rotationDegrees; SourceRuns = sourceRuns; Confidence = PdfInference.Clamp(confidence); Evidence = PdfInference.Snapshot(evidence);
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
    public PdfUnderstandingLine(IReadOnlyList<PdfUnderstandingWord> words, double? confidence = null, IEnumerable<PdfInferenceEvidence>? evidence = null) {
        Guard.NotNull(words, nameof(words));
        if (words.Count == 0) throw new ArgumentException("A line requires at least one word.", nameof(words));
        Words = words;
        Text = string.Join(" ", words.Select(static word => word.Text));
        XStart = words.Min(static word => word.XStart);
        XEnd = words.Max(static word => word.XEnd);
        BaselineY = words.Average(static word => word.BaselineY);
        FontSize = words.Max(static word => word.FontSize);
        RotationDegrees = words.Average(static word => word.RotationDegrees);
        Confidence = PdfInference.Clamp(confidence ?? words.Average(static word => word.Confidence));
        Evidence = PdfInference.Snapshot(evidence);
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
    /// <summary>Creates a semantic classification for a region.</summary>
    public PdfUnderstandingSemanticElement(PdfUnderstandingRegion region, PdfUnderstandingSemanticKind kind, double confidence = 0.5D, IEnumerable<PdfInferenceEvidence>? evidence = null) { Guard.NotNull(region, nameof(region)); Region = region; Kind = kind; Confidence = PdfInference.Clamp(confidence); Evidence = PdfInference.Snapshot(evidence); }
    /// <summary>Classified region.</summary>
    public PdfUnderstandingRegion Region { get; }
    /// <summary>Semantic kind selected by the active stage.</summary>
    public PdfUnderstandingSemanticKind Kind { get; }
    /// <summary>Normalized classification confidence from 0 to 1.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the semantic classification.</summary>
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
    internal PdfUnderstandingPageResult(int pageNumber, IReadOnlyList<PdfTextSpan> runs, IReadOnlyList<PdfUnderstandingWord> words, IReadOnlyList<PdfUnderstandingLine> lines, IReadOnlyList<PdfUnderstandingRegion> regions, IReadOnlyList<PdfUnderstandingRegion> readingOrder, IReadOnlyList<PdfReadingOrderEvidence> readingOrderEvidence, IReadOnlyList<PdfUnderstandingSemanticElement> elements, IReadOnlyList<PdfUnderstandingStageTrace> trace) {
        PageNumber = pageNumber; DecodedRuns = runs; Words = words; Lines = lines; Regions = regions; ReadingOrder = readingOrder; ReadingOrderEvidence = readingOrderEvidence; Elements = elements; Trace = trace;
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
    /// <summary>Stage execution trace.</summary>
    public IReadOnlyList<PdfUnderstandingStageTrace> Trace { get; }
}

/// <summary>Document-wide understanding result.</summary>
public sealed class PdfUnderstandingResult {
    internal PdfUnderstandingResult(IReadOnlyList<PdfUnderstandingPageResult> pages) { Pages = pages; }
    /// <summary>Page results in caller-selected order.</summary>
    public IReadOnlyList<PdfUnderstandingPageResult> Pages { get; }
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
