using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Logical view of a single PDF page.
/// </summary>
public sealed partial class PdfLogicalPage {
    private IReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>>? _elementsByKind;
    private IReadOnlyList<PdfLogicalTextBlock>? _headers;
    private IReadOnlyList<PdfLogicalTextBlock>? _footers;
    private IReadOnlyList<PdfLogicalTextBlock>? _captions;
    private IReadOnlyList<PdfLogicalTextBlock>? _footnotes;

    private PdfLogicalPage(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        PdfPageGeometry geometry,
        IReadOnlyList<IPdfLogicalElement> elements,
        IReadOnlyList<PdfLogicalTextBlock> textBlocks,
        IReadOnlyList<PdfLogicalHeading> headings,
        IReadOnlyList<PdfLogicalParagraph> paragraphs,
        IReadOnlyList<PdfLogicalListItem> listItems,
        IReadOnlyList<PdfLogicalTable> tables,
        int vectorPrimitiveCount,
        int unrepresentedVectorPrimitiveCount,
        IReadOnlyList<PdfLogicalImage> images,
        IReadOnlyList<PdfLogicalLinkAnnotation> links,
        IReadOnlyList<PdfAnnotation> annotations,
        IReadOnlyList<PdfLinkAnnotation> linkAnnotations,
        IReadOnlyList<PdfLogicalFormWidget> formWidgets,
        IReadOnlyList<PdfPageAction> pageActions,
        PdfUnderstandingPageResult analysis) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        Geometry = geometry;
        Elements = elements;
        TextBlocks = textBlocks;
        Headings = headings;
        Paragraphs = paragraphs;
        ListItems = listItems;
        Tables = tables;
        VectorPrimitiveCount = vectorPrimitiveCount;
        UnrepresentedVectorPrimitiveCount = unrepresentedVectorPrimitiveCount;
        Images = images;
        Links = links;
        Annotations = annotations;
        LinkAnnotations = linkAnnotations;
        FormWidgets = formWidgets;
        PageActions = pageActions;
        Analysis = analysis;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Inherited page rotation normalized to 0, 90, 180, or 270.</summary>
    public int RotationDegrees { get; }

    /// <summary>Page boundary boxes and page-level presentation metadata.</summary>
    public PdfPageGeometry Geometry { get; }

    /// <summary>Inherited /MediaBox boundary, when readable.</summary>
    public PdfPageBox? MediaBox => Geometry.MediaBox;

    /// <summary>Inherited /CropBox boundary, when readable.</summary>
    public PdfPageBox? CropBox => Geometry.CropBox;

    /// <summary>Inherited /BleedBox boundary, when readable.</summary>
    public PdfPageBox? BleedBox => Geometry.BleedBox;

    /// <summary>Inherited /TrimBox boundary, when readable.</summary>
    public PdfPageBox? TrimBox => Geometry.TrimBox;

    /// <summary>Inherited /ArtBox boundary, when readable.</summary>
    public PdfPageBox? ArtBox => Geometry.ArtBox;

    /// <summary>Inherited page user-unit scale from /UserUnit, when present and positive.</summary>
    public double? UserUnit => Geometry.UserUnit;

    /// <summary>Page tab order from /Tabs, when present.</summary>
    public string? TabOrder => Geometry.TabOrder;

    /// <summary>Page display duration from /Dur, in seconds, when present.</summary>
    public double? DurationSeconds => Geometry.DurationSeconds;

    /// <summary>Page transition dictionary from /Trans, when present and readable.</summary>
    public PdfPageTransition? Transition => Geometry.Transition;

    /// <summary>True when page-level /Metadata was present.</summary>
    public bool HasPageMetadata => Geometry.HasMetadata;

    /// <summary>True when page-level /PieceInfo was present.</summary>
    public bool HasPieceInfo => Geometry.HasPieceInfo;

    internal (double Width, double Height) GetVisualPageSize() {
        PdfPageBox pageBox = GetVisualBoundaryBox();
        return PdfVisualCoordinateMapper.GetVisualSize(pageBox, RotationDegrees, Geometry.UserUnit ?? 1D);
    }

    internal PdfVisualBounds TransformBoundsToVisual(double left, double bottom, double right, double top) =>
        PdfVisualCoordinateMapper.TransformBounds(GetVisualBoundaryBox(), RotationDegrees, left, bottom, right, top, Geometry.UserUnit ?? 1D);

    internal PdfVisualBounds TransformVisualBoundsToUser(double left, double top, double right, double bottom) =>
        PdfVisualCoordinateMapper.TransformVisualBoundsToUser(GetVisualBoundaryBox(), RotationDegrees, left, top, right, bottom, Geometry.UserUnit ?? 1D);

    /// <summary>
    /// Converts a top-left visual page rectangle into PDF default user-space coordinates.
    /// This accounts for the effective crop box origin and inherited page rotation.
    /// </summary>
    public PdfPageRectangle MapVisualRectangleToUserSpace(double left, double top, double right, double bottom) {
        if (right <= left) throw new ArgumentOutOfRangeException(nameof(right), "Visual rectangle right must be greater than left.");
        if (bottom <= top) throw new ArgumentOutOfRangeException(nameof(bottom), "Visual rectangle bottom must be greater than top.");
        PdfVisualBounds mapped = TransformVisualBoundsToUser(left, top, right, bottom);
        return new PdfPageRectangle(mapped.Left, mapped.Top, mapped.Right, mapped.Bottom);
    }

    /// <summary>
    /// Converts a point from top-left visual page coordinates into PDF default user-space coordinates.
    /// This accounts for the effective crop box origin and inherited page rotation.
    /// </summary>
    public PdfPagePoint MapVisualPointToUserSpace(double x, double y) {
        (double mappedX, double mappedY) = PdfVisualCoordinateMapper.TransformVisualPointToUser(
            GetVisualBoundaryBox(),
            RotationDegrees,
            x,
            y,
            Geometry.UserUnit ?? 1D);
        return new PdfPagePoint(mappedX, mappedY);
    }

    private PdfPageBox GetVisualBoundaryBox() {
        if (Geometry.EffectiveBox is PdfPageBox effectiveBox) return effectiveBox;
        if (Geometry.HasEmptyEffectiveBoxIntersection) throw new InvalidOperationException("The page CropBox does not intersect its MediaBox; visual coordinates cannot be mapped safely.");
        return new PdfPageBox("MediaBox", 0D, 0D, Width, Height);
    }

    /// <summary>Logical elements in extraction order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements { get; }

    /// <summary>
    /// Bounded semantic-stage artifacts, reading-order evidence, classifications, and provenance for this page.
    /// </summary>
    public PdfUnderstandingPageResult Analysis { get; }

    /// <summary>Logical page elements grouped by element kind.</summary>
    public IReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>> ElementsByKind {
        get {
            if (_elementsByKind is not null) {
                return _elementsByKind;
            }

            var grouped = new Dictionary<PdfLogicalElementKind, List<IPdfLogicalElement>>();
            for (int i = 0; i < Elements.Count; i++) {
                IPdfLogicalElement element = Elements[i];
                if (!grouped.TryGetValue(element.Kind, out List<IPdfLogicalElement>? kindElements)) {
                    kindElements = new List<IPdfLogicalElement>();
                    grouped.Add(element.Kind, kindElements);
                }

                kindElements.Add(element);
            }

            var result = new Dictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _elementsByKind = new System.Collections.ObjectModel.ReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>>(result);
            return _elementsByKind;
        }
    }

    /// <summary>True when at least one logical element of the requested kind is present on this page.</summary>
    public bool HasElementKind(PdfLogicalElementKind kind) {
        return ElementsByKind.ContainsKey(kind);
    }

    /// <summary>Returns logical page elements of the requested kind.</summary>
    public IReadOnlyList<IPdfLogicalElement> GetElements(PdfLogicalElementKind kind) {
        return ElementsByKind.TryGetValue(kind, out IReadOnlyList<IPdfLogicalElement>? elements)
            ? elements
            : Array.Empty<IPdfLogicalElement>();
    }

    /// <summary>Line-level text blocks extracted from positioned text spans.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks { get; }

    /// <summary>Heading lines inferred from fused font, geometry, outline, and tagged-PDF evidence.</summary>
    public IReadOnlyList<PdfLogicalHeading> Headings { get; }

    /// <summary>Page-header text identified by the canonical semantic pipeline.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Headers =>
        _headers ??= TextBlocks.Where(static block => block.Kind == PdfLogicalElementKind.Header).ToArray();

    /// <summary>Page-footer text identified by the canonical semantic pipeline.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Footers =>
        _footers ??= TextBlocks.Where(static block => block.Kind == PdfLogicalElementKind.Footer).ToArray();

    /// <summary>Caption text identified by the canonical semantic pipeline.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Captions =>
        _captions ??= TextBlocks.Where(static block => block.Kind == PdfLogicalElementKind.Caption).ToArray();

    /// <summary>Footnote text identified by the canonical semantic pipeline.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Footnotes =>
        _footnotes ??= TextBlocks.Where(static block => block.Kind == PdfLogicalElementKind.Footnote).ToArray();

    /// <summary>Heuristic paragraph groups built from non-table, non-list text lines.</summary>
    public IReadOnlyList<PdfLogicalParagraph> Paragraphs { get; }

    /// <summary>Detected bullet and numbered list items with marker and level hints.</summary>
    public IReadOnlyList<PdfLogicalListItem> ListItems { get; }

    /// <summary>Detected table-like regions.</summary>
    public IReadOnlyList<PdfLogicalTable> Tables { get; }

    /// <summary>Number of visible vector drawing primitives recovered from the source page.</summary>
    public int VectorPrimitiveCount { get; }

    internal int UnrepresentedVectorPrimitiveCount { get; }

    /// <summary>Image XObjects referenced by the page.</summary>
    public IReadOnlyList<PdfLogicalImage> Images { get; }

    /// <summary>URI, named-destination, direct-destination, named-action, and remote GoTo link annotations on the page.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> Links { get; }

    /// <summary>Generic page annotations read from the page, including any primary, additional, or chained actions.</summary>
    public IReadOnlyList<PdfAnnotation> Annotations { get; }

    /// <summary>Number of generic page annotations read from the page.</summary>
    public int AnnotationCount => Annotations.Count;

    /// <summary>True when the page has at least one generic annotation.</summary>
    public bool HasAnnotations => AnnotationCount > 0;

    /// <summary>Simple link annotations read from the page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations { get; }

    /// <summary>AcroForm widget annotations placed on this page.</summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets { get; }

    /// <summary>Page-level additional actions attached to the source page dictionary.</summary>
    public IReadOnlyList<PdfPageAction> PageActions { get; }

    /// <summary>Number of page-level additional actions attached to the source page dictionary.</summary>
    public int PageActionCount => PageActions.Count;

    /// <summary>True when the source page dictionary has page-level additional actions.</summary>
    public bool HasPageActions => PageActionCount > 0;

    internal static PdfLogicalPage From(
        PdfReadDocument document,
        PdfReadPage page,
        int pageNumber,
        PdfTextLayoutOptions? options,
        IReadOnlyList<PdfLogicalFormWidget>? pageFormWidgets = null,
        PdfUnderstandingPageResult? analysis = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var size = page.GetPageSize();
        PdfPageGeometry geometry = page.GetGeometry();
        List<PdfUnderstandingLine>? retainedLines = analysis is null
            ? null
            : GetRetainedProjectionLines(analysis);
        IReadOnlyList<PdfTextSpan>? retainedRuns = analysis is null
            ? null
            : analysis.RestrictLogicalProjectionToReadingOrder
                ? GetRetainedProjectionRuns(retainedLines!, cancellationToken)
                : analysis.DecodedRuns;
        IReadOnlyList<PdfUnderstandingTableCandidate>? projectedTableCandidates = analysis is null
            ? null
            : GetProjectedTableCandidates(analysis, retainedLines!, retainedRuns!, cancellationToken);
        IReadOnlyList<PdfUnderstandingTableCandidate>? nativeStructuredTableCandidates = projectedTableCandidates?
            .Where(static candidate => candidate.SourceKind == PdfLogicalContentSourceKind.Native)
            .ToArray();
        var structured = analysis is null
            ? page.ExtractStructured(options)
            : CreateStructuredProjection(
                page,
                retainedLines!,
                nativeStructuredTableCandidates!,
                analysis,
                cancellationToken);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var semanticByTextBlock = new Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var links = new List<PdfLogicalLinkAnnotation>();
        var formWidgets = new List<PdfLogicalFormWidget>();
        var listLines = new HashSet<string>(structured.ListItems.Select(NormalizeForKindComparison), StringComparer.Ordinal);
        PdfUnderstandingPageResult pageAnalysis = analysis ?? PdfUnderstandingPageResult.Empty(pageNumber);
        SemanticElementIndex semanticIndex = SemanticElementIndex.Create(pageAnalysis.Elements);
        var headingSourceRuns = new HashSet<PdfTextSpan>(
            structured.Headings.SelectMany(static heading => heading.Line.Spans));
        LogicalTableSourceIndex tableSourceIndex = LogicalTableSourceIndex.Create(
            analysis is null ? structured.TablesDetailed : null,
            projectedTableCandidates);

        foreach (var line in structured.LinesDetailed) {
            cancellationToken.ThrowIfCancellationRequested();
            string text = line.Text?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            bool isStructuredHeading = analysis is null && line.Spans.Any(headingSourceRuns.Contains);
            bool isStructuredListItem = analysis is null &&
                (listLines.Contains(NormalizeForKindComparison(text)) || ContentStructureExtractor.IsListItemText(text));
            PdfUnderstandingSemanticElement? semantic = semanticIndex.Find(line.Y, line.XStart, text, line.Spans);
            var kind = ToLogicalKind(semantic, line, isStructuredHeading)
                ?? (isStructuredHeading
                    ? PdfLogicalElementKind.Heading
                    : isStructuredListItem
                    ? PdfLogicalElementKind.ListItem
                    : PdfLogicalElementKind.TextBlock);
            var block = new PdfLogicalTextBlock(
                pageNumber,
                kind,
                text,
                line.XStart,
                line.XEnd,
                line.Y,
                line.FontSize,
                line.Spans,
                line.SourceKind,
                line.Confidence,
                visualBounds: line.VisualBounds,
                isTableContent: tableSourceIndex.Contains(line));
            textBlocks.Add(block);
            if (semantic is not null) semanticByTextBlock.Add(block, semantic);
            elements.Add(block);
        }

        foreach (var row in structured.LeaderRows) {
            if (row.Length < 2) {
                continue;
            }

            var leader = new PdfLogicalLeaderRow(pageNumber, row[0], row[1]);
            elements.Add(leader);
        }

        IEnumerable<PdfLogicalTable> logicalTables = analysis is null
            ? structured.TablesDetailed.Select(table => PdfLogicalTable.From(pageNumber, table))
            : projectedTableCandidates!.Select(table => PdfLogicalTable.From(pageNumber, table));
        foreach (PdfLogicalTable logicalTable in logicalTables) {
            tables.Add(logicalTable);
            elements.Add(logicalTable);
        }

        IReadOnlyList<PdfImagePlacement> imagePlacements = analysis is not null
            ? pageAnalysis.ImagePlacements
            : page.GetImagePlacements(pageNumber);
        foreach (var image in page.GetImages(pageNumber, imagePlacements)) {
            IReadOnlyList<PdfImagePlacement> matchingPlacements = MatchImagePlacements(image, imagePlacements);
            var logicalImage = new PdfLogicalImage(
                image,
                matchingPlacements,
                MatchImageRegions(matchingPlacements, pageAnalysis.ImageRegions));
            images.Add(logicalImage);
            elements.Add(logicalImage);
        }

        IReadOnlyList<PdfLinkAnnotation> readLinkAnnotations = page.GetLinkAnnotations();
        var linkAnnotations = new List<PdfLinkAnnotation>(readLinkAnnotations.Count);
        for (int i = 0; i < readLinkAnnotations.Count; i++) {
            PdfLinkAnnotation linkAnnotation = ResolveLinkDestinationPageNumber(document, readLinkAnnotations[i]);
            linkAnnotations.Add(linkAnnotation);
            var logicalLink = new PdfLogicalLinkAnnotation(pageNumber, linkAnnotation);
            links.Add(logicalLink);
            elements.Add(logicalLink);
        }

        IReadOnlyList<PdfAnnotation> readAnnotations = page.GetAnnotations();
        var annotations = new List<PdfAnnotation>(readAnnotations.Count);
        for (int i = 0; i < readAnnotations.Count; i++) {
            annotations.Add(readAnnotations[i].WithPageNumber(pageNumber));
        }

        if (pageFormWidgets is not null) {
            for (int widgetIndex = 0; widgetIndex < pageFormWidgets.Count; widgetIndex++) {
                PdfLogicalFormWidget logicalWidget = pageFormWidgets[widgetIndex];
                formWidgets.Add(logicalWidget);
                elements.Add(logicalWidget);
            }
        }

        IReadOnlyList<PdfPageAction> readPageActions = page.GetPageActions();
        var pageActions = new List<PdfPageAction>(readPageActions.Count);
        for (int i = 0; i < readPageActions.Count; i++) {
            pageActions.Add(readPageActions[i].WithPageNumber(pageNumber));
        }

        (int vectorPrimitiveCount, int unrepresentedVectorPrimitiveCount) =
            page.GetVisibleVisualPrimitiveCounts(structured.TablesDetailed);
        Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>> textBlockLookup =
            CreateLogicalTextBlockLookup(textBlocks);
        LogicalTextBlockSourceIndex textBlockSourceIndex = LogicalTextBlockSourceIndex.Create(textBlocks);
        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            geometry,
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            BuildHeadings(pageNumber, analysis is null ? structured.Headings : new List<StructuredHeading>(), textBlocks, semanticByTextBlock, semanticIndex, textBlockLookup, textBlockSourceIndex),
            analysis is null
                ? BuildParagraphs(pageNumber, structured.Paragraphs, textBlocks, textBlockSourceIndex)
                : BuildParagraphsFromSemanticRegions(pageNumber, textBlocks, semanticByTextBlock),
            BuildListItems(pageNumber, analysis is null ? structured.ListNodes : new List<StructuredListItem>(), textBlocks, semanticByTextBlock, textBlockLookup, textBlockSourceIndex),
            tables.AsReadOnly(),
            vectorPrimitiveCount,
            unrepresentedVectorPrimitiveCount,
            images.AsReadOnly(),
            links.AsReadOnly(),
            annotations.AsReadOnly(),
            linkAnnotations.AsReadOnly(),
            formWidgets.AsReadOnly(),
            pageActions.AsReadOnly(),
            pageAnalysis);
    }

    private static List<PdfUnderstandingLine> GetRetainedProjectionLines(PdfUnderstandingPageResult analysis) {
        return new List<PdfUnderstandingLine>(analysis.LogicalProjectionLines);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfTextSpan> GetRetainedProjectionRuns(
        List<PdfUnderstandingLine> sourceLines,
        CancellationToken cancellationToken) {
        var retained = new List<PdfTextSpan>();
        var seen = new HashSet<PdfTextSpan>();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            IReadOnlyList<PdfUnderstandingWord> words = sourceLines[lineIndex].Words;
            for (int wordIndex = 0; wordIndex < words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> runs = words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (seen.Add(runs[runIndex])) retained.Add(runs[runIndex]);
                }
            }
        }
        return retained.AsReadOnly();
    }

    private static IReadOnlyList<PdfUnderstandingTableCandidate> GetProjectedTableCandidates(
        PdfUnderstandingPageResult analysis,
        IReadOnlyList<PdfUnderstandingLine> retainedLines,
        IReadOnlyList<PdfTextSpan> retainedRuns,
        CancellationToken cancellationToken) {
        if (!analysis.RestrictLogicalProjectionToReadingOrder || analysis.TableCandidates.Count == 0) {
            return analysis.TableCandidates;
        }

        var retained = new HashSet<PdfTextSpan>(retainedRuns);
        var retainedLineSet = new HashSet<PdfUnderstandingLine>(retainedLines);
        var retainedWords = new HashSet<PdfUnderstandingWord>(
            retainedLines.SelectMany(static line => line.Words));
        var candidates = new List<PdfUnderstandingTableCandidate>(analysis.TableCandidates.Count);
        for (int candidateIndex = 0; candidateIndex < analysis.TableCandidates.Count; candidateIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfUnderstandingTableCandidate candidate = analysis.TableCandidates[candidateIndex];
            if (candidate.SourceKind == PdfLogicalContentSourceKind.Ocr) {
                bool hasSourceLine = candidate.SourceLines.Count > 0;
                bool linesRetained = hasSourceLine && candidate.SourceLines.All(line =>
                    retainedLineSet.Contains(line) || line.Words.All(retainedWords.Contains));
                if (linesRetained) candidates.Add(candidate);
                continue;
            }
            bool hasSourceRun = false;
            bool isRetained = true;
            for (int lineIndex = 0; lineIndex < candidate.SourceLines.Count && isRetained; lineIndex++) {
                IReadOnlyList<PdfUnderstandingWord> words = candidate.SourceLines[lineIndex].Words;
                for (int wordIndex = 0; wordIndex < words.Count && isRetained; wordIndex++) {
                    IReadOnlyList<PdfTextSpan> sourceRuns = words[wordIndex].SourceRuns;
                    for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                        cancellationToken.ThrowIfCancellationRequested();
                        hasSourceRun = true;
                        if (!retained.Contains(sourceRuns[runIndex])) {
                            isRetained = false;
                            break;
                        }
                    }
                }
            }
            if (hasSourceRun && isRetained) candidates.Add(candidate);
        }
        return candidates.Count == analysis.TableCandidates.Count
            ? analysis.TableCandidates
            : candidates.AsReadOnly();
    }

    private static void ReplaceProjectionLines(
        PdfReadPage page,
        StructuredPage structured,
        List<PdfUnderstandingLine> sourceLines,
        CancellationToken cancellationToken) {
        structured.Lines.Clear();
        structured.LinesDetailed.Clear();
        for (int lineIndex = 0; lineIndex < sourceLines.Count; lineIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfUnderstandingLine sourceLine = sourceLines[lineIndex];
            var spans = new List<PdfTextSpan>();
            var seenSpans = new HashSet<PdfTextSpan>();
            for (int wordIndex = 0; wordIndex < sourceLine.Words.Count; wordIndex++) {
                IReadOnlyList<PdfTextSpan> sourceRuns = sourceLine.Words[wordIndex].SourceRuns;
                for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (seenSpans.Add(sourceRuns[runIndex])) spans.Add(sourceRuns[runIndex]);
                }
            }
            string text = sourceLine.Text.Trim();
            if (text.Length > 0) structured.Lines.Add(text);
            structured.LinesDetailed.Add(new StructuredLine {
                Y = sourceLine.BaselineY,
                XStart = sourceLine.XStart,
                XEnd = sourceLine.XEnd,
                Text = sourceLine.Text,
                FontSize = sourceLine.FontSize,
                Confidence = sourceLine.Confidence,
                Spans = Array.AsReadOnly(spans.ToArray()),
                SourceKind = sourceLine.SourceKind,
                VisualBounds = sourceLine.VisualBounds ?? CreateLineVisualBounds(page, sourceLine, spans)
            });
        }
    }

    private static PdfLogicalVisualBounds? CreateLineVisualBounds(
        PdfReadPage page,
        PdfUnderstandingLine line,
        List<PdfTextSpan> spans) {
        double left = double.MaxValue;
        double bottom = double.MaxValue;
        double right = double.MinValue;
        double top = double.MinValue;
        for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
            PdfTextSpan span = spans[spanIndex];
            double advance = Math.Abs(span.Advance) > 0.001D
                ? span.Advance
                : span.FontSize * Math.Max(1, PdfUnicodeScalarAnalysis.CountScalars(span.Text)) * 0.55D;
            IncludeOrientedBounds(span.X, span.Y, advance, span.FontSize, span.RotationDegrees);
        }
        if (spans.Count == 0) {
            for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                PdfUnderstandingWord word = line.Words[wordIndex];
                double radians = word.RotationDegrees * Math.PI / 180D;
                double startX = Math.Cos(radians) >= 0D ? word.XStart : word.XEnd;
                double projectedWidth = Math.Max(0D, word.XEnd - word.XStart);
                double advance = word.Advance is double explicitAdvance && explicitAdvance > 0.001D
                    ? explicitAdvance
                    : Math.Abs(Math.Cos(radians)) > 0.001D && projectedWidth > 0.001D
                    ? projectedWidth / Math.Abs(Math.Cos(radians))
                    : word.FontSize * Math.Max(1, PdfUnicodeScalarAnalysis.CountScalars(word.Text)) * 0.55D;
                IncludeOrientedBounds(startX, word.BaselineY, advance, word.FontSize, word.RotationDegrees);
            }
        }
        if (left == double.MaxValue || right <= left || top <= bottom) return null;
        PdfVisualBounds visual = page.TransformBoundsToVisual(left, bottom, right, top);
        return visual.Right > visual.Left && visual.Bottom > visual.Top
            ? new PdfLogicalVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom)
            : null;

        void IncludeOrientedBounds(double x, double y, double advance, double fontSize, double rotationDegrees) {
            double radians = rotationDegrees * Math.PI / 180D;
            double alongX = Math.Cos(radians);
            double alongY = Math.Sin(radians);
            double normalX = -alongY;
            double normalY = alongX;
            double endX = x + alongX * advance;
            double endY = y + alongY * advance;
            double descent = Math.Max(1D, fontSize * 0.25D);
            double ascent = Math.Max(1D, fontSize);
            Include(x - normalX * descent, y - normalY * descent);
            Include(x + normalX * ascent, y + normalY * ascent);
            Include(endX - normalX * descent, endY - normalY * descent);
            Include(endX + normalX * ascent, endY + normalY * ascent);
        }

        void Include(double x, double y) {
            left = Math.Min(left, x);
            bottom = Math.Min(bottom, y);
            right = Math.Max(right, x);
            top = Math.Max(top, y);
        }
    }

    private static PdfLinkAnnotation ResolveLinkDestinationPageNumber(PdfReadDocument document, PdfLinkAnnotation link) {
        if (link.DestinationPageNumber.HasValue || !link.DestinationPageObjectNumber.HasValue) {
            return link;
        }

        return link.WithDestinationPageNumber(document.GetPageNumberForObject(link.DestinationPageObjectNumber.Value));
    }

    private static IReadOnlyList<PdfLogicalParagraph> BuildParagraphs(
        int pageNumber,
        List<StructuredParagraph> paragraphs,
        List<PdfLogicalTextBlock> textBlocks,
        LogicalTextBlockSourceIndex textBlockSourceIndex) {
        if (paragraphs.Count == 0) {
            return Array.Empty<PdfLogicalParagraph>();
        }

        var textBlockLookup = new Dictionary<(long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>>();
        for (int i = 0; i < textBlocks.Count; i++) {
            PdfLogicalTextBlock block = textBlocks[i];
            if (block.Kind != PdfLogicalElementKind.TextBlock) {
                continue;
            }

            var key = CreateTextBlockLookupKey(block.BaselineY, block.XStart, block.Text);
            if (!textBlockLookup.TryGetValue(key, out Queue<PdfLogicalTextBlock>? blocks)) {
                blocks = new Queue<PdfLogicalTextBlock>();
                textBlockLookup.Add(key, blocks);
            }

            blocks.Enqueue(block);
        }

        var result = new List<PdfLogicalParagraph>(paragraphs.Count);
        var textBlockIndexes = new Dictionary<PdfLogicalTextBlock, int>();
        for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
            textBlockIndexes[textBlocks[blockIndex]] = blockIndex;
        }
        var claimedLines = new HashSet<PdfLogicalTextBlock>();
        for (int i = 0; i < paragraphs.Count; i++) {
            var paragraph = paragraphs[i];
            var lines = new List<PdfLogicalTextBlock>(paragraph.Lines.Count);
            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++) {
                var line = paragraph.Lines[lineIndex];
                var key = CreateTextBlockLookupKey(line.Y, line.XStart, line.Text.Trim());
                PdfLogicalTextBlock? block = textBlockLookup.TryGetValue(key, out Queue<PdfLogicalTextBlock>? blocks) && blocks.Count > 0
                    ? blocks.Dequeue()
                    : textBlockSourceIndex.Find(line, PdfLogicalElementKind.TextBlock);
                if (block is not null && claimedLines.Add(block)) lines.Add(block);
            }

            if (lines.Count > 0) {
                int start = 0;
                for (int lineIndex = 1; lineIndex <= lines.Count; lineIndex++) {
                    bool boundary = lineIndex == lines.Count ||
                        textBlockIndexes[lines[lineIndex]] != textBlockIndexes[lines[lineIndex - 1]] + 1;
                    if (!boundary) continue;
                    result.Add(PdfLogicalParagraph.From(pageNumber, lines.GetRange(start, lineIndex - start)));
                    start = lineIndex;
                }
            }
        }

        return result.AsReadOnly();
    }

    private static StructuredPage CreateStructuredProjection(
        PdfReadPage page,
        List<PdfUnderstandingLine> sourceLines,
        IReadOnlyList<PdfUnderstandingTableCandidate> nativeTableCandidates,
        PdfUnderstandingPageResult analysis,
        CancellationToken cancellationToken) {
        var structured = new StructuredPage();
        ReplaceProjectionLines(page, structured, sourceLines, cancellationToken);
        for (int tableIndex = 0; tableIndex < nativeTableCandidates.Count; tableIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            StructuredTable table = nativeTableCandidates[tableIndex].ToStructuredTable(
                analysis.ConsumeWork,
                analysis.CancellationCheck);
            ContentStructureExtractor.NormalizeDetectedTable(table);
            structured.TablesDetailed.Add(table);
            if (string.Equals(table.Kind, "leaders", StringComparison.OrdinalIgnoreCase)) {
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
                    string[] row = table.Rows[rowIndex];
                    if (row.Length >= 2) ContentStructureExtractor.AddLeaderRow(structured, row[0], row[1]);
                }
            } else {
                structured.Tables.AddRange(table.Rows);
            }
        }

        return structured;
    }

    private static IReadOnlyList<PdfUnderstandingImageRegion> MatchImageRegions(
        IReadOnlyList<PdfImagePlacement> placements,
        IReadOnlyList<PdfUnderstandingImageRegion> regions) {
        if (placements.Count == 0 || regions.Count == 0) return Array.Empty<PdfUnderstandingImageRegion>();
        var result = new List<PdfUnderstandingImageRegion>(placements.Count);
        for (int placementIndex = 0; placementIndex < placements.Count; placementIndex++) {
            PdfImagePlacement placement = placements[placementIndex];
            for (int regionIndex = 0; regionIndex < regions.Count; regionIndex++) {
                if (!ReferenceEquals(regions[regionIndex].Placement, placement)) continue;
                result.Add(regions[regionIndex]);
                break;
            }
        }
        return result.Count == 0 ? Array.Empty<PdfUnderstandingImageRegion>() : result.AsReadOnly();
    }

    private static IReadOnlyList<PdfLogicalParagraph> BuildParagraphsFromSemanticRegions(
        int pageNumber,
        List<PdfLogicalTextBlock> textBlocks,
        Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement> semanticByTextBlock) {
        var result = new List<PdfLogicalParagraph>();
        var current = new List<PdfLogicalTextBlock>();
        PdfUnderstandingSemanticElement? currentSemantic = null;

        for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
            PdfLogicalTextBlock block = textBlocks[blockIndex];
            if (block.Kind != PdfLogicalElementKind.TextBlock) {
                FlushCurrent();
                continue;
            }

            semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? semantic);
            if (semantic?.Kind == PdfUnderstandingSemanticKind.Table) {
                FlushCurrent();
                continue;
            }
            if (current.Count > 0 && !ReferenceEquals(currentSemantic, semantic)) FlushCurrent();
            currentSemantic = semantic;
            current.Add(block);
        }
        FlushCurrent();
        return result.Count == 0 ? Array.Empty<PdfLogicalParagraph>() : result.AsReadOnly();

        void FlushCurrent() {
            if (current.Count == 0) return;
            result.Add(PdfLogicalParagraph.From(pageNumber, current));
            current = new List<PdfLogicalTextBlock>();
            currentSemantic = null;
        }
    }

    private static (long BaselineY, long XStart, string Text) CreateTextBlockLookupKey(double baselineY, double xStart, string text) =>
        (BitConverter.DoubleToInt64Bits(baselineY), BitConverter.DoubleToInt64Bits(xStart), text);

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfLogicalListItem> BuildListItems(
        int pageNumber,
        List<StructuredListItem> listItems,
        List<PdfLogicalTextBlock> textBlocks,
        Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement> semanticByTextBlock,
        Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>> textBlockLookup,
        LogicalTextBlockSourceIndex textBlockSourceIndex) {
        var result = new List<PdfLogicalListItem>(Math.Max(listItems.Count, 4));
        var represented = new HashSet<PdfLogicalTextBlock>();
        var semanticGroups = new Dictionary<PdfUnderstandingSemanticElement, List<PdfLogicalTextBlock>>();
        for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
            PdfLogicalTextBlock block = textBlocks[blockIndex];
            if (block.Kind != PdfLogicalElementKind.ListItem ||
                !semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? semantic) ||
                semantic.Kind != PdfUnderstandingSemanticKind.ListItem ||
                !PreservesListContinuationOwnership(semantic)) continue;
            if (!semanticGroups.TryGetValue(semantic, out List<PdfLogicalTextBlock>? group)) {
                group = new List<PdfLogicalTextBlock>();
                semanticGroups.Add(semantic, group);
            }
            group.Add(block);
        }
        for (int i = 0; i < listItems.Count; i++) {
            var item = listItems[i];
            PdfLogicalTextBlock? block = FindTextBlock(item.Line, textBlockLookup, textBlockSourceIndex, PdfLogicalElementKind.ListItem);
            if (block is not null && !represented.Contains(block)) {
                semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? semantic);
                IReadOnlyList<PdfLogicalTextBlock> lines = semantic is not null && semanticGroups.TryGetValue(semantic, out List<PdfLogicalTextBlock>? group)
                    ? group
                    : new[] { block };
                IReadOnlyList<string> lineTexts = BuildListItemLineTexts(lines, block, item.Text);
                result.Add(new PdfLogicalListItem(
                    pageNumber,
                    semantic?.Level ?? item.Level,
                    item.Marker,
                    string.Join(" ", lineTexts.Where(static part => part.Length > 0)),
                    lines,
                    lineTexts,
                    semantic?.Confidence,
                    semantic?.Evidence));
                for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) represented.Add(lines[lineIndex]);
            }
        }

        for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
            PdfLogicalTextBlock block = textBlocks[blockIndex];
            if (block.Kind != PdfLogicalElementKind.ListItem || represented.Contains(block)) continue;
            semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? semantic);
            IReadOnlyList<PdfLogicalTextBlock> lines = semantic is not null && semanticGroups.TryGetValue(semantic, out List<PdfLogicalTextBlock>? group)
                ? group
                : new[] { block };
            bool parsed = ContentStructureExtractor.TryParseListItemText(
                block.Text,
                out string marker,
                out string text,
                out int level);
            string firstLineText = parsed ? text : block.Text;
            IReadOnlyList<string> lineTexts = BuildListItemLineTexts(lines, block, firstLineText);
            result.Add(new PdfLogicalListItem(
                pageNumber,
                semantic?.Level ?? level,
                parsed ? marker : string.Empty,
                string.Join(" ", lineTexts.Where(static part => part.Length > 0)),
                lines,
                lineTexts,
                semantic?.Confidence,
                semantic?.Evidence));
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) represented.Add(lines[lineIndex]);
        }

        return result.AsReadOnly();
    }

    private static bool PreservesListContinuationOwnership(PdfUnderstandingSemanticElement semantic) {
        if (semantic.Evidence.Any(static evidence =>
                string.Equals(evidence.Code, "semantic.tagged-pdf-role", StringComparison.Ordinal))) return true;
        if (semantic.Region.Lines.Count < 2 ||
            !semantic.Evidence.Any(static evidence =>
                string.Equals(evidence.Code, "semantic.list-marker", StringComparison.Ordinal)) ||
            !ContentStructureExtractor.IsListItemText(semantic.Region.Lines[0].Text)) return false;
        for (int lineIndex = 1; lineIndex < semantic.Region.Lines.Count; lineIndex++) {
            if (ContentStructureExtractor.IsListItemText(semantic.Region.Lines[lineIndex].Text)) return false;
        }
        return true;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> BuildListItemLineTexts(
        IReadOnlyList<PdfLogicalTextBlock> lines,
        PdfLogicalTextBlock anchor,
        string anchorText) {
        var result = new string[lines.Count];
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            result[lineIndex] = ReferenceEquals(lines[lineIndex], anchor)
                ? anchorText
                : lines[lineIndex].Text.Trim();
        }
        return Array.AsReadOnly(result);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfLogicalHeading> BuildHeadings(
        int pageNumber,
        List<StructuredHeading> headings,
        List<PdfLogicalTextBlock> textBlocks,
        Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement> semanticByTextBlock,
        SemanticElementIndex semanticIndex,
        Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>> textBlockLookup,
        LogicalTextBlockSourceIndex textBlockSourceIndex) {
        var result = new List<PdfLogicalHeading>(Math.Max(headings.Count, 4));
        var represented = new HashSet<PdfLogicalTextBlock>();
        for (int i = 0; i < headings.Count; i++) {
            var heading = headings[i];
            PdfLogicalTextBlock? block = FindTextBlock(heading.Line, textBlockLookup, textBlockSourceIndex, PdfLogicalElementKind.Heading);
            if (block is not null) {
                PdfUnderstandingSemanticElement? semantic = semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? aligned)
                    ? aligned
                    : semanticIndex.Find(block.BaselineY, block.XStart, block.Text, block.Spans);
                result.Add(new PdfLogicalHeading(
                    pageNumber,
                    semantic?.Level ?? heading.Level,
                    heading.Text,
                    heading.FontSize,
                    block,
                    semantic?.Confidence ?? 0.82D,
                    semantic?.Evidence));
                represented.Add(block);
            }
        }

        for (int i = 0; i < textBlocks.Count; i++) {
            PdfLogicalTextBlock block = textBlocks[i];
            if (block.Kind != PdfLogicalElementKind.Heading || represented.Contains(block)) continue;
            PdfUnderstandingSemanticElement? semantic = semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? aligned)
                ? aligned
                : semanticIndex.Find(block.BaselineY, block.XStart, block.Text, block.Spans);
            result.Add(new PdfLogicalHeading(
                pageNumber,
                semantic?.Level ?? 1,
                block.Text,
                block.FontSize,
                block,
                semantic?.Confidence ?? 0.65D,
                semantic?.Evidence));
        }

        return result.AsReadOnly();
    }

    private static PdfLogicalElementKind? ToLogicalKind(
        PdfUnderstandingSemanticElement? semantic,
        StructuredLine line,
        bool isStructuredHeading) => semantic?.Kind switch {
        PdfUnderstandingSemanticKind.Heading when isStructuredHeading || SupportsRegionHeadingProjection(semantic, line) => PdfLogicalElementKind.Heading,
        PdfUnderstandingSemanticKind.ListItem => PdfLogicalElementKind.ListItem,
        PdfUnderstandingSemanticKind.Header => PdfLogicalElementKind.Header,
        PdfUnderstandingSemanticKind.Footer => PdfLogicalElementKind.Footer,
        PdfUnderstandingSemanticKind.Caption => PdfLogicalElementKind.Caption,
        PdfUnderstandingSemanticKind.Footnote => PdfLogicalElementKind.Footnote,
        _ => null
    };

    private static bool SupportsRegionHeadingProjection(PdfUnderstandingSemanticElement semantic, StructuredLine line) {
        if (HasExplicitStructuralEvidence(semantic)) return true;
        if (semantic.Region.Lines.Count == 1) return true;
        double[] sizes = semantic.Region.Lines.Select(static candidate => candidate.FontSize).OrderBy(static size => size).ToArray();
        double median = sizes[(sizes.Length - 1) / 2];
        double largest = sizes[sizes.Length - 1];
        return line.FontSize >= largest * 0.95D && line.FontSize >= median * 1.15D;
    }

    private static PdfUnderstandingSemanticElement SelectBestSemanticElement(
        IReadOnlyList<PdfUnderstandingSemanticElement> matches) => matches
            .OrderByDescending(HasExplicitStructuralEvidence)
            .ThenBy(static element => element.Region.Lines.Count)
            .ThenByDescending(static element => element.Confidence)
            .First();

    private static bool HasExplicitStructuralEvidence(PdfUnderstandingSemanticElement element) =>
        element.Evidence.Any(static evidence =>
            string.Equals(evidence.Code, "semantic.outline-heading", StringComparison.Ordinal) ||
            string.Equals(evidence.Code, "semantic.tagged-pdf-role", StringComparison.Ordinal));

    private static Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>>
        CreateLogicalTextBlockLookup(List<PdfLogicalTextBlock> textBlocks) {
        var lookup = new Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>>();
        for (int index = 0; index < textBlocks.Count; index++) {
            PdfLogicalTextBlock block = textBlocks[index];
            var key = (block.Kind, BitConverter.DoubleToInt64Bits(block.BaselineY), BitConverter.DoubleToInt64Bits(block.XStart), block.Text);
            if (!lookup.TryGetValue(key, out Queue<PdfLogicalTextBlock>? blocks)) {
                blocks = new Queue<PdfLogicalTextBlock>();
                lookup.Add(key, blocks);
            }
            blocks.Enqueue(block);
        }
        return lookup;
    }

    private static PdfLogicalTextBlock? FindTextBlock(
        StructuredLine line,
        Dictionary<(PdfLogicalElementKind Kind, long BaselineY, long XStart, string Text), Queue<PdfLogicalTextBlock>> textBlockLookup,
        LogicalTextBlockSourceIndex textBlockSourceIndex,
        PdfLogicalElementKind kind) {
        var key = (kind, BitConverter.DoubleToInt64Bits(line.Y), BitConverter.DoubleToInt64Bits(line.XStart), line.Text.Trim());
        return textBlockLookup.TryGetValue(key, out Queue<PdfLogicalTextBlock>? blocks) && blocks.Count > 0
            ? blocks.Dequeue()
            : textBlockSourceIndex.Find(line, kind);
    }

    private static (long BaselineY, long XStart, string Text) CreateStructuredLineKey(StructuredLine line) =>
        (BitConverter.DoubleToInt64Bits(line.Y), BitConverter.DoubleToInt64Bits(line.XStart), line.Text.Trim());

    private sealed class LogicalTableSourceIndex {
        private readonly HashSet<PdfTextSpan> _sourceRuns = new();
        private readonly HashSet<(PdfLogicalContentSourceKind SourceKind, long Baseline, long XStart, string Text)> _sourceLines = new();

        private LogicalTableSourceIndex() { }

        internal static LogicalTableSourceIndex Create(
            List<StructuredTable>? structuredTables,
            IReadOnlyList<PdfUnderstandingTableCandidate>? candidates) {
            var index = new LogicalTableSourceIndex();
            if (structuredTables is not null) {
                for (int tableIndex = 0; tableIndex < structuredTables.Count; tableIndex++) {
                    IReadOnlyList<PdfTextSpan> runs = structuredTables[tableIndex].SourceRuns;
                    for (int runIndex = 0; runIndex < runs.Count; runIndex++) index._sourceRuns.Add(runs[runIndex]);
                }
            }
            if (candidates is not null) {
                for (int tableIndex = 0; tableIndex < candidates.Count; tableIndex++) {
                    IReadOnlyList<PdfUnderstandingLine> lines = candidates[tableIndex].SourceLines;
                    for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                        PdfUnderstandingLine line = lines[lineIndex];
                        index._sourceLines.Add((
                            line.SourceKind,
                            GetBucket(line.BaselineY, 0.25D),
                            GetBucket(line.XStart, 0.5D),
                            NormalizeForKindComparison(line.Text)));
                        for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                            IReadOnlyList<PdfTextSpan> runs = line.Words[wordIndex].SourceRuns;
                            for (int runIndex = 0; runIndex < runs.Count; runIndex++) index._sourceRuns.Add(runs[runIndex]);
                        }
                    }
                }
            }
            return index;
        }

        internal bool Contains(StructuredLine line) {
            for (int runIndex = 0; runIndex < line.Spans.Count; runIndex++) {
                if (_sourceRuns.Contains(line.Spans[runIndex])) return true;
            }
            return _sourceLines.Contains((
                line.SourceKind,
                GetBucket(line.Y, 0.25D),
                GetBucket(line.XStart, 0.5D),
                NormalizeForKindComparison(line.Text)));
        }

        private static long GetBucket(double value, double width) =>
            checked((long)Math.Round(value / width, MidpointRounding.AwayFromZero));
    }

    private sealed class SemanticElementIndex {
        private readonly Dictionary<(long BaselineBucket, long XBucket, string Text), List<SemanticLineBinding>> _byGeometry;
        private readonly Dictionary<(long BaselineY, long XStart, string Text), PdfUnderstandingSemanticElement> _byExactGeometry;
        private readonly Dictionary<string, HashSet<PdfUnderstandingSemanticElement>> _byText;
        private readonly Dictionary<PdfTextSpan, HashSet<PdfUnderstandingSemanticElement>> _bySourceRun;

        private SemanticElementIndex(
            Dictionary<(long BaselineBucket, long XBucket, string Text), List<SemanticLineBinding>> byGeometry,
            Dictionary<(long BaselineY, long XStart, string Text), PdfUnderstandingSemanticElement> byExactGeometry,
            Dictionary<string, HashSet<PdfUnderstandingSemanticElement>> byText,
            Dictionary<PdfTextSpan, HashSet<PdfUnderstandingSemanticElement>> bySourceRun) {
            _byGeometry = byGeometry;
            _byExactGeometry = byExactGeometry;
            _byText = byText;
            _bySourceRun = bySourceRun;
        }

        internal static SemanticElementIndex Create(IReadOnlyList<PdfUnderstandingSemanticElement> elements) {
            var byGeometry = new Dictionary<(long BaselineBucket, long XBucket, string Text), List<SemanticLineBinding>>();
            var byExactGeometry = new Dictionary<(long BaselineY, long XStart, string Text), PdfUnderstandingSemanticElement>();
            var byText = new Dictionary<string, HashSet<PdfUnderstandingSemanticElement>>(StringComparer.Ordinal);
            var bySourceRun = new Dictionary<PdfTextSpan, HashSet<PdfUnderstandingSemanticElement>>();
            for (int elementIndex = 0; elementIndex < elements.Count; elementIndex++) {
                PdfUnderstandingSemanticElement element = elements[elementIndex];
                for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
                    PdfUnderstandingLine line = element.Region.Lines[lineIndex];
                    string text = NormalizeForKindComparison(line.Text);
                    var key = (GetSemanticBucket(line.BaselineY, 0.25D), GetSemanticBucket(line.XStart, 0.5D), text);
                    if (!byGeometry.TryGetValue(key, out List<SemanticLineBinding>? bindings)) {
                        bindings = new List<SemanticLineBinding>();
                        byGeometry.Add(key, bindings);
                    }
                    bindings.Add(new SemanticLineBinding(element, line.BaselineY, line.XStart));
                    var exactKey = (BitConverter.DoubleToInt64Bits(line.BaselineY), BitConverter.DoubleToInt64Bits(line.XStart), text);
                    if (!byExactGeometry.TryGetValue(exactKey, out PdfUnderstandingSemanticElement? currentBest) ||
                        IsBetterSemanticElement(element, currentBest)) {
                        byExactGeometry[exactKey] = element;
                    }
                    if (!byText.TryGetValue(text, out HashSet<PdfUnderstandingSemanticElement>? textElements)) {
                        textElements = new HashSet<PdfUnderstandingSemanticElement>();
                        byText.Add(text, textElements);
                    }
                    textElements.Add(element);
                    for (int wordIndex = 0; wordIndex < line.Words.Count; wordIndex++) {
                        IReadOnlyList<PdfTextSpan> sourceRuns = line.Words[wordIndex].SourceRuns;
                        for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                            if (!bySourceRun.TryGetValue(sourceRuns[runIndex], out HashSet<PdfUnderstandingSemanticElement>? runElements)) {
                                runElements = new HashSet<PdfUnderstandingSemanticElement>();
                                bySourceRun.Add(sourceRuns[runIndex], runElements);
                            }
                            runElements.Add(element);
                        }
                    }
                }
            }
            return new SemanticElementIndex(byGeometry, byExactGeometry, byText, bySourceRun);
        }

        internal PdfUnderstandingSemanticElement? Find(
            double baselineY,
            double xStart,
            string text,
            IReadOnlyList<PdfTextSpan>? sourceRuns = null) {
            string normalized = NormalizeForKindComparison(text);
            var exactKey = (BitConverter.DoubleToInt64Bits(baselineY), BitConverter.DoubleToInt64Bits(xStart), normalized);
            if (_byExactGeometry.TryGetValue(exactKey, out PdfUnderstandingSemanticElement? exactElement)) return exactElement;
            long baselineBucket = GetSemanticBucket(baselineY, 0.25D);
            long xBucket = GetSemanticBucket(xStart, 0.5D);
            var exact = new HashSet<PdfUnderstandingSemanticElement>();
            for (long baselineOffset = -1; baselineOffset <= 1; baselineOffset++) {
                for (long xOffset = -1; xOffset <= 1; xOffset++) {
                    if (!_byGeometry.TryGetValue(
                            (baselineBucket + baselineOffset, xBucket + xOffset, normalized),
                            out List<SemanticLineBinding>? bindings)) continue;
                    for (int bindingIndex = 0; bindingIndex < bindings.Count; bindingIndex++) {
                        SemanticLineBinding binding = bindings[bindingIndex];
                        if (Math.Abs(binding.BaselineY - baselineY) <= 0.25D &&
                            Math.Abs(binding.XStart - xStart) <= 0.5D) {
                            exact.Add(binding.Element);
                        }
                    }
                }
            }
            if (exact.Count > 0) return SelectBestSemanticElement(exact.ToArray());
            if (sourceRuns is not null && sourceRuns.Count > 0) {
                var sourceMatches = new HashSet<PdfUnderstandingSemanticElement>();
                for (int runIndex = 0; runIndex < sourceRuns.Count; runIndex++) {
                    if (_bySourceRun.TryGetValue(sourceRuns[runIndex], out HashSet<PdfUnderstandingSemanticElement>? runElements)) {
                        sourceMatches.UnionWith(runElements);
                    }
                }
                if (sourceMatches.Count > 0) return SelectBestSemanticElement(sourceMatches.ToArray());
            }
            return _byText.TryGetValue(normalized, out HashSet<PdfUnderstandingSemanticElement>? textMatches) && textMatches.Count == 1
                ? textMatches.First()
                : null;
        }

        private static bool IsBetterSemanticElement(
            PdfUnderstandingSemanticElement candidate,
            PdfUnderstandingSemanticElement current) =>
            HasExplicitStructuralEvidence(candidate) != HasExplicitStructuralEvidence(current)
                ? HasExplicitStructuralEvidence(candidate)
                : candidate.Region.Lines.Count != current.Region.Lines.Count
                    ? candidate.Region.Lines.Count < current.Region.Lines.Count
                    : candidate.Confidence > current.Confidence;

        private static long GetSemanticBucket(double value, double width) {
            if (double.IsNaN(value)) return 0L;
            double bucket = Math.Floor(value / width);
            if (bucket <= long.MinValue) return long.MinValue + 1;
            if (bucket >= long.MaxValue) return long.MaxValue - 1;
            return (long)bucket;
        }
    }

    private sealed class LogicalTextBlockSourceIndex {
        private readonly Dictionary<PdfTextSpan, HashSet<PdfLogicalTextBlock>> _bySourceRun;

        private LogicalTextBlockSourceIndex(Dictionary<PdfTextSpan, HashSet<PdfLogicalTextBlock>> bySourceRun) {
            _bySourceRun = bySourceRun;
        }

        internal static LogicalTextBlockSourceIndex Create(List<PdfLogicalTextBlock> textBlocks) {
            var bySourceRun = new Dictionary<PdfTextSpan, HashSet<PdfLogicalTextBlock>>();
            for (int blockIndex = 0; blockIndex < textBlocks.Count; blockIndex++) {
                PdfLogicalTextBlock block = textBlocks[blockIndex];
                for (int runIndex = 0; runIndex < block.Spans.Count; runIndex++) {
                    PdfTextSpan sourceRun = block.Spans[runIndex];
                    if (!bySourceRun.TryGetValue(sourceRun, out HashSet<PdfLogicalTextBlock>? blocks)) {
                        blocks = new HashSet<PdfLogicalTextBlock>();
                        bySourceRun.Add(sourceRun, blocks);
                    }
                    blocks.Add(block);
                }
            }
            return new LogicalTextBlockSourceIndex(bySourceRun);
        }

        internal PdfLogicalTextBlock? Find(StructuredLine line, PdfLogicalElementKind kind) {
            var matches = new Dictionary<PdfLogicalTextBlock, int>();
            for (int runIndex = 0; runIndex < line.Spans.Count; runIndex++) {
                if (!_bySourceRun.TryGetValue(line.Spans[runIndex], out HashSet<PdfLogicalTextBlock>? blocks)) continue;
                foreach (PdfLogicalTextBlock block in blocks) {
                    if (block.Kind != kind) continue;
                    matches[block] = matches.TryGetValue(block, out int count) ? count + 1 : 1;
                }
            }
            return matches
                .OrderByDescending(static pair => pair.Value)
                .ThenBy(pair => Math.Abs(pair.Key.BaselineY - line.Y))
                .ThenBy(pair => Math.Abs(pair.Key.XStart - line.XStart))
                .Select(static pair => pair.Key)
                .FirstOrDefault();
        }
    }

    private readonly struct SemanticLineBinding {
        internal SemanticLineBinding(PdfUnderstandingSemanticElement element, double baselineY, double xStart) {
            Element = element;
            BaselineY = baselineY;
            XStart = xStart;
        }

        internal PdfUnderstandingSemanticElement Element { get; }
        internal double BaselineY { get; }
        internal double XStart { get; }
    }

    private static string NormalizeForKindComparison(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        var builder = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            if (!char.IsWhiteSpace(text[i])) {
                builder.Append(text[i]);
            }
        }

        return builder.ToString();
    }

    private static IReadOnlyList<PdfImagePlacement> MatchImagePlacements(PdfExtractedImage image, IReadOnlyList<PdfImagePlacement> placements) {
        if (placements.Count == 0) {
            return Array.Empty<PdfImagePlacement>();
        }

        var result = new List<PdfImagePlacement>();
        for (int i = 0; i < placements.Count; i++) {
            PdfImagePlacement placement = placements[i];
            if (placement.PageNumber == image.PageNumber &&
                placement.ObjectNumber == image.ObjectNumber &&
                (image.ObjectNumber > 0 || placement.DirectStreamIdentity == image.DirectStreamIdentity) &&
                string.Equals(placement.ResourceName, image.ResourceName, StringComparison.Ordinal)) {
                result.Add(placement);
            }
        }

        return result.Count == 0 ? Array.Empty<PdfImagePlacement>() : result.AsReadOnly();
    }
}
