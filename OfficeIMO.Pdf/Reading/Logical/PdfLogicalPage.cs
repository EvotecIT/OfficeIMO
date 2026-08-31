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
        return PdfVisualCoordinateMapper.GetVisualSize(pageBox, RotationDegrees);
    }

    internal PdfVisualBounds TransformBoundsToVisual(double left, double bottom, double right, double top) =>
        PdfVisualCoordinateMapper.TransformBounds(GetVisualBoundaryBox(), RotationDegrees, left, bottom, right, top);

    internal PdfVisualBounds TransformVisualBoundsToUser(double left, double top, double right, double bottom) =>
        PdfVisualCoordinateMapper.TransformVisualBoundsToUser(GetVisualBoundaryBox(), RotationDegrees, left, top, right, bottom);

    private PdfPageBox GetVisualBoundaryBox() =>
        CropBox ?? MediaBox ?? new PdfPageBox("MediaBox", 0D, 0D, Width, Height);

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
        PdfUnderstandingPageResult? analysis = null) {
        var size = page.GetPageSize();
        PdfPageGeometry geometry = page.GetGeometry();
        var structured = page.ExtractStructured(options);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var semanticByTextBlock = new Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var links = new List<PdfLogicalLinkAnnotation>();
        var formWidgets = new List<PdfLogicalFormWidget>();
        var listLines = new HashSet<string>(structured.ListItems.Select(NormalizeForKindComparison), StringComparer.Ordinal);
        PdfUnderstandingPageResult pageAnalysis = analysis ?? PdfUnderstandingPageResult.Empty(pageNumber);

        foreach (var line in structured.LinesDetailed) {
            string text = line.Text?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            bool isStructuredHeading = IsStructuredHeadingLine(line, structured.Headings);
            bool isStructuredListItem = listLines.Contains(NormalizeForKindComparison(text)) || ContentStructureExtractor.IsListItemText(text);
            PdfUnderstandingSemanticElement? semantic = FindSemanticElement(line, pageAnalysis);
            var kind = ToLogicalKind(semantic, line, isStructuredHeading, isStructuredListItem)
                ?? (isStructuredHeading
                    ? PdfLogicalElementKind.Heading
                    : isStructuredListItem
                    ? PdfLogicalElementKind.ListItem
                    : PdfLogicalElementKind.TextBlock);
            var block = new PdfLogicalTextBlock(pageNumber, kind, text, line.XStart, line.XEnd, line.Y, line.FontSize, line.Spans);
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

        foreach (var table in structured.TablesDetailed) {
            var logicalTable = PdfLogicalTable.From(pageNumber, table);
            tables.Add(logicalTable);
            elements.Add(logicalTable);
        }

        IReadOnlyList<PdfImagePlacement> imagePlacements = page.GetImagePlacements(pageNumber);
        foreach (var image in page.GetImages(pageNumber, imagePlacements)) {
            var logicalImage = new PdfLogicalImage(image, MatchImagePlacements(image, imagePlacements));
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
        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            geometry,
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            BuildHeadings(pageNumber, structured.Headings, textBlocks, semanticByTextBlock, pageAnalysis),
            BuildParagraphs(pageNumber, structured.Paragraphs, textBlocks),
            BuildListItems(pageNumber, structured.ListNodes, textBlocks),
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

    private static PdfLinkAnnotation ResolveLinkDestinationPageNumber(PdfReadDocument document, PdfLinkAnnotation link) {
        if (link.DestinationPageNumber.HasValue || !link.DestinationPageObjectNumber.HasValue) {
            return link;
        }

        return link.WithDestinationPageNumber(document.GetPageNumberForObject(link.DestinationPageObjectNumber.Value));
    }

    private static IReadOnlyList<PdfLogicalParagraph> BuildParagraphs(int pageNumber, List<StructuredParagraph> paragraphs, List<PdfLogicalTextBlock> textBlocks) {
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
        for (int i = 0; i < paragraphs.Count; i++) {
            var paragraph = paragraphs[i];
            var lines = new List<PdfLogicalTextBlock>(paragraph.Lines.Count);
            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++) {
                var line = paragraph.Lines[lineIndex];
                var key = CreateTextBlockLookupKey(line.Y, line.XStart, line.Text.Trim());
                if (textBlockLookup.TryGetValue(key, out Queue<PdfLogicalTextBlock>? blocks) && blocks.Count > 0) {
                    lines.Add(blocks.Dequeue());
                }
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

    private static (long BaselineY, long XStart, string Text) CreateTextBlockLookupKey(double baselineY, double xStart, string text) =>
        (BitConverter.DoubleToInt64Bits(baselineY), BitConverter.DoubleToInt64Bits(xStart), text);

    private static IReadOnlyList<PdfLogicalListItem> BuildListItems(int pageNumber, List<StructuredListItem> listItems, IReadOnlyList<PdfLogicalTextBlock> textBlocks) {
        if (listItems.Count == 0) {
            return Array.Empty<PdfLogicalListItem>();
        }

        var result = new List<PdfLogicalListItem>(listItems.Count);
        for (int i = 0; i < listItems.Count; i++) {
            var item = listItems[i];
            PdfLogicalTextBlock? block = FindTextBlock(item.Line, textBlocks, PdfLogicalElementKind.ListItem);
            if (block is not null) {
                result.Add(new PdfLogicalListItem(pageNumber, item.Level, item.Marker, item.Text, block));
            }
        }

        return result.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfLogicalHeading> BuildHeadings(
        int pageNumber,
        List<StructuredHeading> headings,
        List<PdfLogicalTextBlock> textBlocks,
        Dictionary<PdfLogicalTextBlock, PdfUnderstandingSemanticElement> semanticByTextBlock,
        PdfUnderstandingPageResult analysis) {
        var result = new List<PdfLogicalHeading>(Math.Max(headings.Count, 4));
        var represented = new HashSet<PdfLogicalTextBlock>();
        for (int i = 0; i < headings.Count; i++) {
            var heading = headings[i];
            PdfLogicalTextBlock? block = FindTextBlock(heading.Line, textBlocks, PdfLogicalElementKind.Heading);
            if (block is not null) {
                PdfUnderstandingSemanticElement? semantic = semanticByTextBlock.TryGetValue(block, out PdfUnderstandingSemanticElement? aligned)
                    ? aligned
                    : FindSemanticElement(block, analysis);
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
                : FindSemanticElement(block, analysis);
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
        bool isStructuredHeading,
        bool isStructuredListItem) => semantic?.Kind switch {
        PdfUnderstandingSemanticKind.Heading when isStructuredHeading || SupportsRegionHeadingProjection(semantic, line) => PdfLogicalElementKind.Heading,
        PdfUnderstandingSemanticKind.ListItem when isStructuredListItem => PdfLogicalElementKind.ListItem,
        PdfUnderstandingSemanticKind.Header => PdfLogicalElementKind.Header,
        PdfUnderstandingSemanticKind.Footer => PdfLogicalElementKind.Footer,
        PdfUnderstandingSemanticKind.Caption => PdfLogicalElementKind.Caption,
        PdfUnderstandingSemanticKind.Footnote => PdfLogicalElementKind.Footnote,
        _ => null
    };

    private static bool SupportsRegionHeadingProjection(PdfUnderstandingSemanticElement semantic, StructuredLine line) {
        if (semantic.Region.Lines.Count == 1) return true;
        double[] sizes = semantic.Region.Lines.Select(static candidate => candidate.FontSize).OrderBy(static size => size).ToArray();
        double median = sizes[sizes.Length / 2];
        double largest = sizes[sizes.Length - 1];
        return line.FontSize >= largest * 0.95D && line.FontSize >= median * 1.15D;
    }

    private static PdfUnderstandingSemanticElement? FindSemanticElement(
        StructuredLine line,
        PdfUnderstandingPageResult analysis) {
        var exactMatches = new List<PdfUnderstandingSemanticElement>();
        PdfUnderstandingSemanticElement? uniqueRegionMatch = null;
        bool ambiguousRegionMatch = false;
        string normalizedLine = NormalizeForKindComparison(line.Text);
        for (int elementIndex = 0; elementIndex < analysis.Elements.Count; elementIndex++) {
            PdfUnderstandingSemanticElement element = analysis.Elements[elementIndex];
            for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
                PdfUnderstandingLine candidate = element.Region.Lines[lineIndex];
                if (Math.Abs(candidate.BaselineY - line.Y) <= 0.25D &&
                    Math.Abs(candidate.XStart - line.XStart) <= 0.5D &&
                    string.Equals(NormalizeForKindComparison(candidate.Text), normalizedLine, StringComparison.Ordinal)) {
                    exactMatches.Add(element);
                    break;
                }
            }

            if (NormalizeForKindComparison(element.Region.Text).Contains(normalizedLine)) {
                if (uniqueRegionMatch is null) {
                    uniqueRegionMatch = element;
                } else {
                    ambiguousRegionMatch = true;
                }
            }
        }

        if (exactMatches.Count > 0) return SelectBestSemanticElement(exactMatches);
        return ambiguousRegionMatch ? null : uniqueRegionMatch;
    }

    private static PdfUnderstandingSemanticElement? FindSemanticElement(
        PdfLogicalTextBlock block,
        PdfUnderstandingPageResult analysis) {
        var exactMatches = new List<PdfUnderstandingSemanticElement>();
        PdfUnderstandingSemanticElement? uniqueTextMatch = null;
        bool ambiguousTextMatch = false;
        string normalizedBlock = NormalizeForKindComparison(block.Text);
        for (int elementIndex = 0; elementIndex < analysis.Elements.Count; elementIndex++) {
            PdfUnderstandingSemanticElement element = analysis.Elements[elementIndex];
            for (int lineIndex = 0; lineIndex < element.Region.Lines.Count; lineIndex++) {
                PdfUnderstandingLine candidate = element.Region.Lines[lineIndex];
                if (!string.Equals(NormalizeForKindComparison(candidate.Text), normalizedBlock, StringComparison.Ordinal)) continue;
                if (Math.Abs(candidate.BaselineY - block.BaselineY) <= 0.25D &&
                    Math.Abs(candidate.XStart - block.XStart) <= 0.5D) {
                    exactMatches.Add(element);
                    break;
                }
                if (uniqueTextMatch is null) {
                    uniqueTextMatch = element;
                } else if (!ReferenceEquals(uniqueTextMatch, element)) {
                    ambiguousTextMatch = true;
                }
            }
        }

        if (exactMatches.Count > 0) return SelectBestSemanticElement(exactMatches);
        return ambiguousTextMatch ? null : uniqueTextMatch;
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

    private static PdfLogicalTextBlock? FindTextBlock(StructuredLine line, IReadOnlyList<PdfLogicalTextBlock> textBlocks, PdfLogicalElementKind kind) {
        for (int i = 0; i < textBlocks.Count; i++) {
            var block = textBlocks[i];
            if (block.Kind == kind &&
                Math.Abs(block.BaselineY - line.Y) <= 0.001 &&
                Math.Abs(block.XStart - line.XStart) <= 0.001 &&
                string.Equals(block.Text, line.Text.Trim(), StringComparison.Ordinal)) {
                return block;
            }
        }

        return null;
    }

    private static bool IsStructuredHeadingLine(StructuredLine line, List<StructuredHeading> headings) {
        for (int i = 0; i < headings.Count; i++) {
            var heading = headings[i];
            if (Math.Abs(heading.Line.Y - line.Y) <= 0.001 &&
                Math.Abs(heading.Line.XStart - line.XStart) <= 0.001 &&
                string.Equals(heading.Text, line.Text.Trim(), StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
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
