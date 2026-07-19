namespace OfficeIMO.Pdf;

/// <summary>
/// Logical view of a single PDF page.
/// </summary>
public sealed class PdfLogicalPage {
    private IReadOnlyDictionary<PdfLogicalElementKind, IReadOnlyList<IPdfLogicalElement>>? _elementsByKind;

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
        IReadOnlyList<PdfLogicalImage> images,
        IReadOnlyList<PdfLogicalLinkAnnotation> links,
        IReadOnlyList<PdfAnnotation> annotations,
        IReadOnlyList<PdfLinkAnnotation> linkAnnotations,
        IReadOnlyList<PdfLogicalFormWidget> formWidgets,
        IReadOnlyList<PdfPageAction> pageActions) {
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
        Images = images;
        Links = links;
        Annotations = annotations;
        LinkAnnotations = linkAnnotations;
        FormWidgets = formWidgets;
        PageActions = pageActions;
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

    /// <summary>Logical elements in extraction order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements { get; }

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

    /// <summary>Heuristic heading lines inferred from text size and geometry.</summary>
    public IReadOnlyList<PdfLogicalHeading> Headings { get; }

    /// <summary>Heuristic paragraph groups built from non-table, non-list text lines.</summary>
    public IReadOnlyList<PdfLogicalParagraph> Paragraphs { get; }

    /// <summary>Detected bullet and numbered list items with marker and level hints.</summary>
    public IReadOnlyList<PdfLogicalListItem> ListItems { get; }

    /// <summary>Detected table-like regions.</summary>
    public IReadOnlyList<PdfLogicalTable> Tables { get; }

    /// <summary>Number of visible vector drawing primitives recovered from the source page.</summary>
    public int VectorPrimitiveCount { get; }

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

    internal static PdfLogicalPage From(PdfReadDocument document, PdfReadPage page, int pageNumber, PdfTextLayoutOptions? options, IReadOnlyList<PdfFormField>? formFields = null) {
        var size = page.GetPageSize();
        PdfPageGeometry geometry = page.GetGeometry();
        var structured = page.ExtractStructured(options);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var links = new List<PdfLogicalLinkAnnotation>();
        var formWidgets = new List<PdfLogicalFormWidget>();
        var listLines = new HashSet<string>(structured.ListItems.Select(NormalizeForKindComparison), StringComparer.Ordinal);

        foreach (var line in structured.LinesDetailed) {
            string text = line.Text?.Trim() ?? string.Empty;
            if (text.Length == 0) {
                continue;
            }

            var kind = IsStructuredHeadingLine(line, structured.Headings)
                ? PdfLogicalElementKind.Heading
                : listLines.Contains(NormalizeForKindComparison(text)) || LooksLikeListItem(text)
                ? PdfLogicalElementKind.ListItem
                : PdfLogicalElementKind.TextBlock;
            var block = new PdfLogicalTextBlock(pageNumber, kind, text, line.XStart, line.XEnd, line.Y, line.FontSize, line.SpanCount);
            textBlocks.Add(block);
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

        if (formFields is not null) {
            for (int fieldIndex = 0; fieldIndex < formFields.Count; fieldIndex++) {
                PdfFormField field = formFields[fieldIndex];
                for (int widgetIndex = 0; widgetIndex < field.Widgets.Count; widgetIndex++) {
                    PdfFormWidget widget = field.Widgets[widgetIndex];
                    if (widget.PageNumber == pageNumber) {
                        var logicalWidget = new PdfLogicalFormWidget(pageNumber, field, widget);
                        formWidgets.Add(logicalWidget);
                        elements.Add(logicalWidget);
                    }
                }
            }
        }

        IReadOnlyList<PdfPageAction> readPageActions = page.GetPageActions();
        var pageActions = new List<PdfPageAction>(readPageActions.Count);
        for (int i = 0; i < readPageActions.Count; i++) {
            pageActions.Add(readPageActions[i].WithPageNumber(pageNumber));
        }

        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            geometry,
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            BuildHeadings(pageNumber, structured.Headings, textBlocks),
            BuildParagraphs(pageNumber, structured.Paragraphs, textBlocks),
            BuildListItems(pageNumber, structured.ListNodes, textBlocks),
            tables.AsReadOnly(),
            page.GetVisualPrimitiveCount(),
            images.AsReadOnly(),
            links.AsReadOnly(),
            annotations.AsReadOnly(),
            linkAnnotations.AsReadOnly(),
            formWidgets.AsReadOnly(),
            pageActions.AsReadOnly());
    }

    private static PdfLinkAnnotation ResolveLinkDestinationPageNumber(PdfReadDocument document, PdfLinkAnnotation link) {
        if (link.DestinationPageNumber.HasValue || !link.DestinationPageObjectNumber.HasValue) {
            return link;
        }

        return link.WithDestinationPageNumber(document.GetPageNumberForObject(link.DestinationPageObjectNumber.Value));
    }

    private static IReadOnlyList<PdfLogicalParagraph> BuildParagraphs(int pageNumber, List<StructuredParagraph> paragraphs, IReadOnlyList<PdfLogicalTextBlock> textBlocks) {
        if (paragraphs.Count == 0) {
            return Array.Empty<PdfLogicalParagraph>();
        }

        var result = new List<PdfLogicalParagraph>(paragraphs.Count);
        for (int i = 0; i < paragraphs.Count; i++) {
            var paragraph = paragraphs[i];
            var lines = new List<PdfLogicalTextBlock>(paragraph.Lines.Count);
            for (int lineIndex = 0; lineIndex < paragraph.Lines.Count; lineIndex++) {
                var line = paragraph.Lines[lineIndex];
                PdfLogicalTextBlock? block = FindTextBlock(line, textBlocks, PdfLogicalElementKind.TextBlock);
                if (block is not null) {
                    lines.Add(block);
                }
            }

            if (lines.Count > 0) {
                result.Add(PdfLogicalParagraph.From(pageNumber, paragraph, lines));
            }
        }

        return result.AsReadOnly();
    }

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

    private static IReadOnlyList<PdfLogicalHeading> BuildHeadings(int pageNumber, List<StructuredHeading> headings, IReadOnlyList<PdfLogicalTextBlock> textBlocks) {
        if (headings.Count == 0) {
            return Array.Empty<PdfLogicalHeading>();
        }

        var result = new List<PdfLogicalHeading>(headings.Count);
        for (int i = 0; i < headings.Count; i++) {
            var heading = headings[i];
            PdfLogicalTextBlock? block = FindTextBlock(heading.Line, textBlocks, PdfLogicalElementKind.Heading);
            if (block is not null) {
                result.Add(new PdfLogicalHeading(pageNumber, heading.Level, heading.Text, heading.FontSize, block));
            }
        }

        return result.AsReadOnly();
    }

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

    private static bool LooksLikeListItem(string text) {
        string trimmed = text.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        char marker = trimmed[0];
        if (marker == '\u2022' || marker == '-' || marker == '*' || marker == '\u25CF') {
            return true;
        }

        int index = 0;
        while (index < trimmed.Length && char.IsDigit(trimmed[index])) {
            index++;
        }

        return index > 0 &&
            index < trimmed.Length &&
            (trimmed[index] == '.' || trimmed[index] == ')');
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
