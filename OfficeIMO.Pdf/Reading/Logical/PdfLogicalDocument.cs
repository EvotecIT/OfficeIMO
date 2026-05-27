namespace OfficeIMO.Pdf;

/// <summary>
/// Logical PDF element categories exposed by the first-party read model.
/// </summary>
public enum PdfLogicalElementKind {
    /// <summary>Line-level text recovered from positioned PDF text spans.</summary>
    TextBlock,
    /// <summary>Heuristic heading line inferred from text size and geometry.</summary>
    Heading,
    /// <summary>Detected bullet or numbered list item.</summary>
    ListItem,
    /// <summary>Detected leader row such as label plus dotted value.</summary>
    LeaderRow,
    /// <summary>Detected table-like region.</summary>
    Table,
    /// <summary>Image XObject referenced by the page.</summary>
    Image,
    /// <summary>URI or named-destination link annotation on the page.</summary>
    LinkAnnotation
}

/// <summary>
/// Common shape for logical page elements extracted from a PDF page.
/// </summary>
public interface IPdfLogicalElement {
    /// <summary>One-based source page number.</summary>
    int PageNumber { get; }

    /// <summary>Element kind.</summary>
    PdfLogicalElementKind Kind { get; }
}

/// <summary>
/// First-party logical read model for a parser-supported PDF.
/// </summary>
public sealed class PdfLogicalDocument {
    private IReadOnlyList<IPdfLogicalElement>? _elements;

    private PdfLogicalDocument(
        PdfMetadata metadata,
        IReadOnlyList<PdfLogicalPage> pages,
        IReadOnlyList<PdfOutlineItem> outlines,
        IReadOnlyList<PdfPageLabel> pageLabels,
        IReadOnlyList<PdfNamedDestination> namedDestinations,
        PdfDocumentOpenAction? openAction,
        PdfViewerPreferences? viewerPreferences,
        IReadOnlyList<PdfFormField> formFields,
        string? catalogPageMode,
        string? catalogPageLayout,
        string? catalogVersion,
        string? catalogLanguage) {
        Metadata = metadata;
        Pages = pages;
        Outlines = outlines;
        PageLabels = pageLabels;
        NamedDestinations = namedDestinations;
        OpenAction = openAction;
        ViewerPreferences = viewerPreferences;
        FormFields = formFields;
        CatalogPageMode = catalogPageMode;
        CatalogPageLayout = catalogPageLayout;
        CatalogVersion = catalogVersion;
        CatalogLanguage = catalogLanguage;
    }

    /// <summary>Document metadata read from the PDF Info dictionary when available.</summary>
    public PdfMetadata Metadata { get; }

    /// <summary>Logical pages in document order.</summary>
    public IReadOnlyList<PdfLogicalPage> Pages { get; }

    /// <summary>Top-level document outline/bookmark entries.</summary>
    public IReadOnlyList<PdfOutlineItem> Outlines { get; }

    /// <summary>Page-label rules discovered from the document catalog.</summary>
    public IReadOnlyList<PdfPageLabel> PageLabels { get; }

    /// <summary>Named destinations discovered from the document catalog.</summary>
    public IReadOnlyList<PdfNamedDestination> NamedDestinations { get; }

    /// <summary>Simple document open action discovered from the document catalog, when supported.</summary>
    public PdfDocumentOpenAction? OpenAction { get; }

    /// <summary>Simple viewer preference entries discovered from the document catalog, when supported.</summary>
    public PdfViewerPreferences? ViewerPreferences { get; }

    /// <summary>Simple AcroForm fields discovered from the document catalog.</summary>
    public IReadOnlyList<PdfFormField> FormFields { get; }

    /// <summary>Catalog page mode, for example UseOutlines or FullScreen, when present.</summary>
    public string? CatalogPageMode { get; }

    /// <summary>Catalog page layout, for example SinglePage or TwoColumnLeft, when present.</summary>
    public string? CatalogPageLayout { get; }

    /// <summary>Catalog PDF version override, for example 1.7, when present.</summary>
    public string? CatalogVersion { get; }

    /// <summary>Catalog language tag, for example en-US or pl-PL, when present.</summary>
    public string? CatalogLanguage { get; }

    /// <summary>Number of pages in the logical document.</summary>
    public int PageCount => Pages.Count;

    /// <summary>True when at least one outline/bookmark entry was read from the catalog.</summary>
    public bool HasOutlines => Outlines.Count > 0;

    /// <summary>True when at least one readable page-label rule was read from the catalog.</summary>
    public bool HasReadablePageLabels => PageLabels.Count > 0;

    /// <summary>True when at least one named destination was read from the catalog.</summary>
    public bool HasNamedDestinations => NamedDestinations.Count > 0;

    /// <summary>True when a simple document open action was read from the catalog.</summary>
    public bool HasReadableOpenAction => OpenAction is not null;

    /// <summary>True when simple viewer preferences were read from the catalog.</summary>
    public bool HasReadableViewerPreferences => ViewerPreferences is not null;

    /// <summary>True when at least one simple AcroForm field was read from the document catalog.</summary>
    public bool HasFormFields => FormFields.Count > 0;

    /// <summary>All logical page elements flattened in page order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements {
        get {
            if (_elements is not null) {
                return _elements;
            }

            var elements = new List<IPdfLogicalElement>();
            for (int i = 0; i < Pages.Count; i++) {
                elements.AddRange(Pages[i].Elements);
            }

            _elements = elements.AsReadOnly();
            return _elements;
        }
    }

    /// <summary>Loads a PDF from bytes and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(byte[] pdf, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(pdf, nameof(pdf));
        return From(PdfReadDocument.Load(pdf), options);
    }

    /// <summary>Loads a PDF from a file path and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(string path, PdfTextLayoutOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return From(PdfReadDocument.Load(path), options);
    }

    /// <summary>Loads a PDF from the current position of a readable stream and returns the logical read model.</summary>
    public static PdfLogicalDocument Load(Stream stream, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        return From(PdfReadDocument.Load(stream), options);
    }

    /// <summary>Builds the logical read model from an already parsed PDF document.</summary>
    public static PdfLogicalDocument From(PdfReadDocument document, PdfTextLayoutOptions? options = null) {
        Guard.NotNull(document, nameof(document));

        var pages = new List<PdfLogicalPage>(document.Pages.Count);
        for (int i = 0; i < document.Pages.Count; i++) {
            pages.Add(PdfLogicalPage.From(document.Pages[i], i + 1, options));
        }

        return new PdfLogicalDocument(
            document.Metadata,
            pages.AsReadOnly(),
            document.Outlines,
            document.PageLabels,
            document.NamedDestinations,
            document.OpenAction,
            document.ViewerPreferences,
            document.FormFields,
            document.CatalogPageMode,
            document.CatalogPageLayout,
            document.CatalogVersion,
            document.CatalogLanguage);
    }
}

/// <summary>
/// Logical view of a single PDF page.
/// </summary>
public sealed class PdfLogicalPage {
    private PdfLogicalPage(
        int pageNumber,
        double width,
        double height,
        int rotationDegrees,
        IReadOnlyList<IPdfLogicalElement> elements,
        IReadOnlyList<PdfLogicalTextBlock> textBlocks,
        IReadOnlyList<PdfLogicalHeading> headings,
        IReadOnlyList<PdfLogicalParagraph> paragraphs,
        IReadOnlyList<PdfLogicalListItem> listItems,
        IReadOnlyList<PdfLogicalTable> tables,
        IReadOnlyList<PdfLogicalImage> images,
        IReadOnlyList<PdfLogicalLinkAnnotation> links,
        IReadOnlyList<PdfLinkAnnotation> linkAnnotations) {
        PageNumber = pageNumber;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
        Elements = elements;
        TextBlocks = textBlocks;
        Headings = headings;
        Paragraphs = paragraphs;
        ListItems = listItems;
        Tables = tables;
        Images = images;
        Links = links;
        LinkAnnotations = linkAnnotations;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Page width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Page height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Inherited page rotation normalized to 0, 90, 180, or 270.</summary>
    public int RotationDegrees { get; }

    /// <summary>Logical elements in extraction order.</summary>
    public IReadOnlyList<IPdfLogicalElement> Elements { get; }

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

    /// <summary>Image XObjects referenced by the page.</summary>
    public IReadOnlyList<PdfLogicalImage> Images { get; }

    /// <summary>URI and named-destination link annotations on the page.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> Links { get; }

    /// <summary>Simple link annotations read from the page.</summary>
    public IReadOnlyList<PdfLinkAnnotation> LinkAnnotations { get; }

    internal static PdfLogicalPage From(PdfReadPage page, int pageNumber, PdfTextLayoutOptions? options) {
        var size = page.GetPageSize();
        var structured = page.ExtractStructured(options);
        var elements = new List<IPdfLogicalElement>();
        var textBlocks = new List<PdfLogicalTextBlock>();
        var tables = new List<PdfLogicalTable>();
        var images = new List<PdfLogicalImage>();
        var links = new List<PdfLogicalLinkAnnotation>();
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
            var block = new PdfLogicalTextBlock(pageNumber, kind, text, line.XStart, line.XEnd, line.Y, line.SpanCount);
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

        foreach (var image in page.GetImages(pageNumber)) {
            var logicalImage = new PdfLogicalImage(image);
            images.Add(logicalImage);
            elements.Add(logicalImage);
        }

        IReadOnlyList<PdfLinkAnnotation> linkAnnotations = page.GetLinkAnnotations();
        for (int i = 0; i < linkAnnotations.Count; i++) {
            var logicalLink = new PdfLogicalLinkAnnotation(pageNumber, linkAnnotations[i]);
            links.Add(logicalLink);
            elements.Add(logicalLink);
        }

        return new PdfLogicalPage(
            pageNumber,
            size.Width,
            size.Height,
            page.GetRotationDegrees(),
            elements.AsReadOnly(),
            textBlocks.AsReadOnly(),
            BuildHeadings(pageNumber, structured.Headings, textBlocks),
            BuildParagraphs(pageNumber, structured.Paragraphs, textBlocks),
            BuildListItems(pageNumber, structured.ListNodes, textBlocks),
            tables.AsReadOnly(),
            images.AsReadOnly(),
            links.AsReadOnly(),
            linkAnnotations);
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
}

/// <summary>
/// Line-level text block extracted from a PDF page.
/// </summary>
public sealed class PdfLogicalTextBlock : IPdfLogicalElement {
    internal PdfLogicalTextBlock(int pageNumber, PdfLogicalElementKind kind, string text, double xStart, double xEnd, double baselineY, int spanCount) {
        PageNumber = pageNumber;
        Kind = kind;
        Text = text;
        XStart = xStart;
        XEnd = xEnd;
        BaselineY = baselineY;
        SpanCount = spanCount;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind { get; }

    /// <summary>Extracted text for the line-level block.</summary>
    public string Text { get; }

    /// <summary>Leftmost X coordinate in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Rightmost X coordinate in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Baseline Y coordinate in PDF points from the bottom of the page.</summary>
    public double BaselineY { get; }

    /// <summary>Number of text spans merged into this block.</summary>
    public int SpanCount { get; }
}

/// <summary>
/// Heuristic heading line inferred from text size and geometry.
/// </summary>
public sealed class PdfLogicalHeading {
    internal PdfLogicalHeading(int pageNumber, int level, string text, double fontSize, PdfLogicalTextBlock line) {
        PageNumber = pageNumber;
        Level = level;
        Text = text;
        FontSize = fontSize;
        Line = line;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Best-effort heading level, where 1 is the largest heading tier.</summary>
    public int Level { get; }

    /// <summary>Heading text.</summary>
    public string Text { get; }

    /// <summary>Representative font size in points.</summary>
    public double FontSize { get; }

    /// <summary>Line-level text block that produced the heading.</summary>
    public PdfLogicalTextBlock Line { get; }
}

/// <summary>
/// Detected bullet or numbered list item.
/// </summary>
public sealed class PdfLogicalListItem {
    internal PdfLogicalListItem(int pageNumber, int level, string marker, string text, PdfLogicalTextBlock line) {
        PageNumber = pageNumber;
        Level = level;
        Marker = marker;
        Text = text;
        Line = line;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Best-effort nesting level, where 1 is the outermost list level.</summary>
    public int Level { get; }

    /// <summary>List marker such as "1", "1.2", "-", "•", or "(a)".</summary>
    public string Marker { get; }

    /// <summary>List item text without the marker.</summary>
    public string Text { get; }

    /// <summary>Line-level text block that produced the list item.</summary>
    public PdfLogicalTextBlock Line { get; }
}

/// <summary>
/// Heuristic paragraph group built from nearby line-level text blocks.
/// </summary>
public sealed class PdfLogicalParagraph {
    private PdfLogicalParagraph(
        int pageNumber,
        string text,
        IReadOnlyList<PdfLogicalTextBlock> lines,
        double xStart,
        double xEnd,
        double yTop,
        double yBottom) {
        PageNumber = pageNumber;
        Text = text;
        Lines = lines;
        XStart = xStart;
        XEnd = xEnd;
        YTop = yTop;
        YBottom = yBottom;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Paragraph text with grouped lines joined by spaces.</summary>
    public string Text { get; }

    /// <summary>Line-level blocks that make up this paragraph.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> Lines { get; }

    /// <summary>Leftmost X coordinate in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Rightmost X coordinate in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Top baseline Y coordinate in PDF points.</summary>
    public double YTop { get; }

    /// <summary>Bottom baseline Y coordinate in PDF points.</summary>
    public double YBottom { get; }

    internal static PdfLogicalParagraph From(int pageNumber, StructuredParagraph paragraph, IReadOnlyList<PdfLogicalTextBlock> lines) {
        return new PdfLogicalParagraph(
            pageNumber,
            paragraph.Text,
            lines.ToArray(),
            paragraph.XStart,
            paragraph.XEnd,
            paragraph.YTop,
            paragraph.YBottom);
    }
}

/// <summary>
/// Detected leader row such as a table-of-contents or label/value row.
/// </summary>
public sealed class PdfLogicalLeaderRow : IPdfLogicalElement {
    internal PdfLogicalLeaderRow(int pageNumber, string label, string value) {
        PageNumber = pageNumber;
        Label = label;
        Value = value;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.LeaderRow;

    /// <summary>Leader row label.</summary>
    public string Label { get; }

    /// <summary>Leader row trailing value.</summary>
    public string Value { get; }
}

/// <summary>
/// Detected table-like region with simple geometry.
/// </summary>
public sealed class PdfLogicalTable : IPdfLogicalElement {
    private PdfLogicalTable(
        int pageNumber,
        string kind,
        double yTop,
        double yBottom,
        IReadOnlyList<PdfLogicalTableColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> rows) {
        PageNumber = pageNumber;
        DetectionKind = kind;
        YTop = yTop;
        YBottom = yBottom;
        Columns = columns;
        Rows = rows;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Table;

    /// <summary>Detection heuristic that produced the table.</summary>
    public string DetectionKind { get; }

    /// <summary>Top Y coordinate of the detected table band.</summary>
    public double YTop { get; }

    /// <summary>Bottom Y coordinate of the detected table band.</summary>
    public double YBottom { get; }

    /// <summary>Detected table columns.</summary>
    public IReadOnlyList<PdfLogicalTableColumn> Columns { get; }

    /// <summary>Extracted table rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    internal static PdfLogicalTable From(int pageNumber, StructuredTable table) {
        var columns = new List<PdfLogicalTableColumn>(table.Columns.Count);
        for (int i = 0; i < table.Columns.Count; i++) {
            columns.Add(new PdfLogicalTableColumn(table.Columns[i].From, table.Columns[i].To));
        }

        var rows = new List<IReadOnlyList<string>>(table.Rows.Count);
        for (int i = 0; i < table.Rows.Count; i++) {
            rows.Add(Array.AsReadOnly((string[])table.Rows[i].Clone()));
        }

        return new PdfLogicalTable(
            pageNumber,
            table.Kind,
            table.YTop,
            table.YBottom,
            columns.AsReadOnly(),
            rows.AsReadOnly());
    }
}

/// <summary>
/// Detected table column geometry.
/// </summary>
public sealed class PdfLogicalTableColumn {
    internal PdfLogicalTableColumn(double from, double to) {
        From = from;
        To = to;
    }

    /// <summary>Left X coordinate in PDF points.</summary>
    public double From { get; }

    /// <summary>Right X coordinate in PDF points.</summary>
    public double To { get; }
}

/// <summary>
/// Image XObject entry in the logical page model.
/// </summary>
public sealed class PdfLogicalImage : IPdfLogicalElement {
    internal PdfLogicalImage(PdfExtractedImage image) {
        SourceImage = image;
    }

    /// <inheritdoc />
    public int PageNumber => SourceImage.PageNumber;

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Image;

    /// <summary>Underlying extracted image payload and metadata.</summary>
    public PdfExtractedImage SourceImage { get; }

    /// <summary>PDF image resource name.</summary>
    public string ResourceName => SourceImage.ResourceName;

    /// <summary>Image width in pixels.</summary>
    public int Width => SourceImage.Width;

    /// <summary>Image height in pixels.</summary>
    public int Height => SourceImage.Height;

    /// <summary>Suggested MIME type when bytes are a complete image file.</summary>
    public string? MimeType => SourceImage.MimeType;
}

/// <summary>
/// Link annotation entry in the logical page model.
/// </summary>
public sealed class PdfLogicalLinkAnnotation : IPdfLogicalElement {
    internal PdfLogicalLinkAnnotation(int pageNumber, PdfLinkAnnotation link) {
        PageNumber = pageNumber;
        SourceLink = link.PageNumber == pageNumber ? link : link.WithPageNumber(pageNumber);
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.LinkAnnotation;

    /// <summary>Underlying parsed link annotation.</summary>
    public PdfLinkAnnotation SourceLink { get; }

    /// <summary>Absolute URI opened by the link annotation, or null for an internal named-destination link.</summary>
    public string? Uri => SourceLink.Uri;

    /// <summary>Named destination opened by the link annotation, or null for a URI link.</summary>
    public string? DestinationName => SourceLink.DestinationName;

    /// <summary>True when the link annotation opens an absolute URI.</summary>
    public bool IsUriLink => SourceLink.IsUriLink;

    /// <summary>True when the link annotation opens an internal named destination.</summary>
    public bool IsNamedDestinationLink => SourceLink.IsNamedDestinationLink;

    /// <summary>Optional annotation contents metadata.</summary>
    public string? Contents => SourceLink.Contents;

    /// <summary>Left edge of the annotation rectangle in PDF points.</summary>
    public double X1 => SourceLink.X1;

    /// <summary>Bottom edge of the annotation rectangle in PDF points.</summary>
    public double Y1 => SourceLink.Y1;

    /// <summary>Right edge of the annotation rectangle in PDF points.</summary>
    public double X2 => SourceLink.X2;

    /// <summary>Top edge of the annotation rectangle in PDF points.</summary>
    public double Y2 => SourceLink.Y2;

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => SourceLink.Width;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => SourceLink.Height;
}
