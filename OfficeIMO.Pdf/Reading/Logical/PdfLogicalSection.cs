namespace OfficeIMO.Pdf;

/// <summary>A heading-owned document section reconstructed from the canonical reading order.</summary>
public sealed class PdfLogicalSection {
    private readonly List<PdfLogicalSection> _children = new();
    private readonly List<PdfLogicalTextBlock> _textBlocks = new();
    private readonly List<PdfLogicalParagraph> _paragraphs = new();
    private readonly List<PdfLogicalListItem> _listItems = new();
    private readonly List<PdfLogicalTable> _tables = new();
    private readonly List<PdfLogicalImage> _images = new();
    private readonly List<PdfLogicalLinkAnnotation> _links = new();
    private readonly List<PdfLogicalFormWidget> _formWidgets = new();
    private readonly HashSet<PdfLogicalImage> _ownedImages = new();

    internal PdfLogicalSection(int index, PdfLogicalHeading heading, PdfLogicalSection? parent) {
        Index = index;
        Heading = heading;
        Parent = parent;
        FirstPageNumber = heading.PageNumber;
        LastPageNumber = heading.PageNumber;
    }

    /// <summary>Zero-based section index in document reading order.</summary>
    public int Index { get; }

    /// <summary>Heading that starts the section.</summary>
    public PdfLogicalHeading Heading { get; }

    /// <summary>Best-evidence hierarchy level, where one is top-level.</summary>
    public int Level => Heading.Level;

    /// <summary>Section title.</summary>
    public string Title => Heading.Text;

    /// <summary>Containing section, or null for a top-level section.</summary>
    public PdfLogicalSection? Parent { get; }

    /// <summary>Direct child sections in document order.</summary>
    public IReadOnlyList<PdfLogicalSection> Children => _children;

    /// <summary>First one-based source page owned by this section.</summary>
    public int FirstPageNumber { get; }

    /// <summary>Last one-based source page reached by this section or its descendants.</summary>
    public int LastPageNumber { get; private set; }

    /// <summary>Direct ungrouped text blocks owned by this section.</summary>
    public IReadOnlyList<PdfLogicalTextBlock> TextBlocks => _textBlocks;

    /// <summary>Direct paragraphs owned by this section.</summary>
    public IReadOnlyList<PdfLogicalParagraph> Paragraphs => _paragraphs;

    /// <summary>Direct list items owned by this section.</summary>
    public IReadOnlyList<PdfLogicalListItem> ListItems => _listItems;

    /// <summary>Direct tables owned by this section.</summary>
    public IReadOnlyList<PdfLogicalTable> Tables => _tables;

    /// <summary>
    /// Direct placement-local images owned by this section.
    /// Repeated uses of one page image resource are represented separately with one placement each.
    /// </summary>
    public IReadOnlyList<PdfLogicalImage> Images => _images;

    /// <summary>Direct links owned by this section.</summary>
    public IReadOnlyList<PdfLogicalLinkAnnotation> Links => _links;

    /// <summary>Direct form widgets owned by this section.</summary>
    public IReadOnlyList<PdfLogicalFormWidget> FormWidgets => _formWidgets;

    internal void AddChild(PdfLogicalSection child) {
        _children.Add(child);
        Touch(child.LastPageNumber);
    }

    internal void Add(PdfLogicalTextBlock value) { _textBlocks.Add(value); Touch(value.PageNumber); }
    internal void Add(PdfLogicalParagraph value) { _paragraphs.Add(value); Touch(value.PageNumber); }
    internal void Add(PdfLogicalListItem value) { _listItems.Add(value); Touch(value.PageNumber); }
    internal void Add(PdfLogicalTable value) { _tables.Add(value); Touch(value.PageNumber); }
    internal bool Add(PdfLogicalImage value) { if (!_ownedImages.Add(value)) return false; _images.Add(value); Touch(value.PageNumber); return true; }
    internal void Add(PdfLogicalLinkAnnotation value) { _links.Add(value); Touch(value.PageNumber); }
    internal void Add(PdfLogicalFormWidget value) { _formWidgets.Add(value); Touch(value.PageNumber); }
    internal void IncludeDescendantPage(int pageNumber) => Touch(pageNumber);

    private void Touch(int pageNumber) {
        LastPageNumber = pageNumber;
        Parent?.IncludeDescendantPage(pageNumber);
    }
}
