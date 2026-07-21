using AngleSharp.Dom;

namespace OfficeIMO.Html;

/// <summary>Typed semantic HTML block categories interpreted once by the shared core.</summary>
public enum HtmlSemanticBlockKind {
    /// <summary>Heading block.</summary>
    Heading,
    /// <summary>Paragraph or address block.</summary>
    Paragraph,
    /// <summary>Preformatted source or code block.</summary>
    Code,
    /// <summary>Quoted block.</summary>
    Quote,
    /// <summary>Ordered, unordered, or definition list.</summary>
    List,
    /// <summary>One list item.</summary>
    ListItem,
    /// <summary>Tabular content.</summary>
    Table,
    /// <summary>Image content.</summary>
    Image,
    /// <summary>Audio, video, or embedded media.</summary>
    Media,
    /// <summary>Form or form control.</summary>
    Form,
    /// <summary>Footnote, endnote, or note-like content.</summary>
    Note,
    /// <summary>Other retained semantic content.</summary>
    Other
}

/// <summary>Canonical typed semantic representation shared by HTML target adapters.</summary>
public sealed class HtmlSemanticDocument {
    internal HtmlSemanticDocument(
        string title,
        string language,
        IReadOnlyDictionary<string, string> metadata,
        IReadOnlyList<HtmlSemanticSection> sections,
        IReadOnlyList<HtmlSemanticBlock> rootTables,
        IReadOnlyList<HtmlSemanticResource> resources) {
        Title = title;
        Language = language;
        Metadata = metadata;
        Sections = sections;
        RootTables = rootTables;
        Resources = resources;
    }

    /// <summary>Normalized document title.</summary>
    public string Title { get; }

    /// <summary>Declared document language.</summary>
    public string Language { get; }

    /// <summary>Normalized document metadata keyed by lower-case name.</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; }

    /// <summary>Ordered semantic sections used as pages, sheets, or slides by generic adapters.</summary>
    public IReadOnlyList<HtmlSemanticSection> Sections { get; }

    /// <summary>Top-level tables, excluding tables nested in another table.</summary>
    public IReadOnlyList<HtmlSemanticBlock> RootTables { get; }

    /// <summary>Resources referenced by retained semantic blocks in source order.</summary>
    public IReadOnlyList<HtmlSemanticResource> Resources { get; }
}

/// <summary>One shared semantic section.</summary>
public sealed class HtmlSemanticSection {
    internal HtmlSemanticSection(string title, IReadOnlyList<HtmlSemanticBlock> blocks, HtmlSemanticSourceLocation? sourceLocation) {
        Title = title;
        Blocks = blocks;
        SourceLocation = sourceLocation;
    }

    /// <summary>Section title selected by shared heading, label, id, and document-title rules.</summary>
    public string Title { get; }

    /// <summary>Ordered blocks contained by this section.</summary>
    public IReadOnlyList<HtmlSemanticBlock> Blocks { get; }

    /// <summary>Source provenance for the section container or first retained block.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }
}

/// <summary>One typed semantic block with optional list, table, resource, form, and rich-run data.</summary>
public sealed class HtmlSemanticBlock {
    internal HtmlSemanticBlock(
        HtmlSemanticBlockKind kind,
        string text,
        int level,
        bool ordered,
        IReadOnlyList<HtmlSemanticRun> runs,
        IReadOnlyList<HtmlSemanticBlock> children,
        HtmlSemanticTable? table,
        HtmlSemanticResource? resource,
        IReadOnlyList<HtmlSemanticResource> inlineResources,
        HtmlSemanticFormControl? formControl,
        HtmlComputedStyle? style,
        HtmlSemanticSourceLocation? sourceLocation,
        IElement sourceElement) {
        Kind = kind;
        Text = text;
        Level = level;
        Ordered = ordered;
        Runs = runs;
        Children = children;
        Table = table;
        Resource = resource;
        InlineResources = inlineResources;
        FormControl = formControl;
        Style = style;
        SourceLocation = sourceLocation;
        SourceElement = sourceElement;
    }

    /// <summary>Semantic block category.</summary>
    public HtmlSemanticBlockKind Kind { get; }

    /// <summary>Normalized plain-text projection.</summary>
    public string Text { get; }

    /// <summary>Heading level or nested list depth; zero when not applicable.</summary>
    public int Level { get; }

    /// <summary>Whether a list is ordered.</summary>
    public bool Ordered { get; }

    /// <summary>Rich editable text runs.</summary>
    public IReadOnlyList<HtmlSemanticRun> Runs { get; }

    /// <summary>Nested list items or retained child blocks.</summary>
    public IReadOnlyList<HtmlSemanticBlock> Children { get; }

    /// <summary>Typed table data when <see cref="Kind"/> is <see cref="HtmlSemanticBlockKind.Table"/>.</summary>
    public HtmlSemanticTable? Table { get; }

    /// <summary>Typed resource when the block references an image or media object.</summary>
    public HtmlSemanticResource? Resource { get; }

    /// <summary>Resources embedded inside this block, retained independently from its text runs.</summary>
    public IReadOnlyList<HtmlSemanticResource> InlineResources { get; }

    /// <summary>Typed form state when the block represents a form control.</summary>
    public HtmlSemanticFormControl? FormControl { get; }

    /// <summary>Computed style snapshot from the shared CSS engine.</summary>
    public HtmlComputedStyle? Style { get; }

    /// <summary>Source provenance.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }

    internal IElement SourceElement { get; }
}

/// <summary>Editable rich-text run shared by native adapters.</summary>
public sealed class HtmlSemanticRun {
    internal HtmlSemanticRun(string text, string? hyperlink, bool bold, bool italic, bool underline,
        bool strikethrough, bool superscript, bool subscript, HtmlComputedStyle? style,
        HtmlSemanticSourceLocation? sourceLocation,
        bool isLineBreak) {
        Text = text;
        Hyperlink = hyperlink;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Strikethrough = strikethrough;
        Superscript = superscript;
        Subscript = subscript;
        Style = style;
        SourceLocation = sourceLocation;
        IsLineBreak = isLineBreak;
    }

    /// <summary>Run text.</summary>
    public string Text { get; }
    /// <summary>Policy-normalized hyperlink, when present.</summary>
    public string? Hyperlink { get; }
    /// <summary>Whether the run is bold.</summary>
    public bool Bold { get; }
    /// <summary>Whether the run is italic.</summary>
    public bool Italic { get; }
    /// <summary>Whether the run is underlined.</summary>
    public bool Underline { get; }
    /// <summary>Whether the run is struck through.</summary>
    public bool Strikethrough { get; }
    /// <summary>Whether the run is superscript.</summary>
    public bool Superscript { get; }
    /// <summary>Whether the run is subscript.</summary>
    public bool Subscript { get; }
    /// <summary>Computed style snapshot.</summary>
    public HtmlComputedStyle? Style { get; }
    /// <summary>Source provenance.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }
    /// <summary>Whether this run represents an explicit HTML line break.</summary>
    public bool IsLineBreak { get; }
}

/// <summary>Typed semantic table.</summary>
public sealed class HtmlSemanticTable {
    internal HtmlSemanticTable(string caption, IReadOnlyList<HtmlSemanticTableRow> rows) {
        Caption = caption;
        Rows = rows;
    }

    /// <summary>Resolved table caption or shared fallback title.</summary>
    public string Caption { get; }
    /// <summary>Rows in source order.</summary>
    public IReadOnlyList<HtmlSemanticTableRow> Rows { get; }
}

/// <summary>One semantic table row.</summary>
public sealed class HtmlSemanticTableRow {
    internal HtmlSemanticTableRow(IReadOnlyList<HtmlSemanticTableCell> cells, HtmlSemanticSourceLocation? sourceLocation) {
        Cells = cells;
        SourceLocation = sourceLocation;
    }

    /// <summary>Cells in source order.</summary>
    public IReadOnlyList<HtmlSemanticTableCell> Cells { get; }
    /// <summary>Source provenance.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }
}

/// <summary>One semantic table cell.</summary>
public sealed class HtmlSemanticTableCell {
    internal HtmlSemanticTableCell(string text, bool isHeader, int rowSpan, int columnSpan,
        IReadOnlyList<HtmlSemanticRun> runs, IReadOnlyList<HtmlSemanticResource> resources,
        HtmlComputedStyle? style, HtmlSemanticSourceLocation? sourceLocation) {
        Text = text;
        IsHeader = isHeader;
        RowSpan = rowSpan;
        ColumnSpan = columnSpan;
        Runs = runs;
        Resources = resources;
        Style = style;
        SourceLocation = sourceLocation;
    }

    /// <summary>Normalized cell text.</summary>
    public string Text { get; }
    /// <summary>Whether the source used a header cell.</summary>
    public bool IsHeader { get; }
    /// <summary>Source row span, normalized to at least one.</summary>
    public int RowSpan { get; }
    /// <summary>Source column span, normalized to at least one.</summary>
    public int ColumnSpan { get; }
    /// <summary>Rich cell runs.</summary>
    public IReadOnlyList<HtmlSemanticRun> Runs { get; }
    /// <summary>Resources embedded in the cell.</summary>
    public IReadOnlyList<HtmlSemanticResource> Resources { get; }
    /// <summary>Computed cell style.</summary>
    public HtmlComputedStyle? Style { get; }
    /// <summary>Source provenance.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }
}

/// <summary>Policy-normalized resource reference interpreted from a semantic element.</summary>
public sealed class HtmlSemanticResource {
    internal HtmlSemanticResource(HtmlResourceKind kind, string source, string alternateText,
        string mediaType, double? widthPixels, double? heightPixels,
        HtmlSemanticSourceLocation? sourceLocation) {
        Kind = kind;
        Source = source;
        AlternateText = alternateText;
        MediaType = mediaType;
        WidthPixels = widthPixels;
        HeightPixels = heightPixels;
        SourceLocation = sourceLocation;
    }

    /// <summary>Resource kind.</summary>
    public HtmlResourceKind Kind { get; }
    /// <summary>Policy-normalized source value.</summary>
    public string Source { get; }
    /// <summary>Accessible alternate text.</summary>
    public string AlternateText { get; }
    /// <summary>Declared or data-URI media type.</summary>
    public string MediaType { get; }
    /// <summary>Resolved CSS or HTML width in pixels, when explicitly available.</summary>
    public double? WidthPixels { get; }
    /// <summary>Resolved CSS or HTML height in pixels, when explicitly available.</summary>
    public double? HeightPixels { get; }
    /// <summary>Source provenance.</summary>
    public HtmlSemanticSourceLocation? SourceLocation { get; }
}

/// <summary>Typed HTML form-control state.</summary>
public sealed class HtmlSemanticFormControl {
    internal HtmlSemanticFormControl(string type, string name, string value, bool isChecked, bool isDisabled) {
        Type = type;
        Name = name;
        Value = value;
        IsChecked = isChecked;
        IsDisabled = isDisabled;
    }

    /// <summary>Normalized control type.</summary>
    public string Type { get; }
    /// <summary>Control name.</summary>
    public string Name { get; }
    /// <summary>Current value.</summary>
    public string Value { get; }
    /// <summary>Checked or selected state.</summary>
    public bool IsChecked { get; }
    /// <summary>Disabled state.</summary>
    public bool IsDisabled { get; }
}
