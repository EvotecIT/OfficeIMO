namespace OfficeIMO.Xps;

/// <summary>The semantic elements defined by native StoryFragments markup.</summary>
public enum XpsStructureKind {
    /// <summary>An extension whose semantics are not reconstructed.</summary>
    Unknown,
    /// <summary>An arbitrary grouping of blocks.</summary>
    Section,
    /// <summary>A paragraph of referenced page content.</summary>
    Paragraph,
    /// <summary>A table.</summary>
    Table,
    /// <summary>A group of table rows.</summary>
    TableRowGroup,
    /// <summary>A table row.</summary>
    TableRow,
    /// <summary>A table cell.</summary>
    TableCell,
    /// <summary>A list.</summary>
    List,
    /// <summary>A list item.</summary>
    ListItem,
    /// <summary>A group of content forming one figure.</summary>
    Figure,
    /// <summary>A reference to a named page element.</summary>
    NamedElement
}

/// <summary>The native type of a story fragment.</summary>
public enum XpsStoryFragmentType {
    /// <summary>Document body content.</summary>
    Content,
    /// <summary>Page header content.</summary>
    Header,
    /// <summary>Page footer content.</summary>
    Footer
}

/// <summary>A resolved native page reference, without inferring text from glyph identifiers or graphics.</summary>
public sealed class XpsNamedContent {
    internal XpsNamedContent(string name, string elementName, string pagePart, int pageIndex, string text, IReadOnlyList<int> glyphOrdinals,
        XpsGraphicDescription? description = null) {
        Name = name; ElementName = elementName; PagePartName = pagePart; PageIndex = pageIndex; Text = text;
        GlyphOrdinals = glyphOrdinals; Description = description;
    }
    /// <summary>The native Name attribute.</summary>
    public string Name { get; }
    /// <summary>FixedPage, Canvas, Path or Glyphs.</summary>
    public string ElementName { get; }
    /// <summary>The referenced native page part.</summary>
    public string PagePartName { get; }
    /// <summary>The zero-based page occurrence in the document sequence when the snapshot was read.</summary>
    public int PageIndex { get; }
    /// <summary>Literal UnicodeString content in the referenced element's native order.</summary>
    public string Text { get; }
    /// <summary>The authored AutomationProperties.Name on a Path or Canvas, when present.</summary>
    public string? AccessibilityName => Description?.Name;
    /// <summary>The authored AutomationProperties.HelpText on a Path or Canvas, when present.</summary>
    public string? AccessibilityHelpText => Description?.HelpText;
    internal XpsGraphicDescription? Description { get; }
    internal IReadOnlyList<int> GlyphOrdinals { get; }
}

/// <summary>An immutable native semantic node. Merged nodes can contain references from several pages.</summary>
public sealed class XpsStructureNode {
    internal XpsStructureNode(XpsStructureKind kind, IReadOnlyList<XpsStructureNode> children,
        XpsNamedContent? content = null, XpsNamedContent? marker = null, int rowSpan = 1, int columnSpan = 1, string? nameReference = null) {
        Kind = kind; Children = children; Content = content; Marker = marker; RowSpan = rowSpan; ColumnSpan = columnSpan;
        NameReference = nameReference;
    }
    /// <summary>The native semantic role.</summary>
    public XpsStructureKind Kind { get; }
    /// <summary>Child nodes in native reading order.</summary>
    public IReadOnlyList<XpsStructureNode> Children { get; }
    /// <summary>Resolved page content for a NamedElement, or null for a container/unresolved reference.</summary>
    public XpsNamedContent? Content { get; }
    /// <summary>The authored NameReference, including when its target is unresolved.</summary>
    public string? NameReference { get; }
    /// <summary>A list item's separately referenced marker, when present.</summary>
    public XpsNamedContent? Marker { get; }
    /// <summary>Number of rows occupied by a table cell.</summary>
    public int RowSpan { get; }
    /// <summary>Number of columns occupied by a table cell.</summary>
    public int ColumnSpan { get; }
    /// <summary>Literal text in this node's reading order, without inserted layout whitespace.</summary>
    public string Text => Content?.Text ?? string.Concat(Children.Select(c => c.Text));
    internal int TextLength => Content?.Text.Length ?? Children.Sum(c => c.TextLength);
}

/// <summary>A native story fragment and its continuation boundaries.</summary>
public sealed class XpsStoryFragment {
    internal XpsStoryFragment(string pagePart, int pageIndex, string? storyName, string? fragmentName,
        XpsStoryFragmentType type, bool breakBefore, bool breakAfter, IReadOnlyList<XpsStructureNode> blocks) {
        PagePartName = pagePart; PageIndex = pageIndex; StoryName = storyName; FragmentName = fragmentName;
        Type = type; BreakBefore = breakBefore; BreakAfter = breakAfter; Blocks = blocks;
    }
    /// <summary>The native page part containing this fragment.</summary>
    public string PagePartName { get; }
    /// <summary>The zero-based page occurrence in the sequence.</summary>
    public int PageIndex { get; }
    /// <summary>The story association, if authored.</summary>
    public string? StoryName { get; }
    /// <summary>The optional fragment identifier.</summary>
    public string? FragmentName { get; }
    /// <summary>Body, header or footer content.</summary>
    public XpsStoryFragmentType Type { get; }
    /// <summary>A leading StoryBreak prevents merging with the previous fragment.</summary>
    public bool BreakBefore { get; }
    /// <summary>A trailing StoryBreak prevents merging with the next fragment.</summary>
    public bool BreakAfter { get; }
    /// <summary>Native semantic blocks in fragment order.</summary>
    public IReadOnlyList<XpsStructureNode> Blocks { get; }
}

/// <summary>A detached read result for one page's relationship-owned StoryFragments part.</summary>
public sealed class XpsPageStructure {
    internal XpsPageStructure(bool hasNativeStructure, IReadOnlyList<XpsStoryFragment> fragments, IReadOnlyList<string> diagnostics, bool hasUnresolvedNames = false) {
        HasNativeStructure = hasNativeStructure; Fragments = fragments; Diagnostics = diagnostics; HasUnresolvedNames = hasUnresolvedNames;
    }
    internal bool HasUnresolvedNames { get; }
    /// <summary>Whether the page has a native StoryFragments relationship.</summary>
    public bool HasNativeStructure { get; }
    /// <summary>Fragments in their native part order.</summary>
    public IReadOnlyList<XpsStoryFragment> Fragments { get; }
    /// <summary>Unknown semantics or unresolved content that prevent a complete reconstruction.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Whether all present native semantics were reconstructed.</summary>
    public bool IsComplete => Diagnostics.Count == 0;
}

/// <summary>A story assembled in DocumentStructure reference order, or a page-local unassociated fragment.</summary>
public sealed class XpsLogicalStory {
    internal XpsLogicalStory(string? name, XpsStoryFragmentType type, IReadOnlyList<XpsStoryFragment> fragments, IReadOnlyList<XpsStructureNode> blocks) {
        Name = name; Type = type; Fragments = fragments.ToList().AsReadOnly(); Blocks = blocks;
    }
    /// <summary>The native story name, when associated.</summary>
    public string? Name { get; }
    /// <summary>Body, header or footer content.</summary>
    public XpsStoryFragmentType Type { get; }
    /// <summary>Fragments in their document reading order.</summary>
    public IReadOnlyList<XpsStoryFragment> Fragments { get; }
    /// <summary>Blocks with cross-fragment continuation applied.</summary>
    public IReadOnlyList<XpsStructureNode> Blocks { get; }
}

/// <summary>An immutable native logical-structure snapshot. Unstructured pages are not given inferred semantics.</summary>
public sealed class XpsLogicalStructure {
    internal XpsLogicalStructure(IReadOnlyList<XpsPageStructure> pages, IReadOnlyList<XpsLogicalStory> stories, IReadOnlyList<string> diagnostics) {
        Pages = pages; Stories = stories; Diagnostics = diagnostics;
    }
    /// <summary>Page occurrence snapshots in sequence order.</summary>
    public IReadOnlyList<XpsPageStructure> Pages { get; }
    /// <summary>Declared stories followed by unassociated page-local fragments.</summary>
    public IReadOnlyList<XpsLogicalStory> Stories { get; }
    /// <summary>Unresolved references or semantics that prevent complete logical reconstruction.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
    /// <summary>Whether any page has native StoryFragments metadata.</summary>
    public bool HasNativeStructure => Pages.Any(p => p.HasNativeStructure);
    /// <summary>Whether all present native semantics were reconstructed.</summary>
    public bool IsComplete => Diagnostics.Count == 0;
}
