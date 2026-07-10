namespace OfficeIMO.OpenDocument;

/// <summary>ODT block kinds exposed in document reading order.</summary>
public enum OdtContentBlockKind {
    /// <summary>Normal paragraph.</summary>
    Paragraph,
    /// <summary>Outline heading.</summary>
    Heading,
    /// <summary>Table.</summary>
    Table
}

/// <summary>A paragraph, heading, or table in ODT reading order.</summary>
public sealed class OdtContentBlock {
    private OdtContentBlock(OdtContentBlockKind kind, OdtParagraph? paragraph, OdtTable? table,
        bool isListItem = false, bool? isOrderedList = null, int listLevel = 0) {
        Kind = kind;
        Paragraph = paragraph;
        Table = table;
        IsListItem = isListItem;
        IsOrderedList = isOrderedList;
        ListLevel = listLevel;
    }

    /// <summary>Block kind.</summary>
    public OdtContentBlockKind Kind { get; }
    /// <summary>Paragraph or heading value when applicable.</summary>
    public OdtParagraph? Paragraph { get; }
    /// <summary>Table value when applicable.</summary>
    public OdtTable? Table { get; }
    /// <summary>True when the paragraph is contained by an ODF list item.</summary>
    public bool IsListItem { get; }
    /// <summary>Whether the containing list is ordered, or null for non-list blocks.</summary>
    public bool? IsOrderedList { get; }
    /// <summary>Zero-based nesting level for list paragraphs.</summary>
    public int ListLevel { get; }

    internal static OdtContentBlock FromParagraph(OdtParagraph paragraph, bool isListItem = false,
        bool? isOrderedList = null, int listLevel = 0) => new OdtContentBlock(
        paragraph.IsHeading ? OdtContentBlockKind.Heading : OdtContentBlockKind.Paragraph, paragraph, null,
        isListItem, isOrderedList, listLevel);

    internal static OdtContentBlock FromTable(OdtTable table) => new OdtContentBlock(OdtContentBlockKind.Table, null, table);
}
